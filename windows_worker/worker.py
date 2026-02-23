from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional

import boto3
import redis
import requests


QUEUE_NAME = "visa_jobs"


def _env(name: str, default: Optional[str] = None) -> str:
    v = os.environ.get(name, default)
    if v is None or v == "":
        raise RuntimeError(f"Missing env var: {name}")
    return v


def _optional_env(name: str, default: str) -> str:
    v = os.environ.get(name)
    return v if v else default


def _is_windows() -> bool:
    return os.name == "nt"


def log(msg: str) -> None:
    # IMPORTANT: no PII in logs.
    sys.stdout.write(msg + "\n")
    sys.stdout.flush()


@dataclass(frozen=True)
class Config:
    backend_base_url: str
    internal_api_key: str
    redis_url: str

    s3_endpoint_url: str
    s3_bucket: str
    s3_access_key_id: str
    s3_secret_access_key: str
    s3_region: str

    work_root: Path
    runner_script: Path


def load_config() -> Config:
    return Config(
        backend_base_url=_env("BACKEND_BASE_URL"),
        internal_api_key=_env("INTERNAL_API_KEY"),
        redis_url=_env("REDIS_URL"),
        s3_endpoint_url=_env("S3_ENDPOINT_URL"),
        s3_bucket=_env("S3_BUCKET"),
        s3_access_key_id=_env("S3_ACCESS_KEY_ID"),
        s3_secret_access_key=_env("S3_SECRET_ACCESS_KEY"),
        s3_region=_optional_env("S3_REGION", "us-east-1"),
        work_root=Path(_optional_env("WORK_ROOT", r"C:\temp\visa_jobs")),
        runner_script=Path(_optional_env("RUNNER_SCRIPT", r"excel_runner\runner.py")),
    )


def s3_client(cfg: Config):
    return boto3.client(
        "s3",
        endpoint_url=cfg.s3_endpoint_url,
        aws_access_key_id=cfg.s3_access_key_id,
        aws_secret_access_key=cfg.s3_secret_access_key,
        region_name=cfg.s3_region,
    )


def ensure_bucket(s3, bucket: str) -> None:
    try:
        s3.head_bucket(Bucket=bucket)
    except Exception:
        try:
            s3.create_bucket(Bucket=bucket)
        except Exception:
            # might already exist or no perms
            pass


def s3_download(s3, *, bucket: str, key: str, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    s3.download_file(bucket, key, str(dst))


def s3_upload(s3, *, bucket: str, key: str, src: Path, content_type: str) -> None:
    extra = {"ContentType": content_type}
    s3.upload_file(str(src), bucket, key, ExtraArgs=extra)


def backend_headers(cfg: Config) -> Dict[str, str]:
    return {"X-Internal-Api-Key": cfg.internal_api_key}


def backend_get_job_payload(cfg: Config, job_id: str) -> Dict[str, Any]:
    url = f"{cfg.backend_base_url.rstrip('/')}/internal/jobs/{job_id}/payload"
    r = requests.get(url, headers=backend_headers(cfg), timeout=30)
    r.raise_for_status()
    return r.json()


def backend_complete(cfg: Config, job_id: str, files: list[dict]) -> None:
    url = f"{cfg.backend_base_url.rstrip('/')}/internal/jobs/{job_id}/complete"
    r = requests.post(url, headers=backend_headers(cfg), json={"files": files}, timeout=30)
    r.raise_for_status()


def backend_fail(cfg: Config, job_id: str, message: str) -> None:
    url = f"{cfg.backend_base_url.rstrip('/')}/internal/jobs/{job_id}/fail"
    r = requests.post(url, headers=backend_headers(cfg), json={"error_message": message}, timeout=30)
    r.raise_for_status()


def build_runner_payload(
    *,
    template_path: Path,
    work_dir: Path,
    internal_payload: Dict[str, Any],
    assets_local: Dict[str, Path],
) -> Dict[str, Any]:
    client = internal_payload["client"]
    job = internal_payload["job"]
    export = internal_payload["export"]
    company = internal_payload["company"]

    return {
        "template_path": str(template_path),
        "work_dir": str(work_dir),
        "client": {
            "full_name": client["full_name"],
            "passport_no": client["passport_no"],
            "dob": client["dob"],
            "address": client["address"],
            "country_display": client["country_display"],
        },
        "job": {
            "currency_symbol": job["currency_symbol"],
            "fx_rate": job["fx_rate"],
            "salary_rub": job["salary_rub"],
            "position": job["position"],
            "contract_start_date": job["contract_start_date"],
            "contract_number": job["contract_number"],
        },
        "company": {
            "selected_company_name": company["selected_company_name"],
            "assets": {
                "logo_png": str(assets_local["logo_png"]),
                "seal_png": str(assets_local["seal_png"]),
                "director_sign_png": str(assets_local["director_sign_png"]),
                "client_sign_png": str(assets_local["client_sign_png"]),
            },
        },
        "export": export,
    }


def run_excel_runner(cfg: Config, payload_path: Path) -> Dict[str, Any]:
    runner = cfg.runner_script
    if not runner.exists():
        # try relative to repo root (worker is in windows_worker/)
        runner2 = (Path(__file__).resolve().parents[1] / runner).resolve()
        if runner2.exists():
            runner = runner2
        else:
            raise RuntimeError(f"Runner script not found: {cfg.runner_script}")

    cmd = [sys.executable, str(runner), "--payload", str(payload_path)]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    out = (proc.stdout or "").strip()
    err = (proc.stderr or "").strip()

    if proc.returncode != 0:
        # Avoid printing full runner output (may include paths).
        raise RuntimeError(f"excel_runner failed (code={proc.returncode})")

    try:
        return json.loads(out)
    except Exception:
        # Sometimes other output appears; try last JSON object.
        last_brace = out.rfind("{")
        if last_brace >= 0:
            return json.loads(out[last_brace:])
        raise RuntimeError("excel_runner returned non-JSON output")


def make_bundle_zip(output_dir: Path, *, pdf_names: list[str], zip_name: str = "bundle.zip") -> Path:
    zip_path = output_dir / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in pdf_names:
            p = output_dir / name
            if not p.exists():
                raise RuntimeError(f"Missing expected PDF: {p}")
            zf.write(p, arcname=name)
    return zip_path


def process_job(cfg: Config, s3, job_id: str, package_id: str) -> None:
    log(f"[job] start job_id={job_id} package_id={package_id}")

    work_dir = (cfg.work_root / job_id).resolve()
    if work_dir.exists():
        # clean stale dir
        shutil.rmtree(work_dir)
    (work_dir / "assets").mkdir(parents=True, exist_ok=True)
    (work_dir / "output").mkdir(parents=True, exist_ok=True)

    internal = backend_get_job_payload(cfg, job_id)
    template_key = internal["template_key"]

    assets_keys = internal["company"]["assets"]
    logo_key = assets_keys["logo_key"]
    seal_key = assets_keys["seal_key"]
    director_key = assets_keys["director_sign_key"]
    client_key = assets_keys["client_sign_key"]

    template_path = work_dir / "template.xlsm"
    s3_download(s3, bucket=cfg.s3_bucket, key=template_key, dst=template_path)

    # Download assets to known local names
    logo_path = work_dir / "assets" / "logo.png"
    seal_path = work_dir / "assets" / "seal.png"
    director_path = work_dir / "assets" / "director_sign.png"
    client_path = work_dir / "assets" / "client_sign.png"
    s3_download(s3, bucket=cfg.s3_bucket, key=logo_key, dst=logo_path)
    s3_download(s3, bucket=cfg.s3_bucket, key=seal_key, dst=seal_path)
    s3_download(s3, bucket=cfg.s3_bucket, key=director_key, dst=director_path)
    s3_download(s3, bucket=cfg.s3_bucket, key=client_key, dst=client_path)

    runner_payload = build_runner_payload(
        template_path=template_path,
        work_dir=work_dir,
        internal_payload=internal,
        assets_local={
            "logo_png": logo_path,
            "seal_png": seal_path,
            "director_sign_png": director_path,
            "client_sign_png": client_path,
        },
    )

    payload_path = work_dir / "payload.json"
    payload_path.write_text(json.dumps(runner_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    result = run_excel_runner(cfg, payload_path)
    if result.get("status") != "ok":
        raise RuntimeError("excel_runner returned error status")

    output_dir = Path(result["output_dir"])
    pdf_files = list(result.get("pdf_files") or [])
    if len(pdf_files) == 0:
        raise RuntimeError("No PDFs produced by runner")

    # ZIP bundle (exact filenames inside archive)
    zip_path = make_bundle_zip(output_dir, pdf_names=pdf_files, zip_name="bundle.zip")

    # Upload PDFs
    files_for_backend: list[dict] = []
    pdf_map = {
        "Contract.pdf": ("contract", "Contract_v1.pdf"),
        "Bank_Statement_6m.pdf": ("bank", "Bank_Statement_6m_v1.pdf"),
        "Insurance.pdf": ("insurance", "Insurance_v1.pdf"),
        "Salary_Certificate.pdf": ("salary", "Salary_Certificate_v1.pdf"),
    }

    for pdf_name in pdf_files:
        doc_type, storage_name = pdf_map.get(pdf_name, (None, None))
        if not doc_type:
            # skip unexpected file names
            continue
        src = output_dir / pdf_name
        key = f"packages/{package_id}/{storage_name}"
        s3_upload(s3, bucket=cfg.s3_bucket, key=key, src=src, content_type="application/pdf")
        files_for_backend.append(
            {
                "doc_type": doc_type,
                "version": 1,
                "filename": pdf_name,
                "storage_key": key,
                "content_type": "application/pdf",
            }
        )

    # Upload bundle
    bundle_key = f"packages/{package_id}/bundle_v1.zip"
    s3_upload(s3, bucket=cfg.s3_bucket, key=bundle_key, src=zip_path, content_type="application/zip")
    files_for_backend.append(
        {
            "doc_type": "bundle",
            "version": 1,
            "filename": "bundle.zip",
            "storage_key": bundle_key,
            "content_type": "application/zip",
        }
    )

    backend_complete(cfg, job_id, files_for_backend)
    log(f"[job] done job_id={job_id} package_id={package_id}")


def main() -> int:
    if not _is_windows():
        log("ERROR: windows_worker must run on Windows")
        return 2

    cfg = load_config()
    r = redis.Redis.from_url(cfg.redis_url, decode_responses=True)
    s3 = s3_client(cfg)
    ensure_bucket(s3, cfg.s3_bucket)

    log("[worker] started, waiting for jobs...")
    while True:
        item = r.brpop(QUEUE_NAME, timeout=5)
        if not item:
            continue
        _queue, msg = item
        try:
            data = json.loads(msg)
            job_id = str(data["job_id"])
            package_id = str(data["package_id"])
        except Exception:
            log("[worker] invalid queue message (skipped)")
            continue

        try:
            process_job(cfg, s3, job_id, package_id)
        except Exception as e:
            # Do not include PII in message. Keep it short.
            msg = f"worker_failed: {type(e).__name__}"
            try:
                backend_fail(cfg, job_id, msg)
            except Exception:
                pass
            log(f"[job] error job_id={job_id} package_id={package_id} ({type(e).__name__})")
            # small backoff to avoid tight-loop on repeated failures
            time.sleep(1)


if __name__ == "__main__":
    raise SystemExit(main())

