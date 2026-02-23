from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict

from .io import copy_file, ensure_dir


class AssetsError(RuntimeError):
    pass


@dataclass(frozen=True)
class PreparedAssets:
    assets_dir: Path
    logo_path: Path
    seal_path: Path
    director_sign_path: Path
    client_sign_path: Path


def prepare_assets(work_dir: Path, payload_assets: Dict[str, str]) -> PreparedAssets:
    """
    Copy provided PNG assets into `work_dir/assets/` and return absolute paths.

    Payload keys:
      - logo_png
      - seal_png
      - director_sign_png
      - client_sign_png
    """
    assets_dir = work_dir / "assets"
    ensure_dir(assets_dir)

    def _req(key: str) -> Path:
        v = payload_assets.get(key)
        if not v:
            raise AssetsError(f"Missing company.assets.{key} in payload")
        p = Path(v)
        if not p.exists():
            raise AssetsError(f"Asset file not found: {key} -> {p}")
        return p

    src_logo = _req("logo_png")
    src_seal = _req("seal_png")
    src_dir = _req("director_sign_png")
    src_client = _req("client_sign_png")

    dst_logo = assets_dir / "logo.png"
    dst_seal = assets_dir / "seal.png"
    dst_dir = assets_dir / "director_sign.png"
    dst_client = assets_dir / "client_sign.png"

    copy_file(src_logo, dst_logo)
    copy_file(src_seal, dst_seal)
    copy_file(src_dir, dst_dir)
    copy_file(src_client, dst_client)

    return PreparedAssets(
        assets_dir=assets_dir.resolve(),
        logo_path=dst_logo.resolve(),
        seal_path=dst_seal.resolve(),
        director_sign_path=dst_dir.resolve(),
        client_sign_path=dst_client.resolve(),
    )

