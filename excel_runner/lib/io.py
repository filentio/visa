from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, Optional, TextIO


class RunnerIOError(RuntimeError):
    pass


def is_windows() -> bool:
    return os.name == "nt"


def read_payload(payload_arg: str) -> Dict[str, Any]:
    """
    Read JSON payload from a file or stdin.

    `--payload -` means stdin.
    """
    if payload_arg == "-":
        raw = sys.stdin.read()
        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            raise RunnerIOError(f"Invalid JSON in stdin: {e}") from e

    path = Path(payload_arg)
    try:
        raw = path.read_text(encoding="utf-8")
    except Exception as e:
        raise RunnerIOError(f"Failed to read payload file: {path}: {e}") from e

    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise RunnerIOError(f"Invalid JSON in file {path}: {e}") from e


def ensure_empty_dir(path: Path) -> None:
    """
    Create directory, removing it first if it exists.
    """
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def copy_file(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        if src.resolve() == dst.resolve():
            return
    except Exception:
        pass
    shutil.copy2(src, dst)


def write_json_stdout(obj: Dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(obj, ensure_ascii=False, indent=2) + "\n")


def short_exc(e: BaseException, *, limit: int = 4000) -> str:
    s = repr(e)
    if len(s) > limit:
        return s[:limit] + "...(truncated)"
    return s

