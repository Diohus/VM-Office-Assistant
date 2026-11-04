"""Kho file theo SHA-256. Không ghi đè: một hash luôn ứng với đúng một nội dung."""
from __future__ import annotations

import hashlib
import os
import re
import tempfile

from .config import STORE_DIR

_SHA = re.compile(r"[0-9a-f]{64}")


def _path(sha: str, suffix: str):
    if not _SHA.fullmatch(sha):
        raise ValueError("sha không hợp lệ")
    folder = STORE_DIR / sha[:2]
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{sha}{suffix}"


def _put(path, data: bytes) -> None:
    if path.exists():
        return
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    with os.fdopen(fd, "wb") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, path)


def put_new(sha: str, data: bytes, grid_json: bytes) -> None:
    if hashlib.sha256(data).hexdigest() != sha:
        raise ValueError("Nội dung không khớp hash")
    _put(_path(sha, ".xlsx"), data)
    _put(_path(sha, ".grid.json"), grid_json)


def read(sha: str) -> bytes:
    data = _path(sha, ".xlsx").read_bytes()
    if hashlib.sha256(data).hexdigest() != sha:
        raise OSError("Bản trong kho không còn khớp hash")
    return data


def read_grid(sha: str) -> bytes:
    return _path(sha, ".grid.json").read_bytes()
