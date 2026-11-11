"""Kho file theo SHA-256. File xlsx không bao giờ bị ghi đè: một hash luôn ứng với đúng một nội dung."""
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


def _write(path, data: bytes) -> None:
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    with os.fdopen(fd, "wb") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, path)


def put_new(sha: str, data: bytes, grid_json: bytes) -> None:
    if hashlib.sha256(data).hexdigest() != sha:
        raise ValueError("Nội dung không khớp hash")
    file_path = _path(sha, ".xlsx")
    if not file_path.exists():
        _write(file_path, data)
    # Dữ liệu lưới là kết quả dịch, không phải nguồn sự thật: bộ dịch đổi thì phải ghi lại, nếu không app dùng bản dịch cũ.
    _write(_path(sha, ".grid.json"), grid_json)


def put_file(sha: str, data: bytes) -> None:
    """File nguồn của lát 1 (Q133): chỉ lưu file, không lưu lưới đã đọc, vì lưới chứa mọi giá trị kể cả cột người bệnh."""
    if hashlib.sha256(data).hexdigest() != sha:
        raise ValueError("Nội dung không khớp hash")
    path = _path(sha, ".xlsx")
    if not path.exists():
        _write(path, data)


def put_original(sha: str, data: bytes) -> None:
    """Bản gốc của file đời cũ. Chỉ lưu để không mất, không đưa cho thư viện đọc nào."""
    path = _path(sha, ".original")
    if not path.exists():
        _write(path, data)


def read(sha: str) -> bytes:
    data = _path(sha, ".xlsx").read_bytes()
    if hashlib.sha256(data).hexdigest() != sha:
        raise OSError("Bản trong kho không còn khớp hash")
    return data


def read_grid(sha: str) -> bytes:
    return _path(sha, ".grid.json").read_bytes()
