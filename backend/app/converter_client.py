"""Gửi việc cho worker LibreOffice qua thư mục spool riêng. Worker tắt thì báo ngay, không để người dùng chờ."""
from __future__ import annotations

import json
import os
import tempfile
import time
import uuid

from .config import CONVERTER_ALIVE_SECONDS, CONVERTER_SPOOL_DIR, CONVERTER_TIMEOUT_SECONDS

POLL_SECONDS = 0.05


class ConverterUnavailable(Exception):
    pass


class ConversionFailed(Exception):
    pass


class ConverterTimeout(Exception):
    """Worker còn sống nhưng việc chạy quá hạn: file lớn, xuất quá lâu. Khác với worker không chạy (T-DEV-10)."""


def _write_atomic(path, data: bytes) -> None:
    fd, tmp = tempfile.mkstemp(dir=CONVERTER_SPOOL_DIR, suffix=".tmp")
    with os.fdopen(fd, "wb") as fh:
        fh.write(data)
    os.replace(tmp, path)


def _alive() -> bool:
    try:
        return time.time() - (CONVERTER_SPOOL_DIR / ".alive").stat().st_mtime < CONVERTER_ALIVE_SECONDS
    except OSError:
        return False


def _run(kind: str, data: bytes, outputs: tuple[str, ...], **options) -> tuple[dict, dict[str, bytes]]:
    if not _alive():
        raise ConverterUnavailable
    job_id = uuid.uuid4().hex
    source, job, result = (CONVERTER_SPOOL_DIR / f"{job_id}{s}" for s in (".in", ".lojob", ".out.json"))
    files = {name: CONVERTER_SPOOL_DIR / f"{job_id}.out.{name}" for name in outputs}
    _write_atomic(source, data)
    _write_atomic(job, json.dumps({"kind": kind, **options}).encode())
    deadline = time.monotonic() + CONVERTER_TIMEOUT_SECONDS
    try:
        while not result.exists():
            if time.monotonic() > deadline:
                raise ConverterTimeout
            if not _alive():
                raise ConverterUnavailable
            time.sleep(POLL_SECONDS)
        info = json.loads(result.read_text(encoding="utf8"))
        if not info.get("ok"):
            if info.get("reason") == "timeout":
                raise ConverterTimeout
            raise ConversionFailed(info.get("reason"))
        return info, {name: path.read_bytes() for name, path in files.items()}
    finally:
        for path in (source, job, result, *files.values()):
            path.unlink(missing_ok=True)


def xls_to_xlsx(data: bytes) -> bytes:
    return _run("xls_to_xlsx", data, ("xlsx",))[1]["xlsx"]


def pdf(data: bytes, decimal: str = ".") -> tuple[bytes, bytes, float]:
    """Trả (PDF, bản xlsx LibreOffice đã tính để đọc số, số giây). decimal: dấu thập phân của số in trong PDF."""
    info, files = _run("pdf", data, ("pdf", "xlsx"), decimal=decimal)
    return files["pdf"], files["xlsx"], info["seconds"]
