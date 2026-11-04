"""Gửi file cho container reader qua thư mục spool và chờ kết quả. Kết quả được coi là dữ liệu không đáng tin."""
from __future__ import annotations

import json
import os
import tempfile
import time
import uuid

from pydantic import BaseModel, ConfigDict, ValidationError

from .config import MAX_READER_OUTPUT_BYTES, READER_TIMEOUT_SECONDS, SPOOL_DIR

POLL_SECONDS = 0.05


class ReaderUnavailable(Exception):
    pass


class _Rejected(BaseModel):
    model_config = ConfigDict(extra="ignore")
    ok: bool
    code: str
    message: str
    detail: dict = {}


class _Accepted(BaseModel):
    model_config = ConfigDict(extra="ignore")
    ok: bool
    warnings: list[dict]
    has_macro: bool
    cells: int
    protected_sheets: list[str]
    unshown: list[dict]
    external: list[dict]
    blocked_formulas: list[dict]
    grid: dict


def _write_atomic(path, data: bytes) -> None:
    fd, tmp = tempfile.mkstemp(dir=SPOOL_DIR, suffix=".tmp")
    with os.fdopen(fd, "wb") as fh:
        fh.write(data)
    os.replace(tmp, path)


def read(data: bytes, filename: str) -> dict:
    job_id = uuid.uuid4().hex
    source, job, out = SPOOL_DIR / f"{job_id}.xlsx", SPOOL_DIR / f"{job_id}.job", SPOOL_DIR / f"{job_id}.out.json"
    _write_atomic(source, data)
    _write_atomic(job, filename.encode("utf8"))   # ghi sau cùng: reader chỉ nhận việc khi file đã đủ
    deadline = time.monotonic() + READER_TIMEOUT_SECONDS
    try:
        while not out.exists():
            if time.monotonic() > deadline:
                raise ReaderUnavailable
            time.sleep(POLL_SECONDS)
        if out.stat().st_size > MAX_READER_OUTPUT_BYTES:
            raise ReaderUnavailable
        raw = json.loads(out.read_text(encoding="utf8"))
        model = _Accepted if raw.get("ok") else _Rejected
        return model.model_validate(raw).model_dump()
    except (ValidationError, ValueError) as exc:
        raise ReaderUnavailable from exc
    finally:
        for path in (source, job, out):
            path.unlink(missing_ok=True)
