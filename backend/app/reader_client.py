"""Gửi việc cho container reader qua thư mục spool và chờ kết quả. Kết quả được coi là dữ liệu không đáng tin."""
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


def _run(data: bytes, job: dict) -> tuple[dict, bytes | None]:
    job_id = uuid.uuid4().hex
    source, job_file, out, out_file = (SPOOL_DIR / f"{job_id}{suffix}" for suffix in (".xlsx", ".job", ".out.json", ".out.xlsx"))
    _write_atomic(source, data)
    _write_atomic(job_file, json.dumps(job, ensure_ascii=False).encode("utf8"))   # ghi sau cùng: reader chỉ nhận việc khi file đã đủ
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
        return model.model_validate(raw).model_dump(), out_file.read_bytes() if raw.get("ok") and out_file.exists() else None
    except (ValidationError, ValueError) as exc:
        raise ReaderUnavailable from exc
    finally:
        for path in (source, job_file, out, out_file):
            path.unlink(missing_ok=True)


def values(data: bytes) -> list[dict]:
    job_id = uuid.uuid4().hex
    source, job, out = SPOOL_DIR / f"{job_id}.xlsx", SPOOL_DIR / f"{job_id}.job", SPOOL_DIR / f"{job_id}.out.json"
    _write_atomic(source, data)
    _write_atomic(job, json.dumps({"kind": "values"}).encode("utf8"))
    deadline = time.monotonic() + READER_TIMEOUT_SECONDS
    try:
        while not out.exists():
            if time.monotonic() > deadline:
                raise ReaderUnavailable
            time.sleep(POLL_SECONDS)
        raw = json.loads(out.read_text(encoding="utf8"))
        if not raw.get("ok"):
            raise ReaderUnavailable
        return raw["cells"]
    except ValueError as exc:
        raise ReaderUnavailable from exc
    finally:
        for path in (source, job, out):
            path.unlink(missing_ok=True)


# Lô lát 1 tối đa 5 MB (vmai kiểm); thêm khoảng dư cho phần bọc kết quả.
MAX_FILL_OUTPUT_BYTES = 6 * 1024 * 1024


class FillRejected(Exception):
    """vmai từ chối tạo lô (FillError) hoặc tiến trình con không xong. Câu `message` do vmai hoặc reader soạn."""

    def __init__(self, code: str, message: str, detail: dict):
        super().__init__(code)
        self.code, self.message, self.detail = code, message, detail


def fill(template: bytes, template_name: str, source: bytes, source_name: str, spec_id: str) -> dict:
    """Tạo lô đề xuất trong reader (không mạng, không key): biểu mẫu ở <id>.xlsx, nguồn ở <id>.src.xlsx. Trả lô dạng dict."""
    job_id = uuid.uuid4().hex
    tpl, src, job, out = (SPOOL_DIR / f"{job_id}{s}" for s in (".xlsx", ".src.xlsx", ".job", ".out.json"))
    _write_atomic(tpl, template)
    _write_atomic(src, source)
    _write_atomic(job, json.dumps({"kind": "fill", "filename": template_name, "source_name": source_name, "spec_id": spec_id}, ensure_ascii=False).encode("utf8"))
    deadline = time.monotonic() + READER_TIMEOUT_SECONDS
    try:
        while not out.exists():
            if time.monotonic() > deadline:
                raise ReaderUnavailable
            time.sleep(POLL_SECONDS)
        if out.stat().st_size > MAX_FILL_OUTPUT_BYTES:
            raise ReaderUnavailable
        raw = json.loads(out.read_text(encoding="utf8"))
        if raw.get("ok") is True and isinstance(raw.get("batch"), dict):
            return raw["batch"]
        rejected = _Rejected.model_validate(raw)
        raise FillRejected(rejected.code, rejected.message, rejected.detail)
    except (ValidationError, ValueError) as exc:
        raise ReaderUnavailable from exc
    finally:
        for path in (tpl, src, job, out):
            path.unlink(missing_ok=True)


def read(data: bytes, filename: str) -> dict:
    return _run(data, {"kind": "read", "filename": filename})[0]


def write(data: bytes, filename: str, edits: list[dict], document_uuid: str) -> tuple[dict, bytes | None]:
    return _run(data, {"kind": "write", "filename": filename, "edits": edits, "document_uuid": document_uuid})


# Lát 2 (Q135): việc chat chỉ có JSON, không có file. Kết quả nhỏ (vùng tối đa 3.000 ô).
MAX_CHAT_OUTPUT_BYTES = 4 * 1024 * 1024


def _json_job(job: dict) -> dict:
    job_id = uuid.uuid4().hex
    job_file, out = SPOOL_DIR / f"{job_id}.job", SPOOL_DIR / f"{job_id}.out.json"
    _write_atomic(job_file, json.dumps(job, ensure_ascii=False).encode("utf8"))
    deadline = time.monotonic() + READER_TIMEOUT_SECONDS
    try:
        while not out.exists():
            if time.monotonic() > deadline:
                raise ReaderUnavailable
            time.sleep(POLL_SECONDS)
        if out.stat().st_size > MAX_CHAT_OUTPUT_BYTES:
            raise ReaderUnavailable
        raw = json.loads(out.read_text(encoding="utf8"))
        if not isinstance(raw, dict) or not isinstance(raw.get("ok"), bool):
            raise ReaderUnavailable
        return raw
    except ValueError as exc:
        raise ReaderUnavailable from exc
    finally:
        for path in (job_file, out):
            path.unlink(missing_ok=True)


class ChatRejected(Exception):
    """vmai từ chối đầu vào (ChatInputError) hoặc không kiểm được. Chỉ có mã và danh sách ngắn, không có nội dung ô."""

    def __init__(self, code: str, items: list[str]):
        super().__init__(code)
        self.code, self.items = code, items


def _code(raw: dict, default: str) -> str:
    code = raw.get("code")
    return code if isinstance(code, str) and code.startswith("CHAT_") and len(code) <= 40 else default


def chat_prepare(request: dict) -> dict:
    """ChatPrepared dạng dict. Lỗi đầu vào: ChatRejected."""
    raw = _json_job({"kind": "chat_prepare", "request": request})
    if raw["ok"] and isinstance(raw.get("prepared"), dict):
        return raw["prepared"]
    items = raw.get("items") if isinstance(raw.get("items"), list) else []
    raise ChatRejected(_code(raw, "CHAT_BAD_INPUT"), [str(i)[:40] for i in items][:20])


def chat_check(context: dict, raw_text: str, llm_call: dict, masked: list) -> dict:
    """ChatEditPlan dạng dict."""
    raw = _json_job({"kind": "chat_check", "context": context, "raw_text": raw_text, "llm_call": llm_call, "masked": masked})
    if raw["ok"] and isinstance(raw.get("plan"), dict):
        return raw["plan"]
    raise ChatRejected(_code(raw, "CHAT_CHECK_FAILED"), [])
