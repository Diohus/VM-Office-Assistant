"""API backend. Không có mạng ra ngoài và không giữ khóa AI. Không tự mở file tải lên: mọi việc đọc và ghi file đi qua reader."""
from __future__ import annotations

import hashlib
import json
import re
import uuid
from contextlib import asynccontextmanager
from typing import Literal
from urllib.parse import quote

from fastapi import FastAPI, HTTPException, Request
from fastapi.exceptions import RequestValidationError
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel, Field

from . import chat, comparison, converter_client, db, fill, llm_client, reader_client, store
from .config import MAX_UPLOAD_BYTES

MAX_EDITS = 5000


@asynccontextmanager
async def lifespan(_: FastAPI):
    await run_in_threadpool(db.migrate)
    yield


app = FastAPI(title="VM-Office-Assistant", docs_url=None, redoc_url=None, openapi_url=None, lifespan=lifespan)


_INVALID = "App không xử lý được yêu cầu này nên chưa làm gì. Bản đang lưu và file của anh/chị không bị đổi. Hãy tải lại trang rồi thử lại; nếu vẫn lỗi, báo người quản lý app."


@app.exception_handler(RequestValidationError)
async def _invalid_request(_request: Request, exc: RequestValidationError) -> JSONResponse:
    """Thay câu lỗi mặc định của FastAPI. Câu mặc định trả lại nguyên giá trị đã gửi, tức số trong ô: không bao giờ trả lại (Q124).

    Chỉ nêu trường nào sai và loại lỗi, không có giá trị, không có `ctx`, không ghi log.
    """
    errors = exc.errors()
    fields = [{"field": ".".join(str(part) for part in e.get("loc", ())), "type": e.get("type")} for e in errors[:20]]
    if any(tuple(e.get("loc", ())) == ("body", "edits") and e.get("type") == "too_long" for e in errors):
        return _problem(413, "EDITS_TOO_MANY", f"Lần lưu này có quá {MAX_EDITS} ô sửa nên app chưa ghi. Bản đang lưu không bị đổi, các sửa vẫn còn trên màn hình. Muốn lưu, cần bỏ hết các sửa này rồi làm lại từng phần: bấm Xuất, chọn Bỏ thay đổi, rồi bấm Hủy. Sau đó dán hoặc gõ lại từng phần, mỗi phần không quá {MAX_EDITS} ô, và bấm Lưu sau mỗi phần.", limit=MAX_EDITS)
    return _problem(422, "INVALID_REQUEST", _INVALID, fields=fields)

_CONTROL = re.compile(r"[\x00-\x1f\x7f]")
_BUSY = "App chưa xử lý được vì hệ thống đang bận hoặc chưa sẵn sàng. File của anh/chị không bị đổi. Hãy thử lại sau ít phút."
_LEGACY_CONVERTED = "Đây là file đời cũ, app đã chuyển sang định dạng mới để mở, và sẽ lưu thành .xlsx. Bản gốc còn nguyên. Trình bày có thể lệch một chút."
_LEGACY_FAILED = "App chưa mở được file đời cũ này. Bản gốc không bị đổi. Hãy thử mở bằng Excel, lưu thành .xlsx rồi tải lại."
_PDF_DOWN = "App chưa xuất được PDF vì phần xử lý PDF đang không chạy. Biểu mẫu của anh/chị không bị đổi. Anh/chị vẫn xuất được file Excel, hoặc thử lại sau."
_PDF_SLOW = "File này lớn nên xuất PDF quá lâu, app đã dừng sau 2 phút. Biểu mẫu của anh/chị không bị đổi. Anh/chị vẫn xuất được file Excel; muốn có PDF thì mở file Excel bằng Excel rồi in ra PDF."
_PDF_FAILED = "App chưa xuất được PDF từ file này. Biểu mẫu của anh/chị không bị đổi. Anh/chị vẫn xuất được file Excel."
_STALE = "Bản này đã có thay đổi mới hơn nên app chưa ghi. Bản đang lưu không bị đổi. Hãy mở lại bản mới nhất rồi sửa lại."


class Edit(BaseModel):
    sheet: str = Field(max_length=255)
    cell: str = Field(max_length=12)
    value: str | int | float | bool | None = None
    formula: str | None = Field(default=None, max_length=8192)


class CellRef(BaseModel):
    sheet: str = Field(max_length=255)
    cell: str = Field(max_length=12)


class SaveRequest(BaseModel):
    base_version: int
    author: str = Field(min_length=1, max_length=120)
    edits: list[Edit] = Field(min_length=1, max_length=MAX_EDITS)
    # Lát 2 (Q135): ô do AI sửa qua khung chat, chưa lưu. Phải nằm trong `edits`; lịch sử bản ghi rõ.
    ai_cells: list[CellRef] = Field(default_factory=list, max_length=MAX_EDITS)


class RestoreRequest(BaseModel):
    base_version: int
    author: str = Field(min_length=1, max_length=120)


class GridValue(BaseModel):
    sheet: str = Field(max_length=255)
    cell: str = Field(max_length=12)
    value: str | int | float | bool | None = None


class PdfRequest(BaseModel):
    version: int | None = None
    decimal: Literal[".", ","] = "."  # dấu thập phân của app (frontend/src/numberFormat.js, Q122), PDF in theo dấu này
    grid_values: list[GridValue] = Field(max_length=200_000)


def _clean_name(raw: str) -> str:
    name = _CONTROL.sub("", raw.replace("\\", "/").rsplit("/", 1)[-1]).strip()
    return name[:200] or "khong_ten.xlsx"


def _problem(status: int, code: str, message: str, **detail) -> JSONResponse:
    return JSONResponse({"code": code, "message": message, "detail": detail}, status_code=status)


async def _read_limited(request: Request) -> bytes | None:
    """Đọc thân yêu cầu, dừng ngay khi vượt giới hạn thay vì đọc hết vào bộ nhớ."""
    chunks, size = [], 0
    async for chunk in request.stream():
        size += len(chunk)
        if size > MAX_UPLOAD_BYTES:
            return None
        chunks.append(chunk)
    return b"".join(chunks)


def _meta(result: dict) -> dict:
    grid = result["grid"]
    return {
        "warnings": result["warnings"],
        "unshown": result["unshown"],
        "external": result["external"],
        "blocked_formulas": result["blocked_formulas"],
        "protected_sheets": result["protected_sheets"],
        "hidden_sheets": [s["name"] for s in grid["sheets"] if s["hidden"]],
        "cells": result["cells"],
    }


def _public(row: dict) -> dict:
    return {
        "document_id": str(row["document_id"]),
        "name": row["name"],
        "version": row["version"],
        "sha256": row["sha256"],
        "size": row["size"],
        "created_at": row["created_at"].isoformat(),
        "author": row["author"],
        "summary": row["summary"],
        **_first_flags(row),
        **row["meta"],
    }


def _first_flags(row: dict) -> dict:
    first = _first_sha(row["document_id"])
    # fill_supported: lát 1 chỉ điền biểu mẫu QC mẫu (Q133, PM 11/11). So SHA bản tải lên đầu tiên.
    return {"written": row["sha256"] != first, "fill_supported": first in fill.QC_TEMPLATE_SHA256S}


def _first_sha(document_id: uuid.UUID) -> str:
    """Bản nào khác bản tải lên đầu tiên thì đã bị bộ ghi bỏ giá trị công thức, nên phải có cảnh báo khi xuất."""
    return db.version(document_id, 1)["sha256"]


def _version_row(row: dict) -> dict:
    return {"version": row["version"], "created_at": row["created_at"].isoformat(), "author": row["author"], "summary": row["summary"], "restored_from": row["restored_from"],
            "ai_cells": len((row.get("meta") or {}).get("ai_cells", [])),
            # K3 (5.4): lịch sử bản nêu ô nào do AI sửa qua chat ("sheet!A1"), không chỉ số ô. Chỉ địa chỉ ô, không có giá trị.
            "ai_cell_list": list((row.get("meta") or {}).get("ai_cells", []))}


@app.get("/api/health")
def health() -> dict:
    db.ping()
    # Bản thử (trần khác mặc định hoặc đồng hồ thử): giao diện hiện nhãn "Bản thử" để không ai nhầm với bản thật.
    return {"status": "ok", "chat_test_settings": chat.load_settings().is_test}


@app.post("/api/documents", status_code=201)
async def upload(request: Request, filename: str):
    name = _clean_name(filename)
    data = await _read_limited(request)
    if data is None:
        return _problem(413, "FILE_TOO_LARGE", "File vượt giới hạn 25 MB nên app không mở. File của anh/chị không bị đổi. Hãy tách file hoặc hỏi người quản lý app.", limit="file")
    if not data:
        return _problem(422, "FILE_CORRUPT", "File không có nội dung nên app không mở được. File gốc của anh/chị không bị đổi. Hãy kiểm tra lại file rồi chọn lại.")
    sha = hashlib.sha256(data).hexdigest()
    try:
        result = await run_in_threadpool(reader_client.read, data, name)
    except reader_client.ReaderUnavailable:
        return _problem(503, "READER_UNAVAILABLE", _BUSY)
    original = None
    if not result["ok"] and result["code"] == "FILE_LEGACY_XLS":
        # File đời cũ: chuyển trong container không mạng, giữ bản gốc. Bản chuyển vẫn phải qua cổng đầu vào như mọi file.
        try:
            converted = await run_in_threadpool(converter_client.xls_to_xlsx, data)
            checked = await run_in_threadpool(reader_client.read, converted, name.rsplit(".", 1)[0] + ".xlsx")
        except (converter_client.ConverterUnavailable, converter_client.ConversionFailed, converter_client.ConverterTimeout, reader_client.ReaderUnavailable):
            return _problem(422, "FILE_LEGACY_FAILED", _LEGACY_FAILED)
        if not checked["ok"]:
            return _problem(422, "FILE_LEGACY_FAILED", _LEGACY_FAILED)
        original = sha
        await run_in_threadpool(store.put_original, sha, data)
        data, result, name = converted, checked, name.rsplit(".", 1)[0] + ".xlsx"
        sha = hashlib.sha256(data).hexdigest()
    if not result["ok"]:
        status = 413 if result["code"] == "FILE_TOO_LARGE" else 422
        return _problem(status, result["code"], result["message"], **result.get("detail", {}))
    await run_in_threadpool(store.put_new, sha, data, json.dumps(result["grid"], ensure_ascii=False).encode("utf8"))
    meta = _meta(result)
    if original is not None:
        meta["warnings"] = [{"code": "FILE_LEGACY_CONVERTED", "message": _LEGACY_CONVERTED}, *meta["warnings"]]
        meta["legacy_original_sha256"] = original
    row = await run_in_threadpool(db.create_document, name, sha, len(data), meta)
    return _public(row)


@app.get("/api/documents")
def recent() -> list[dict]:
    return [{"document_id": str(r["document_id"]), "name": r["name"], "version": r["version"], "updated_at": r["created_at"].isoformat(), "author": r["author"]} for r in db.recent()]


def _key(document_id: str) -> uuid.UUID:
    try:
        return uuid.UUID(document_id)
    except ValueError as exc:
        raise HTTPException(404) from exc


def _find(document_id: str, version: int | None = None) -> dict:
    key = _key(document_id)
    row = db.latest(key) if version is None else db.version(key, version)
    if row is None:
        raise HTTPException(404)
    return row


@app.get("/api/documents/{document_id}")
def get_document(document_id: str, version: int | None = None) -> dict:
    return _public(_find(document_id, version))


@app.get("/api/documents/{document_id}/grid")
def get_grid(document_id: str, version: int | None = None) -> Response:
    row = _find(document_id, version)
    return Response(store.read_grid(row["sha256"]), media_type="application/json")


@app.get("/api/documents/{document_id}/versions")
def list_versions(document_id: str) -> list[dict]:
    rows = db.versions(_key(document_id))
    if not rows:
        raise HTTPException(404)
    return [_version_row(r) for r in rows]


@app.post("/api/documents/{document_id}/versions", status_code=201)
def save_version(document_id: str, body: SaveRequest):
    """Ghi các ô vào bản sao, tạo bản mới. File của các bản trước không bao giờ bị ghi đè."""
    base = _find(document_id, body.base_version)
    if db.latest(_key(document_id))["version"] != body.base_version:
        return _problem(409, "STALE_VERSION", _STALE)
    edits = [e.model_dump() for e in body.edits]
    ai = {(c.sheet, c.cell.upper()) for c in body.ai_cells}
    edited = {(e["sheet"], e["cell"].upper()) for e in edits}
    if not ai <= edited:
        return _problem(422, "AI_CELLS_INVALID", _INVALID)
    try:
        result, data = reader_client.write(store.read(base["sha256"]), base["name"], edits, document_id)
    except reader_client.ReaderUnavailable:
        return _problem(503, "READER_UNAVAILABLE", _BUSY)
    if not result["ok"]:
        return _problem(422, result["code"], result["message"], **result.get("detail", {}))
    if data is None:
        return _problem(503, "READER_UNAVAILABLE", _BUSY)
    sha = hashlib.sha256(data).hexdigest()
    store.put_new(sha, data, json.dumps(result["grid"], ensure_ascii=False).encode("utf8"))
    meta = _meta(result)
    summary = f"Sửa {len(edits)} ô"
    if ai:
        meta["ai_cells"] = sorted(f"{sheet}!{cell}" for sheet, cell in ai)
        summary += f", {len(ai)} ô do AI sửa qua chat"
    for e in edits:
        e["by_ai"] = (e["sheet"], e["cell"].upper()) in ai
    try:
        row = db.add_version(_key(document_id), body.base_version, sha, len(data), meta, body.author, summary, edits)
    except db.StaleVersion:
        return _problem(409, "STALE_VERSION", _STALE)
    return _public(row)


@app.post("/api/documents/{document_id}/versions/{version}/restore", status_code=201)
def restore_version(document_id: str, version: int, body: RestoreRequest):
    """Tạo bản mới có nội dung bản cũ. Lịch sử không mất."""
    old = _find(document_id, version)
    try:
        row = db.add_version(_key(document_id), body.base_version, old["sha256"], old["size"], old["meta"], body.author, f"Khôi phục bản {version}", restored_from=version)
    except db.StaleVersion:
        return _problem(409, "STALE_VERSION", _STALE)
    return _public(row)


@app.post("/api/documents/{document_id}/pdf")
def export_pdf(document_id: str, body: PdfRequest):
    """Xuất PDF do LibreOffice tạo. Số LibreOffice tính được phải khớp số trên lưới, lệch thì không xuất (Q48c)."""
    row = _find(document_id, body.version)
    data = store.read(row["sha256"])
    try:
        pdf, computed, seconds = converter_client.pdf(data, body.decimal)
        cells = reader_client.values(computed)
    except converter_client.ConverterTimeout:
        return _problem(504, "PDF_TOO_SLOW", _PDF_SLOW)
    except converter_client.ConverterUnavailable:
        return _problem(503, "PDF_UNAVAILABLE", _PDF_DOWN)
    except (converter_client.ConversionFailed, reader_client.ReaderUnavailable):
        return _problem(422, "PDF_FAILED", _PDF_FAILED)
    mismatches = comparison.compare(cells, [v.model_dump() for v in body.grid_values])
    if mismatches:
        return _problem(409, "PDF_MISMATCH", f"Số trong PDF khác số trên bảng ở {len(mismatches)} ô, nên app chưa xuất PDF. Biểu mẫu của anh/chị không bị đổi. Hãy kiểm tra các ô được nêu, hoặc xuất file Excel.", cells=mismatches[:200])
    stem = row["name"].rpartition(".")[0] or row["name"]
    return Response(pdf, media_type="application/pdf", headers={
        "Content-Disposition": f"attachment; filename=\"export.pdf\"; filename*=UTF-8''{quote(stem + '_xuat.pdf')}",
        "X-Convert-Seconds": str(seconds),
    })


@app.get("/api/documents/{document_id}/export")
def export(document_id: str, version: int | None = None) -> Response:
    """Xuất đúng từng byte của bản được chọn. Chưa sửa thì giống hệt file đã tải lên."""
    row = _find(document_id, version)
    data = store.read(row["sha256"])
    stem, dot, ext = row["name"].rpartition(".")
    base = stem if dot else row["name"]
    out_name = f"{base}_xuat.{ext if dot else 'xlsx'}"
    return Response(data, media_type="application/octet-stream", headers={
        "Content-Disposition": f"attachment; filename=\"export.xlsx\"; filename*=UTF-8''{quote(out_name)}",
        "X-Content-SHA256": row["sha256"],
    })



# ---- Lát 1 nối AI (Q133, T-DEV-09). Không ghi log tên file hay nội dung lô. ----
_FILL_ONLY_QC = "Bản thử này chỉ điền được biểu mẫu QC mẫu (GIA-QC-01_mau.xlsx). Biểu mẫu của anh/chị không bị đổi. Hãy mở biểu mẫu QC mẫu đội gửi kèm rồi bấm Điền từ file."
_FILL_SOURCE_XLSX = "Bản thử này chỉ điền từ file Excel .xlsx. File của anh/chị không bị đổi. Hãy mở file nguồn bằng Excel, lưu thành .xlsx rồi chọn lại."
_FILL_SOURCE_GONE = "Không tìm thấy file nguồn đã chọn. Biểu mẫu của anh/chị không bị đổi. Hãy chọn lại file nguồn."
_FILL_INVALID = "App không dùng được đề xuất vừa tạo nên chưa hiện gì. Biểu mẫu của anh/chị không bị đổi. Hãy thử lại; nếu vẫn lỗi, báo người quản lý app."
_FILL_CONFLICT = "Một số ô của biểu mẫu này đang chờ duyệt trong một lần điền khác nên app chưa mở lần điền này. Biểu mẫu của anh/chị không bị đổi. Hãy duyệt xong hoặc bỏ lần điền đang mở, rồi điền lại."
_FILL_CHANGED = "Lần điền này vừa được cập nhật ở chỗ khác (có thể ở tab khác) nên app chưa ghi. Bản đang lưu không bị đổi. Hãy tải lại trang rồi duyệt tiếp."
_FILL_WRONG_DOC = "Lần điền này thuộc một biểu mẫu khác với biểu mẫu đang mở nên app chưa ghi. Không biểu mẫu nào bị đổi. Hãy mở lại đúng biểu mẫu rồi Ghi."
_FILL_GONE = "Lần điền này không còn hoặc đã xong. Biểu mẫu của anh/chị không bị đổi. Hãy bấm Điền từ file để tạo lần mới."


class FillRequest(BaseModel):
    source_id: str = Field(max_length=64)
    base_version: int


class Decision(BaseModel):
    change_id: str = Field(max_length=64)
    action: Literal["accept", "reject", "red_accept", "red_choice", "red_typed", "red_blank"]
    index: int | None = None
    text: str | None = Field(default=None, max_length=64)
    seen: str | None = Field(default=None, max_length=64)


class ApplyRequest(BaseModel):
    document_id: str = Field(min_length=1, max_length=64)  # tài liệu đang hiện trên màn hình; phải là tài liệu của lô (PO/BA 11/11)
    base_version: int
    author: str = Field(min_length=1, max_length=120)
    decisions: list[Decision] = Field(min_length=1, max_length=MAX_EDITS)


def _fill_supported(key: uuid.UUID) -> bool:
    return _first_sha(key) in fill.QC_TEMPLATE_SHA256S


def _batch_key(batch_id: str) -> uuid.UUID:
    try:
        return uuid.UUID(batch_id)
    except ValueError as exc:
        raise HTTPException(404) from exc


def _batch_public(row: dict) -> dict:
    return {"batch_id": str(row["batch_id"]), "document_id": str(row["document_id"]), "status": row["status"],
            "source_name": row["source_name"], "base_version": row["base_version"], "batch": row["batch"], "decisions": row["decisions"]}


@app.post("/api/documents/{document_id}/sources", status_code=201)
async def upload_source(document_id: str, request: Request, filename: str):
    """File nguồn: qua cùng cổng đầu vào như biểu mẫu. Chỉ lưu file và SHA; không trả nội dung, không lưu lưới của nguồn."""
    key = _key(document_id)
    if await run_in_threadpool(db.latest, key) is None:
        raise HTTPException(404)
    if not await run_in_threadpool(_fill_supported, key):
        return _problem(422, "FILL_TEMPLATE_UNSUPPORTED", _FILL_ONLY_QC)
    name = _clean_name(filename)
    data = await _read_limited(request)
    if data is None:
        return _problem(413, "FILE_TOO_LARGE", "File vượt giới hạn 25 MB nên app không mở. File của anh/chị không bị đổi. Hãy tách file hoặc hỏi người quản lý app.", limit="file")
    if not data:
        return _problem(422, "FILE_CORRUPT", "File không có nội dung nên app không mở được. File gốc của anh/chị không bị đổi. Hãy kiểm tra lại file rồi chọn lại.")
    try:
        result = await run_in_threadpool(reader_client.read, data, name)
    except reader_client.ReaderUnavailable:
        return _problem(503, "READER_UNAVAILABLE", _BUSY)
    if not result["ok"]:
        if result["code"] == "FILE_LEGACY_XLS":
            return _problem(422, "FILL_SOURCE_XLSX_ONLY", _FILL_SOURCE_XLSX)
        status = 413 if result["code"] == "FILE_TOO_LARGE" else 422
        return _problem(status, result["code"], result["message"], **result.get("detail", {}))
    sha = hashlib.sha256(data).hexdigest()
    await run_in_threadpool(store.put_file, sha, data)
    return await run_in_threadpool(db.add_source, key, name, sha, len(data))


@app.post("/api/documents/{document_id}/fill", status_code=201)
def create_fill(document_id: str, body: FillRequest):
    """Tạo lô đề xuất (draft) trong reader. Lô chưa mở không tính D10 và bị thay khi tạo lô mới (6.1)."""
    key = _key(document_id)
    base = _find(document_id, body.base_version)
    if not _fill_supported(key):
        return _problem(422, "FILL_TEMPLATE_UNSUPPORTED", _FILL_ONLY_QC)
    if db.latest(key)["version"] != body.base_version:
        return _problem(409, "STALE_VERSION", _STALE)
    try:
        src = db.source(key, uuid.UUID(body.source_id))
    except ValueError:
        src = None
    if src is None:
        return _problem(404, "FILL_SOURCE_GONE", _FILL_SOURCE_GONE)
    try:
        batch = reader_client.fill(store.read(base["sha256"]), base["name"], store.read(src["sha256"]), src["name"], fill.SPEC_ID)
    except reader_client.FillRejected as exc:
        return _problem(422, exc.code, exc.message)
    except reader_client.ReaderUnavailable:
        return _problem(503, "READER_UNAVAILABLE", _BUSY)
    grid = json.loads(store.read_grid(base["sha256"]))
    try:
        targets, decisions = fill.check_batch(batch, grid, base["meta"].get("hidden_sheets", []))
    except fill.BatchInvalid:
        return _problem(422, "FILL_INVALID", _FILL_INVALID)
    batch_id = db.create_batch(key, src["source_id"], body.base_version, batch, targets, decisions)
    return _batch_public(db.batch(batch_id))


@app.post("/api/fill/{batch_id}/open")
def open_fill(batch_id: str):
    """Mở lô để duyệt. D10: ô đang chờ ở lô khác đang mở thì từ chối, kể cả hai tab cùng lúc."""
    key = _batch_key(batch_id)
    try:
        db.open_batch(key)
    except LookupError:
        return _problem(404, "FILL_GONE", _FILL_GONE)
    except db.FillConflict as exc:
        return _problem(409, "FILL_CONFLICT", _FILL_CONFLICT, cells=exc.cells[:200])
    return _batch_public(db.batch(key))


@app.delete("/api/fill/{batch_id}")
def cancel_fill(batch_id: str) -> dict:
    """Hủy lô chưa mở: xóa hẳn (6.1). Lô đang mở: bỏ các ô còn chờ, lô xong, ô đã ghi giữ nguyên (N5)."""
    return db.drop_batch(_batch_key(batch_id))


@app.get("/api/fill/{batch_id}")
def get_fill(batch_id: str) -> dict:
    row = db.batch(_batch_key(batch_id))
    if row is None:
        raise HTTPException(404)
    return _batch_public(row)


@app.get("/api/documents/{document_id}/fill")
def list_fill(document_id: str) -> list[dict]:
    """Lô đang mở của tài liệu, để mở lại khi còn ô đỏ chờ (3.8a)."""
    return [{"batch_id": str(r["batch_id"]), "source_name": r["source_name"], "created_at": r["created_at"].isoformat()} for r in db.open_batches(_key(document_id))]


@app.post("/api/fill/{batch_id}/apply")
def apply_fill(batch_id: str, body: ApplyRequest):
    """Ghi các ô người dùng đã nhận thành bản mới. Chỉ nhận change_id và lựa chọn; số lấy từ lô đã lưu (4.6)."""
    key = _batch_key(batch_id)
    row = db.batch(key)
    if row is None or row["status"] != "open":
        return _problem(409, "FILL_GONE", _FILL_GONE)
    document_id = str(row["document_id"])
    if body.document_id != document_id:
        return _problem(409, "FILL_WRONG_DOCUMENT", _FILL_WRONG_DOC)
    latest = _find(document_id)
    if latest["version"] != body.base_version:
        return _problem(409, "STALE_VERSION", _STALE)
    current_grid = json.loads(store.read_grid(latest["sha256"]))
    try:
        decisions, edits, review_again, counts = fill.decide(row["batch"], row["decisions"], [d.model_dump() for d in body.decisions], current_grid)
    except fill.DecisionInvalid as exc:
        return _problem(422, "FILL_DECISION_INVALID", exc.message, change_id=exc.change_id)
    except fill.BatchInvalid:
        return _problem(422, "FILL_INVALID", _FILL_INVALID)
    write = None
    if edits:
        try:
            result, data = reader_client.write(store.read(latest["sha256"]), latest["name"], edits, document_id)
        except reader_client.ReaderUnavailable:
            return _problem(503, "READER_UNAVAILABLE", _BUSY)
        if not result["ok"]:
            return _problem(422, result["code"], result["message"], **result.get("detail", {}))
        if data is None:
            return _problem(503, "READER_UNAVAILABLE", _BUSY)
        sha = hashlib.sha256(data).hexdigest()
        store.put_new(sha, data, json.dumps(result["grid"], ensure_ascii=False).encode("utf8"))
        write = {"sha": sha, "size": len(data), "meta": _meta(result), "author": body.author,
                 "summary": fill.summary(row["source_name"], counts, len(edits)), "edits": edits}
    try:
        version_row = db.apply_batch(key, row["document_id"], body.base_version, decisions, write, row["decisions"])
    except db.StaleVersion:
        return _problem(409, "STALE_VERSION", _STALE)
    except db.FillConflict as exc:
        return _problem(409, "FILL_CONFLICT", _FILL_CONFLICT, cells=exc.cells[:200])
    except db.BatchChanged:
        return _problem(409, "FILL_BATCH_CHANGED", _FILL_CHANGED)
    except LookupError:
        return _problem(409, "FILL_GONE", _FILL_GONE)
    pending = sum(1 for d in decisions.values() if d.get("state") == "pending")
    return {"version": _public(version_row) if version_row else None, "written": len(edits), "pending": pending,
            "review_again": review_again, "decisions": decisions}


# ---------- Lát 2: khung chat (Q135), hợp đồng docs/contracts/chat_edit_plan.md ----------

_CHAT_TEXT = {
    "CHAT_PREVIEW_EXPIRED": "Bản xem trước đã hết hạn hoặc đã gửi. Bảng không đổi. Anh/chị gửi lại yêu cầu.",
    "CHAT_BUSY": "AI đang làm yêu cầu trước của tài liệu này. Anh/chị chờ xong rồi gửi tiếp.",
    "CHAT_TURN_LIMIT": "Yêu cầu này vượt hạn mức chi phí một lần gửi. Anh/chị chọn vùng nhỏ hơn rồi gửi lại.",
    "CHAT_DAY_LIMIT": "Hôm nay đã dùng hết hạn mức gọi AI. Khung chat dùng lại được từ 0 giờ ngày mai (giờ Việt Nam). Anh/chị vẫn sửa tay, Hoàn tác và Lưu bình thường.",
    "CHAT_MODEL": "Chưa gọi được AI lần này. Bảng không đổi. Anh/chị thử lại sau.",
}
_CHAT_STATUS = {"CHAT_BUSY": 409, "CHAT_TURN_LIMIT": 422, "CHAT_DAY_LIMIT": 429, "CHAT_NO_PRICE": 503}


class ChatCellIn(BaseModel):
    a1: str = Field(max_length=12)
    t: Literal["number", "date", "boolean", "text"]
    v: str | None = Field(default=None, max_length=2000)
    f: str | None = Field(default=None, max_length=8192)
    display: str | None = Field(default=None, max_length=2000)


class ChatSheetIn(BaseModel):
    name: str = Field(max_length=255)
    region: str = Field(max_length=32)
    cells: list[ChatCellIn] = Field(max_length=3000)
    merges: list[str] = Field(default_factory=list, max_length=3000)
    # ChatEditPlan 0.2 (N2): vùng chỉ để đọc chứa region (tiêu đề, cột Ngày). vmai kiểm và chặn ghi ngoài region.
    context_region: str | None = Field(default=None, max_length=32)


class ChatPreviewRequest(BaseModel):
    base_version: int
    active_sheet: str = Field(max_length=255)
    sheets: list[ChatSheetIn] = Field(min_length=1, max_length=20)
    message: str = Field(max_length=4000)  # vmai báo CHAT_MESSAGE_TOO_LONG từ 1.000 ký tự


def _chat_problem(status: int, code: str, message_key: str | None = None, **detail) -> JSONResponse:
    return _problem(status, code, _CHAT_TEXT.get(message_key or code, _CHAT_TEXT["CHAT_MODEL"]), **detail)


@app.post("/api/documents/{document_id}/chat/preview")
def chat_preview(document_id: str, body: ChatPreviewRequest):
    """Bước 1: reader che ô người, người bệnh và dựng đúng chuỗi sẽ gửi. Bản này chỉ giữ trong bộ nhớ, 10 phút."""
    key = _key(document_id)
    _find(document_id, body.base_version)
    request = body.model_dump()
    request["base_version"] = str(body.base_version)
    # context_region chỉ chuyển khi có (ChatSheet của vmai cấm trường lạ; trình duyệt không gửi thì giữ đúng như 0.1).
    for sheet in request["sheets"]:
        if sheet.get("context_region") is None:
            sheet.pop("context_region", None)
    try:
        prepared = reader_client.chat_prepare(request)
    except reader_client.ChatRejected as exc:
        # Câu báo do giao diện chọn theo mã (AI Engineer viết, PO/BA duyệt). `items` chỉ là ô, số ô, không có giá trị.
        return _problem(422, exc.code, _INVALID, items=exc.items)
    except reader_client.ReaderUnavailable:
        return _problem(503, "READER_UNAVAILABLE", _BUSY)
    # Trần lượt, trần ngày: báo ngay ở bước xem trước, cùng mã và câu như bước gửi. Không tạo bản xem trước.
    settings = chat.load_settings()
    try:
        chat.TURNS.peek(prepared["estimated_cost_usd"], db.chat_spent(chat.vn_day(settings.now())), settings)
    except chat.BudgetRefused as exc:
        return _chat_problem(_CHAT_STATUS.get(exc.code, 422), exc.code)
    preview_id, item = chat.PREVIEWS.put(key, body.base_version, prepared)
    return {
        "preview_id": preview_id,
        "prompt": prepared["prompt"],
        "prompt_sha256": item.prompt_sha256,
        "masked": prepared["masked"],
        "cell_count": prepared["cell_count"],
        "region_cells": prepared["region_cells"],
        "estimated_cost_usd": prepared["estimated_cost_usd"],
        "expires_in": chat.PREVIEW_SECONDS,
        "chat_test_settings": chat.load_settings().is_test,
    }


def _record(item: chat.Preview, at, status: str, reason: str | None, cost: float, plan: dict | None = None, llm: dict | None = None) -> None:
    prepared = item.prepared
    db.add_chat_call({
        "document_id": item.document_id, "at": at, "vn_day": chat.vn_day(at), "status": status, "reason": reason,
        "cells_sent": prepared["cell_count"], "masked_groups": len(prepared["masked"]),
        "ops": len(plan["ops"]) if plan else 0,
        "rejected_codes": [str(r.get("reason"))[:40] for r in plan["rejected_ops"]] if plan else [],
        "input_tokens": (llm or {}).get("input_tokens"), "output_tokens": (llm or {}).get("output_tokens"),
        "estimated_cost_usd": prepared["estimated_cost_usd"] or 0.0, "cost_usd": cost,
    })


@app.post("/api/chat/{preview_id}/send")
def chat_send(preview_id: str):
    """Bước 2: kiểm trần, gửi đúng bản đã xem cho llm, rồi reader chạy 7 cổng trên kết quả. Lưu số đếm, không lưu nội dung."""
    item = chat.PREVIEWS.take(preview_id[:64])
    if item is None:
        return _chat_problem(404, "CHAT_PREVIEW_EXPIRED")
    prepared = item.prepared
    settings = chat.load_settings()
    at = settings.now()
    estimated = prepared["estimated_cost_usd"]
    try:
        chat.TURNS.start(item.document_id, estimated, db.chat_spent(chat.vn_day(at)), settings)
    except chat.BudgetRefused as exc:
        return _chat_problem(_CHAT_STATUS[exc.code], exc.code)
    try:
        try:
            out = llm_client.call(prepared["system"], prepared["prompt"], prepared["response_schema"], prepared["schema_name"], prepared["max_output_tokens"])
        except llm_client.LlmFailed as exc:
            _record(item, at, "model_error", exc.code, chat.cost_of_failure(exc.code, estimated, exc.not_sent))
            return _chat_problem(502, exc.code, "CHAT_MODEL")
        # LlmCall của vmai không có reasoning_tokens, prompt_sha256 (extra=forbid).
        llm_call = {"purpose": "chat_edit", "provider": "openai", "model": out["model"] or "gpt-6-luna", "request_id": out["request_id"],
                    "input_tokens": out["input_tokens"], "output_tokens": out["output_tokens"], "latency_ms": out["latency_ms"]}
        try:
            plan = reader_client.chat_check(prepared["context"], out["text"], llm_call, prepared["masked"])
        except (reader_client.ChatRejected, reader_client.ReaderUnavailable):
            _record(item, at, "check_failed", "CHAT_CHECK_FAILED", estimated, llm=out)
            return _problem(503, "READER_UNAVAILABLE", _BUSY)
        cost = (plan.get("llm_call") or {}).get("cost_usd")
        cost = float(cost) if isinstance(cost, (int, float)) and not isinstance(cost, bool) else float(estimated)
        _record(item, at, str(plan.get("status"))[:20], plan.get("reason"), cost, plan, out)
    finally:
        chat.TURNS.finish(item.document_id)
    return {"plan": plan, "base_version": item.base_version, "prompt_sha256": item.prompt_sha256, "cost_usd": cost,
            "chat_test_settings": settings.is_test}
