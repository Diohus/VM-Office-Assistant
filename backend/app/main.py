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

from . import comparison, converter_client, db, reader_client, store
from .config import MAX_UPLOAD_BYTES

MAX_EDITS = 5000


@asynccontextmanager
async def lifespan(_: FastAPI):
    await run_in_threadpool(db.migrate)
    yield


app = FastAPI(title="VM-Office-Assistant", docs_url=None, redoc_url=None, openapi_url=None, lifespan=lifespan)


_INVALID = "App không xử lý được yêu cầu này nên chưa làm gì. Bản đang lưu và file của bạn không bị đổi. Hãy tải lại trang rồi thử lại; nếu vẫn lỗi, báo người quản lý app."


@app.exception_handler(RequestValidationError)
async def _invalid_request(_request: Request, exc: RequestValidationError) -> JSONResponse:
    """Thay câu lỗi mặc định của FastAPI. Câu mặc định trả lại nguyên giá trị đã gửi, tức số trong ô: không bao giờ trả lại (Q124).

    Chỉ nêu trường nào sai và loại lỗi, không có giá trị, không có `ctx`, không ghi log.
    """
    errors = exc.errors()
    fields = [{"field": ".".join(str(part) for part in e.get("loc", ())), "type": e.get("type")} for e in errors[:20]]
    if any(tuple(e.get("loc", ())) == ("body", "edits") and e.get("type") == "too_long" for e in errors):
        return _problem(413, "EDITS_TOO_MANY", f"Lần lưu này có quá {MAX_EDITS} ô sửa nên app chưa ghi. Bản đang lưu không bị đổi. Hãy lưu sau mỗi lần sửa ít hơn {MAX_EDITS} ô.", limit=MAX_EDITS)
    return _problem(422, "INVALID_REQUEST", _INVALID, fields=fields)

_CONTROL = re.compile(r"[\x00-\x1f\x7f]")
_BUSY = "App chưa xử lý được vì hệ thống đang bận hoặc chưa sẵn sàng. File của bạn không bị đổi. Hãy thử lại sau ít phút."
_LEGACY_CONVERTED = "Đây là file đời cũ, app đã chuyển sang định dạng mới để mở, và sẽ lưu thành .xlsx. Bản gốc còn nguyên. Trình bày có thể lệch một chút."
_LEGACY_FAILED = "App chưa mở được file đời cũ này. Bản gốc không bị đổi. Hãy thử mở bằng Excel, lưu thành .xlsx rồi tải lại."
_PDF_DOWN = "App chưa xuất được PDF vì phần xử lý PDF đang không chạy. Biểu mẫu của bạn không bị đổi. Bạn vẫn xuất được file Excel, hoặc thử lại sau."
_PDF_FAILED = "App chưa xuất được PDF từ file này. Biểu mẫu của bạn không bị đổi. Bạn vẫn xuất được file Excel."
_STALE = "Bản này đã có thay đổi mới hơn nên app chưa ghi. Bản đang lưu không bị đổi. Hãy mở lại bản mới nhất rồi sửa lại."


class Edit(BaseModel):
    sheet: str = Field(max_length=255)
    cell: str = Field(max_length=12)
    value: str | int | float | bool | None = None
    formula: str | None = Field(default=None, max_length=8192)


class SaveRequest(BaseModel):
    base_version: int
    author: str = Field(min_length=1, max_length=120)
    edits: list[Edit] = Field(min_length=1, max_length=MAX_EDITS)


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
        "written": row["sha256"] != _first_sha(row["document_id"]),
        **row["meta"],
    }


def _first_sha(document_id: uuid.UUID) -> str:
    """Bản nào khác bản tải lên đầu tiên thì đã bị bộ ghi bỏ giá trị công thức, nên phải có cảnh báo khi xuất."""
    return db.version(document_id, 1)["sha256"]


def _version_row(row: dict) -> dict:
    return {"version": row["version"], "created_at": row["created_at"].isoformat(), "author": row["author"], "summary": row["summary"], "restored_from": row["restored_from"]}


@app.get("/api/health")
def health() -> dict:
    db.ping()
    return {"status": "ok"}


@app.post("/api/documents", status_code=201)
async def upload(request: Request, filename: str):
    name = _clean_name(filename)
    data = await _read_limited(request)
    if data is None:
        return _problem(413, "FILE_TOO_LARGE", "File vượt giới hạn 25 MB nên app không mở. File của bạn không bị đổi. Hãy tách file hoặc hỏi người quản lý app.", limit="file")
    if not data:
        return _problem(422, "FILE_CORRUPT", "File không có nội dung nên app không mở được. File gốc của bạn không bị đổi. Hãy kiểm tra lại file rồi chọn lại.")
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
        except (converter_client.ConverterUnavailable, converter_client.ConversionFailed, reader_client.ReaderUnavailable):
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
    try:
        row = db.add_version(_key(document_id), body.base_version, sha, len(data), _meta(result), body.author, f"Sửa {len(edits)} ô", edits)
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
    except converter_client.ConverterUnavailable:
        return _problem(503, "PDF_UNAVAILABLE", _PDF_DOWN)
    except (converter_client.ConversionFailed, reader_client.ReaderUnavailable):
        return _problem(422, "PDF_FAILED", _PDF_FAILED)
    mismatches = comparison.compare(cells, [v.model_dump() for v in body.grid_values])
    if mismatches:
        return _problem(409, "PDF_MISMATCH", f"Số trong PDF khác số trên bảng ở {len(mismatches)} ô, nên app chưa xuất PDF. Biểu mẫu của bạn không bị đổi. Hãy kiểm tra các ô được nêu, hoặc xuất file Excel.", cells=mismatches[:200])
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
