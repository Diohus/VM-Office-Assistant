"""API backend. Không có mạng ra ngoài và không giữ khóa AI. Không tự mở file tải lên: mọi việc đọc file đi qua reader."""
from __future__ import annotations

import hashlib
import json
import re
import uuid
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import JSONResponse, Response

from . import db, reader_client, store
from .config import MAX_UPLOAD_BYTES


@asynccontextmanager
async def lifespan(_: FastAPI):
    await run_in_threadpool(db.migrate)
    yield


app = FastAPI(title="VM-Office-Assistant", docs_url=None, redoc_url=None, openapi_url=None, lifespan=lifespan)

_CONTROL = re.compile(r"[\x00-\x1f\x7f]")


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


def _public(row: dict) -> dict:
    meta = row["meta"]
    return {
        "document_id": str(row["document_id"]),
        "name": row["name"],
        "version": row["version"],
        "sha256": row["sha256"],
        "size": row["size"],
        "created_at": row["created_at"].isoformat(),
        "warnings": meta["warnings"],
        "unshown": meta["unshown"],
        "external": meta["external"],
        "blocked_formulas": meta["blocked_formulas"],
        "protected_sheets": meta["protected_sheets"],
        "hidden_sheets": meta["hidden_sheets"],
        "cells": meta["cells"],
    }


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
        return _problem(503, "READER_UNAVAILABLE", "App chưa đọc được file vì hệ thống đang bận hoặc chưa sẵn sàng. File của bạn không bị đổi. Hãy thử lại sau ít phút.")
    if not result["ok"]:
        status = 413 if result["code"] == "FILE_TOO_LARGE" else 422
        return _problem(status, result["code"], result["message"], **result.get("detail", {}))
    grid = result["grid"]
    await run_in_threadpool(store.put_new, sha, data, json.dumps(grid, ensure_ascii=False).encode("utf8"))
    meta = {
        "warnings": result["warnings"],
        "unshown": result["unshown"],
        "external": result["external"],
        "blocked_formulas": result["blocked_formulas"],
        "protected_sheets": result["protected_sheets"],
        "hidden_sheets": [s["name"] for s in grid["sheets"] if s["hidden"]],
        "cells": result["cells"],
    }
    row = await run_in_threadpool(db.create_document, name, sha, len(data), meta)
    return _public(row)


@app.get("/api/documents")
def recent() -> list[dict]:
    return [{"document_id": str(r["document_id"]), "name": r["name"], "created_at": r["created_at"].isoformat()} for r in db.recent()]


def _find(document_id: str) -> dict:
    try:
        key = uuid.UUID(document_id)
    except ValueError as exc:
        raise HTTPException(404) from exc
    row = db.latest(key)
    if row is None:
        raise HTTPException(404)
    return row


@app.get("/api/documents/{document_id}")
def get_document(document_id: str) -> dict:
    return _public(_find(document_id))


@app.get("/api/documents/{document_id}/grid")
def get_grid(document_id: str) -> Response:
    row = _find(document_id)
    return Response(store.read_grid(row["sha256"]), media_type="application/json")


@app.get("/api/documents/{document_id}/export")
def export(document_id: str) -> Response:
    """Xuất lại đúng từng byte của bản đang lưu. Chưa sửa thì giống hệt file đã tải lên."""
    row = _find(document_id)
    data = store.read(row["sha256"])
    stem, dot, ext = row["name"].rpartition(".")
    base = stem if dot else row["name"]
    out_name = f"{base}_xuat.{ext if dot else 'xlsx'}"
    from urllib.parse import quote

    return Response(data, media_type="application/octet-stream", headers={
        "Content-Disposition": f"attachment; filename=\"export.xlsx\"; filename*=UTF-8''{quote(out_name)}",
        "X-Content-SHA256": row["sha256"],
    })
