"""Tiến trình con xử lý đúng một việc rồi thoát. Việc là `read` (mở file) hoặc `write` (ghi ô vào bản sao)."""
from __future__ import annotations

import json
import os
import pathlib
import re
import sys

import openpyxl

from . import writer
from .job import process

_VOLATILE = re.compile(r"\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(", re.I)
_WRITE_FAILED = "App chưa ghi được vào file này. Bản đang lưu của anh/chị không bị đổi. Hãy thử lại, hoặc hỏi người quản lý app."


def _read(source: pathlib.Path, job: dict) -> dict:
    return process(source.read_bytes(), job["filename"], source)


def _values(source: pathlib.Path) -> dict:
    """Số LibreOffice tính được của mọi ô công thức, để so với lưới trước khi xuất PDF (Q48c)."""
    formulas = openpyxl.load_workbook(source)
    computed = openpyxl.load_workbook(source, data_only=True)
    out = []
    for ws in formulas.worksheets:
        values = computed[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if c.data_type == "f":
                    text = c.value.text if hasattr(c.value, "text") else str(c.value)
                    value = values[c.coordinate].value
                    out.append({"sheet": ws.title, "cell": c.coordinate, "value": value.isoformat() if hasattr(value, "isoformat") else value, "volatile": bool(_VOLATILE.search(text))})
    return {"ok": True, "cells": out}


def _write(source: pathlib.Path, job: dict, out_file: pathlib.Path) -> dict:
    try:
        written = writer.apply(source.read_bytes(), job["edits"], job["document_uuid"])
    except writer.WriteRefused as refused:
        return {"ok": False, "code": "WRITE_REFUSED", "message": "Có ô không ghi được nên app chưa ghi gì. Bản đang lưu của anh/chị không bị đổi. Hãy xem các ô được nêu.", "detail": {"cells": refused.cells}}
    except writer.PatchError:
        return {"ok": False, "code": "WRITE_FAILED", "message": _WRITE_FAILED, "detail": {}}
    tmp = out_file.with_name("dang_ghi_" + out_file.name)  # thư viện đọc xlsx xét đuôi file nên phải giữ .xlsx
    tmp.write_bytes(written["data"])
    # Đọc lại chính file vừa ghi: nếu không đọc được thì không có bản mới nào được tạo.
    result = process(written["data"], job["filename"], tmp)
    if not result["ok"]:
        tmp.unlink(missing_ok=True)
        return {"ok": False, "code": "WRITE_FAILED", "message": _WRITE_FAILED, "detail": {"reason": result["code"]}}
    os.replace(tmp, out_file)
    return result


def _fill(template: pathlib.Path, job: dict) -> dict:
    """Lát 1 (Q133): tạo lô đề xuất bằng vmai, không model, tối đa vàng. Nguồn nằm cạnh biểu mẫu, tên cố định <id>.src.xlsx.
    Không ghi log tên file hay nội dung. Lỗi của vmai đã có câu tiếng Việt và không chứa tên file (FillError)."""
    from vmai import FillError, fill_from_excel
    from vmai.fill import user_message

    source = template.with_name(template.name.removesuffix(".xlsx") + ".src.xlsx")
    try:
        batch = fill_from_excel(template, source, spec_id=str(job["spec_id"]), template_name=str(job["filename"]), source_name=str(job["source_name"]))
    except FillError as exc:
        return {"ok": False, "code": exc.code, "message": exc.message, "detail": {"reason": str(exc.detail)}}
    except Exception as exc:  # noqa: BLE001 - lỗi lạ: chỉ trả tên loại lỗi, không trả nội dung
        return {"ok": False, "code": "FILL_FAILED", "message": user_message("FILL_FAILED"), "detail": {"reason": type(exc).__name__}}
    return {"ok": True, "batch": json.loads(batch.model_dump_json())}


def main() -> int:
    source, job_path, out = map(pathlib.Path, sys.argv[1:4])
    job = json.loads(job_path.read_text(encoding="utf8"))
    if job["kind"] == "write":
        result = _write(source, job, out.with_name(out.name.replace(".out.json", ".out.xlsx")))
    elif job["kind"] == "values":
        result = _values(source)
    elif job["kind"] == "fill":
        result = _fill(source, job)
    else:
        result = _read(source, job)
    tmp = out.with_name(out.name + ".tmp")
    tmp.write_text(json.dumps(result, ensure_ascii=False), encoding="utf8")
    os.replace(tmp, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
