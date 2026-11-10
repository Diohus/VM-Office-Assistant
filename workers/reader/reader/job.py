"""Xử lý một file: cổng đầu vào, đếm ô, kiểm kê, dịch. Kết quả là JSON, không có gì khác rời khỏi đây."""
from __future__ import annotations

import io
import zipfile

from . import gate, inventory, translate
from .ooxml import Package


def process(data: bytes, filename: str, path) -> dict:
    try:
        accepted = gate.check(data, filename)
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            try:
                pkg = Package(zf)
                cells = inventory.count_cells(pkg)
                if cells > gate.MAX_CELLS:
                    raise gate.Rejected("FILE_TOO_LARGE", f"File có {cells:,} ô dữ liệu".replace(",", ".") + ", vượt giới hạn 200.000 ô nên app không mở. File của anh/chị không bị đổi. Hãy tách file hoặc hỏi người quản lý app.", limit="cells", cells=cells)
                unshown = inventory.unshown(pkg)
                external = inventory.external_refs(pkg)
                protected = inventory.protected_sheets(pkg)
                out = translate.translate(path, pkg)
            except gate.Rejected:
                raise
            except Exception as exc:  # noqa: BLE001 - mọi lỗi của thư viện đọc trên file lạ đều là "không đọc được"
                raise gate.Rejected("FILE_CORRUPT", gate.MSG_CORRUPT, reason=type(exc).__name__) from exc
    except gate.Rejected as rej:
        return {"ok": False, "code": rej.code, "message": rej.message, "detail": rej.detail}
    return {
        "ok": True,
        "warnings": accepted.warnings,
        "has_macro": accepted.has_macro,
        "cells": cells,
        "protected_sheets": protected,
        "unshown": unshown + out["unshown_extra"],
        "external": external,
        "blocked_formulas": out["blocked_formulas"],
        "grid": out["grid"],
    }
