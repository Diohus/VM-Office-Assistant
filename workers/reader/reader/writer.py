"""Ghi các ô đã được duyệt vào bản sao của biểu mẫu, theo Q48. Chạy trong reader: không mạng, không khóa.

Sau khi ghi, mọi ô công thức bỏ giá trị đã lưu và file bật tính lại khi mở. Không số nào của bộ tính khác được ghi vào file,
nên Protected View thấy ô trống chứ không thấy số cũ. Số lưu trong biểu đồ cũng bị bỏ vì lý do đó (Q45, Q48d).
"""
from __future__ import annotations

import re

from lxml import etree
from vmai.security import check_formula

from .xlsxpatch import C, CT, PKG, Q, PatchError, XlsxPatcher

CUSTOM_PART = "docProps/custom.xml"
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
UUID_NAME = "vm_document_uuid"
_ADDR = re.compile(r"[A-Z]{1,3}[1-9][0-9]{0,6}")


class WriteRefused(Exception):
    """Ô không được phép ghi. `cells` nêu từng ô và lý do, để app báo đúng chỗ."""

    def __init__(self, cells: list[dict]) -> None:
        super().__init__("write refused")
        self.cells = cells


# Hàm đời mới phải có tiền tố trong file, nếu không Excel báo #NAME? cho tới khi gõ lại công thức.
_FUTURE = re.compile(r'(?<![\w.])(STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|PERCENTILE\.INC|PERCENTILE\.EXC|QUARTILE\.INC|QUARTILE\.EXC|RANK\.EQ|RANK\.AVG|MODE\.SNGL|IFS|SWITCH|MAXIFS|MINIFS|CONCAT|TEXTJOIN|XLOOKUP|IFNA|DAYS|ISFORMULA|FORMULATEXT)(?=\()', re.I)


def _with_future_prefix(formula: str) -> str:
    parts = re.split(r'("(?:[^"]|"")*")', formula.lstrip("="))
    return "".join(p if p.startswith('"') else _FUTURE.sub(lambda m: "_xlfn." + m.group(1).upper(), p) for p in parts)


def _locked_styles(p: XlsxPatcher) -> list[bool]:
    """Theo chỉ số kiểu ô: ô có bị khóa khi sheet bật bảo vệ không. Mặc định của Excel là có."""
    root = p._tree("xl/styles.xml")
    xfs = root.find(Q("cellXfs"))
    out = []
    for xf in xfs if xfs is not None else []:
        prot = xf.find(Q("protection"))
        out.append(not (prot is not None and prot.get("locked") in ("0", "false")))
    return out or [True]


def _refusal(p: XlsxPatcher, edit: dict, locked: list[bool]) -> str | None:
    sheet, cell = edit.get("sheet"), edit.get("cell")
    if sheet not in p.sheets:
        return "Không có sheet này trong file."
    if not isinstance(cell, str) or not _ADDR.fullmatch(cell):
        return "Địa chỉ ô không hợp lệ."
    root = p._tree(p.sheets[sheet])
    protection = root.find(Q("sheetProtection"))
    if protection is not None and protection.get("sheet") in ("1", "true"):
        el = next((c for c in root.iter(Q("c")) if c.get("r") == cell), None)
        style = int(el.get("s") or 0) if el is not None else 0
        if locked[style] if style < len(locked) else True:
            return "Ô đang bị khóa bảo vệ trong file."
    formula = edit.get("formula")
    if formula is not None:
        check = check_formula("=" + formula.lstrip("="))
        if not check.allowed and check.block_on_open:
            return "Công thức này gọi ra ngoài file nên app không ghi."
    return None


def _drop_cached_values(p: XlsxPatcher) -> int:
    count = 0
    for part in p.sheets.values():
        root = p._touch(part)
        for c in root.iter(Q("c")):
            if c.find(Q("f")) is None:
                continue
            v = c.find(Q("v"))
            if v is not None:
                c.remove(v)
            c.attrib.pop("t", None)
            count += 1
    p._force_recalc()
    return count


def _drop_chart_caches(p: XlsxPatcher) -> int:
    count = 0
    for name in p.order:
        if not re.fullmatch(r"xl/charts/chart\d+\.xml", name):
            continue
        root = p._touch(name)
        for tag in ("numCache", "strCache", "multiLvlStrCache"):
            for el in list(root.iter(f"{{{C}}}{tag}")):
                el.getparent().remove(el)
                count += 1
    return count


def _embed_uuid(p: XlsxPatcher, document_uuid: str) -> None:
    """Chỉ nhúng mã định danh, không có đường dẫn hay nội dung nguồn."""
    if CUSTOM_PART in p.parts:
        root = p._touch(CUSTOM_PART)
        for prop in root:
            if prop.get("name") == UUID_NAME:
                prop[0].text = document_uuid
                return
        pid = max((int(x.get("pid")) for x in root), default=1) + 1
    else:
        root = etree.Element(f"{{{CUSTOM_NS}}}Properties", nsmap={None: CUSTOM_NS, "vt": VT_NS})
        p.parts[CUSTOM_PART] = b""
        p.trees[CUSTOM_PART] = root
        p.order.append(CUSTOM_PART)
        p.dirty.add(CUSTOM_PART)
        pid = 2
        ct = p._touch("[Content_Types].xml")
        etree.SubElement(ct, f"{{{CT}}}Override", PartName=f"/{CUSTOM_PART}", ContentType=CUSTOM_TYPE)
        rels = p._touch("_rels/.rels")
        etree.SubElement(rels, f"{{{PKG}}}Relationship", Id="rIdVmUuid", Type=CUSTOM_REL, Target=CUSTOM_PART)
    prop = etree.SubElement(root, f"{{{CUSTOM_NS}}}property", fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", pid=str(pid), name=UUID_NAME)
    etree.SubElement(prop, f"{{{VT_NS}}}lpwstr").text = document_uuid


def apply(data: bytes, edits: list[dict], document_uuid: str) -> dict:
    """`edits`: [{sheet, cell, value} hoặc {sheet, cell, formula}]. Một ô bị từ chối thì không ghi ô nào."""
    p = XlsxPatcher(data)
    locked = _locked_styles(p)
    refused = [{"sheet": e.get("sheet"), "cell": e.get("cell"), "reason": why} for e in edits if (why := _refusal(p, e, locked))]
    if refused:
        raise WriteRefused(refused)
    for e in edits:
        if e.get("formula") is not None:
            p.set_cell(e["sheet"], e["cell"], formula=_with_future_prefix(e["formula"]))
        else:
            # Chữ bắt đầu bằng = + - @ vẫn là ô chữ: bộ vá chỉ tạo công thức khi được gọi với formula.
            p.set_cell(e["sheet"], e["cell"], value=e.get("value"))
    dropped = _drop_cached_values(p)
    charts = _drop_chart_caches(p)
    _embed_uuid(p, document_uuid)
    return {"data": p.save(), "formulas_cleared": dropped, "chart_caches_cleared": charts, "warnings": p.warnings}


__all__ = ["apply", "WriteRefused", "PatchError"]
