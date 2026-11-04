"""Dịch xlsx sang dữ liệu lưới trung lập, không phụ thuộc thư viện lưới nào (Q20). Chỉ để xem, không bao giờ dùng để tạo file.

Công thức gọi ra ngoài không được đưa vào: ô đó chỉ có giá trị Excel đã lưu.
"""
from __future__ import annotations

import base64
import datetime as dt
import re
from types import SimpleNamespace

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel

from .formula_policy import blocked_reason
from .ooxml import A, R, X14, XDR, XM, Colors, Package, q

SCHEMA = 1
MIN_ROWS, MIN_COLS = 60, 16
_CF_TRANSLATED = {"expression", "cellIs"}
_CF_SHOWN_ELSEWHERE = {"dataBar", "colorScale", "iconSet"}


def _plain(value):
    if isinstance(value, dt.datetime):
        return to_excel(value)
    if isinstance(value, dt.date):
        return to_excel(dt.datetime(value.year, value.month, value.day))
    if isinstance(value, dt.time):
        return to_excel(value)
    if isinstance(value, dt.timedelta):
        return value.total_seconds() / 86400
    return value


class _Styles:
    """Bảng kiểu dùng chung, để biểu mẫu nhiều ô không lặp cùng một kiểu hàng chục nghìn lần."""

    def __init__(self, colors: Colors) -> None:
        self.colors = colors
        self.table: list[dict] = []
        self.index: dict[str, int] = {}

    def of(self, cell) -> int | None:
        st = self._style(cell)
        if not st:
            return None
        sig = repr(sorted(st.items(), key=lambda kv: kv[0]))
        if sig not in self.index:
            self.index[sig] = len(self.table)
            self.table.append(st)
        return self.index[sig]

    def _style(self, cell) -> dict:
        st: dict = {}
        f = cell.font
        if f.name:
            st["ff"] = f.name
        if f.sz:
            st["fs"] = f.sz
        if f.b:
            st["bl"] = 1
        if f.i:
            st["it"] = 1
        if f.u:
            st["ul"] = 1
        if f.strike:
            st["st"] = 1
        color = self.colors.resolve(f.color)
        if color:
            st["fc"] = color
        if cell.fill.fill_type == "solid":
            bg = self.colors.resolve(cell.fill.fgColor)
            if bg:
                st["bg"] = bg
        al = cell.alignment
        if al.horizontal:
            st["ha"] = al.horizontal
        if al.vertical:
            st["va"] = al.vertical
        if al.wrap_text:
            st["wrap"] = 1
        if cell.number_format and cell.number_format != "General":
            st["nf"] = cell.number_format
        borders = {}
        for side, key in ((cell.border.top, "t"), (cell.border.bottom, "b"), (cell.border.left, "l"), (cell.border.right, "r")):
            if side is not None and side.style:
                borders[key] = [side.style, self.colors.resolve(side.color) or "#000000"]
        if borders:
            st["bd"] = borders
        return st


class _Formulas:
    """Đổi công thức về dạng A1 để lưới hiểu: bỏ tiền tố _xlfn., đổi tên vùng và tham chiếu Table."""

    def __init__(self, wb) -> None:
        self.names = {k: v.attr_text for k, v in wb.defined_names.items()}
        self.sheet_names = {ws.title: {k: v.attr_text for k, v in ws.defined_names.items()} for ws in wb.worksheets}
        self.tables: dict[str, dict] = {}
        for ws in wb.worksheets:
            for t in ws.tables.values():
                lo, hi = t.ref.split(":")
                first_col = column_index_from_string(re.match(r"[A-Z]+", lo).group())
                self.tables[t.name] = {
                    "first": int(re.search(r"\d+", lo).group()) + (t.headerRowCount if t.headerRowCount is not None else 1),
                    "last": int(re.search(r"\d+", hi).group()) - (t.totalsRowCount or 0),
                    "sheet": ws.title,
                    "cols": {tc.name: get_column_letter(first_col + i) for i, tc in enumerate(t.tableColumns)},
                }

    def _operand(self, value: str, sheet: str, row: int) -> str:
        scoped = self.sheet_names.get(sheet, {})
        if value in scoped:
            return scoped[value]
        if value in self.names:
            return self.names[value]
        m = re.fullmatch(r"(\w+)\[\[#This Row\],\[([^\]]+)\]\]", value)
        if m and m.group(1) in self.tables and m.group(2) in self.tables[m.group(1)]["cols"]:
            return f"{self.tables[m.group(1)]['cols'][m.group(2)]}{row}"
        m = re.fullmatch(r"(\w+)\[([^\[\]#]+)\]", value)
        if m and m.group(1) in self.tables and m.group(2) in self.tables[m.group(1)]["cols"]:
            t = self.tables[m.group(1)]
            col = t["cols"][m.group(2)]
            return f"${col}${t['first']}:${col}${t['last']}"
        return value

    def to_a1(self, text: str, sheet: str, row: int) -> str:
        tok = Tokenizer(text)
        parts = ["="]
        for t in tok.items:
            value = t.value
            if t.type == "OPERAND" and t.subtype == "RANGE":
                value = self._operand(value, sheet, row)
            elif t.type == "FUNC" and t.subtype == "OPEN":
                value = value.replace("_xlfn._xlws.", "").replace("_xlfn.", "").replace("_xlws.", "")
            parts.append(value)
        return "".join(parts)


def _color_el(colors: Colors, el) -> str | None:
    if el is None:
        return None
    theme, indexed, tint = el.get("theme"), el.get("indexed"), float(el.get("tint") or 0)
    if el.get("rgb"):
        ns = SimpleNamespace(type="rgb", rgb=el.get("rgb"), theme=None, tint=0, indexed=None)
    elif theme is not None:
        ns = SimpleNamespace(type="theme", rgb=None, theme=int(theme), tint=tint, indexed=None)
    elif indexed is not None:
        ns = SimpleNamespace(type="indexed", rgb=None, theme=None, tint=0, indexed=int(indexed))
    else:
        return None
    return colors.resolve(ns)


def _dxf_style(colors: Colors, dxf) -> dict:
    out: dict = {}
    if dxf is None:
        return out
    font = dxf.find(q("font"))
    if font is not None:
        c = _color_el(colors, font.find(q("color")))
        if c:
            out["fc"] = c
        if font.find(q("b")) is not None:
            out["bl"] = 1
        if font.find(q("i")) is not None:
            out["it"] = 1
    fill = dxf.find(q("fill"))
    if fill is not None:
        pattern = fill.find(q("patternFill"))
        if pattern is not None:
            # Trong dxf, màu nền nằm ở bgColor, khác với kiểu ô thường.
            c = _color_el(colors, pattern.find(q("bgColor"))) or _color_el(colors, pattern.find(q("fgColor")))
            if c:
                out["bg"] = c
    return out


def _sheet_rules(pkg: Package, colors: Colors, root, sheet: str, formulas: _Formulas, blocked: list[dict]) -> tuple[list[dict], list[dict], list[dict]]:
    """Định dạng có điều kiện và danh sách chọn, cả phần mở rộng x14 mà openpyxl bỏ qua."""
    dxfs = []
    if "xl/styles.xml" in pkg.names:
        node = pkg.xml("xl/styles.xml").find(q("dxfs"))
        dxfs = list(node) if node is not None else []
    cf, dv, other = [], [], []

    def add_cf(sqref, rule, texts, dxf):
        kind = rule.get("type")
        if kind in _CF_SHOWN_ELSEWHERE:
            return
        if kind not in _CF_TRANSLATED:
            other.append({"kind": "conditional_format", "sheet": sheet, "cell": sqref, "detail": kind})
            return
        for text in texts:
            reason = blocked_reason(text or "")
            if reason:
                blocked.append({"sheet": sheet, "cell": sqref, "where": "conditional_format", **reason})
                return
        cf.append({"sqref": sqref, "type": kind, "operator": rule.get("operator"), "formulas": texts, "style": _dxf_style(colors, dxf),
                   "priority": int(rule.get("priority") or 0), "stop": rule.get("stopIfTrue") in ("1", "true")})

    for el in root.iter(q("conditionalFormatting")):
        for rule in el.iter(q("cfRule")):
            dxf_id = rule.get("dxfId")
            dxf = dxfs[int(dxf_id)] if dxf_id is not None and int(dxf_id) < len(dxfs) else None
            add_cf(el.get("sqref"), rule, [f.text for f in rule.findall(q("formula"))], dxf)
    for el in root.iter(q("conditionalFormatting", X14)):
        sq = el.find(q("sqref", XM))
        for rule in el.iter(q("cfRule", X14)):
            add_cf(sq.text if sq is not None else None, rule, [f.text for f in rule.findall(q("f", XM))], rule.find(q("dxf", X14)))

    def add_dv(sqref, kind, f1):
        if f1 and blocked_reason(f1):
            blocked.append({"sheet": sheet, "cell": sqref, "where": "data_validation", **blocked_reason(f1)})
            return
        dv.append({"sqref": sqref, "type": kind, "f1": f1})

    for el in root.iter(q("dataValidation")):
        f1 = el.find(q("formula1"))
        add_dv(el.get("sqref"), el.get("type"), f1.text if f1 is not None else None)
    for el in root.iter(q("dataValidation", X14)):
        sq = el.find(q("sqref", XM))
        add_dv(sq.text if sq is not None else None, el.get("type"), next((x.text for x in el.iter(q("f", XM))), None))
    return cf, dv, other


def _images(pkg: Package, part: str) -> tuple[list[dict], int]:
    images, skipped = [], 0
    for rel in pkg.rels(part).values():
        if rel["type"] != "drawing" or rel["external"] or rel["target"] not in pkg.names:
            continue
        rels = pkg.rels(rel["target"])
        for anchor in pkg.xml(rel["target"]):
            blip = next(anchor.iter(q("blip", A)), None)
            start = anchor.find(q("from", XDR))
            if blip is None or start is None or anchor.find(q("pic", XDR)) is None:
                continue
            target = rels.get(blip.get(q("embed", R)))
            if target is None or target["external"] or target["target"] not in pkg.names:
                skipped += 1      # ảnh trỏ ra ngoài: không tải, đã có trong danh sách trỏ ra ngoài
                continue
            data = pkg.zf.read(target["target"])
            mime = "image/png" if data[:8] == b"\x89PNG\r\n\x1a\n" else "image/jpeg" if data[:3] == b"\xff\xd8\xff" else None
            if mime is None:
                skipped += 1
                continue
            ext = anchor.find(q("ext", XDR))
            images.append({"row": int(start.find(q("row", XDR)).text), "col": int(start.find(q("col", XDR)).text),
                           "cx": int(ext.get("cx")) if ext is not None else None, "cy": int(ext.get("cy")) if ext is not None else None,
                           "src": f"data:{mime};base64," + base64.b64encode(data).decode()})
    return images, skipped


def translate(path, pkg: Package) -> dict:
    wf = openpyxl.load_workbook(path, data_only=False, keep_vba=False)
    wv = openpyxl.load_workbook(path, data_only=True, keep_vba=False)
    colors = Colors(pkg)
    styles = _Styles(colors)
    formulas = _Formulas(wf)
    parts = {s["name"]: s["part"] for s in pkg.sheets}
    blocked: list[dict] = []
    other_unshown: list[dict] = []
    sheets = []
    for ws in wf.worksheets:
        vs = wv[ws.title]
        cells = []
        for row in ws.iter_rows():
            for c in row:
                val = c.value
                if val is None and not c.has_style:
                    continue
                item: dict = {"r": c.row - 1, "c": c.column - 1}
                sid = styles.of(c) if c.has_style else None
                if sid is not None:
                    item["s"] = sid
                if hasattr(val, "text"):      # công thức mảng
                    val = val.text
                elif val is not None and val.__class__.__name__ == "DataTableFormula":
                    other_unshown.append({"kind": "data_table", "sheet": ws.title, "cell": c.coordinate, "detail": None})
                    item["v"] = _plain(vs[c.coordinate].value)
                    cells.append(item)
                    continue
                if isinstance(val, str) and c.data_type == "f":
                    cached = _plain(vs[c.coordinate].value)
                    try:
                        a1 = formulas.to_a1(val, ws.title, c.row)
                    except Exception:  # noqa: BLE001 - công thức lạ: giữ giá trị đã lưu và báo, không để hỏng cả file
                        a1 = None
                    reason = blocked_reason(val) or (blocked_reason(a1) if a1 else None)
                    if reason:
                        blocked.append({"sheet": ws.title, "cell": c.coordinate, "where": "cell", **reason})
                        item["v"] = cached
                        item["blocked"] = 1
                    elif a1 is None:
                        other_unshown.append({"kind": "formula_unparsed", "sheet": ws.title, "cell": c.coordinate, "detail": None})
                        item["v"] = cached
                    else:
                        item["f"] = a1
                        item["excel_v"] = cached
                elif val is not None:
                    item["v"] = _plain(val)
                cells.append(item)
        root = pkg.xml(parts[ws.title])
        cf, dv, other = _sheet_rules(pkg, colors, root, ws.title, formulas, blocked)
        other_unshown.extend(other)
        images, skipped = _images(pkg, parts[ws.title])
        if skipped:
            other_unshown.append({"kind": "image_other", "sheet": ws.title, "cell": None, "detail": str(skipped)})
        freeze = None
        if ws.freeze_panes:
            fc = ws[ws.freeze_panes]
            freeze = {"rows": fc.row - 1, "cols": fc.column - 1}
        sheets.append({
            "name": ws.title,
            "hidden": ws.sheet_state != "visible",
            "rows": max(ws.max_row, MIN_ROWS),
            "cols": max(ws.max_column, MIN_COLS),
            "cells": cells,
            "merges": [[m.min_row - 1, m.min_col - 1, m.max_row - 1, m.max_col - 1] for m in ws.merged_cells.ranges],
            "col_groups": [[d.min - 1, d.max - 1, d.width, 1 if d.hidden else 0] for d in ws.column_dimensions.values() if d.min and (d.width or d.hidden)],
            "row_dims": {str(k - 1): [d.height, 1 if d.hidden else 0] for k, d in ws.row_dimensions.items() if d.height or d.hidden},
            "freeze": freeze,
            "cf": cf,
            "dv": dv,
            "images": images,
        })
    return {"grid": {"schema": SCHEMA, "styles": styles.table, "sheets": sheets}, "blocked_formulas": blocked, "unshown_extra": other_unshown}
