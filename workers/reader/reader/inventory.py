"""Kiểm kê những gì file có mà lưới chưa hiện được, kèm vị trí, và những gì trỏ ra ngoài file.

Lưới không hiện không có nghĩa là mất: file gốc được giữ nguyên từng byte, các mục này vẫn còn khi xuất.
"""
from __future__ import annotations

from lxml import etree

from .ooxml import A, MAIN, X14, XDR, XM, Package, q

_XDR_PIC, _XDR_SP, _XDR_GRAPHIC, _XDR_CONN, _XDR_GROUP = (q(t, XDR) for t in ("pic", "sp", "graphicFrame", "cxnSp", "grpSp"))
MAX_COMMENT_CHARS = 2000
_CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _col(n: int) -> str:
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _anchor_cell(anchor) -> str | None:
    start = anchor.find(q("from", XDR))
    if start is None:
        return None
    return f"{_col(int(start.find(q('col', XDR)).text))}{int(start.find(q('row', XDR)).text) + 1}"


def _item(kind: str, sheet: str, cell: str | None, detail: str | None = None) -> dict:
    return {"kind": kind, "sheet": sheet, "cell": cell, "detail": detail}


def _comment_text(el) -> str:
    """Nội dung nhận xét để đọc trong app (Q58). Là dữ liệu trong file, chỉ hiện dạng chữ."""
    return "".join(t.text or "" for t in el.iter(q("t")))[:MAX_COMMENT_CHARS]


def _drawing_items(pkg: Package, part: str, sheet: str) -> list[dict]:
    out = []
    for anchor in pkg.xml(part):
        cell = _anchor_cell(anchor)
        if anchor.find(_XDR_GRAPHIC) is not None:
            is_chart = any(el.get("uri") == _CHART_URI for el in anchor.iter(q("graphicData", A)))
            out.append(_item("chart" if is_chart else "object", sheet, cell))
        elif anchor.find(_XDR_SP) is not None or anchor.find(_XDR_CONN) is not None or anchor.find(_XDR_GROUP) is not None:
            out.append(_item("shape", sheet, cell))
        # Ảnh (xdr:pic) lưới hiện được nên không liệt kê ở đây.
    return out


def unshown(pkg: Package) -> list[dict]:
    items: list[dict] = []
    for sheet in pkg.sheets:
        part = sheet["part"]
        if not part or part not in pkg.names:
            continue
        name = sheet["name"]
        rels = pkg.rels(part)
        for rel in rels.values():
            if rel["external"]:
                continue
            kind, target = rel["type"], rel["target"]
            if kind == "drawing" and target in pkg.names:
                items.extend(_drawing_items(pkg, target, name))
            elif kind == "comments" and target in pkg.names:
                items.extend(_item("comment", name, el.get("ref"), _comment_text(el)) for el in pkg.xml(target).iter(q("comment")))
            elif kind == "threadedComment":
                items.append(_item("threaded_comment", name, None))
            elif kind == "pivotTable" and target in pkg.names:
                loc = pkg.xml(target).find(q("location"))
                items.append(_item("pivot", name, loc.get("ref") if loc is not None else None))
            elif kind == "table" and target in pkg.names:
                root = pkg.xml(target)
                items.append(_item("table_style", name, root.get("ref"), root.get("displayName")))
            elif kind in ("control", "oleObject"):
                items.append(_item("object", name, None))
        root = pkg.xml(part)
        for cf in root.iter(q("conditionalFormatting")):
            for rule in cf.iter(q("cfRule")):
                if rule.get("type") in ("dataBar", "colorScale", "iconSet"):
                    items.append(_item({"dataBar": "data_bar", "colorScale": "color_scale", "iconSet": "icon_set"}[rule.get("type")], name, cf.get("sqref")))
        for cf in root.iter(q("conditionalFormatting", X14)):
            sq = cf.find(q("sqref", XM))
            for rule in cf.iter(q("cfRule", X14)):
                if rule.get("type") in ("dataBar", "colorScale", "iconSet"):
                    items.append(_item({"dataBar": "data_bar", "colorScale": "color_scale", "iconSet": "icon_set"}[rule.get("type")], name, sq.text if sq is not None else None))
        for spark in root.iter(q("sparklineGroup", X14)):
            for sq in spark.iter(q("sqref", XM)):
                items.append(_item("sparkline", name, sq.text))
    # Thanh dữ liệu được khai cả ở phần chính lẫn phần x14 của cùng một quy tắc: chỉ báo một lần.
    seen, unique = set(), []
    for it in items:
        key = (it["kind"], it["sheet"], it["cell"], it["detail"])
        if it["cell"] is None or key not in seen:
            seen.add(key)
            unique.append(it)
    return unique


def external_refs(pkg: Package) -> list[dict]:
    """Điểm trỏ ra ngoài ở cấp gói. Đích được giữ dạng chữ để báo, app không bao giờ mở chúng."""
    out = []
    for rel in pkg.external_relations():
        kind = rel["type"]
        if kind in ("externalLinkPath", "externalLinkLong", "xlExternalLinkPath"):
            out.append({"kind": "external_link", "part": rel["part"], "target": rel["target"]})
        elif kind == "hyperlink":
            out.append({"kind": "hyperlink", "part": rel["part"], "target": rel["target"]})
        else:
            out.append({"kind": "external_resource", "part": rel["part"], "target": rel["target"]})
    for name in sorted(pkg.names):
        if name == "xl/connections.xml":
            out.append({"kind": "data_connection", "part": name, "target": None})
        elif name.startswith("xl/queryTables/"):
            out.append({"kind": "query_table", "part": name, "target": None})
        elif name.startswith("customXml/") and name.endswith("item1.xml"):
            # Power Query lưu ở DataMashup: có thể gọi nguồn ngoài khi Excel làm mới.
            if b"DataMashup" in pkg.zf.read(name):
                out.append({"kind": "data_connection", "part": name, "target": None})
    return out


def protected_sheets(pkg: Package) -> list[str]:
    out = []
    for sheet in pkg.sheets:
        if sheet["part"] and sheet["part"] in pkg.names:
            root = pkg.xml(sheet["part"])
            el = root.find(q("sheetProtection"))
            if el is not None and el.get("sheet") in ("1", "true"):
                out.append(sheet["name"])
    return out


def count_cells(pkg: Package) -> int:
    """Đếm ô có dữ liệu bằng cách đọc luồng, không dựng cả cây, để từ chối file quá lớn trước khi dịch."""
    total = 0
    tag_c, tag_v, tag_f, tag_is = q("c"), q("v"), q("f"), q("is")
    for sheet in pkg.sheets:
        part = sheet["part"]
        if not part or part not in pkg.names:
            continue
        with pkg.zf.open(part) as fh:
            for _, el in etree.iterparse(fh, events=("end",), tag=tag_c, resolve_entities=False, no_network=True, huge_tree=False):
                if el.find(tag_v) is not None or el.find(tag_f) is not None or el.find(tag_is) is not None:
                    total += 1
                el.clear()
    return total


__all__ = ["unshown", "external_refs", "protected_sheets", "count_cells", "MAIN"]
