"""Đọc cấu trúc gói xlsx bằng XML: quan hệ giữa các phần, các sheet, màu theme. An toàn với XML độc (cổng đã chặn DTD)."""
from __future__ import annotations

import posixpath
import zipfile

from lxml import etree

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False, recover=False)


def q(tag: str, ns: str = MAIN) -> str:
    return f"{{{ns}}}{tag}"


class Package:
    def __init__(self, zf: zipfile.ZipFile) -> None:
        self.zf = zf
        self.names = set(zf.namelist())
        self.workbook = self.xml("xl/workbook.xml")
        wb_rels = self.rels("xl/workbook.xml")
        self.sheets: list[dict] = []
        for el in self.workbook.iter(q("sheet")):
            rel = wb_rels.get(el.get(q("id", R)))
            self.sheets.append({"name": el.get("name"), "state": el.get("state") or "visible", "part": rel["target"] if rel else None})

    def xml(self, part: str):
        return etree.fromstring(self.zf.read(part), PARSER)

    def rels(self, part: str) -> dict[str, dict]:
        """Quan hệ của một phần. Quan hệ External giữ nguyên chuỗi đích, không bao giờ được mở."""
        rel_part = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rel_part not in self.names:
            return {}
        out = {}
        for el in self.xml(rel_part):
            external = el.get("TargetMode") == "External"
            target = el.get("Target") or ""
            if not external:
                target = target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
            out[el.get("Id")] = {"type": (el.get("Type") or "").rsplit("/", 1)[-1], "target": target, "external": external}
        return out

    def external_relations(self) -> list[dict]:
        """Mọi quan hệ trỏ ra ngoài file, ở bất kỳ phần nào: liên kết ngoài, ảnh từ xa, siêu liên kết."""
        out = []
        for name in sorted(self.names):
            if not name.endswith(".rels"):
                continue
            for el in self.xml(name):
                if el.get("TargetMode") == "External":
                    out.append({"part": name, "type": (el.get("Type") or "").rsplit("/", 1)[-1], "target": el.get("Target") or ""})
        return out


class Colors:
    """Đổi màu Excel (rgb, theme + tint, indexed) sang #RRGGBB."""

    # Thứ tự clrScheme trong theme khác thứ tự chỉ số theme của Excel: dk1/lt1 và dk2/lt2 bị đổi chỗ.
    _ORDER = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
    _DEFAULT_THEME = {"lt1": "FFFFFF", "dk1": "000000", "lt2": "E7E6E6", "dk2": "44546A", "accent1": "4472C4", "accent2": "ED7D31",
                      "accent3": "A5A5A5", "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47", "hlink": "0563C1", "folHlink": "954F72"}

    def __init__(self, pkg: Package) -> None:
        theme = dict(self._DEFAULT_THEME)
        for name in pkg.names:
            if name.startswith("xl/theme/theme") and name.endswith(".xml"):
                root = pkg.xml(name)
                scheme = root.find(f".//{q('clrScheme', A)}")
                if scheme is not None:
                    for el in scheme:
                        key = etree.QName(el).localname
                        child = el[0] if len(el) else None
                        if child is not None and key in theme:
                            # sysClr có val là tên màu hệ thống ("windowText"), mã màu nằm ở lastClr.
                            theme[key] = child.get("lastClr") or child.get("val") or theme[key]
                break
        self.theme = [theme[k].upper() for k in self._ORDER]

    @staticmethod
    def _tint(rgb: str, tint: float) -> str:
        import colorsys

        r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
        r, g, b = colorsys.hls_to_rgb(h, min(max(l, 0), 1), s)
        return "".join(f"{round(x * 255):02X}" for x in (r, g, b))

    def resolve(self, color) -> str | None:
        """Nhận đối tượng màu của openpyxl."""
        if color is None:
            return None
        kind = color.type
        if kind == "rgb" and isinstance(color.rgb, str):
            return "#" + color.rgb[-6:].upper()
        if kind == "theme" and color.theme is not None and 0 <= color.theme < len(self.theme):
            rgb = self.theme[color.theme]
            return "#" + (self._tint(rgb, color.tint) if color.tint else rgb)
        if kind == "indexed":
            from openpyxl.styles.colors import COLOR_INDEX

            idx = color.indexed
            if idx is not None and 0 <= idx < len(COLOR_INDEX) and idx < 64:
                return "#" + COLOR_INDEX[idx][-6:].upper()
        return None
