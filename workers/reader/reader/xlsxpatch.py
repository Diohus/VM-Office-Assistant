"""Vá trực tiếp XML của file xlsx: sửa ô, thêm dòng, thêm cột mà không dựng lại cả file.

Phần nào không bị đụng tới thì giữ nguyên từng byte.
Những gì không xử lý được sẽ được đẩy vào `warnings` hoặc báo lỗi, không làm lặng lẽ.
"""
from __future__ import annotations

import io
import posixpath
import re
import zipfile

from lxml import etree
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string as col2n
from openpyxl.utils import get_column_letter as n2col

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
X = "urn:schemas-microsoft-com:office:excel"
Q = lambda t: f"{{{NS}}}{t}"

MAX_ROW, MAX_COL = 1048576, 16384
MAX_UNCOMPRESSED = 200 * 1024 * 1024
_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False, remove_blank_text=False)
_CELL = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")
_COLREF = re.compile(r"^(\$?)([A-Z]{1,3})$")
_ROWREF = re.compile(r"^(\$?)(\d+)$")
_SPLIT_SHEET = re.compile(r"^(?P<sheet>'(?:[^']|'')+'|[^'!]+)!(?P<ref>.+)$")


class PatchError(Exception):
    pass


def _shift_one(idx, at, n):
    return idx + n if idx >= at else idx


def shift_ref(ref, axis, at, n):
    """Dời một địa chỉ không có tên sheet (A1, $A$1, A1:B2, A:B, 1:3). Trả None nếu không phải địa chỉ."""
    parts = ref.split(":")
    if len(parts) > 2:
        return None
    out = []
    for p in parts:
        m = _CELL.match(p)
        if m:
            cd, col, rd, row = m.groups()
            if axis == "row":
                out.append(f"{cd}{col}{rd}{_shift_one(int(row), at, n)}")
            else:
                out.append(f"{cd}{n2col(_shift_one(col2n(col), at, n))}{rd}{row}")
            continue
        m = _COLREF.match(p)
        if m and len(parts) == 2:
            d, col = m.groups()
            out.append(f"{d}{col}" if axis == "row" else f"{d}{n2col(_shift_one(col2n(col), at, n))}")
            continue
        m = _ROWREF.match(p)
        if m and len(parts) == 2:
            d, row = m.groups()
            out.append(f"{d}{_shift_one(int(row), at, n)}" if axis == "row" else f"{d}{row}")
            continue
        return None
    return ":".join(out)


def _shift_sqref(text, axis, at, n):
    return " ".join(shift_ref(p, axis, at, n) or p for p in text.split())


def _shift_sqref_inherit(text, axis, at, n):
    """Như _shift_sqref, nhưng vùng kết thúc ngay trước chỗ chèn được nới ra.

    Dòng/cột mới nhận định dạng của dòng/cột liền trước, kể cả định dạng có điều kiện và data validation.
    """
    out = []
    for part in text.split():
        shifted = shift_ref(part, axis, at, n) or part
        last = shifted.split(":")[-1]
        m = _CELL.match(last)
        if m:
            idx = int(m.group(4)) if axis == "row" else col2n(m.group(2))
            if idx == at - 1:
                new_last = f"{m.group(2)}{idx + n}" if axis == "row" else f"{n2col(idx + n)}{m.group(4)}"
                shifted = f"{shifted.split(':')[0]}:{new_last}"
        out.append(shifted)
    return " ".join(out)


def _sheet_name(prefix):
    if prefix.startswith("'"):
        return prefix[1:-1].replace("''", "'")
    return prefix


def shift_formula(formula, host_sheet, target_sheet, axis, at, n, warnings=None):
    """Dời địa chỉ trỏ tới target_sheet trong công thức. formula không có dấu = đầu."""
    tok = Tokenizer("=" + formula)
    out = []
    for t in tok.items:
        v = t.value
        if t.type == Token.FUNC and warnings is not None and v.upper() in ("INDIRECT(", "OFFSET(", "ADDRESS("):
            warnings.append(f"{v[:-1]} không dời tự động được: {formula}")
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            m = _SPLIT_SHEET.match(v)
            if m:
                sheet = m.group("sheet")
                if sheet.startswith("[") or sheet.startswith("'["):
                    pass                                      # liên kết ngoài, không đụng
                elif ":" in sheet:
                    if warnings is not None and target_sheet in _sheet_name(sheet).split(":"):
                        warnings.append(f"Tham chiếu 3 chiều chưa hỗ trợ dời: {formula}")
                elif _sheet_name(sheet) == target_sheet:
                    shifted = shift_ref(m.group("ref"), axis, at, n)
                    if shifted:
                        v = f"{sheet}!{shifted}"
            elif host_sheet == target_sheet:
                shifted = shift_ref(v, axis, at, n)
                if shifted:
                    v = shifted
        out.append(v)
    return "".join(out)


class XlsxPatcher:
    def __init__(self, data: bytes):
        self.warnings: list[str] = []
        self.zin = zipfile.ZipFile(io.BytesIO(data))
        total = sum(i.file_size for i in self.zin.infolist())
        if total > MAX_UNCOMPRESSED:
            raise PatchError(f"File giải nén quá lớn: {total} byte")
        self.order = [i.filename for i in self.zin.infolist()]
        self.parts = {n: self.zin.read(n) for n in self.order}
        self.trees: dict[str, etree._Element] = {}
        self.dirty: set[str] = set()
        self.removed: set[str] = set()
        wb = self._tree("xl/workbook.xml")
        rels = self._rels("xl/workbook.xml")
        self.sheets = {}
        for s in wb.find(Q("sheets")):
            target = rels[s.get(f"{{{RNS}}}id")]
            self.sheets[s.get("name")] = target
        if "xl/pivotCache/pivotCacheDefinition1.xml" in self.parts or any("pivot" in n for n in self.order):
            self.warnings.append("File có pivot: chưa hỗ trợ dời nguồn của pivot")
        if any(n.endswith("vbaProject.bin") for n in self.order):
            self.warnings.append("File có macro: giữ nguyên, không chạy, không dời tham chiếu trong macro")
        self._expanded: set[str] = set()
        self._pending_fill: list[tuple[str, str, str]] = []

    # ---- tiện ích gói ----
    def _tree(self, name):
        if name not in self.trees:
            self.trees[name] = etree.fromstring(self.parts[name], _PARSER)
        return self.trees[name]

    def _touch(self, name):
        self.dirty.add(name)
        return self._tree(name)

    def _rels(self, part):
        d, f = posixpath.split(part)
        rn = f"{d}/_rels/{f}.rels"
        if rn not in self.parts:
            return {}
        root = etree.fromstring(self.parts[rn], _PARSER)
        return {r.get("Id"): posixpath.normpath(posixpath.join(d, r.get("Target"))) if not r.get("Target").startswith("/")
                else r.get("Target").lstrip("/") for r in root if r.get("TargetMode") != "External"}

    def _sheet_rel_parts(self, sheet, kind):
        return [t for t in self._rels(self.sheets[sheet]).values() if kind in t]

    def save(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name in self.order:
                if name in self.removed:
                    continue
                if name in self.dirty:
                    data = etree.tostring(self.trees[name], xml_declaration=True, encoding="UTF-8", standalone=True)
                else:
                    data = self.parts[name]
                z.writestr(name, data)
        return buf.getvalue()

    def _drop_calc_chain(self):
        """Excel tự dựng lại calcChain; để sót mục sai thì mới báo hỏng file."""
        name = "xl/calcChain.xml"
        if name in self.parts and name not in self.removed:
            self.removed.add(name)
            ct = self._touch("[Content_Types].xml")
            for e in list(ct):
                if e.get("PartName") == "/xl/calcChain.xml":
                    ct.remove(e)
            rn = "xl/_rels/workbook.xml.rels"
            rels = self._touch(rn)
            for e in list(rels):
                if e.get("Target", "").endswith("calcChain.xml"):
                    rels.remove(e)

    def _force_recalc(self):
        wb = self._touch("xl/workbook.xml")
        calc = wb.find(Q("calcPr"))
        if calc is None:
            # Thứ tự phần tử trong workbook.xml là bắt buộc: calcPr đứng ngay sau definedNames, hoặc các mục trước nó.
            calc = etree.Element(Q("calcPr"))
            anchor = next((wb.find(Q(t)) for t in ("definedNames", "externalReferences", "functionGroups", "sheets") if wb.find(Q(t)) is not None), None)
            anchor.addnext(calc)
        calc.set("fullCalcOnLoad", "1")

    # ---- công thức dùng chung ----
    def _expand_shared(self, sheet):
        if sheet in self._expanded:
            return
        self._expanded.add(sheet)
        root = self._tree(self.sheets[sheet])
        masters = {}
        cells = list(root.iter(Q("f")))
        for f in cells:
            if f.get("t") == "shared" and f.text:
                masters[f.get("si")] = (f.getparent().get("r"), f.text)
        if not any(f.get("t") == "shared" for f in cells):
            return
        self._touch(self.sheets[sheet])
        for f in cells:
            if f.get("t") == "shared":
                origin, text = masters[f.get("si")]
                dest = f.getparent().get("r")
                f.text = Translator("=" + text, origin=origin).translate_formula(dest)[1:]
                for a in ("t", "si", "ref"):
                    f.attrib.pop(a, None)

    # ---- sửa ô ----
    def _row(self, root, r, create=True):
        data = root.find(Q("sheetData"))
        after = None
        for row in data:
            rr = int(row.get("r"))
            if rr == r:
                return row
            if rr > r:
                after = row
                break
        if not create:
            return None
        new = etree.Element(Q("row"), r=str(r))
        if after is not None:
            after.addprevious(new)
        else:
            data.append(new)
        return new

    def _cell(self, row, addr, create=True):
        col = col2n(re.match(r"[A-Z]+", addr).group())
        after = None
        for c in row:
            cc = col2n(re.match(r"[A-Z]+", c.get("r")).group())
            if cc == col:
                return c
            if cc > col:
                after = c
                break
        if not create:
            return None
        row.attrib.pop("spans", None)
        new = etree.Element(Q("c"), r=addr)
        if after is not None:
            after.addprevious(new)
        else:
            row.append(new)
        return new

    def set_cell(self, sheet, addr, value=None, formula=None):
        """Ghi giá trị hoặc công thức, giữ nguyên kiểu định dạng (thuộc tính s) của ô."""
        self._expand_shared(sheet)
        root = self._touch(self.sheets[sheet])
        r = int(re.search(r"\d+", addr).group())
        c = self._cell(self._row(root, r), addr)
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop("t", None)
        if formula is not None:
            etree.SubElement(c, Q("f")).text = formula.lstrip("=")
            self._drop_calc_chain()
        elif isinstance(value, bool):
            c.set("t", "b")
            etree.SubElement(c, Q("v")).text = "1" if value else "0"
        elif isinstance(value, (int, float)):
            etree.SubElement(c, Q("v")).text = repr(value)
        elif value is not None:
            c.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(c, Q("is")), Q("t"))
            t.text = str(value)
            if t.text != t.text.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        self._force_recalc()

    # ---- chèn dòng, cột ----
    def insert_rows(self, sheet, at, n=1):
        self._insert(sheet, "row", at, n)

    def insert_cols(self, sheet, at, n=1):
        self._insert(sheet, "col", at, n)

    def _insert(self, sheet, axis, at, n):
        for s in self.sheets:
            self._expand_shared(s)
        self._drop_calc_chain()
        self._force_recalc()
        root = self._touch(self.sheets[sheet])
        if axis == "row":
            self._move_rows(root, at, n)
        else:
            self._move_cols(root, at, n)
        self._fix_sheet_ranges(root, sheet, axis, at, n)
        self._fix_drawings(sheet, axis, at, n)
        self._fix_comments(sheet, axis, at, n)
        self._fix_tables(sheet, axis, at, n)
        for s, part in self.sheets.items():
            self._fix_formulas(self._tree(part), s, sheet, axis, at, n, part)
        self._fix_defined_names(sheet, axis, at, n)
        self._fix_charts(sheet, axis, at, n)
        self._fix_table_formulas(sheet, axis, at, n)
        for s, addr, text in self._pending_fill:
            self.set_cell(s, addr, formula=text)
        self._pending_fill.clear()

    def _move_rows(self, root, at, n):
        data = root.find(Q("sheetData"))
        template = None
        for row in data:
            r = int(row.get("r"))
            if r == at - 1:
                template = row
            if r >= at:
                row.set("r", str(r + n))
                for c in row:
                    c.set("r", re.sub(r"\d+", str(r + n), c.get("r")))
        # Dòng mới nhận định dạng của dòng ngay trên, giống hành vi mặc định của Excel.
        if template is not None:
            for k in range(n):
                new = etree.Element(Q("row"), {a: v for a, v in template.attrib.items() if a != "spans"})
                new.set("r", str(at + k))
                for c in template:
                    if c.get("s"):
                        col = re.match(r"[A-Z]+", c.get("r")).group()
                        etree.SubElement(new, Q("c"), r=f"{col}{at + k}", s=c.get("s"))
                template.addnext(new) if k == 0 else data[data.index(template) + k].addnext(new)

    def _move_cols(self, root, at, n):
        for row in root.find(Q("sheetData")):
            row.attrib.pop("spans", None)
            cells = list(row)
            for c in cells:
                m = re.match(r"([A-Z]+)(\d+)", c.get("r"))
                idx = col2n(m.group(1))
                if idx >= at:
                    c.set("r", f"{n2col(idx + n)}{m.group(2)}")
            left = next((c for c in cells if col2n(re.match(r"[A-Z]+", c.get("r")).group()) == at - 1), None)
            if left is not None and left.get("s"):
                r = re.search(r"\d+", left.get("r")).group()
                prev = left
                for k in range(n):
                    new = etree.Element(Q("c"), r=f"{n2col(at + k)}{r}", s=left.get("s"))
                    prev.addnext(new)
                    prev = new
        cols = root.find(Q("cols"))
        if cols is None:
            return
        inherit = None
        for col in list(cols):
            lo, hi = int(col.get("min")), int(col.get("max"))
            if lo >= at:
                col.set("min", str(lo + n))
                col.set("max", str(min(hi + n, MAX_COL)))
            elif hi >= at:
                col.set("max", str(min(hi + n, MAX_COL)))
            elif hi == at - 1:
                inherit = col
        if inherit is not None:
            # Cột mới lấy độ rộng của cột bên trái.
            new = etree.Element(Q("col"), dict(inherit.attrib))
            new.set("min", str(at))
            new.set("max", str(at + n - 1))
            inherit.addnext(new)

    def _fix_sheet_ranges(self, root, sheet, axis, at, n):
        dim = root.find(Q("dimension"))
        if dim is not None:
            first, _, last = dim.get("ref").partition(":")
            last = last or first
            m = _CELL.match(last)
            if axis == "row":
                last = f"{m.group(2)}{int(m.group(4)) + n}"
            else:
                last = f"{n2col(col2n(m.group(2)) + n)}{m.group(4)}"
            dim.set("ref", f"{first}:{last}")
        for e in root.iter(Q("mergeCell"), Q("hyperlink"), Q("autoFilter"), Q("ignoredError"), Q("protectedRange")):
            attr = "sqref" if e.get("sqref") else "ref"
            e.set(attr, _shift_sqref(e.get(attr), axis, at, n))
        for e in root.iter(Q("conditionalFormatting"), Q("dataValidation")):
            if e.get("sqref"):
                e.set("sqref", _shift_sqref_inherit(e.get("sqref"), axis, at, n))
        for e in root.iter(Q("selection")):
            if e.get("sqref"):
                e.set("sqref", _shift_sqref(e.get("sqref"), axis, at, n))
            if e.get("activeCell"):
                e.set("activeCell", shift_ref(e.get("activeCell"), axis, at, n) or e.get("activeCell"))
        for e in root.iter(f"{{{XM}}}sqref"):
            e.text = _shift_sqref_inherit(e.text, axis, at, n)
        brks = root.find(Q("rowBreaks" if axis == "row" else "colBreaks"))
        if brks is not None:
            for b in brks:
                if int(b.get("id")) >= at - 1:
                    b.set("id", str(int(b.get("id")) + n))
        pane = next(root.iter(Q("pane")), None)
        if pane is not None:
            attr = "ySplit" if axis == "row" else "xSplit"
            if pane.get(attr) and at <= int(float(pane.get(attr))):
                pane.set(attr, str(int(float(pane.get(attr))) + n))
            if pane.get("topLeftCell"):
                pane.set("topLeftCell", shift_ref(pane.get("topLeftCell"), axis, at, n) or pane.get("topLeftCell"))

    def _fix_formulas(self, root, host, target, axis, at, n, part):
        touched = False

        def fix(text):
            nonlocal touched
            new = shift_formula(text, host, target, axis, at, n, self.warnings)
            touched = touched or new != text
            return new

        for f in root.iter(Q("f")):
            if f.text:
                f.text = fix(f.text)
            if f.get("t") == "array" and f.get("ref") and host == target:
                f.set("ref", shift_ref(f.get("ref"), axis, at, n))
                touched = True
            if f.get("t") == "dataTable":
                self.warnings.append("Data table chưa hỗ trợ dời")
        for e in root.iter(Q("formula"), Q("formula1"), Q("formula2"), f"{{{XM}}}f"):
            if e.text:
                e.text = fix(e.text)
        if touched:
            self._touch(part)

    def _fix_defined_names(self, target, axis, at, n):
        wb = self._tree("xl/workbook.xml")
        names = wb.find(Q("definedNames"))
        if names is None:
            return
        for d in names:
            new = shift_formula(d.text, None, target, axis, at, n, self.warnings)
            if new != d.text:
                d.text = new
                self._touch("xl/workbook.xml")

    def _fix_drawings(self, sheet, axis, at, n):
        idx_tag = "row" if axis == "row" else "col"
        for part in self._sheet_rel_parts(sheet, "drawings/drawing"):
            root = self._touch(part)
            for anchor in list(root):
                start = anchor.find(f"{{{XDR}}}from")
                end = anchor.find(f"{{{XDR}}}to")
                if start is None:
                    continue
                s = start.find(f"{{{XDR}}}{idx_tag}")
                moved = int(s.text) + 1 >= at
                if moved:
                    s.text = str(int(s.text) + n)
                # editAs="oneCell"/"absolute": ảnh không co giãn, nên mép dưới chỉ đi theo khi cả ảnh bị dời.
                if end is not None and (moved or anchor.get("editAs", "twoCell") == "twoCell"):
                    t = end.find(f"{{{XDR}}}{idx_tag}")
                    if int(t.text) + 1 >= at:
                        t.text = str(int(t.text) + n)

    def _fix_comments(self, sheet, axis, at, n):
        for part in self._sheet_rel_parts(sheet, "comments"):
            root = self._touch(part)
            for c in root.iter(Q("comment")):
                c.set("ref", shift_ref(c.get("ref"), axis, at, n))
        for part in self._sheet_rel_parts(sheet, "vmlDrawing"):
            root = self._touch(part)
            tag, a_i = ("Row", (2, 6)) if axis == "row" else ("Column", (0, 4))
            for e in root.iter(f"{{{X}}}{tag}"):
                if int(e.text) + 1 >= at:
                    e.text = str(int(e.text) + n)
            for e in root.iter(f"{{{X}}}Anchor"):
                nums = [int(x) for x in e.text.split(",")]
                for i in a_i:
                    if nums[i] + 1 >= at:
                        nums[i] += n
                e.text = ", ".join(map(str, nums))

    def _fix_tables(self, sheet, axis, at, n):
        for part in self._sheet_rel_parts(sheet, "tables/table"):
            root = self._touch(part)
            old = root.get("ref")
            root.set("ref", shift_ref(old, axis, at, n))
            m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", old)
            if axis == "row" and int(m.group(2)) < at <= int(m.group(4)):
                # Dòng mới trong Table: Excel tự chép công thức của các cột tính, phải làm giống vậy.
                first_col = col2n(m.group(1))
                for i, tc in enumerate(root.iter(Q("tableColumn"))):
                    calc = tc.find(Q("calculatedColumnFormula"))
                    if calc is not None and calc.text:
                        self._pending_fill.append((sheet, f"{n2col(first_col + i)}{at}", calc.text))
            for af in root.iter(Q("autoFilter")):
                af.set("ref", shift_ref(af.get("ref"), axis, at, n))
            if axis == "col":
                self.warnings.append("Chèn cột vào bảng (table): chưa thêm tableColumn, phải kiểm tra")

    def _fix_table_formulas(self, target, axis, at, n):
        for name in self.order:
            if re.match(r"xl/tables/table\d*\.xml$", name):
                root = self._tree(name)
                host = next((s for s in self.sheets if name in self._sheet_rel_parts(s, "tables/table")), None)
                for e in root.iter(Q("calculatedColumnFormula"), Q("totalsRowFormula")):
                    if e.text:
                        new = shift_formula(e.text, host, target, axis, at, n, self.warnings)
                        if new != e.text:
                            e.text = new
                            self._touch(name)

    def _fix_charts(self, target, axis, at, n):
        for name in self.order:
            if re.match(r"xl/charts/chart\d*\.xml$", name):
                root = self._tree(name)
                for f in root.iter(f"{{{C}}}f"):
                    new = shift_formula(f.text, None, target, axis, at, n, self.warnings)
                    if new != f.text:
                        f.text = new
                        self._touch(name)
