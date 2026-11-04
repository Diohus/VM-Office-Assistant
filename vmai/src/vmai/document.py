"""Đọc tài liệu hãng thành dòng có vị trí. Không mạng, không key, không sửa chữ.

Cần `vmai[pdf]`. Trang quét không tự đọc ở đây: `lines_from_tokens` nhận token từ OCR nội bộ.
"""
from __future__ import annotations

import re
from collections import Counter
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict

from .errors import SourceFileError
from .mapping import normalize_text
from .template import sha256_file

SCAN_IMAGE_COVERAGE = 0.6
REPEAT_PAGE_SHARE = 0.6
REPEAT_MIN_PAGES = 3
EDGE_BAND = 0.1  # dòng lặp chỉ tính khi nằm sát mép trên, mép dưới
TWO_COLUMN_MIN_LINES = 8
_NUMBERED = re.compile(r"\d+")


class DocLine(BaseModel):
    model_config = ConfigDict(frozen=True)

    index: int
    page: int
    text: str
    x0: float
    top: float
    x1: float
    bottom: float
    font_size: float | None = None
    bold: bool = False
    source: Literal["text", "ocr"] = "text"
    uncertain: bool = False  # dòng OCR có chữ số, hoặc thiếu lần đọc thứ hai: màn song song phải tô để so với ảnh (Q70)
    digits_differ: bool = False  # hai OCR đọc chữ số khác nhau: tô đậm hơn, xem trước
    in_table: bool = False
    repeated: bool = False  # đầu trang, chân trang lặp lại


class DocTable(BaseModel):
    model_config = ConfigDict(frozen=True)

    page: int
    bbox: tuple[float, float, float, float]
    rows: tuple[tuple[str, ...], ...]


class DocImage(BaseModel):
    model_config = ConfigDict(frozen=True)

    page: int
    bbox: tuple[float, float, float, float]


class DocPage(BaseModel):
    model_config = ConfigDict(frozen=True)

    page: int
    kind: Literal["text", "scanned", "empty"]
    width: float
    height: float
    two_column: bool = False


class DocumentFacts(BaseModel):
    model_config = ConfigDict(frozen=True)

    file_name: str
    sha256: str
    pages: tuple[DocPage, ...]
    lines: tuple[DocLine, ...]
    tables: tuple[DocTable, ...] = ()
    images: tuple[DocImage, ...] = ()

    @property
    def scanned_pages(self) -> list[int]:
        return [p.page for p in self.pages if p.kind == "scanned"]


def _group_words(words: list[dict], tolerance: float = 2.5) -> list[list[dict]]:
    rows: list[list[dict]] = []
    for word in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
        if rows and abs(rows[-1][0]["top"] - word["top"]) <= tolerance:
            rows[-1].append(word)
        else:
            rows.append([word])
    return [sorted(r, key=lambda w: w["x0"]) for r in rows]


def _split_columns(row: list[dict], gap: float) -> list[list[dict]]:
    """Tách một hàng thành nhiều đoạn khi có khoảng trống lớn, ví dụ hai cột chữ."""
    parts = [[row[0]]]
    for word in row[1:]:
        if word["x0"] - parts[-1][-1]["x1"] > gap:
            parts.append([word])
        else:
            parts[-1].append(word)
    return parts


def _is_two_column(segments: list[list[dict]], width: float) -> bool:
    if len(segments) < TWO_COLUMN_MIN_LINES:
        return False
    left = sum(s[-1]["x1"] < width * 0.52 for s in segments)
    right = sum(s[0]["x0"] > width * 0.48 for s in segments)
    return left >= TWO_COLUMN_MIN_LINES / 2 and right >= TWO_COLUMN_MIN_LINES / 2 and (left + right) / len(segments) > 0.8


def read_pdf_lines(path: Path | str) -> DocumentFacts:
    """Đọc lớp chữ. Trang có ảnh phủ phần lớn là trang quét, lớp chữ của nó không được tin nên bị bỏ."""
    import pdfplumber
    from pdfminer.pdfdocument import PDFPasswordIncorrect

    path = Path(path)
    if not path.read_bytes().startswith(b"%PDF"):
        raise SourceFileError("FILE_CORRUPT", "File hỏng hoặc không phải PDF")
    pages: list[DocPage] = []
    lines: list[DocLine] = []
    tables: list[DocTable] = []
    images: list[DocImage] = []
    try:
        pdf = pdfplumber.open(path)
    except Exception as exc:
        # pdfplumber bọc lỗi của pdfminer, nên phải nhìn cả lỗi gốc để phân biệt mật khẩu với file hỏng.
        wrapped = isinstance(exc, PDFPasswordIncorrect) or "PDFPasswordIncorrect" in repr(exc) or any(isinstance(a, PDFPasswordIncorrect) for a in exc.args)
        if wrapped:
            raise SourceFileError("FILE_PASSWORD", "PDF có mật khẩu, chưa đọc được") from exc
        raise SourceFileError("FILE_CORRUPT", f"Không mở được PDF: {type(exc).__name__}") from exc
    with pdf:
        for number, page in enumerate(pdf.pages, start=1):
            width, height = float(page.width), float(page.height)
            area = width * height or 1.0
            coverage = sum(float(i["width"] * i["height"]) for i in page.images) / area
            words = page.extract_words(extra_attrs=["fontname", "size"], use_text_flow=False)
            if coverage >= SCAN_IMAGE_COVERAGE:
                pages.append(DocPage(page=number, kind="scanned", width=width, height=height))
                continue
            if not words:
                pages.append(DocPage(page=number, kind="empty", width=width, height=height))
                continue
            for image in page.images:
                images.append(DocImage(page=number, bbox=(float(image["x0"]), float(image["top"]), float(image["x1"]), float(image["bottom"]))))
            found = page.find_tables()
            bboxes = [t.bbox for t in found]
            for table in found:
                rows = tuple(tuple((c or "").strip() for c in row) for row in table.extract())
                tables.append(DocTable(page=number, bbox=tuple(float(v) for v in table.bbox), rows=rows))

            segments: list[list[dict]] = []
            for row in _group_words(words):
                segments.extend(_split_columns(row, gap=max(width * 0.05, 18)))
            two_column = _is_two_column(segments, width)
            pages.append(DocPage(page=number, kind="text", width=width, height=height, two_column=two_column))
            if two_column:
                mid = width * 0.5
                segments.sort(key=lambda s: (s[0]["x0"] >= mid, s[0]["top"], s[0]["x0"]))
            else:
                # Hàng có khoảng trống lớn vẫn là một dòng: giữ nguyên thứ tự từ trái sang.
                merged: list[list[dict]] = []
                for row in _group_words(words):
                    merged.append(row)
                segments = merged
            for segment in segments:
                text = " ".join(w["text"] for w in segment)
                sizes = Counter(round(float(w["size"]), 1) for w in segment)
                fontnames = "".join(str(w.get("fontname", "")) for w in segment[:1]).lower()
                center_y = (segment[0]["top"] + segment[0]["bottom"]) / 2
                center_x = (segment[0]["x0"] + segment[-1]["x1"]) / 2
                in_table = any(b[0] <= center_x <= b[2] and b[1] <= center_y <= b[3] for b in bboxes)
                lines.append(
                    DocLine(
                        index=len(lines),
                        page=number,
                        text=text,
                        x0=float(segment[0]["x0"]),
                        top=float(segment[0]["top"]),
                        x1=float(segment[-1]["x1"]),
                        bottom=float(segment[0]["bottom"]),
                        font_size=sizes.most_common(1)[0][0],
                        bold="bold" in fontnames or "black" in fontnames,
                        in_table=in_table,
                    )
                )
    return DocumentFacts(file_name=path.name, sha256=sha256_file(path), pages=tuple(pages), lines=tuple(mark_repeated(lines, pages)), tables=tuple(tables), images=tuple(images))


def _repeat_key(text: str) -> str:
    return _NUMBERED.sub("#", normalize_text(text))


def mark_repeated(lines: list[DocLine], pages: list[DocPage]) -> list[DocLine]:
    """Đầu trang, chân trang lặp: cùng chữ (không kể số) ở sát mép trên nhiều trang."""
    heights = {p.page: p.height for p in pages}
    seen: dict[str, set[int]] = {}
    for line in lines:
        height = heights.get(line.page, 0) or 1
        if line.top < height * EDGE_BAND or line.bottom > height * (1 - EDGE_BAND):
            seen.setdefault(_repeat_key(line.text), set()).add(line.page)
    total = len(pages)
    repeated = {k for k, v in seen.items() if len(v) >= REPEAT_MIN_PAGES and len(v) / max(total, 1) >= REPEAT_PAGE_SHARE and k}
    out = []
    for line in lines:
        height = heights.get(line.page, 0) or 1
        edge = line.top < height * EDGE_BAND or line.bottom > height * (1 - EDGE_BAND)
        out.append(line.model_copy(update={"repeated": True}) if edge and _repeat_key(line.text) in repeated else line)
    return out


def scale_tokens(tokens: list[dict], factor: float) -> list[dict]:
    """Đổi tọa độ OCR (pixel của ảnh) sang đơn vị của trang. Phải làm trước khi ghép dòng: mark_repeated so với chiều cao trang."""
    return [{**t, "bbox": [v * factor for v in t["bbox"]]} for t in tokens]


def lines_from_tokens(tokens: list[dict], page: int, start_index: int = 0) -> list[DocLine]:
    """Ghép token của OCR thành dòng. Token: {text, bbox: [x0, y0, x1, y1]}. Mọi dòng mang source = ocr."""
    if not tokens:
        return []
    heights = sorted(t["bbox"][3] - t["bbox"][1] for t in tokens)
    tolerance = max(heights[len(heights) // 2] * 0.6, 4)
    rows: list[list[dict]] = []
    for token in sorted(tokens, key=lambda t: (t["bbox"][1] + t["bbox"][3]) / 2):
        center = (token["bbox"][1] + token["bbox"][3]) / 2
        if rows and abs((rows[-1][0]["bbox"][1] + rows[-1][0]["bbox"][3]) / 2 - center) <= tolerance:
            rows[-1].append(token)
        else:
            rows.append([token])
    out = []
    for row in rows:
        row.sort(key=lambda t: t["bbox"][0])
        out.append(
            DocLine(
                index=start_index + len(out),
                page=page,
                text=" ".join(t["text"] for t in row),
                x0=float(row[0]["bbox"][0]),
                top=float(min(t["bbox"][1] for t in row)),
                x1=float(row[-1]["bbox"][2]),
                bottom=float(max(t["bbox"][3] for t in row)),
                # Trên ảnh không có cỡ chữ thật: lấy chiều cao dòng làm gần đúng.
                font_size=round(float(max(t["bbox"][3] - t["bbox"][1] for t in row)), 1),
                source="ocr",
            )
        )
    return out


def mark_uncertain(lines: list[DocLine], second: list[dict]) -> list[DocLine]:
    """Q70: mọi dòng có chữ số ở trang quét đều phải so với ảnh, không chỉ dòng hai OCR lệch.

    Hai OCR có thể cùng đọc sai, nên khớp nhau không đủ để bỏ qua. `digits_differ` chỉ để xếp thứ tự xem.
    Thông số máy nằm ở chữ số. Chữ có dấu khác nhau giữa hai engine là chuyện thường, nên không tính.
    """
    out = []
    for line in lines:
        center = (line.top + line.bottom) / 2
        overlapping = [t for t in second if t["bbox"][1] <= center <= t["bbox"][3] and t["bbox"][2] >= line.x0 and t["bbox"][0] <= line.x1]
        overlapping.sort(key=lambda t: t["bbox"][0])
        other = " ".join(t["text"] for t in overlapping)
        digits_a, digits_b = re.findall(r"\d+", line.text), re.findall(r"\d+", other)
        out.append(line.model_copy(update={"uncertain": bool(digits_a or digits_b) or not overlapping, "digits_differ": digits_a != digits_b}))
    return out
