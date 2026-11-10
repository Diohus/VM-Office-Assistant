"""Đọc bảng từ PDF có lớp chữ thành `SourceFacts` (Q98). Sau cờ `experimental_pdf_sources`, mặc định tắt.

PDF lớp chữ có thể xanh (R13), nên chữ đọc được phải đúng là chữ người ta nhìn thấy. Trang có dấu hiệu lớp chữ khác thứ hiển thị
thì từ chối cả file, không đoán và không tụt màu ngầm. Kiểm tra này chỉ bắt các dấu hiệu đã nghĩ ra, xem docs/ai/pdf_report.md.
"""
from __future__ import annotations

import re
from pathlib import Path

from .errors import FeatureNotReleased, SourceFileError
from .table_read import source_from_tokens
from .template import sha256_file

READER_ID = "pdf_text_code"  # theo docs/contracts/proposal_schema.md
MAX_PAGES = 50
SCAN_IMAGE_COVERAGE = 0.5  # cùng ngưỡng với trang quét ở document.py
MIN_VISIBLE_SIZE = 3.0  # giả định: chữ nhỏ hơn thì mắt không đọc được
WHITE_LEVEL = 0.95
OVERLAP_SHARE = 0.5
RENDER_DPI = 100
INK_CONTRAST = 40  # giả định: chênh với nền từ mức này trên thang 255 mới coi là có nét
_HIDDEN_RENDER_MODE = re.compile(rb"(?:^|[\s\]>)])[37]\s+Tr\b")


def _untrusted(page: int, why: str) -> SourceFileError:
    return SourceFileError("PDF_TEXT_LAYER_UNTRUSTED", f"Trang {page}: {why} Chữ trong file có thể khác thứ nhìn thấy, chưa đọc tự động được.")


def _streams(page_obj) -> list[bytes]:
    """Nội dung trang và các Form XObject nó dùng, để tìm chế độ vẽ chữ. pdfminer không cho biết chế độ này theo từng ký tự."""
    from pdfminer.pdftypes import PDFStream, resolve1

    out: list[bytes] = []
    seen: set[int] = set()

    def visit(obj, resources) -> None:
        stream = resolve1(obj)
        if isinstance(stream, PDFStream) and id(stream) not in seen:
            seen.add(id(stream))
            out.append(stream.get_data())
            walk(stream.attrs.get("Resources"))

    def walk(resources) -> None:
        resources = resolve1(resources)
        if not isinstance(resources, dict):
            return
        xobjects = resolve1(resources.get("XObject"))
        if isinstance(xobjects, dict):
            for item in xobjects.values():
                visit(item, None)

    for content in page_obj.contents:
        visit(content, None)
    walk(page_obj.resources)
    return out


def _is_white(color) -> bool:
    if color is None:
        return False
    values = color if isinstance(color, (tuple, list)) else (color,)
    if not values or not all(isinstance(v, (int, float)) for v in values):
        return False
    if len(values) == 4:  # CMYK: trắng là không có mực
        return all(v <= 1 - WHITE_LEVEL for v in values)
    return all(v >= WHITE_LEVEL for v in values)


def _overlaps(a: dict, b: dict) -> bool:
    width = min(a["x1"], b["x1"]) - max(a["x0"], b["x0"])
    height = min(a["bottom"], b["bottom"]) - max(a["top"], b["top"])
    if width <= 0 or height <= 0:
        return False
    smaller = min((a["x1"] - a["x0"]) * (a["bottom"] - a["top"]), (b["x1"] - b["x0"]) * (b["bottom"] - b["top"])) or 1.0
    return width * height / smaller >= OVERLAP_SHARE


def _invisible_word(page, words: list[dict]) -> dict | None:
    """Dựng trang thành ảnh rồi xem chỗ có chữ có nét không. Bắt được chữ bị che bằng ô màu vẽ sau, vốn không lộ ra trong lớp chữ."""
    gray = page.to_image(resolution=RENDER_DPI, antialias=False).original.convert("L")
    scale = RENDER_DPI / 72
    for word in words:
        box = (max(int(float(word["x0"]) * scale), 0), max(int(float(word["top"]) * scale), 0), int(float(word["x1"]) * scale) + 1, int(float(word["bottom"]) * scale) + 1)
        darkest, lightest = gray.crop(box).getextrema()
        if lightest - darkest < INK_CONTRAST:
            return word
    return None


def _check_page(number: int, page, words: list[dict]) -> None:
    width, height = float(page.width), float(page.height)
    coverage = sum(float(i["width"] * i["height"]) for i in page.images) / (width * height or 1.0)
    if coverage >= SCAN_IMAGE_COVERAGE:
        raise _untrusted(number, "có ảnh phủ phần lớn trang, giống trang quét (R11).")
    for stream in _streams(page.page_obj):
        if _HIDDEN_RENDER_MODE.search(stream):
            raise _untrusted(number, "có chữ ẩn.")
    for char in page.chars:
        if not char["text"].strip():
            continue
        if _is_white(char.get("non_stroking_color")):
            raise _untrusted(number, "có chữ màu trắng.")
        if float(char["size"]) < MIN_VISIBLE_SIZE:
            raise _untrusted(number, "có chữ quá nhỏ để đọc bằng mắt.")
        if char["x1"] < 0 or char["x0"] > width or char["bottom"] < 0 or char["top"] > height:
            raise _untrusted(number, "có chữ nằm ngoài trang.")
    for image in page.images:
        if any(_overlaps(word, image) for word in words):
            raise _untrusted(number, "có ảnh đè lên chữ.")
    # Hai ký tự khác nhau cùng một chỗ: extract_words có thể gộp chúng thành một từ, nên phải xét từng ký tự.
    visible = [c for c in page.chars if c["text"].strip()]
    at: dict[tuple[int, int], str] = {}
    for char in visible:
        spot = (round(float(char["x0"])), round(float(char["top"])))
        if at.setdefault(spot, char["text"]) != char["text"]:
            raise _untrusted(number, "có hai chữ khác nhau đè lên nhau.")
    if words and _invisible_word(page, words) is not None:
        raise _untrusted(number, "có chữ trong file mà không nhìn thấy trên trang.")
    ordered = sorted(words, key=lambda w: (w["top"], w["x0"]))
    for i, word in enumerate(ordered):
        for other in ordered[i + 1 :]:
            if other["top"] > word["bottom"]:
                break
            if other["text"] != word["text"] and _overlaps(word, other):
                raise _untrusted(number, "có hai chữ khác nhau đè lên nhau.")


def read_pdf_table(path: Path | str, *, experimental_pdf_sources: bool = False):
    """PDF lớp chữ thành `SourceFacts` loại `pdf_text`. Không có lớp chữ hoặc không có bảng thì báo, không đoán."""
    if not experimental_pdf_sources:
        raise FeatureNotReleased("Đọc bảng từ PDF")
    import pdfplumber

    path = Path(path)
    if not path.read_bytes().startswith(b"%PDF"):
        raise SourceFileError("FILE_CORRUPT", "File hỏng hoặc không phải PDF")
    try:
        pdf = pdfplumber.open(path)
    except Exception as exc:
        if "PDFPasswordIncorrect" in repr(exc) or "PDFPasswordIncorrect" in repr(exc.__cause__) or any("PDFPasswordIncorrect" in repr(a) for a in exc.args):
            raise SourceFileError("FILE_PASSWORD", "PDF có mật khẩu, chưa đọc được") from exc
        raise SourceFileError("FILE_CORRUPT", f"Không mở được PDF: {type(exc).__name__}") from exc
    pages = []
    with pdf:
        if len(pdf.pages) > MAX_PAGES:
            raise SourceFileError("FILE_TOO_LARGE", f"PDF quá {MAX_PAGES} trang")
        for number, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(use_text_flow=False)
            _check_page(number, page, words)
            pages.append((number, [{"text": w["text"], "bbox": [float(w["x0"]), float(w["top"]), float(w["x1"]), float(w["bottom"])]} for w in words], None))
    if not any(tokens for _, tokens, _ in pages):
        raise SourceFileError("PDF_NO_TEXT_LAYER", "PDF không có lớp chữ")
    usable = [p for p in pages if p[1]]
    return source_from_tokens(path.name, sha256_file(path), usable, kind="pdf_text", first_reader=READER_ID, second_reader=None)
