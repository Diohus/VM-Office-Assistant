"""Dựng bảng từ token có tọa độ, không giả định bố cục, rồi đưa về `SourceFacts` như một file Excel.

Nhờ vậy ảnh và PDF đi chung đường với Excel: infer_spec, plan_headers, propose. Màu đọc do `reads` quyết định.
"""
from __future__ import annotations

import re
from statistics import median

from openpyxl.utils import get_column_letter

from .errors import SourceFileError
from .mapping import normalize_text, split_header_unit
from .numbers import ParseStatus, parse_number
from .source_xlsx import CellRead, SourceCell, SourceFacts, SourceSheet
from .template import classify_text

MIN_TABLE_ROWS = 3
MIN_TABLE_COLUMNS = 2
_DATE_RE = re.compile(r"\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{4}")
_DECIMAL_RE = re.compile(r"\d+[.,]\d+")
SURE_MIN_TOKENS = 10
SURE_MIN_SHARE = 0.25


def _rows(tokens: list[dict]) -> list[list[dict]]:
    heights = sorted(t["bbox"][3] - t["bbox"][1] for t in tokens)
    tolerance = max(heights[len(heights) // 2] * 0.6, 3)
    ordered = sorted(tokens, key=lambda t: (t["bbox"][1] + t["bbox"][3]) / 2)
    rows: list[list[dict]] = []
    for token in ordered:
        center = (token["bbox"][1] + token["bbox"][3]) / 2
        if rows and abs(median((t["bbox"][1] + t["bbox"][3]) / 2 for t in rows[-1]) - center) <= tolerance:
            rows[-1].append(token)
        else:
            rows.append([token])
    return [sorted(r, key=lambda t: t["bbox"][0]) for r in rows]


def _merge_cells(row: list[dict], gap: float) -> list[dict]:
    """Gộp các token gần nhau trong một hàng thành ô."""
    cells = [dict(row[0])]
    for token in row[1:]:
        # Hai số đứng sát nhau là hai cột, trừ khi số sau đúng là nhóm nghìn như `1 234`. Gộp vào thì bảng dày không dựng được, hoặc ra số sai.
        two_numbers = _is_clean_number(cells[-1]["text"]) and _is_clean_number(token["text"]) and not re.fullmatch(r"\d{3}", token["text"])
        if token["bbox"][0] - cells[-1]["bbox"][2] <= gap and not two_numbers:
            box = cells[-1]["bbox"]
            cells[-1] = {"text": f"{cells[-1]['text']} {token['text']}", "bbox": [box[0], min(box[1], token["bbox"][1]), token["bbox"][2], max(box[3], token["bbox"][3])]}
        else:
            cells.append(dict(token))
    return cells


def _column_lines(rows: list[list[dict]], gap: float) -> list[float]:
    """Tâm các cột: gom tâm của các ô theo chiều ngang trên mọi hàng có từ hai ô."""
    centers = sorted((c["bbox"][0] + c["bbox"][2]) / 2 for row in rows if len(row) >= 2 for c in row)
    if not centers:
        return []
    groups: list[list[float]] = [[centers[0]]]
    for x in centers[1:]:
        if x - groups[-1][-1] <= gap * 1.5:
            groups[-1].append(x)
        else:
            groups.append([x])
    # Cột thật phải có mặt ở một phần đáng kể các hàng, để chữ rải rác không tạo cột giả.
    wide_rows = sum(len(r) >= 2 for r in rows)
    return [median(g) for g in groups if len(g) >= max(2, wide_rows * 0.25)]


def tokens_to_grid(tokens: list[dict]) -> tuple[list[list[dict | None]], list[float]]:
    """Trả về bảng theo hàng và cột. Ô trống là None."""
    if not tokens:
        return [], []
    heights = sorted(t["bbox"][3] - t["bbox"][1] for t in tokens)
    gap = heights[len(heights) // 2] * 0.9
    rows = [_merge_cells(r, gap) for r in _rows(tokens)]
    lines = _column_lines(rows, gap * 2)
    if len(lines) < MIN_TABLE_COLUMNS:
        return [], lines
    grid: list[list[dict | None]] = []
    half = median(b - a for a, b in zip(lines, lines[1:])) / 2 if len(lines) > 1 else 50.0
    for row in rows:
        out: list[dict | None] = [None] * len(lines)
        conflict = False
        for cell in row:
            center = (cell["bbox"][0] + cell["bbox"][2]) / 2
            index = min(range(len(lines)), key=lambda i: abs(lines[i] - center))
            if abs(lines[index] - center) > half * 1.2:
                conflict = True
                continue
            if out[index] is None:
                out[index] = cell
            else:
                conflict = True
        # Hàng có ô ngoài cột, như tiêu đề trang hay dòng tổng ghi liền: bỏ khỏi bảng, không ép vào cột.
        if sum(c is not None for c in out) >= MIN_TABLE_COLUMNS and not conflict:
            grid.append(out)
    return grid, lines


def _strong_tokens(tokens: list[dict]) -> int:
    """Token là ngày hoặc số thập phân. Đọc ngược thì hiếm khi ra dạng này, còn số nguyên như 118, 811 thì vẫn ra."""
    return sum(bool(_DATE_RE.fullmatch(t["text"]) or _DECIMAL_RE.fullmatch(t["text"])) for t in tokens)


def orientation_sure(tokens: list[dict]) -> bool:
    strong = _strong_tokens(tokens)
    return strong >= SURE_MIN_TOKENS and strong >= len(tokens) * SURE_MIN_SHARE


def orientation_score(tokens: list[dict]) -> int:
    """Điểm cho một hướng ảnh, tính trên token của engine không tự lật chữ (Tesseract).

    RapidOCR tự lật chữ nên đọc ảnh ngược vẫn ra đúng chữ, không dùng để xét hướng được.
    Bảng toàn số nguyên không có token mạnh, lúc đó chỉ còn số ô và tiêu đề nằm trên dòng số: kém chắc hơn.
    """
    strong = _strong_tokens(tokens)
    grid, _ = tokens_to_grid(tokens)
    cells = 0
    seen_number = False
    above = 0
    for row in grid:
        present = [c for c in row if c is not None]
        cells += len(present)
        has_number = any(_is_clean_number(c["text"]) or _DATE_RE.fullmatch(c["text"].replace(" ", "")) for c in present)
        if has_number:
            seen_number = True
        elif not seen_number:
            above += 1
    # Ngược chiều thì dòng số đứng trước tiêu đề, `above` bằng 0.
    score = strong * 10 + cells
    return score if above else score // 2


def keep_one_header_row(grid: list[list[dict | None]]) -> list[list[dict | None]]:
    """Chỉ giữ một hàng tiêu đề ngay trên hàng dữ liệu đầu tiên. Hàng cao hơn, như tiêu đề trang bị OCR cắt thành nhiều ô, bị bỏ.

    Tầng trên của tiêu đề hai tầng do `compose_two_tier` dựng lại từ vị trí. Không dựng được thì cột không ghép được và phải hỏi, không sinh số sai.
    """

    def is_data(row: list[dict | None]) -> bool:
        clean = sum(c is not None and (_is_clean_number(c["text"]) or bool(_DATE_RE.fullmatch(c["text"].replace(" ", "")))) for c in row)
        return clean >= MIN_TABLE_COLUMNS

    first = next((i for i, row in enumerate(grid) if is_data(row)), None)
    return grid if first is None or first < 2 else grid[first - 1 :]


def _squeezed(text: str | None) -> str:
    return normalize_text(text or "").replace(" ", "")


def _upper_tier(header: list[dict | None], first: list[dict]) -> list[dict]:
    """Các ô của hàng ngay trên hàng tiêu đề. Hàng này thường đã bị loại khỏi bảng vì ô nằm giữa hai cột."""
    cells = [c for c in header if c is not None]
    top = min(c["bbox"][1] for c in cells)
    height = median(c["bbox"][3] - c["bbox"][1] for c in cells)
    gap = height * 0.9
    above = [r for r in (_merge_cells(row, gap) for row in _rows(first)) if max(c["bbox"][3] for c in r) <= top + height * 0.3]
    if not above:
        return []
    nearest = max(above, key=lambda r: max(c["bbox"][3] for c in r))
    return nearest if top - max(c["bbox"][3] for c in nearest) <= height * 2 else []


def compose_two_tier(grid: list[list[dict | None]], first: list[dict], second: list[dict] | None) -> list[list[dict | None]]:
    """Ghép tầng trên vào tiêu đề khi tầng dưới lặp tên, như `Mức 1`, `Mức 2` dưới từng tên xét nghiệm. Ảnh không có ô gộp nên phải suy từ vị trí.

    Ghép sai tầng là số vào nhầm cột mà vẫn trông hợp lý, nên chỉ ghép khi cấu trúc không còn cách hiểu khác: mỗi ô tầng trên trùm
    một số cột bằng nhau, liền nhau, các nhóm cùng thứ tự tên, và lần đọc thứ hai đọc ra cùng chữ. Thiếu một điều thì giữ nguyên,
    cột không ghép được sẽ phải hỏi.
    """
    if not grid:
        return grid
    header = grid[0]
    names = [_squeezed(c["text"]) if c is not None else "" for c in header]
    repeated = [i for i, name in enumerate(names) if name and names.count(name) > 1]
    if not repeated or any(b - a != 1 for a, b in zip(repeated, repeated[1:])):
        return grid
    upper = _upper_tier(header, first)
    if len(upper) < 2 or any(_is_clean_number(c["text"]) for c in upper):
        return grid
    centers = [(c["bbox"][0] + c["bbox"][2]) / 2 for c in upper]
    groups: dict[int, list[int]] = {i: [] for i in range(len(upper))}
    for column in repeated:
        cell = header[column]
        middle = (cell["bbox"][0] + cell["bbox"][2]) / 2
        distances = sorted((abs(center - middle), i) for i, center in enumerate(centers))
        if len(distances) > 1 and distances[1][0] - distances[0][0] < 1.0:
            return grid  # cách đều hai ô tầng trên: không biết thuộc ô nào
        groups[distances[0][1]].append(column)
    sizes = {len(g) for g in groups.values()}
    if len(sizes) != 1 or sizes.pop() < 2:
        return grid
    pattern = [names[c] for c in groups[0]]
    for i, columns in groups.items():
        left, right = header[columns[0]]["bbox"][0], header[columns[-1]]["bbox"][2]
        if [names[c] for c in columns] != pattern or len(set(pattern)) != len(pattern) or not left <= centers[i] <= right:
            return grid
    out = list(header)
    for i, columns in groups.items():
        word = upper[i]["text"]
        again = _second_text(upper[i], second) if second is not None else None
        if second is not None and _squeezed(again) != _squeezed(word):
            continue  # hai lần đọc không cùng chữ: nhóm này giữ nguyên và phải hỏi
        name, unit = split_header_unit(word)
        suffix = f" ({unit})" if unit else ""
        for column in columns:
            lower = header[column]
            composed = {**lower, "text": f"{name} {lower['text']}{suffix}"}
            if second is not None:
                lower_again = _second_text(lower, second)
                # Giữ đơn vị của lần đọc thứ hai, để đơn vị đọc lệch vẫn bị bắt ở bước sau.
                again_name, again_unit = split_header_unit(again)
                composed["second_text"] = None if lower_again is None else f"{again_name} {lower_again}" + (f" ({again_unit})" if again_unit else "")
            out[column] = composed
    return [out, *grid[1:]]


def _second_text(cell: dict, second: list[dict]) -> str | None:
    """Chữ của lần đọc thứ hai đúng vị trí ô. Token thứ hai tràn nhiều ô thì coi như không đọc được ô này."""
    x0, y0, x1, y1 = cell["bbox"]
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    width = x1 - x0
    hits = [t for t in second if t["bbox"][0] <= cx <= t["bbox"][2] and t["bbox"][1] <= cy <= t["bbox"][3]]
    if not hits:
        hits = [t for t in second if t["bbox"][1] <= cy <= t["bbox"][3] and x0 - 2 <= (t["bbox"][0] + t["bbox"][2]) / 2 <= x1 + 2]
    if not hits:
        return None
    if any(t["bbox"][2] - t["bbox"][0] > max(width * 3, 150) for t in hits):
        return None
    return "".join(t["text"] for t in sorted(hits, key=lambda t: t["bbox"][0])).replace(" ", "")


def _is_clean_number(text: str) -> bool:
    return parse_number(text.replace(" ", "")).status is not ParseStatus.INVALID


def printed_lines(*readings: list[dict] | None) -> tuple[str, ...]:
    """Từng hàng chữ của mỗi lần đọc, kể cả hàng bị loại khỏi bảng, để kiểm dòng tổng in trên phiếu."""
    lines: list[str] = []
    for tokens in readings:
        if tokens:
            lines.extend(" ".join(t["text"] for t in row) for row in _rows(tokens))
    return tuple(lines)


def grid_to_sheet(name: str, grid: list[list[dict | None]], *, page: int, first_reader: str, second_reader: str | None, second_tokens: list[dict] | None, lines: tuple[str, ...] = ()) -> SourceSheet:
    rows = []
    for r, row in enumerate(grid, start=1):
        cells = []
        for c, cell in enumerate(row, start=1):
            a1 = f"{get_column_letter(c)}{r}"
            if cell is None:
                cells.append(SourceCell(a1=a1, kind="empty"))
                continue
            kind, value = classify_text(cell["text"])
            second = cell["second_text"] if "second_text" in cell else (_second_text(cell, second_tokens) if second_tokens is not None else None)
            cells.append(
                SourceCell(
                    a1=a1,
                    kind=kind,
                    value=value,
                    read=CellRead(page=page, bbox=tuple(cell["bbox"]), reader=first_reader, second_reader=second_reader, second_text=second),
                )
            )
        rows.append(tuple(cells))
    return SourceSheet(name=name, rows=tuple(rows), printed_lines=lines)


def source_from_tokens(
    file_name: str,
    sha256: str,
    pages: list[tuple[int, list[dict], list[dict] | None]],
    *,
    kind: str,
    first_reader: str,
    second_reader: str | None,
    low_quality: bool = False,
    handwriting_declared: bool = False,
    scanned_text_layer: bool = False,
) -> SourceFacts:
    """`pages`: (số trang, token lần đọc đầu, token lần đọc hai). Không dựng được bảng nào thì báo, không đoán.

    Ảnh thiếu lần đọc thứ hai thì báo lỗi, không tụt im lặng về một lần đọc (Q83).
    """
    if kind in ("image", "pdf_image") and (not second_reader or any(second is None for _, _, second in pages)):
        raise SourceFileError("OCR_SECOND_READ_MISSING", "Thiếu lần đọc thứ hai. Ảnh phải được đọc bằng hai engine.")
    sheets = []
    for page, first, second in pages:
        grid = compose_two_tier(keep_one_header_row(tokens_to_grid(first)[0]), first, second)
        if len(grid) >= MIN_TABLE_ROWS:
            sheets.append(grid_to_sheet(f"Trang {page}", grid, page=page, first_reader=first_reader, second_reader=second_reader, second_tokens=second, lines=printed_lines(first, second)))
    if not sheets:
        raise SourceFileError("NO_MATCHING_TABLE", "Không tìm thấy bảng trong tài liệu")
    return SourceFacts(file_name=file_name, sha256=sha256, sheets=tuple(sheets), has_macros=False, hidden_sheets=0, kind=kind, low_quality=low_quality, handwriting_declared=handwriting_declared, scanned_text_layer=scanned_text_layer)
