"""Đo Word đầu mục (Q54) trên tài liệu giả. Không gọi mạng.

Bộ đoán không dùng LLM đo được thật. Độ đúng đầu mục của model là CHƯA ĐO: chưa có key.
"""
from __future__ import annotations

import difflib
import json
import re
from pathlib import Path

from docx import Document
from PIL import Image

from vmai.document import DocumentFacts, lines_from_tokens, mark_repeated, mark_uncertain, read_pdf_lines, scale_tokens
from vmai.errors import SourceFileError
from vmai.outline import build_outline, check_written_paragraphs, guess_headings

ROOT = Path(__file__).resolve().parents[2]
CASES = ROOT / "fixtures" / "eval" / "generated" / "word"
OUT = ROOT / "vmai" / "out"


def _close(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a, b).ratio()


def heading_scores(gold: list[dict], outline, fuzzy: bool = False) -> dict:
    found = [(b.text, b.level) for b in outline.blocks if b.kind == "heading"]
    right = level_right = 0
    used: set[int] = set()
    for item in gold:
        best, best_i = 0.0, None
        for i, (text, _) in enumerate(found):
            if i in used:
                continue
            score = _close(item["text"], text) if fuzzy else float(item["text"] == text)
            if score > best:
                best, best_i = score, i
        if best_i is not None and best >= (0.7 if fuzzy else 1.0):
            used.add(best_i)
            right += 1
            level_right += int(found[best_i][1] == item["level"])
    return {
        "gold": len(gold),
        "found": len(found),
        "right": right,
        "level_right": level_right,
        "precision": right / len(found) if found else None,
        "recall": right / len(gold) if gold else None,
    }


def write_and_check(outline, path: Path) -> dict:
    """Ghi thử ra Word rồi đọc lại. Bản ghi của sản phẩm là việc của Dev; đây chỉ để thử phép so chữ tới cuối."""
    doc = Document()
    for block in outline.blocks:
        if block.kind == "heading":
            doc.add_heading(f"{block.number} {block.text}" if block.number else block.text, level=min(block.level, 9))
        elif block.kind == "paragraph":
            doc.add_paragraph(block.text)
        elif block.kind == "table":
            table = doc.add_table(rows=len(block.rows), cols=max(len(r) for r in block.rows))
            for r, row in enumerate(block.rows):
                for c, cell in enumerate(row):
                    table.cell(r, c).text = cell
    doc.save(path)
    check = check_written_paragraphs(outline, [p.text for p in Document(path).paragraphs])
    return check.model_dump()


def scan_case(directory: Path, answer: dict) -> dict:
    """Trang quét: Tesseract đọc chính vì đọc tiếng Việt tốt hơn, RapidOCR so chữ số."""
    tokens = json.loads((directory / "ocr_tokens.json").read_text(encoding="utf-8"))
    base = read_pdf_lines(directory / "source.pdf")
    lines = []
    heights = {p.page: p.height for p in base.pages}
    for page in base.scanned_pages:
        with Image.open(directory / f"page{page}.jpg") as image:
            factor = heights[page] / image.height
        first = lines_from_tokens(scale_tokens(tokens["tesseract"][str(page)], factor), page, len(lines))
        lines.extend(mark_uncertain(first, scale_tokens(tokens["rapidocr"][str(page)], factor)))
    lines = mark_repeated(lines, list(base.pages))
    document = DocumentFacts(file_name=base.file_name, sha256=base.sha256, pages=base.pages, lines=tuple(lines))
    outline = build_outline(document, guess_headings(document.lines), "heuristic")
    # Số là thứ nguy hiểm nhất ở tài liệu máy: dòng đọc sai chữ số có được tô không?
    # Mỗi dòng OCR được so với đúng dòng đã in gần nó nhất. Phải khớp đủ các nhóm chữ số.
    gold_lines = answer["rendered_lines"]
    wrong_digits = flagged = right_digit_lines = false_alarm = 0
    differ_hits: list[bool] = []
    samples = []
    for line in lines:
        if line.repeated:
            continue
        best = max(gold_lines, key=lambda g: _close(line.text, g))
        want = re.findall(r"\d+", best)
        got = re.findall(r"\d+", line.text)
        if not want and not got:
            continue
        if want != got:
            wrong_digits += 1
            flagged += int(line.uncertain)
            differ_hits.append(line.digits_differ)
            if not line.uncertain:
                samples.append({"read": line.text[:70], "printed": best[:70]})
        else:
            right_digit_lines += 1
            false_alarm += int(line.uncertain)
    read_text = " ".join(l.text for l in lines)
    gold_text = " ".join(item["text"] for item in answer["headings"]) + " " + " ".join(answer["body"])
    gold_numbers = sorted(t for t in gold_text.split() if any(c.isdigit() for c in t))
    read_numbers = sorted(t for t in read_text.split() if any(c.isdigit() for c in t))
    return {
        "pages_scanned": len(base.scanned_pages),
        "lines": len(lines),
        "dropped_lines": len(outline.dropped_lines),
        "dropped_groups": [{"sample": g.sample[:60], "pages": list(g.pages)} for g in outline.dropped_groups],
        "uncertain_lines": sum(l.uncertain for l in lines),
        "digits_differ_lines": sum(l.digits_differ for l in lines),
        "wrong_digit_lines_flagged_by_disagreement_only": sum(1 for s in differ_hits if s),
        "lines_with_wrong_digits": wrong_digits,
        "wrong_digit_lines_flagged": flagged,
        "right_digit_lines": right_digit_lines,
        "right_digit_lines_flagged_needlessly": false_alarm,
        "unflagged_samples": samples[:8],
        "text_similarity": round(_close(sorted(gold_text.split()).__str__(), sorted(read_text.split()).__str__()), 3),
        "number_tokens_gold_only": len(set(gold_numbers) - set(read_numbers)),
        "headings": heading_scores(answer["headings"], outline, fuzzy=True),
        "check_identical": outline.check.identical,
    }


def main() -> None:
    OUT.mkdir(exist_ok=True)
    results = {}
    for directory in sorted(p for p in CASES.iterdir() if p.is_dir()):
        answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
        name = directory.name
        try:
            if answer.get("scanned"):
                results[name] = scan_case(directory, answer) if (directory / "ocr_tokens.json").exists() else {"skipped": "chưa chạy ocr_word_pages.py"}
                continue
            document = read_pdf_lines(directory / "source.pdf")
        except SourceFileError as exc:
            results[name] = {"error": exc.code}
            continue
        outline = build_outline(document, guess_headings(document.lines), "heuristic")
        results[name] = {
            "headings": heading_scores(answer["headings"], outline),
            "text_identical": outline.check.identical,
            "text_missing": list(outline.check.missing),
            "dropped_lines": len(outline.dropped_lines),
            "dropped_groups": [{"sample": g.sample, "pages": list(g.pages)} for g in outline.dropped_groups],
            "tables": len(document.tables),
            "two_column_pages": sum(p.two_column for p in document.pages),
            "written_check": write_and_check(outline, OUT / f"{name}.docx")["identical"],
        }
        # Thứ tự thân bài so với đáp án: phát hiện đọc sai thứ tự, điều phép so chữ nội bộ không thấy được.
        body_gold = [t for item in answer["body"] for t in item.split()]
        body_read = [t for b in outline.blocks if b.kind == "paragraph" for t in b.text.split()]
        results[name]["reading_order_similarity"] = round(difflib.SequenceMatcher(None, body_gold, body_read, autojunk=False).ratio(), 3)
    for name, r in results.items():
        print(name, json.dumps(r, ensure_ascii=False))
    print("\nĐây là bộ đoán không dùng LLM trên tài liệu tự nghĩ. Độ đúng đầu mục của model: CHƯA ĐO.")
    (OUT / "eval_word.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
