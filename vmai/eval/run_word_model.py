"""Đo đầu mục Word bằng model trên tài liệu giả D01-D10, mỗi ca một lần (Q87). So với bộ đoán không LLM.

Không có --confirm-network thì chỉ chạy bộ đoán. D09 phải dừng trước khi gửi, D08 có mật khẩu không đọc được.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _env import load_key  # noqa: E402
from run_word import CASES, OUT, heading_scores  # noqa: E402
from vmai.document import DocumentFacts, lines_from_tokens, mark_repeated, mark_uncertain, read_pdf_lines, scale_tokens  # noqa: E402
from vmai.errors import SourceFileError  # noqa: E402
from vmai.llm.factory import create_provider, load_profile  # noqa: E402
from vmai.llm.keys import read_api_key  # noqa: E402
from vmai.outline import assign_headings, build_outline, guess_headings  # noqa: E402
from vmai.phi import ExternalSendPolicy, PhiDetected  # noqa: E402

SEND = ExternalSendPolicy(allow_documents=True)


def load_document(directory: Path, answer: dict) -> DocumentFacts:
    base = read_pdf_lines(directory / "source.pdf")
    if not answer.get("scanned"):
        return base
    tokens = json.loads((directory / "ocr_tokens.json").read_text(encoding="utf-8"))
    heights = {p.page: p.height for p in base.pages}
    lines = []
    for page in base.scanned_pages:
        with Image.open(directory / f"page{page}.jpg") as image:
            factor = heights[page] / image.height
        first = lines_from_tokens(scale_tokens(tokens["tesseract"][str(page)], factor), page, len(lines))
        lines.extend(mark_uncertain(first, scale_tokens(tokens["rapidocr"][str(page)], factor)))
    return DocumentFacts(file_name=base.file_name, sha256=base.sha256, pages=base.pages, lines=tuple(mark_repeated(lines, list(base.pages))))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", default=None)
    parser.add_argument("--confirm-network", action="store_true", help="Gọi model thật, tốn phí, gửi dòng của tài liệu giả")
    parser.add_argument("--only", default=None, help="Chỉ chạy các tài liệu có tên bắt đầu bằng chuỗi này")
    args = parser.parse_args()
    profile = load_profile(args.profile)
    if args.confirm_network:
        load_key(profile.api_key_env)
        read_api_key(profile)

    results = {}
    total = {"calls": 0, "input_tokens": 0, "output_tokens": 0, "reasoning_tokens": 0, "usd_by_price_table": 0.0}
    for directory in sorted(p for p in CASES.iterdir() if p.is_dir() and (not args.only or p.name.startswith(args.only))):
        answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
        fuzzy = bool(answer.get("scanned"))
        try:
            document = load_document(directory, answer)
        except (SourceFileError, FileNotFoundError) as exc:
            results[directory.name] = {"error": getattr(exc, "code", type(exc).__name__)}
            print(directory.name, results[directory.name], flush=True)
            continue
        guessed = build_outline(document, guess_headings(document.lines), "heuristic")
        row: dict = {"pages": len(document.pages), "heuristic": heading_scores(answer["headings"], guessed, fuzzy=fuzzy)}
        if args.confirm_network:
            provider = create_provider(profile)
            try:
                marks, calls, records = assign_headings(provider, document, SEND)
            except PhiDetected as exc:
                row["model"] = {"stopped": "PhiDetected", "kinds": sorted({f.kind for f in exc.findings}), "calls": len(provider.calls)}
            else:
                outline = build_outline(document, marks, "model")
                reasoning = sum(r.reasoning_tokens or 0 for r in provider.calls)
                row["model"] = {
                    **heading_scores(answer["headings"], outline, fuzzy=fuzzy),
                    "text_identical": outline.check.identical,
                    "marks": [{"line": m.line, "level": m.level} for m in marks],  # chỉ số dòng và cấp, không có chữ
                    "level_jumps": len(outline.level_jumps),
                    "calls": len(calls),
                    "input_tokens": sum(c.input_tokens or 0 for c in calls),
                    "output_tokens": sum(c.output_tokens or 0 for c in calls),
                    "reasoning_tokens": reasoning,
                    "latency_ms": [c.latency_ms for c in calls],
                    "usd_by_price_table": round(provider.usd_used, 6),
                    "sent_lines": sum(r.line_count for r in records),
                    "sent_chars": sum(r.char_count for r in records),
                }
                total["calls"] += len(calls)
                total["input_tokens"] += row["model"]["input_tokens"]
                total["output_tokens"] += row["model"]["output_tokens"]
                total["reasoning_tokens"] += reasoning
                total["usd_by_price_table"] += provider.usd_used
        results[directory.name] = row
        print(directory.name, json.dumps(row, ensure_ascii=False), flush=True)
    print("Tổng:", json.dumps(total, ensure_ascii=False))
    print("Tài liệu giả rất đều, mỗi ca chạy một lần. Không dự báo được trên tài liệu hãng thật.")
    OUT.mkdir(exist_ok=True)
    suffix = (profile.name if args.confirm_network else "heuristic") + (f"_{args.only}" if args.only else "")
    (OUT / f"word_model_{suffix}.json").write_text(json.dumps({"total": total, "cases": results}, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
