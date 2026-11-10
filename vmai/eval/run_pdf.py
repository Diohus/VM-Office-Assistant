"""Đo đọc bảng PDF có lớp chữ (Q98). Không gọi dịch vụ ngoài. Cờ PDF chỉ bật trong lúc đo.

Chạy: eval/run_pdf.py            bộ tự dựng, số lạc quan
      eval/run_pdf.py --set <bộ> --out <thư mục>   bộ độc lập: <bộ>/pdf/<ca>/{source.pdf,template.xlsx,answer.json}
Màn hình chỉ in số tổng theo bố cục. Thoát mã 1 nếu có ô xanh: PDF chưa so với OCR thì tối đa vàng.
"""
from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

from vmai import infer_spec, plan_headers, propose, read_form, read_source
from vmai.errors import VmaiError
from vmai.pdf_table import read_pdf_table

ROOT = Path(__file__).resolve().parents[2]
OWN = ROOT / "fixtures" / "eval" / "generated" / "pdf"
OUT = ROOT / "vmai" / "out"


def _same(a: str, b: str) -> bool:
    try:
        return Decimal(a) == Decimal(b)
    except InvalidOperation:
        return a == b


def score_case(directory: Path) -> dict:
    answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
    gold = {c["a1"]: c["value"] for c in answer["cells"]}
    out = {"name": directory.name, "layout": str(answer.get("layout", answer.get("case", "(không ghi)"))), "gold": len(gold), "status": None, "questions": [], "error": None, "colors": {}, "wrong_by_color": {}, "missed": len(gold)}
    try:
        source = read_pdf_table(directory / "source.pdf", experimental_pdf_sources=True)
        guess = infer_spec(read_source(directory / "template.xlsx"), source, frozenset({"LEVEL_NAME_DIFFERS"}), experimental_pdf_sources=True)
        out["status"], out["questions"] = guess.status, [q.code for q in guess.questions]
        if guess.spec is None:
            return out
        form = read_form(directory / "template.xlsx", guess.spec)
        batch = propose(form, source, plan_headers(form, source), experimental_pdf_sources=True)
    except VmaiError as exc:
        out["error"] = exc.code
        return out
    colors, wrong, seen = Counter(), Counter(), set()
    for change in batch.changes:
        a1, color = change.target.a1_at_write, change.confidence.final_color.value
        colors[color] += 1
        seen.add(a1)
        if a1 not in gold or not _same(change.value.normalized, gold[a1]):
            wrong[color] += 1
    out.update(colors=dict(colors), wrong_by_color=dict(wrong), missed=len(set(gold) - seen))
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--set", type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    if args.set is not None and args.out is None:
        parser.error("--set cần --out")
    base = OWN if args.set is None else args.set.resolve() / "pdf"
    target = OUT if args.set is None else args.out.resolve()
    if args.set is not None and (target == OUT.resolve() or OUT.resolve() in target.parents):
        raise SystemExit("Bộ ngoài không được ghi vào vmai/out")
    rows = [score_case(d) for d in sorted(base.iterdir()) if (d / "answer.json").exists()]
    groups = defaultdict(list)
    for row in rows:
        groups[row["layout"]].append(row)
    print(f"{'bố cục':24}{'file':>5}{'ok':>4}{'hỏi':>4}{'lỗi':>4}{'đáp án':>7}{'vàng':>6}{'đỏ':>5}{'xanh':>5}{'sai-vàng':>9}{'sai-đỏ':>7}{'sai-xanh':>9}{'sót':>6}")
    total = Counter()
    for name in sorted(groups):
        c, w = Counter(), Counter()
        for row in groups[name]:
            c.update(row["colors"])
            w.update(row["wrong_by_color"])
        items = groups[name]
        s = {"files": len(items), "ok": sum(r["status"] in ("ok", "partial") for r in items), "asked": sum(r["status"] == "ask" for r in items), "errors": sum(bool(r["error"]) for r in items), "gold": sum(r["gold"] for r in items), "yellow": c["yellow"], "red": c["red"], "green": c["green"], "wrong_yellow": w["yellow"], "wrong_red": w["red"], "wrong_green": w["green"], "missed": sum(r["missed"] for r in items)}
        total.update(s)
        print(f"{name:24}{s['files']:>5}{s['ok']:>4}{s['asked']:>4}{s['errors']:>4}{s['gold']:>7}{s['yellow']:>6}{s['red']:>5}{s['green']:>5}{s['wrong_yellow']:>9}{s['wrong_red']:>7}{s['wrong_green']:>9}{s['missed']:>6}")
    questions = Counter(q for r in rows for q in r["questions"])
    errors = Counter(r["error"] for r in rows if r["error"])
    print("Tổng:", dict(total))
    print("Mã lý do hỏi lại:", dict(questions), "| lỗi:", dict(errors))
    print("PDF mô phỏng. CHƯA kiểm chứng với PDF thật." if args.set is None else "Bộ độc lập, vẫn là dữ liệu giả.")
    target.mkdir(parents=True, exist_ok=True)
    (target / ("eval_pdf.json" if args.set is None else "eval_pdf_external.json")).write_text(json.dumps({"total": dict(total), "questions": dict(questions), "errors": dict(errors), "cases": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
    return 1 if total["green"] else 0


if __name__ == "__main__":
    import _console

    _console.utf8()
    sys.exit(main())
