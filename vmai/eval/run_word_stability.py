"""Đo độ ổn định đầu mục của model (Q88): D03, D04, D10, mỗi ca 3 lần. Ba lần vẫn là mẫu nhỏ.

Lưu từng dòng (số dòng, cấp) của mỗi lần. Tài liệu giả, chỉ gửi dòng của tài liệu giả.
"""
from __future__ import annotations

import argparse
import difflib
import json
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _env import load_key  # noqa: E402
from run_word import CASES, OUT  # noqa: E402
from run_word_model import SEND, load_document  # noqa: E402
from vmai.llm.factory import create_provider, load_profile  # noqa: E402
from vmai.llm.keys import read_api_key  # noqa: E402
from vmai.outline import HeadingMark, assign_headings, build_outline  # noqa: E402

DOCS = ("D03_chi_co_so", "D04_khong_gi_ca", "D10_quet")
RUNS = 3
MATCH_MIN = 0.7


def score(answer: dict, document, marks) -> dict:
    """So từng đầu mục đáp án với dòng gần nhất. Trang quét: đầu mục đúng chỗ nhưng chữ hỏng thì tách riêng."""
    fuzzy = bool(answer.get("scanned"))
    lines = {l.index: l for l in document.lines if not l.repeated}
    by_line = {m.line: m.level for m in marks}
    rendered = answer["headings"]
    right = level_right = 0
    used: set[int] = set()
    for gold in rendered:
        scored = [(difflib.SequenceMatcher(None, gold["text"], lines[i].text).ratio() if fuzzy else float(gold["text"] == lines[i].text), i) for i in by_line if i not in used]
        if not scored:
            continue
        best, index = max(scored)
        if best >= (MATCH_MIN if fuzzy else 1.0):
            used.add(index)
            right += 1
            level_right += int(by_line[index] == gold["level"])
    extra = [{"text": lines[i].text[:70], "level": by_line[i]} for i in by_line if i not in used and i in lines]
    return {"found": len(by_line), "right": right, "level_right": level_right, "gold": len(rendered), "unmatched_marks": extra, "level_counts": {str(k): list(by_line.values()).count(k) for k in sorted(set(by_line.values()))}}


def rescore(profile_name: str) -> None:
    """Chấm lại câu trả lời model đã lưu theo Q89. Không gọi mạng, không tốn phí. Model không được hỏi lại."""
    saved = json.loads((OUT / f"word_stability_{profile_name}.json").read_text(encoding="utf-8"))
    rows = {}
    for name, case in saved["cases"].items():
        directory = CASES / name
        answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
        document = load_document(directory, answer)
        rows[name] = []
        for run in case["runs"]:
            marks = [HeadingMark(line=line, level=level) for line, level in run["marks"]]
            outline = build_outline(document, marks, "model")
            final = [HeadingMark(line=b.line_indices[0], level=b.level) for b in outline.blocks if b.kind == "heading"]
            row = score(answer, document, final)
            by_source = Counter(b.level_source for b in outline.blocks if b.kind == "heading")
            rows[name].append({k: row[k] for k in ("found", "right", "level_right", "gold")} | {"before_level_right": run["level_right"], "level_source": dict(by_source), "text_identical": outline.check.identical})
            print(name, rows[name][-1], flush=True)
    (OUT / f"word_stability_{profile_name}_q89.json").write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", default=None)
    parser.add_argument("--confirm-network", action="store_true")
    parser.add_argument("--rescore", action="store_true", help="chấm lại kết quả đã lưu, không gọi mạng")
    args = parser.parse_args()
    profile = load_profile(args.profile)
    if args.rescore:
        rescore(profile.name)
        return
    print(f"{profile.provider}/{profile.model}: {len(DOCS)} tài liệu, mỗi tài liệu {RUNS} lần.")
    if not args.confirm_network:
        print("Chưa có --confirm-network: không gọi mạng.")
        return
    load_key(profile.api_key_env)
    read_api_key(profile)
    results: dict = {}
    total = {"calls": 0, "input_tokens": 0, "output_tokens": 0, "usd_by_price_table": 0.0}
    for name in DOCS:
        directory = CASES / name
        answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
        document = load_document(directory, answer)
        runs = []
        for attempt in range(RUNS):
            provider = create_provider(profile)
            marks, calls, _ = assign_headings(provider, document, SEND)
            outline = build_outline(document, marks, "model")
            row = score(answer, document, marks)
            row.update(attempt=attempt + 1, text_identical=outline.check.identical, marks=[[m.line, m.level] for m in marks], input_tokens=sum(c.input_tokens or 0 for c in calls), output_tokens=sum(c.output_tokens or 0 for c in calls), usd=provider.usd_used)
            runs.append(row)
            total["calls"] += len(calls)
            total["input_tokens"] += row["input_tokens"]
            total["output_tokens"] += row["output_tokens"]
            total["usd_by_price_table"] += provider.usd_used
            print(name, attempt + 1, {k: row[k] for k in ("found", "right", "level_right", "gold", "level_counts", "text_identical")}, "thừa:", [e["text"][:40] for e in row["unmatched_marks"]][:12], flush=True)
        sets = [{tuple(m) for m in r["marks"]} for r in runs]
        results[name] = {"runs": runs, "same_marks_in_all_runs": len(set.intersection(*sets)), "marks_in_any_run": len(set.union(*sets))}
        print(name, "giống nhau ở cả 3 lần:", results[name]["same_marks_in_all_runs"], "trên", results[name]["marks_in_any_run"], flush=True)
    print("Tổng:", json.dumps(total, ensure_ascii=False))
    OUT.mkdir(exist_ok=True)
    (OUT / f"word_stability_{profile.name}.json").write_text(json.dumps({"total": total, "cases": results}, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
