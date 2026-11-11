"""Đo đọc ảnh cho biểu mẫu lạ: OCR nội bộ, dựng bảng tổng quát, infer_spec, propose. Không gọi dịch vụ ngoài.

Ảnh là mô phỏng, chưa phải ảnh chụp thật. Tesseract đọc trước, RapidOCR đọc lần hai.
Chạy: eval/run_images.py read   (đọc ảnh, lưu token) rồi eval/run_images.py score
"""
from __future__ import annotations

import argparse
import json
import os
import time
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

from vmai import infer_spec, plan_headers, propose, read_form, read_source
from vmai.errors import VmaiError
from vmai.table_read import source_from_tokens

ROOT = Path(__file__).resolve().parents[2]
IMAGES = ROOT / "fixtures" / "eval" / "generated" / "images"
OUT = ROOT / "vmai" / "out"
CACHE = OUT / "ocr_cache"
VARIANTS = ("clean", "phone", "skew2", "skew5", "rot90", "rot180", "perspective", "dark", "shadow", "blur1.5", "blur3", "lowres", "screen", "teams")


def cases() -> list[Path]:
    out = []
    for d in sorted(IMAGES.iterdir()):
        if (d / "answer.json").exists() and d.name.rsplit("_s", 1)[0] in VARIANTS:
            out.append(d)
    return out


def read_all() -> None:
    from vmai.ocr import RapidOcr, TesseractOcr, prepare_image
    from vmai.table_read import orientation_score, orientation_sure

    os.environ.setdefault("VMAI_TESSERACT_DOCKER", "vmai-ocr-do")
    rapid, tess = RapidOcr(), TesseractOcr()
    CACHE.mkdir(parents=True, exist_ok=True)

    def scorer(rgb) -> tuple[int, bool]:
        tokens = tess.read(rgb)  # RapidOCR tự lật chữ nên không xét hướng được
        return orientation_score(tokens), orientation_sure(tokens)

    for directory in cases():
        target = CACHE / f"{directory.name}.json"
        if target.exists():
            continue
        page = prepare_image(directory / "page.jpg", scorer)
        started = time.perf_counter()
        first = tess.read(page.rgb)
        middle = time.perf_counter()
        second = rapid.read(page.rgb)
        target.write_text(
            json.dumps(
                {
                    "readers": [tess.reader_id, rapid.reader_id],
                    "first": first,
                    "second": second,
                    "rotation": page.rotation,
                    "low_quality": page.low_quality,
                    "edge_ratio": page.edge_ratio,
                    "sharpness": page.sharpness,
                    "brightness": page.brightness,
                    "seconds": [middle - started, time.perf_counter() - middle],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        print(directory.name, page.rotation, round(page.sharpness), round(page.brightness), len(first), len(second), flush=True)


def _same(a: str | None, b: str | None) -> bool:
    if a is None or b is None:
        return a == b  # ô đỏ không tự chọn số (X04) không khớp đáp án nào
    try:
        return Decimal(a) == Decimal(b)
    except InvalidOperation:
        return a == b


def score_case(directory: Path, cache: Path = CACHE, ask_model=None) -> dict:
    """`ask_model(plan) -> (ánh xạ, các lần gọi)`: chỉ dùng khi đo có model, chỉ nhận tiêu đề."""
    answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
    data = json.loads((cache / f"{directory.name}.json").read_text(encoding="utf-8"))
    gold = {c["a1"]: c["value"] for c in answer["cells"]}
    out = {"variant": answer.get("variant", answer.get("case")), "gold": len(gold), "status": None, "questions": [], "colors": {}, "right": 0, "wrong_by_color": {}, "missed": len(gold), "error": None, "low_quality": data["low_quality"], "seconds": data["seconds"]}
    try:
        source = source_from_tokens(
            directory.name, "x", [(1, data["first"], data["second"])], kind="image", first_reader=data["readers"][0], second_reader=data["readers"][1], low_quality=data["low_quality"]
        )
        guess = infer_spec(read_source(directory / "template.xlsx"), source, frozenset({"LEVEL_NAME_DIFFERS"}), experimental_image_sources=True)
        out["status"], out["questions"] = guess.status, [q.code for q in guess.questions]
        out["unmapped"] = list(guess.unmapped_columns)
        if guess.spec is None:
            return out
        form = read_form(directory / "template.xlsx", guess.spec)
        plan = plan_headers(form, source)
        mapping, calls = ask_model(plan) if ask_model and plan.llm_headers else (None, ())
        batch = propose(form, source, plan, mapping, calls, experimental_image_sources=True)
    except VmaiError as exc:
        out["error"] = exc.code
        return out
    out["notices"] = sorted({r.code for c in batch.changes for r in c.confidence.reasons} | {n.reason for n in batch.notices})
    colors, wrong = Counter(), Counter()
    seen = set()
    for change in batch.changes:
        a1, color = change.target.a1_at_write, change.confidence.final_color.value
        colors[color] += 1
        seen.add(a1)
        if a1 in gold and _same((change.value.normalized if change.value else None), gold[a1]):
            out["right"] += 1
        else:
            wrong[color] += 1
    out["colors"], out["wrong_by_color"], out["missed"] = dict(colors), dict(wrong), len(set(gold) - seen)
    return out


def score_all() -> None:
    results = [score_case(d) for d in cases() if (CACHE / f"{d.name}.json").exists()]
    by_variant = defaultdict(list)
    for r in results:
        by_variant[r["variant"]].append(r)
    print(f"{'biến thể':12}{'ảnh':>4}{'hỏi':>5}{'1 phần':>7}{'ô đáp án':>9}{'vàng':>6}{'đỏ':>5}{'xanh':>6}{'sai-vàng':>9}{'sai-đỏ':>7}{'sót':>6}  lỗi")
    total = Counter()
    rows = {}
    for variant in VARIANTS:
        items = by_variant.get(variant)
        if not items:
            continue
        c, w = Counter(), Counter()
        for r in items:
            c.update(r["colors"])
            w.update(r["wrong_by_color"])
        row = {
            "images": len(items),
            "asked": sum(r["status"] == "ask" for r in items),
            "partial": sum(r["status"] == "partial" for r in items),
            "errors": [r["error"] for r in items if r["error"]],
            "gold": sum(r["gold"] for r in items),
            "yellow": c["yellow"],
            "red": c["red"],
            "green": c["green"],
            "wrong_yellow": w["yellow"],
            "wrong_red": w["red"],
            "wrong_green": w["green"],
            "missed": sum(r["missed"] for r in items),
        }
        rows[variant] = row
        total.update({k: v for k, v in row.items() if isinstance(v, int)})
        print(f"{variant:12}{row['images']:>4}{row['asked']:>5}{row['partial']:>7}{row['gold']:>9}{row['yellow']:>6}{row['red']:>5}{row['green']:>6}{row['wrong_yellow']:>9}{row['wrong_red']:>7}{row['missed']:>6}  {','.join(sorted(set(row['errors'])))}")
    print("Tổng:", dict(total))
    print("Ảnh mô phỏng, biểu mẫu tự nghĩ. CHƯA kiểm chứng với ảnh chụp thật.")
    (OUT / "eval_images.json").write_text(json.dumps({"rows": rows, "total": dict(total), "cases": results}, ensure_ascii=False, indent=1), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=("read", "score"))
    args = parser.parse_args()
    read_all() if args.mode == "read" else score_all()


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
