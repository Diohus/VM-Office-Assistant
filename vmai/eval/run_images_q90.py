"""Q90: đo phần ảnh trên các bộ chưa đo và bộ chưa dùng để chỉnh. Không gọi dịch vụ ngoài.

Chạy: eval/run_images_q90.py read   rồi   eval/run_images_q90.py score
Bộ `old`: A05, A08, A10 và A09, A09neg (dò dấu hiệu người bệnh). Bộ `blind`: fixtures/eval/generated/images_q90.
Bộ ngoài (Q96): thêm `--set <thư mục> --out <thư mục>`, xem `measure_external`. Không ghi gì vào vmai/out, không in tên ca.
Ảnh mô phỏng, chưa phải ảnh chụp thật.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import statistics
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from run_images import CACHE, IMAGES, OUT, ROOT, score_case  # noqa: E402
from vmai.phi import Token, scan_tokens  # noqa: E402

BLIND = ROOT / "fixtures" / "eval" / "generated" / "images_q90"
BLIND_CACHE = OUT / "ocr_cache_q90"


def old_dirs() -> list[Path]:
    return [d for d in sorted(IMAGES.iterdir()) if d.name.startswith(("A05", "A08", "A09", "A10"))]


def blind_dirs() -> list[Path]:
    return [d for d in sorted(BLIND.iterdir()) if (d / "answer.json").exists()]


def read(directories: list[Path], cache: Path, *, quiet: bool = False) -> None:
    from vmai.ocr import RapidOcr, TesseractOcr, prepare_image
    from vmai.table_read import orientation_score, orientation_sure

    os.environ.setdefault("VMAI_TESSERACT_DOCKER", "vmai-ocr-do")
    rapid, tess = RapidOcr(), TesseractOcr()
    cache.mkdir(parents=True, exist_ok=True)
    for directory in directories:
        target = cache / f"{directory.name}.json"
        if target.exists():
            continue
        tries = [0]

        def scorer(rgb):
            tries[0] += 1
            tokens = tess.read(rgb)
            return orientation_score(tokens), orientation_sure(tokens)

        started = time.perf_counter()
        page = prepare_image(directory / "page.jpg", scorer)
        prepared = time.perf_counter()
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
                    "seconds": [middle - prepared, time.perf_counter() - middle],
                    "prepare_seconds": prepared - started,
                    "orientation_tries": tries[0],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        if not quiet:
            print(directory.name, page.rotation, tries[0], round(prepared - started, 1), flush=True)


def _summary(rows: list[dict]) -> dict:
    colors, wrong = Counter(), Counter()
    for r in rows:
        colors.update(r["colors"])
        wrong.update(r["wrong_by_color"])
    return {
        "images": len(rows),
        "ok": sum(r["status"] == "ok" for r in rows),
        "partial": sum(r["status"] == "partial" for r in rows),
        "asked": sum(r["status"] == "ask" for r in rows),
        "errors": sum(bool(r["error"]) for r in rows),
        "gold": sum(r["gold"] for r in rows),
        "yellow": colors["yellow"],
        "red": colors["red"],
        "green": colors["green"],
        "wrong_yellow": wrong["yellow"],
        "wrong_red": wrong["red"],
        "wrong_green": wrong["green"],
        "missed": sum(r["missed"] for r in rows),
    }


def _print_table(title: str, groups: dict) -> dict:
    print(f"\n{title}")
    print(f"{'nhóm':24}{'ảnh':>4}{'ok':>4}{'1phần':>6}{'hỏi':>4}{'lỗi':>4}{'đáp án':>7}{'vàng':>6}{'đỏ':>5}{'xanh':>5}{'sai-vàng':>9}{'sai-đỏ':>7}{'sót':>6}")
    out = {}
    for name in sorted(groups):
        s = out[name] = _summary(groups[name])
        print(f"{name:24}{s['images']:>4}{s['ok']:>4}{s['partial']:>6}{s['asked']:>4}{s['errors']:>4}{s['gold']:>7}{s['yellow']:>6}{s['red']:>5}{s['green']:>5}{s['wrong_yellow']:>9}{s['wrong_red']:>7}{s['missed']:>6}")
    return out


def score_blind() -> dict:
    manifest = json.loads((BLIND / "manifest.json").read_text(encoding="utf-8"))
    # Mã băm tính trên chuỗi LF lúc sinh; trên Windows file được ghi thành CRLF nên phải quy đổi trước khi so.
    changed = [d.name for d in blind_dirs() if hashlib.sha256((d / "answer.json").read_bytes().replace(b"\r\n", b"\n")).hexdigest() != manifest[d.name]]
    print("Đáp án đổi so với manifest:", changed or "không")
    rows = []
    for d in blind_dirs():
        if not (BLIND_CACHE / f"{d.name}.json").exists():
            continue
        answer = json.loads((d / "answer.json").read_text(encoding="utf-8"))
        data = json.loads((BLIND_CACHE / f"{d.name}.json").read_text(encoding="utf-8"))
        row = score_case(d, BLIND_CACHE)
        row.update(name=d.name, layout=answer["layout"], rotation=data["rotation"], prepare_seconds=data["prepare_seconds"], orientation_tries=data["orientation_tries"])
        rows.append(row)
    by_layout, by_variant, by_both = defaultdict(list), defaultdict(list), defaultdict(list)
    for r in rows:
        by_layout[r["layout"]].append(r)
        by_variant[r["variant"]].append(r)
        by_both[f"{r['layout'][:2]} {r['variant']}"].append(r)
    result = {
        "layout": _print_table("Theo bố cục", by_layout),
        "variant": _print_table("Theo loại ảnh", by_variant),
        "both": _print_table("Theo bố cục và loại ảnh", by_both),
        "total": _summary(rows),
    }
    print("\nTổng:", result["total"])
    print("Câu hỏi:", Counter(q for r in rows for q in r["questions"]), "| lỗi:", Counter(r["error"] for r in rows if r["error"]))
    for r in rows:
        if r["wrong_by_color"].get("yellow") or r["wrong_by_color"].get("green"):
            print("SAI KHÔNG ĐỎ:", r["name"], r["wrong_by_color"])
    expected = {"rot90": 270, "rot180": 180}
    print("Hướng sai:", [(r["name"], r["rotation"]) for r in rows if r["rotation"] != expected.get(r["variant"], 0)])
    for name, group in (("xoay 90, 180", [r for r in rows if r["variant"] in expected]), ("không xoay", [r for r in rows if r["variant"] not in expected])):
        print(f"Chuẩn bị ảnh, {name}: trung vị {statistics.median(r['prepare_seconds'] for r in group):.1f} giây, tối đa {max(r['prepare_seconds'] for r in group):.1f} giây, số lần đọc xét hướng {Counter(r['orientation_tries'] for r in group)}")
    (OUT / "eval_images_q90.json").write_text(json.dumps({**result, "cases": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
    return result


def _answers_changed(base: Path, manifest: dict) -> int:
    """Số đáp án không khớp manifest, chấp nhận cả mã băm trên file gốc lẫn trên xuống dòng LF."""
    changed = 0
    for name, digest in manifest.items():
        data = (base / name / "answer.json").read_bytes()
        changed += digest not in (hashlib.sha256(data).hexdigest(), hashlib.sha256(data.replace(b"\r\n", b"\n")).hexdigest())
    return changed


def measure_external(root: Path, out: Path, mode: str, profile_name: str | None = None, confirm_network: bool = False) -> int:
    """Bộ độc lập do người khác giữ (Q96): `<root>/images/<ca>/{page.jpg,template.xlsx,answer.json}`, `<root>/images/manifest.json`.

    Mọi ca đều được đo, không lọc theo tên. Cache và kết quả chỉ nằm trong `out`, không nằm trong vmai/out.
    Màn hình chỉ in số tổng theo bố cục, loại ảnh và mã lý do, để người giữ bộ gửi đi được mà không lộ ca nào.
    Trả về mã thoát: 1 nếu có số sai bị xanh.
    Có `profile_name`: tiêu đề chưa ghép được thì hỏi model (Q85), chỉ gửi tiêu đề, ghi số lần gọi, token, chi phí theo bảng giá.
    Không có `confirm_network` thì chỉ đếm những gì SẼ gửi, không gọi mạng.
    """
    base = root.resolve() / "images"
    out = out.resolve()
    if out == OUT.resolve() or OUT.resolve() in out.parents:
        raise SystemExit("Bộ ngoài không được ghi vào vmai/out")
    directories = [d for d in sorted(base.iterdir()) if (d / "answer.json").exists()]
    cache = out / "ocr_cache"
    if mode == "read":
        read(directories, cache, quiet=True)
        print(f"Đã đọc {len(directories)} ảnh. Cache ở thư mục --out.")
        return 0
    manifest = json.loads((base / "manifest.json").read_text(encoding="utf-8"))
    provider, profile, would_send = None, None, 0
    if profile_name:
        from _env import load_key
        from vmai.llm.factory import create_provider, load_profile
        from vmai.llm.keys import read_api_key

        profile = load_profile(profile_name)
        if confirm_network:
            if profile.api_key_env:  # profile giả không cần khóa
                load_key(profile.api_key_env)
                read_api_key(profile)
            provider = create_provider(profile)

    def ask_model(plan):
        nonlocal would_send
        from vmai.llm_tasks import map_headers, sendable_headers

        headers = sendable_headers(list(plan.llm_headers))
        would_send += len(headers)
        if provider is None or not headers:
            return None, ()
        mapping, call = map_headers(provider, headers, plan.llm_fields)
        return mapping, (call,)

    rows = []
    for d in directories:
        if not (cache / f"{d.name}.json").exists():
            continue
        answer = json.loads((d / "answer.json").read_text(encoding="utf-8"))
        data = json.loads((cache / f"{d.name}.json").read_text(encoding="utf-8"))
        row = score_case(d, cache, ask_model if profile_name else None)
        row.update(name=d.name, layout=str(answer.get("layout", "(không ghi)")), variant=str(row["variant"] or "(không ghi)"), rotation=data["rotation"])
        rows.append(row)
    by_layout, by_variant = defaultdict(list), defaultdict(list)
    for r in rows:
        by_layout[r["layout"]].append(r)
        by_variant[r["variant"]].append(r)
    result = {"layout": _print_table("Theo bố cục", by_layout), "variant": _print_table("Theo loại ảnh", by_variant), "total": _summary(rows)}
    result["questions"] = dict(Counter(q for r in rows for q in r["questions"]))
    result["errors"] = dict(Counter(r["error"] for r in rows if r["error"]))
    result["answers_changed"] = _answers_changed(base, manifest)
    result["images_in_manifest"] = len(manifest)
    print("\nTổng:", result["total"])
    print("Mã lý do hỏi lại:", result["questions"], "| lỗi:", result["errors"])
    print(f"Ảnh đã chấm: {len(rows)}/{len(manifest)}. Đáp án đổi so với manifest: {result['answers_changed']}.")
    if profile_name:
        result["model"] = {"profile": profile.name, "model": profile.model, "confirmed_network": bool(provider), "headers_would_send": would_send, "calls": len(getattr(provider, "calls", [])), "tokens": getattr(provider, "tokens_used", 0), "usd_by_price_table": getattr(provider, "usd_used", 0.0)}
        print("Model:", result["model"], "(chi phí tính theo bảng giá, không phải hóa đơn)")
    out.mkdir(parents=True, exist_ok=True)
    (out / "eval_external.json").write_text(json.dumps({**result, "cases": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
    return 1 if result["total"]["wrong_green"] else 0


def score_old() -> None:
    rows = {}
    phi = defaultdict(dict)
    for d in old_dirs():
        if not (CACHE / f"{d.name}.json").exists():
            continue
        answer = json.loads((d / "answer.json").read_text(encoding="utf-8"))
        data = json.loads((CACHE / f"{d.name}.json").read_text(encoding="utf-8"))
        if answer["case"] in ("A09", "A09neg"):
            found = {}
            for key in ("first", "second"):
                found[key] = bool(scan_tokens(Token(text=t["text"], x=t["bbox"][0], y=t["bbox"][1]) for t in data[key]))
            phi[answer["case"]][answer["variant"]] = found
            continue
        answer.setdefault("variant", answer["case"])
        row = score_case(d, CACHE)
        rows[d.name] = {k: row[k] for k in ("status", "questions", "error", "gold", "colors", "wrong_by_color", "missed")} | {"expected_errors": answer.get("expected_errors"), "expected_notices": answer.get("expected_notices"), "notices": row.get("notices"), "rotation": data["rotation"]}
        print(d.name, rows[d.name])
    for case, by_variant in phi.items():
        either = sum(v["first"] or v["second"] for v in by_variant.values())
        print(case, f"có dò thấy ở ít nhất một engine: {either}/{len(by_variant)}", {k: (v["first"], v["second"]) for k, v in by_variant.items()})
    (OUT / "eval_images_old_sets.json").write_text(json.dumps({"cases": rows, "phi": phi}, ensure_ascii=False, indent=1), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=("read", "score"))
    parser.add_argument("--set", type=Path, help="Bộ ảnh ngoài repo (Q96)")
    parser.add_argument("--out", type=Path, help="Nơi ghi cache và kết quả của bộ ngoài, bắt buộc khi có --set")
    parser.add_argument("--profile", help="Profile model cho bộ ngoài (Q85). Không có --confirm-network thì không gọi mạng")
    parser.add_argument("--confirm-network", action="store_true", help="Cho phép gọi dịch vụ ngoài, tốn phí. Chỉ gửi tiêu đề của dữ liệu giả")
    args = parser.parse_args()
    if args.set is not None:
        if args.out is None:
            parser.error("--set cần --out")
        sys.exit(measure_external(args.set, args.out, args.mode, args.profile, args.confirm_network))
    if args.mode == "read":
        read(old_dirs(), CACHE)
        read(blind_dirs(), BLIND_CACHE)
    else:
        score_old()
        if score_blind()["total"]["wrong_green"]:
            sys.exit(1)


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
