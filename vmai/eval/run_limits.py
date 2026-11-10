"""Ca giới hạn của `vmai` (T-AI-08): lô lớn, tiêu đề lạ, nhiều nguồn. [MẪU GIẢ] Mọi số đều giả.

Chỉ chạy trong thư viện, chưa nối vào app. Không gọi mạng, không gọi model, không sửa logic.
Tự tạo biểu mẫu và nguồn giả trong vmai/out/limits/ (gitignore), in bảng, lưu vmai/out/eval_limits.json.

    .venv/Scripts/python.exe eval/run_limits.py            # đủ cỡ: 3, 50, 200, 500 xét nghiệm
    .venv/Scripts/python.exe eval/run_limits.py --quick    # bỏ cỡ 500

Thoát mã 1 nếu có số sai bị xanh ở bất kỳ ca nào.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import random
import sys
import time
import tracemalloc
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal, InvalidOperation
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from vmai.profile import TemplateProfile

from vmai import MissingFields, SourceFileError, load_spec, plan_headers, propose, read_form, read_source

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "vmai" / "out"
WORK = OUT / "limits"
START = dt.date(2026, 9, 1)
DAYS = 30
HEADERS = ("Ngày chạy", "Mã XN", "Mức", "Kết quả", "Đơn vị", "Lô")


def tests(n: int, separated: bool = False) -> list[dict]:
    """separated=False: đích ngẫu nhiên 2-200, miền ±3SD chồng nhau nhiều (giống biểu mẫu thật nhiều xét nghiệm).
    separated=True: mỗi cột một miền riêng, không chồng, để Q17 cho phép xanh và thử bẫy trên ô có thể xanh."""
    rng = random.Random(1000 + n)
    out = []
    for i in range(n):
        if separated:
            out.append({"code": f"T{i:03d}", "name": f"Xét nghiệm giả {i:03d}", "unit": "mmol/L", "levels": {"1": (100.0 * (2 * i + 1), 2.0), "2": (100.0 * (2 * i + 2), 2.0)}})
            continue
        target = round(rng.uniform(2, 200), 1)
        out.append({"code": f"T{i:03d}", "name": f"Xét nghiệm giả {i:03d}", "unit": "mmol/L", "levels": {"1": (target, round(target * 0.03, 2)), "2": (round(target * 2.5, 1), round(target * 0.06, 2))}})
    return out


def spec_for(ts: list[dict], form_cols: dict[tuple[str, str], str]) -> dict:
    return {
        "spec_id": f"GIA-LIM-{len(ts)}",
        "version": 1,
        "table": {
            "sheet": "KQ",
            "anchor": "A6",
            "header_rows": [6, 7],
            "first_row": 8,
            "last_row": 8 + DAYS - 1,
            "key_columns": [{"column": "A", "dim": "date", "kind": "date_sequence", "start_cell": "B4"}],
            "value_columns": [{"column": col, "field": "result", "dims": {"code": code, "level": lvl}} for (code, lvl), col in form_cols.items()],
            "manual_columns": [],
        },
        "date_fields": ["date"],
        "period": {"cell": "B4", "field": "date"},
        "unit_field": "unit",
        "lot_field": "lot",
        "source_fields": {
            "date": {"names": ["Ngày chạy"], "aliases": ["Ngày"], "required": True},
            "code": {"names": ["Mã XN"], "aliases": ["Xét nghiệm"], "required": True},
            "level": {"names": ["Mức"], "aliases": ["Level"], "required": True},
            "result": {"names": ["Kết quả"], "aliases": ["KQ", "Giá trị"], "required": True},
            "unit": {"names": ["Đơn vị"], "aliases": ["Unit"]},
            "lot": {"names": ["Lô"], "aliases": ["Số lô"]},
        },
        "dim_values": {"code": {t["code"]: [t["code"]] for t in ts}, "level": {"1": ["L1", "1"], "2": ["L2", "2"]}},
        "limits": {"sheet": "GioiHan", "first_row": 2, "dims": {"code": "A", "level": "C"}, "target": "E", "sd": "F", "unit": "K", "lot": "D"},
    }


def make_form(ts: list[dict], path: Path) -> dict:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KQ"
    ws["A1"] = "[MẪU GIẢ] BIỂU MẪU GIỚI HẠN"
    ws["A4"], ws["B4"] = "Tháng:", dt.datetime(2026, 9, 1)
    ws["A6"] = "Ngày"
    cols, c = {}, 2
    for t in ts:
        for lvl in ("1", "2"):
            letter = get_column_letter(c)
            ws[f"{letter}6"] = f"{t['name']} ({t['unit']})"
            ws[f"{letter}7"] = f"Mức {lvl}"
            cols[(t["code"], lvl)] = letter
            c += 1
    ws["A8"] = "=$B$4"
    for r in range(9, 8 + DAYS):
        ws[f"A{r}"] = f"=A{r - 1}+1"
    lim = wb.create_sheet("GioiHan")
    lim.append(["Mã", "Tên", "Mức", "Lô giả", "Đích", "SD", None, None, None, None, "Đơn vị"])
    for t in ts:
        for lvl, (target, sd) in t["levels"].items():
            lim.append([t["code"], t["name"], int(lvl), f"GIA-L{lvl}-001", target, sd, None, None, None, None, t["unit"]])
    wb.save(path)
    return cols


def make_source(ts: list[dict], path: Path, *, headers=HEADERS, seed: int = 1, trap_rate: float = 0.0, days: range = range(DAYS), shift: float = 0.0) -> tuple[dict, dict]:
    """Trả về (đáp án theo (ngày, mã, mức), loại bẫy theo khóa). Bẫy: đổi đơn vị mg/dL x18, lệch dấu phẩy x10."""
    rng = random.Random(seed)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KetQua"
    ws.append(list(headers))
    truth, traps = {}, {}
    for d in days:
        day = START + dt.timedelta(days=d)
        for t in ts:
            for lvl, (target, sd) in t["levels"].items():
                value = round(target + shift * sd + rng.uniform(-1.5, 1.5) * sd, 2)
                key = (day.isoformat(), t["code"], lvl)
                truth[key] = value
                shown, unit = value, t["unit"]
                if trap_rate and rng.random() < trap_rate:
                    if rng.random() < 0.5:
                        shown, unit, traps[key] = round(value * 18, 1), "mg/dL", "don_vi"
                    else:
                        shown, traps[key] = round(value * 10, 2), "dau_phay"
                ws.append([dt.datetime.combine(day, dt.time()), t["code"], f"L{lvl}", shown, unit, f"GIA-L{lvl}-001"])
    wb.save(path)
    return truth, traps


def score(batch, truth: dict, traps: dict, cols: dict) -> dict:
    by_letter = {v: k for k, v in cols.items()}
    colors, wrong, trap_colors = Counter(), Counter(), Counter()
    samples = []
    for ch in batch.changes:
        a1 = ch.target.a1_at_write
        letter = a1.rstrip("0123456789")
        row = int(a1[len(letter):])
        code, lvl = by_letter[letter]
        key = ((START + dt.timedelta(days=row - 8)).isoformat(), code, lvl)
        color = ch.confidence.final_color.value
        colors[color] += 1
        value = ch.value.normalized if ch.value else None
        if key in traps:
            trap_colors[f"{traps[key]}:{color}"] += 1
        try:
            same = value is not None and Decimal(value) == Decimal(str(truth[key]))
        except InvalidOperation:  # đề xuất là chữ, không phải số: tính là sai
            same = False
        if not same:
            wrong[color] += 1
            if len(samples) < 3:
                samples.append({"a1": a1, "got": value, "want": truth[key], "color": color, "trap": traps.get(key)})
    return {"proposals": len(batch.changes), "colors": dict(colors), "wrong_by_color": dict(wrong), "wrong_green": wrong.get("green", 0), "traps": len(traps), "trap_colors": dict(trap_colors), "samples": samples}


def confirmed_profile(form, ts: list[dict]) -> TemplateProfile:
    """Hồ sơ như sau khi người dùng đã xác nhận cách ghép ở tháng trước. Chỉ khi có hồ sơ mới lên được xanh."""
    return TemplateProfile(
        profile_id="gia-lim", version=1, structure_fingerprint=form.fingerprint, source_headers=HEADERS,
        column_map={"Ngày chạy": "date", "Mã XN": "code", "Mức": "level", "Kết quả": "result", "Đơn vị": "unit", "Lô": "lot"},
        dim_maps={"code": {t["code"]: t["code"] for t in ts}, "level": {"L1": "1", "L2": "2"}},
        confirmed_by="tester.gia", confirmed_at="2026-08-30T08:00:00+07:00",
    )


def timed(fn, *args, **kw):
    t = time.perf_counter()
    r = fn(*args, **kw)
    return r, round(time.perf_counter() - t, 2)


def case_large(n: int, separated: bool = False) -> dict:
    ts = tests(n, separated)
    name = f"lon_{n}" + ("_tach" if separated else "")
    form_path, src_path = WORK / f"{name}_form.xlsx", WORK / f"{name}_source.xlsx"
    cols = make_form(ts, form_path)
    truth, traps = make_source(ts, src_path, seed=n, trap_rate=0.01)
    spec = load_spec(spec_for(ts, cols))
    tracemalloc.start()
    form, t_form = timed(read_form, form_path, spec)
    source, t_src = timed(read_source, src_path)
    plan = plan_headers(form, source)
    batch, t_prop = timed(propose, form, source, plan)
    peak = tracemalloc.get_traced_memory()[1]
    tracemalloc.stop()
    first = score(batch, truth, traps, cols)
    batch2, t_prop2 = timed(propose, form, source, plan_headers(form, source, confirmed_profile(form, ts)))
    r = score(batch2, truth, traps, cols)
    r["wrong_green"] += first["wrong_green"]
    r.update(first_time_colors=first["colors"], first_time_trap_colors=first["trap_colors"], s_propose_with_profile=t_prop2)
    r.update(case=name, separated=separated, tests=n, cells=len(cols) * DAYS, source_rows=len(truth), source_mb=round(src_path.stat().st_size / 1e6, 2),
             s_read_form=t_form, s_read_source=t_src, s_propose=t_prop, peak_mb_python=round(peak / 1e6, 1), batch_json_mb=round(len(batch.model_dump_json()) / 1e6, 2))
    return r


HEADER_VARIANTS = {
    "chuan": HEADERS,
    "dong_nghia_khai_bao": ("Ngày", "Xét nghiệm", "Level", "KQ", "Unit", "Số lô"),
    "khong_dau": ("Ngay chay", "Ma XN", "Muc", "Ket qua", "Don vi", "Lo"),
    "chu_hoa_khoang_trang": (" NGÀY CHẠY ", "MÃ  XN", "mức", "KẾT QUẢ ", "Đơn vị", "Lô"),
    "tieng_anh_chua_khai": ("Run date", "Analyte", "QC level", "Value", "Units", "Lot no"),
    "ten_cot_vo_nghia": ("Cột 1", "Cột 2", "Cột 3", "Cột 4", "Cột 5", "Cột 6"),
    "doi_ten_ket_qua_sang_cot_khac": ("Ngày chạy", "Mã XN", "Mức", "Kết quả cũ", "Đơn vị", "Kết quả"),
}


def case_headers() -> list[dict]:
    ts = tests(3)
    form_path = WORK / "td_form.xlsx"
    cols = make_form(ts, form_path)
    form = read_form(form_path, load_spec(spec_for(ts, cols)))
    rows = []
    for name, headers in HEADER_VARIANTS.items():
        src = WORK / f"td_{name}.xlsx"
        truth, traps = make_source(ts, src, headers=headers, seed=7)
        source = read_source(src)
        row = {"case": f"tieu_de_{name}", "headers": list(headers), "llm_headers": []}
        try:
            plan = plan_headers(form, source)
            row["llm_headers"] = list(plan.llm_headers)
            batch = propose(form, source, plan)  # không có model: header_map None
            row.update(outcome="có đề xuất", **score(batch, truth, traps, cols))
        except (MissingFields, SourceFileError) as exc:
            row.update(outcome=f"hỏi người dùng ({exc.code})", proposals=0, wrong_green=0)
        rows.append(row)
    return rows


def case_many_sources() -> list[dict]:
    ts = tests(3)
    form_path = WORK / "nn_form.xlsx"
    cols = make_form(ts, form_path)
    spec = load_spec(spec_for(ts, cols))
    form = read_form(form_path, spec)
    rows = []

    # 1. Mười nguồn khác nhau chạy song song trong luồng, so với chạy lần lượt.
    paths = []
    for i in range(10):
        p = WORK / f"nn_song_song_{i}.xlsx"
        make_source(ts, p, seed=100 + i, trap_rate=0.05)
        paths.append(p)

    def one(p: Path) -> str:
        b = propose(form, read_source(p), plan_headers(form, read_source(p)))
        return json.dumps([(c.target.a1_at_write, c.value.normalized if c.value else None, c.confidence.final_color.value) for c in b.changes])

    seq, t_seq = timed(lambda: [one(p) for p in paths])
    with ThreadPoolExecutor(max_workers=10) as pool:
        par, t_par = timed(lambda: list(pool.map(one, paths)))
    rows.append({"case": "muoi_nguon_song_song", "sources": 10, "same_as_sequential": seq == par, "s_sequential": t_seq, "s_parallel_threads": t_par})

    # 2. Hai nguồn cho cùng một biểu mẫu, trùng ngày 11-20, khác giá trị. Thư viện xử lý từng lô riêng.
    a, b = WORK / "nn_nguon_a.xlsx", WORK / "nn_nguon_b.xlsx"
    ta, _ = make_source(ts, a, seed=1, days=range(0, 20))
    tb, _ = make_source(ts, b, seed=2, days=range(10, 30), shift=0.5)
    for with_profile in (False, True):
        prof = confirmed_profile(form, ts) if with_profile else None
        ba = propose(form, read_source(a), plan_headers(form, read_source(a), prof))
        bb = propose(form, read_source(b), plan_headers(form, read_source(b), prof))
        ca = {c.target.a1_at_write: (c.value.normalized, c.confidence.final_color.value) for c in ba.changes}
        cb = {c.target.a1_at_write: (c.value.normalized, c.confidence.final_color.value) for c in bb.changes}
        overlap = set(ca) & set(cb)
        conflict = [k for k in overlap if Decimal(ca[k][0]) != Decimal(cb[k][0])]
        both_green = [k for k in conflict if ca[k][1] == "green" and cb[k][1] == "green"]
        rows.append({"case": "hai_nguon_trung_ngay" + ("_co_ho_so" if with_profile else ""), "cells_a": len(ca), "cells_b": len(cb), "overlap": len(overlap),
                     "conflicting_values": len(conflict), "conflict_both_green": len(both_green), "notices_batch_b": sorted({n.reason for n in bb.notices}),
                     "note": "Hai lô tạo cùng lúc, chưa ghi lô nào. Mỗi lô tự đúng với nguồn của nó; vmai không biết lô kia. Chặn hai lô cùng chờ cho một ô là việc của app."})

    # 2b. Thứ tự thật: ghi lô A vào biểu mẫu, đọc lại, rồi mới tạo lô B. Ô đã có số phải duyệt từng ô (Q14, D6), kèm số cũ.
    prof = confirmed_profile(form, ts)
    ba = propose(form, read_source(a), plan_headers(form, read_source(a), prof))
    wb = openpyxl.load_workbook(form_path)
    for c in ba.changes:
        wb["KQ"][c.target.a1_at_write] = float(c.value.normalized)
    filled = WORK / "nn_form_da_ghi_a.xlsx"
    wb.save(filled)
    form_a = read_form(filled, spec)
    bb = propose(form_a, read_source(b), plan_headers(form_a, read_source(b), confirmed_profile(form_a, ts)))
    written = {c.target.a1_at_write for c in ba.changes}
    on_filled = Counter(f"{c.confidence.final_color.value}/{c.review.unit}/{'có số cũ' if c.previous.value is not None else 'không số cũ'}" for c in bb.changes if c.target.a1_at_write in written)
    rows.append({"case": "nguon_b_sau_khi_ghi_a", "cells_on_filled": sum(on_filled.values()), "filled_cells_by_color_review_previous": dict(on_filled),
                 "all_filled_reviewed_per_cell": all(k.split("/")[1] == "cell" for k in on_filled)})

    # 3. Một nguồn có hai bản ghi cùng khóa, giá trị khác nhau (dán trùng).
    dup = WORK / "nn_trung_khoa.xlsx"
    truth, _ = make_source(ts, dup, seed=3)
    wb = openpyxl.load_workbook(dup)
    ws = wb.active
    for r in list(ws.iter_rows(min_row=2, max_row=13, values_only=True)):
        r = list(r)
        r[3] = round(r[3] * 1.05, 2)
        ws.append(r)
    wb.save(dup)
    bd = propose(form, read_source(dup), plan_headers(form, read_source(dup)))
    sd = score(bd, truth, {}, cols)
    rows.append({"case": "mot_nguon_trung_khoa", "duplicated_records": 12, **sd, "notices": sorted({n.reason for n in bd.notices})})
    return rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--quick", action="store_true", help="Bỏ cỡ 500 xét nghiệm")
    args = parser.parse_args()
    WORK.mkdir(parents=True, exist_ok=True)
    sizes = (3, 50, 200) if args.quick else (3, 50, 200, 500)
    print("[MẪU GIẢ] Ca giới hạn, chỉ trong thư viện vmai, chưa nối vào app. Không gọi model.\n")
    large = [case_large(n) for n in sizes] + [case_large(n, separated=True) for n in sizes if n >= 50]
    print("Bảng dưới: lần có hồ sơ đã xác nhận. Lần đầu (chưa có hồ sơ) xem cột cuối.")
    print(f"{'ca':16}{'ô':>7}{'dòng nguồn':>11}{'MB':>6}{'đọc mẫu s':>10}{'đọc nguồn s':>12}{'propose s':>10}{'RAM MB':>8}{'xanh':>7}{'vàng':>6}{'đỏ':>6}{'sai-xanh':>9}  bẫy theo màu | lần đầu")
    for r in large:
        c = r["colors"]
        print(f"{r['case']:16}{r['cells']:>7}{r['source_rows']:>11}{r['source_mb']:>6}{r['s_read_form']:>10}{r['s_read_source']:>12}{r['s_propose']:>10}{r['peak_mb_python']:>8}"
              f"{c.get('green', 0):>7}{c.get('yellow', 0):>6}{c.get('red', 0):>6}{r['wrong_green']:>9}  {r['trap_colors']} | {r['first_time_colors']}")
    headers = case_headers()
    print("\nTiêu đề lạ (không có model, header_map rỗng):")
    for r in headers:
        print(f"  {r['case']:40} {r['outcome']:32} đề xuất={r['proposals']:>4} màu={r.get('colors', {})} sai={r.get('wrong_by_color', {})} cần model={r['llm_headers']}")
    many = case_many_sources()
    print("\nNhiều nguồn:")
    for r in many:
        print("  " + json.dumps(r, ensure_ascii=False))
    wrong_green = sum(r["wrong_green"] for r in large + headers) + sum(r.get("wrong_green", 0) for r in many)
    (OUT / "eval_limits.json").write_text(json.dumps({"date": dt.date.today().isoformat(), "large": large, "headers": headers, "many_sources": many, "wrong_green": wrong_green}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nSai bị xanh tổng: {wrong_green}. Bộ tự nghĩ, dữ liệu giả, CHƯA kiểm chứng với mẫu thật.")
    return 1 if wrong_green else 0


if __name__ == "__main__":
    import _console

    _console.utf8()
    sys.exit(main())
