"""Chấm vmai trên fixtures/eval/generated/xlsx và trên fixture của Dev. In bảng, lưu vmai/out/eval_xlsx.json.

Không gọi mạng. Provider giả chỉ kiểm dữ liệu đi ra, không phải kết quả chất lượng.
"""
from __future__ import annotations

import json
import sys
import tempfile
from collections import Counter
from decimal import Decimal, InvalidOperation
from pathlib import Path

import vmai
from vmai import MissingFields, SourceFileError, load_spec, plan_headers, propose, read_form, read_source
from vmai.profile import TemplateProfile

ROOT = Path(__file__).resolve().parents[2]
GENERATED = ROOT / "fixtures" / "eval" / "generated"
CASE_DIRS = (GENERATED / "xlsx", GENERATED / "xlsx_lat1")
TEMPLATES = ROOT / "fixtures" / "templates"
SPECS = Path(vmai.__file__).parent / "specs"
OUT = ROOT / "vmai" / "out"
SPEC = load_spec(SPECS / "GIA-QC-01.yaml")


def load_profile(answer: dict, fingerprint: str) -> TemplateProfile | None:
    raw = answer.get("profile")
    if not raw:
        return None
    return TemplateProfile(
        profile_id=raw["profile_id"],
        version=raw["version"],
        structure_fingerprint=fingerprint,
        source_headers=tuple(raw["source_headers"]),
        column_map=raw["column_map"],
        dim_maps={"code": raw["test_map"], "level": raw["level_map"]},
        confirmed_by=raw["confirmed_by"],
        confirmed_at=raw["confirmed_at"],
    )


def _same(a: str | None, b: str | None) -> bool:
    if a is None or b is None:
        return False
    try:
        return Decimal(a) == Decimal(b)
    except InvalidOperation:
        return a == b


def score(batch, answer: dict) -> dict:
    # Đáp án cũ có cột Người chạy (H). Theo quy tắc mới AI không điền cột này, nên đề xuất ở H tính là sai.
    gold = {c["a1"]: c for c in answer.get("cells", []) if not c["a1"].startswith("H")}
    by_color: Counter[str] = Counter()
    errors: Counter[str] = Counter()
    mismatches, samples = [], []
    proposed = set()
    for change in batch.changes:
        a1, color = change.target.a1_at_write, change.confidence.final_color.value
        proposed.add(a1)
        by_color[color] += 1
        expected = gold.get(a1)
        value = change.value.normalized if change.value else None
        if expected is None:
            correct = False
        elif expected.get("accept_values") and value is None:
            # X04 (D10): một Change đỏ, không tự chọn; đúng khi mọi số khả dĩ đều có trong evidence để người dùng chọn.
            shown = [e.normalized for e in change.evidence]
            correct = color == "red" and all(any(_same(s, v) for s in shown) for v in expected["accept_values"])
        elif expected.get("accept_values"):
            correct = any(_same(value, v) for v in expected["accept_values"])
        elif expected.get("value") is None:
            correct = color == "red"  # đáp án là không có giá trị đúng để tự điền; chỉ đỏ là chấp nhận được
        else:
            correct = _same(value, expected["value"])
        if not correct:
            errors[color] += 1
            samples.append({"a1": a1, "got": value, "want": expected.get("value") if expected else None, "color": color})
        if expected is not None and color != expected.get("expected_color"):
            mismatches.append({"a1": a1, "got": color, "want": expected.get("expected_color")})
    notices = {n.reason for n in batch.notices}
    return {
        "case": answer["case"],
        "proposals": len(batch.changes),
        "missed": len(set(gold) - proposed),
        "errors": sum(errors.values()),
        "by_color": dict(by_color),
        "errors_by_color": dict(errors),
        "color_mismatches": mismatches,
        "samples": samples[:5],
        "missing_notices": [n for n in answer.get("expected_notices", []) if n not in notices],
        "no_proposal_ok": (not answer.get("expect_no_proposal")) or not batch.changes,
    }


def run_case(directory: Path, **options) -> dict:
    answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
    form = read_form(directory / "template.xlsx", SPEC)
    source = read_source(directory / "source.xlsx")
    plan = plan_headers(form, source, load_profile(answer, form.fingerprint))
    if answer.get("expected_error"):
        try:
            propose(form, source, plan)
        except MissingFields as exc:
            return {"case": answer["case"], "proposals": 0, "missed": 0, "errors": 0, "by_color": {}, "errors_by_color": {}, "color_mismatches": [], "samples": [], "missing_notices": [] if exc.code == answer["expected_error"] else [answer["expected_error"]], "no_proposal_ok": True}
    return score(propose(form, source, plan, **options), answer)


def q30_matrix() -> list[dict]:
    """Q30: so giữ Q17 và nới Q17 trên các biến thể đảo hai cột khác đơn vị, miền chồng."""
    rows = []
    for directory in sorted((GENERATED / "xlsx_lat1").glob("X20_*")):
        for label, options in (("Giữ Q17", {}), ("Nới Q17", {"measure_relax_units": True})):
            r = run_case(directory, **options)
            rows.append({"case": directory.name, "config": label, "errors": r["errors"], "by_color": r["by_color"], "errors_by_color": r["errors_by_color"]})
    return rows


def summarize(results: list[dict]) -> dict:
    totals, errors = Counter(), Counter()
    for r in results:
        totals.update(r["by_color"])
        errors.update(r["errors_by_color"])
    green = totals["green"]
    return {
        "proposals": sum(totals.values()),
        "green": green,
        "yellow": totals["yellow"],
        "red": totals["red"],
        "false_green": errors["green"],
        # Quy tắc ba: khi 0 lỗi, cận trên độ tin cậy 95% xấp xỉ 3/N.
        "false_green_upper_95": (3 / green) if green and not errors["green"] else None,
        "errors_yellow": errors["yellow"],
        "errors_red": errors["red"],
    }


def dev_fixtures() -> dict:
    """Chạy trên file do Excel thật lưu của Dev. Có file thì chạy, không có thì báo."""
    out: dict = {}
    if not TEMPLATES.exists():
        return {"skipped": "không có fixtures/templates"}
    form = read_form(TEMPLATES / "GIA-QC-01_mau.xlsx", SPEC)
    source = read_source(TEMPLATES / "GIA-NGUON_2026-09.xlsx")
    batch = propose(form, source, plan_headers(form, source))
    colors = Counter(c.confidence.final_color.value for c in batch.changes)
    out["GIA-QC-01_mau + GIA-NGUON"] = {"proposals": len(batch.changes), "by_color": dict(colors), "notices": sorted({n.reason for n in batch.notices})}
    for name, code in (("GIA-X19_co_mat_khau.xlsx", "FILE_PASSWORD"), ("GIA-QC-01_da_dien_doi_cu.xls", "FILE_LEGACY_XLS"), ("GIA-X19_hong.xlsx", "FILE_CORRUPT")):
        try:
            read_source(TEMPLATES / name)
            out[name] = {"expected": code, "got": "NO_ERROR", "ok": False}
        except SourceFileError as exc:
            out[name] = {"expected": code, "got": exc.code, "ok": exc.code == code}
    macro = read_source(TEMPLATES / "GIA-X19_khong_co_vba.xlsm")
    out["GIA-X19_khong_co_vba.xlsm"] = {"has_macros_flag": macro.has_macros}

    kho = load_spec(SPECS / "GIA-KHO-02.yaml")
    form = read_form(TEMPLATES / "GIA-KHO-02_tieu_hao.xlsx", kho)
    out["GIA-KHO-02 mô tả"] = {"rows": len(form.rows), "columns": [(c.letter, list(c.col_path)) for c in form.columns], "existing": len(form.existing)}
    return out


def other_machine_source(path: Path, bias: float = 0.01, seed: int = 7) -> Path:
    """Q118, T-AI-13 (b): nguồn X01_m0 nhưng là số của máy khác, cùng tên cột, cùng tháng. [MẪU GIẢ]"""
    import random

    import openpyxl

    wb = openpyxl.load_workbook(CASE_DIRS[0] / "X01_m0" / "source.xlsx")
    ws = wb.worksheets[0]
    col = [c.value for c in ws[1]].index("Kết quả") + 1
    rng = random.Random(seed)
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)):
            nv = round(v * (1 + bias * rng.choice((-1, 1)) * rng.uniform(0.5, 1.5)), 2 if v < 100 else 0)
            ws.cell(r, col, nv if nv != v else v + (0.01 if v < 100 else 1))
    wb.save(path)
    return path


# Q131: rủi ro Owner chấp nhận. Xanh = chép đúng số từ file người dùng đã chọn; chọn nhầm file máy khác không ghi mã máy thì
# số sai vẫn xanh. Ghim đúng con số và phạm vi: đổi số hoặc file có dấu hiệu mã máy thì báo lỗi, để rủi ro không lan lặng lẽ.
Q131_ACCEPTED = {"proposals": 180, "wrong_green": 180}
_MACHINE_WORDS = ("máy", "may ", "thiết bị", "thiet bi", "instrument", "analyzer", "analyser", "device")


def machine_marks(path: Path) -> list[str]:
    """Chữ trong file nguồn trông như ghi mã máy: tên sheet hoặc ô có các từ trong `_MACHINE_WORDS`. Chỉ để ghim phạm vi Q131."""
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True)
    found = [ws.title for ws in wb if any(w in ws.title.lower() for w in _MACHINE_WORDS)]
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            found += [v for v in row if isinstance(v, str) and any(w in v.lower() for w in _MACHINE_WORDS)]
    wb.close()
    return found


def known_leak_q118(tmp: Path) -> dict:
    """Rủi ro Owner chấp nhận (Q131), trước là ca lọt Q118: biểu mẫu khai sẵn, hồ sơ đã xác nhận, nạp nhầm file máy khác
    không ghi mã máy. Không tính vào tổng sai xanh, cũng không tính là đạt."""
    directory = CASE_DIRS[0] / "X01_m0"
    answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
    gold = {c["a1"]: c["value"] for c in answer["cells"] if not c["a1"].startswith("H")}
    form = read_form(directory / "template.xlsx", SPEC)
    path = other_machine_source(tmp / "x01_khac.xlsx")
    source = read_source(path)
    batch = propose(form, source, plan_headers(form, source, load_profile(answer, form.fingerprint)))
    wrong_green = sum(
        1 for c in batch.changes if c.confidence.final_color.value == "green" and not _same(c.value.normalized, str(gold.get(c.target.a1_at_write)))
    )
    return {"case": "Q118_file_may_khac", "proposals": len(batch.changes), "wrong_green": wrong_green, "machine_marks": machine_marks(path)}


def _qc_values(path: Path) -> dict[tuple, float]:
    import openpyxl

    ws = openpyxl.load_workbook(path).worksheets[0]
    head = [c.value for c in ws[1]]
    i_day, i_code, i_level, i_value = (head.index(h) for h in ("Ngày chạy", "Mã XN", "Mức", "Kết quả"))
    return {(r[i_day], r[i_code], r[i_level]): r[i_value] for r in ws.iter_rows(min_row=2, values_only=True) if r[i_value] is not None}


def _write_qc(path: Path, values: dict[tuple, float]) -> Path:
    """Chép nguồn X01_m0, thay cột Kết quả bằng `values`. Giữ ngày tháng này. Dòng không có trong `values` bị bỏ (lô ngắn)."""
    import openpyxl

    wb = openpyxl.load_workbook(CASE_DIRS[0] / "X01_m0" / "source.xlsx")
    ws = wb.worksheets[0]
    head = [c.value for c in ws[1]]
    i_day, i_code, i_level, i_value = (head.index(h) for h in ("Ngày chạy", "Mã XN", "Mức", "Kết quả"))
    kept = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        key = (row[i_day], row[i_code], row[i_level])
        if key in values:
            row = list(row)
            row[i_value] = values[key]
            kept.append(row)
    ws.delete_rows(2, ws.max_row)
    for row in kept:
        ws.append(row)
    wb.save(path)
    return path


def stale_month_cases(tmp: Path) -> list[dict]:
    """Q119, T-AI-14: bẫy nạp lại file tháng trước, và nguồn đúng có nhiều số trùng tự nhiên. [MẪU GIẢ]

    Tháng trước = số đã duyệt của X01_m0. Tháng này đúng = X01_m1 (cùng bố cục, số khác).
    """
    import random

    m0, m1 = CASE_DIRS[0] / "X01_m0", CASE_DIRS[0] / "X01_m1"
    answer = json.loads((m0 / "answer.json").read_text(encoding="utf-8"))
    form = read_form(m0 / "template.xlsx", SPEC)
    profile = load_profile(answer, form.fingerprint)
    old, new = _qc_values(m0 / "source.xlsx"), _qc_values(m1 / "source.xlsx")
    days = sorted({k[0] for k in old})
    limits = {(c, l): (lim.target, lim.sd) for col in form.columns if col.limit for c, l in [(col.dims["code"], "L" + col.dims["level"])] for lim in [col.limit]}

    def run(name: str, values: dict, last: dict, expect_alarm: bool) -> dict:
        source = read_source(_write_qc(tmp / f"{name}.xlsx", values))
        plan = plan_headers(form, source, profile)
        base = propose(form, source, plan)
        last_month = {c.target.a1_at_write: c.value.normalized for c in propose(form, read_source(_write_qc(tmp / f"{name}_truoc.xlsx", last)), plan).changes}
        batch = propose(form, source, plan, previous_month=last_month)
        alarm = "SOURCE_LOOKS_LIKE_LAST_MONTH" in {n.reason for n in batch.notices}
        cells_same = sum(1 for c in batch.changes if last_month.get(c.target.a1_at_write) == c.value.normalized) / max(len(batch.changes), 1)
        # Đo độc lập với code quy tắc: tỷ lệ ngày mà mọi cột không cố định (một giá trị < 80% tháng trước) giống tháng trước.
        by_col: dict[str, list[str]] = {}
        for a1, v in last_month.items():
            by_col.setdefault(a1.rstrip("0123456789"), []).append(v)
        moving = {k for k, vs in by_col.items() if Counter(vs).most_common(1)[0][1] < 0.8 * len(vs)}
        day_rows: dict[str, list[bool]] = {}
        for c in batch.changes:
            letter = c.target.a1_at_write.rstrip("0123456789")
            if letter in moving:
                day_rows.setdefault(c.target.a1_at_write[len(letter):], []).append(last_month.get(c.target.a1_at_write) == c.value.normalized)
        rows_same = sum(all(v) for v in day_rows.values()) / max(len(day_rows), 1)
        colors, base_colors = Counter(c.confidence.final_color.value for c in batch.changes), Counter(c.confidence.final_color.value for c in base.changes)
        # Q121, kiểm độc lập với code: ô nào thuộc ngày trùng cả dòng thì không được xanh; ngoài các ngày đó, màu giữ như khi không kiểm
        # (trừ khi cả lô bị báo). "stale_green" là ô xanh nằm ở ngày trùng: phải bằng 0 ở mọi ca.
        same_rows = {r for r, v in day_rows.items() if v and all(v)}
        row_of = lambda a1: a1[len(a1.rstrip("0123456789")):]  # noqa: E731
        base_color = {c.target.a1_at_write: c.confidence.final_color.value for c in base.changes}
        stale_green = sum(1 for c in batch.changes if row_of(c.target.a1_at_write) in same_rows and c.confidence.final_color.value == "green")
        outside_changed = sum(1 for c in batch.changes if row_of(c.target.a1_at_write) not in same_rows and c.confidence.final_color.value != base_color[c.target.a1_at_write])
        row_yellow = sum(1 for c in batch.changes if row_of(c.target.a1_at_write) in same_rows and base_color[c.target.a1_at_write] == "green")
        ok = alarm == expect_alarm and stale_green == 0 and (colors.get("green", 0) == 0 if expect_alarm else outside_changed == 0)
        return {"case": name, "expect_alarm": expect_alarm, "alarm": alarm, "cells_same_share": round(cells_same, 2), "rows_same_share": round(rows_same, 2),
                "moving_columns": len(moving), "same_rows": len(same_rows), "row_yellow_cells": row_yellow, "stale_green_cells": stale_green,
                "colors": dict(colors), "without_check": dict(base_colors), "ok": ok}

    rows = []
    for edited in (0, 10, 20):
        reload = dict(old)
        for d in days[:edited]:
            for key in [k for k in old if k[0] == d]:
                reload[key] = new[key]
        rows.append(run(f"nap_lai_thang_truoc_sua_{edited}_dong", reload, old, True))
    rows.append(run("nguon_dung_thang_moi", new, old, False))
    rng = random.Random(119)

    # Trùng tự nhiên (1): 4 cột cố định, 2 cột GLU ít biến động (giá trị hay gặp nhất khoảng 75% số ngày, dưới mức "cố định" 80%).
    # Nhiều dòng trùng cả dòng (trên 30%), nhưng chỉ còn 2 cột thay đổi nên kiểm phải tự tắt. Ca này giữ điều kiện "≤ 2 cột".
    def fixed(vals: dict) -> dict:
        out = {}
        for k in vals:
            target, sd = limits[(k[1], k[2])]
            out[k] = round(target + (sd if rng.random() < 0.25 else 0), 2) if k[1] == "GLU" else target
        return out

    rows.append(run("trung_tu_nhien_4_cot_co_dinh", fixed(new), fixed(old), False))
    # Trùng tự nhiên (2): số làm tròn theo bậc SD, mỗi ô chỉ còn vài giá trị, nhiều ô trùng tháng trước.

    def coarse(vals: dict) -> dict:
        out = {}
        for k, v in vals.items():
            target, sd = limits[(k[1], k[2])]
            step = rng.choice((-1, 0, 0, 1))
            out[k] = round(target + step * sd, 2)
        return out

    rows.append(run("trung_tu_nhien_lam_tron_theo_sd", coarse(new), coarse(old), False))

    # Trùng tự nhiên (3): 3 cột ít biến động (mỗi ô chỉ 2 giá trị), 3 cột cố định. Có vài ngày trùng cả dòng nhưng dưới 30%.
    # Ca này để ngưỡng quá thấp (ví dụ 0) phải hỏng test.
    def few(vals: dict) -> dict:
        out = {}
        for k, v in vals.items():
            target, sd = limits[(k[1], k[2])]
            out[k] = round(target + rng.choice((0, 0, 1)) * sd, 2) if (k[1], k[2]) in (("GLU", "L1"), ("GLU", "L2"), ("CHOL", "L1")) else target
        return out

    rows.append(run("trung_tu_nhien_3_cot_it_bien_dong", few(new), few(old), False))

    # Q120: lô 3 ngày. Tháng trước vẫn đủ 30 ngày; nguồn tháng này chỉ có 3 ngày đầu.
    first3 = set(days[:3])
    pick = lambda vals: {k: v for k, v in vals.items() if k[0] in first3}  # noqa: E731
    rows.append(run("lo_3_ngay_nguon_dung", pick(new), old, False))
    one_same = pick(new)
    for key in [k for k in old if k[0] == days[0]]:
        one_same[key] = old[key]  # 1/3 ngày trùng cả dòng: 33% nhưng chưa đủ 2 ngày, không báo (Q120)
    rows.append(run("lo_3_ngay_dung_1_ngay_trung", one_same, old, False))
    rows.append(run("lo_3_ngay_file_cu_nguyen", pick(old), old, True))
    edited = pick(old)
    for key in [k for k in new if k[0] in set(days[1:3])]:
        edited[key] = new[key]
    # Q121: trước đây sót (6 ô xanh mang số tháng trước). Nay dòng trùng tự vàng, không cần ngưỡng.
    rows.append(run("lo_3_ngay_file_cu_sua_2_ngay", edited, old, False))
    return rows


def main() -> None:
    OUT.mkdir(exist_ok=True)
    results = [run_case(d) for base in CASE_DIRS for d in sorted(base.iterdir()) if d.is_dir() and d.name != "X19"]
    print(f"{'case':30}{'đề xuất':>8}{'sai':>5}{'xanh':>6}{'vàng':>6}{'đỏ':>5}{'sai-xanh':>9}{'lệch màu':>9}  notice thiếu")
    for r in results:
        by = r["by_color"]
        print(
            f"{r['case']:30}{r['proposals']:>8}{r['errors']:>5}{by.get('green', 0):>6}{by.get('yellow', 0):>6}{by.get('red', 0):>5}"
            f"{r['errors_by_color'].get('green', 0):>9}{len(r['color_mismatches']):>9}  {','.join(r['missing_notices'])}"
        )
    regular = [r for r in results if not r["case"].startswith("X20")]
    wrong_profile = [r for r in results if r["case"].startswith("X20")]
    summary = {"regular": summarize(regular), "x20": summarize(wrong_profile)}
    print("\nKhông tính X20:", json.dumps(summary["regular"], ensure_ascii=False))
    print("Chỉ X20:", json.dumps(summary["x20"], ensure_ascii=False))
    matrix = q30_matrix()
    print("\nQ30:")
    for row in matrix:
        e = row["errors_by_color"]
        print(f"{row['case']:32}{row['config']:10} sai={row['errors']:>4} sai-xanh={e.get('green', 0):>4} sai-vàng={e.get('yellow', 0):>4} sai-đỏ={e.get('red', 0):>4}")
    fixtures = dev_fixtures()
    print("\nFixture của Dev:", json.dumps(fixtures, ensure_ascii=False))
    (OUT / "eval_xlsx.json").write_text(json.dumps({"cases": results, "summary": summary, "q30": matrix, "dev_fixtures": fixtures}, ensure_ascii=False, indent=1), encoding="utf-8")
    bad = [r["case"] for r in regular if r["color_mismatches"] or r["missing_notices"] or not r["no_proposal_ok"] or r["errors"]]
    bad += [r["case"] for r in wrong_profile if r["errors_by_color"].get("green", 0)]
    with tempfile.TemporaryDirectory() as tmp:
        stale = stale_month_cases(Path(tmp))
    print("\nNghi nạp nhầm file tháng trước (Q119-Q121; dòng trùng luôn vàng; ngưỡng cả lô đặt trên phân bố tự nghĩ):")
    for r in stale:
        print(f"  {r['case']:36} báo={str(r['alarm']):5} mong={str(r['expect_alarm']):5} ô trùng={r['cells_same_share']:.0%} dòng vàng={r['row_yellow_cells']:>3} ô xanh ở ngày trùng={r['stale_green_cells']} màu={r['colors']} {'' if r['ok'] else 'SAI'}")
    bad += [r["case"] for r in stale if not r["ok"]]
    with tempfile.TemporaryDirectory() as tmp:
        leak = known_leak_q118(Path(tmp))
    # Q131: in riêng, không gộp vào số 0 sai xanh, không tính là đạt.
    print(f"\nRỦI RO OWNER CHẤP NHẬN, Q131: {leak['wrong_green']}/{leak['proposals']} ô xanh sai khi nạp nhầm file máy khác không ghi mã máy. Không tính vào số sai xanh.")
    if {k: leak[k] for k in Q131_ACCEPTED} != Q131_ACCEPTED:
        # Tăng là có đường mới; giảm thì phải ghi lại quyết định, sửa Q131_ACCEPTED và xfail trong tests/test_known_leaks.py.
        bad.append(f"Rủi ro Q131 đổi: {leak['wrong_green']}/{leak['proposals']}, đã ghim {Q131_ACCEPTED['wrong_green']}/{Q131_ACCEPTED['proposals']}")
    if leak["machine_marks"]:
        bad.append(f"Ca rủi ro Q131 có dấu hiệu mã máy, không còn đúng phạm vi Owner chấp nhận: {sorted(set(leak['machine_marks']))[:3]}")
    outside = sum(r["errors_by_color"].get("green", 0) for r in results)
    print(f"Tổng: {outside} sai xanh ngoài rủi ro đã chấp nhận; 1 rủi ro Owner chấp nhận (Q131)")
    print("\nCó vấn đề:", bad or "không")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
