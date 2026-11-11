"""Xuất lô thật làm fixture cho phần 3 (T-AI-06, T-AI-07, Q105). Không gọi model, không gọi mạng. Cờ ảnh, PDF chỉ bật trong lúc xuất.

Chạy: .venv/Scripts/python.exe eval/export_batches.py     Ra: fixtures/eval/generated/batches/

Mọi lô do `propose` sinh từ dữ liệu giả, không viết tay. Mỗi ca kèm đúng file biểu mẫu và nguồn, SHA khớp với lô, nằm trong `files/`.
`index.json` là mục lục: ca nào, file nào, phủ mã nào, chưa phủ mã nào. Xuất lại cho ra đúng các file cũ, gồm cả bản Excel dựng thêm.

Ca bắt buộc của Dev (MTG-002): X01 có hồ sơ, X01b, X04, X05, X11, một lô có ô đỏ trong cột thống kê; US-1.13 (`SpecGuess`, hồ sơ); X14, X15.
"""
from __future__ import annotations

import io
import json
import re
import shutil
import sys
import zipfile
from collections import Counter
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from run_xlsx import CASE_DIRS, SPEC, load_profile  # noqa: E402
from vmai import confirmation_rows, infer_spec, plan_headers, propose, read_form, read_source, source_options  # noqa: E402
from vmai.contract import Page  # noqa: E402
from vmai.errors import VmaiError  # noqa: E402
from vmai.infer import profile_from_guess  # noqa: E402
from vmai.pdf_table import read_pdf_table  # noqa: E402
from vmai.table_read import source_from_tokens  # noqa: E402
from vmai.template import sha256_file  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
GENERATED = ROOT / "fixtures" / "eval" / "generated"
TARGET = GENERATED / "batches"
LAYOUTS = GENERATED / "layouts"
IMAGES = GENERATED / "images"
FORM_WITH_FORMULAS = ROOT / "fixtures" / "templates" / "GIA-QC-01_mau.xlsx"  # có dòng thống kê B38-B46; biểu mẫu của eval thì không
REQUIRED = {"X01_m0": "Xanh, có hồ sơ", "X01b": "Vàng, chưa có hồ sơ", "X04": "Đỏ, nguồn có hai kết quả", "X05": "Đỏ, khác đơn vị", "X11": "Ô có sẵn số, duyệt từng ô (D6)"}
_FIXED_TIME = "2026-11-05T00:00:00+00:00"
_FIXED_ZIP = (2026, 11, 5, 0, 0, 0)
_PERSON = "tester.gia"
# Bản chụp mã từ docs/ai/confidence_rules.md (kho công khai không có docs/). test_rules kiểm bản chụp khớp tài liệu.
_DOC_CODES = json.loads((ROOT / "vmai" / "src" / "vmai" / "data" / "doc_codes.json").read_text(encoding="utf-8"))["confidence_rules_codes"]
_RULE_CODES = sorted(c for c in _DOC_CODES if re.fullmatch(r"(?:READ|MAP|CHECK|SOURCE|PROFILE|PERSON)_[A-Z0-9_]+", c))


def _rel(path: Path, base: Path) -> str:
    return path.relative_to(base).as_posix()


def _codes(batch) -> set[str]:
    return {r.code for c in batch.changes for r in c.confidence.reasons} | {n.reason for n in batch.notices}


def _freeze(batch, name: str):
    return batch.model_copy(update={"proposal_id": f"mau-{name}", "created_at": _FIXED_TIME})


def _write_json(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def _save_stable(workbook, path: Path) -> Path:
    """Lưu Excel không mang thời điểm lưu, để mã băm không đổi giữa các lần xuất."""
    path.parent.mkdir(parents=True, exist_ok=True)
    raw = io.BytesIO()
    workbook.save(raw)
    out = io.BytesIO()
    with zipfile.ZipFile(raw) as old, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as new:
        for item in old.infolist():
            data = old.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*", rb"\g<1>2026-11-05T00:00:00Z", data)
            info = zipfile.ZipInfo(item.filename, _FIXED_ZIP)
            info.compress_type = zipfile.ZIP_DEFLATED
            new.writestr(info, data)
    path.write_bytes(out.getvalue())
    return path


class Export:
    def __init__(self, target: Path):
        self.target = target
        self.files = target / "files"

    def copy(self, case: str, source: Path, name: str | None = None) -> Path:
        out = self.files / case / (name or source.name)
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, out)
        return out

    def rel(self, path: Path) -> str:
        return _rel(path, self.target)

    def entry(self, name: str, batch, template: Path, source: Path, note: str | None = None) -> dict:
        _write_json(self.target / f"{name}.json", _freeze(batch, name).model_dump_json(indent=1))
        kind = batch.sources[0].kind
        return {
            "file": f"{name}.json",
            "note": note,
            "template": self.rel(template),
            "source": self.rel(source),
            "template_sha256": batch.template.template_sha256,
            "source_file_sha256": batch.sources[0].source_file_sha256,
            "kind": getattr(kind, "value", str(kind)),
            "colors": dict(Counter(c.confidence.final_color.value for c in batch.changes)),
            "reasons": sorted(_codes(batch)),
        }


def excel_batches() -> dict[str, tuple[object, Path, Path]]:
    out = {}
    for base in CASE_DIRS:
        for directory in sorted(base.iterdir()):
            answer_path = directory / "answer.json"
            if not answer_path.exists() or not (directory / "source.xlsx").exists():
                continue
            answer = json.loads(answer_path.read_text(encoding="utf-8"))
            try:
                form, source = read_form(directory / "template.xlsx", SPEC), read_source(directory / "source.xlsx")
                batch = propose(form, source, plan_headers(form, source, load_profile(answer, form.fingerprint)))
            except VmaiError:
                continue
            out[directory.name] = (batch, directory / "template.xlsx", directory / "source.xlsx")
    return out


def image_batch(name: str, base: Path = IMAGES, cache_dir: str = "ocr_cache"):
    cache = ROOT / "vmai" / "out" / cache_dir / f"{name}.json"
    if not cache.exists():
        cache = ROOT / "fixtures" / "eval" / "ocr_tokens" / f"{name}.json"  # token chép sẵn cho test, xem make_fixtures.py
    if not cache.exists():
        return None
    data = json.loads(cache.read_text(encoding="utf-8"))
    directory = base / name
    page = directory / "page.jpg"
    source = source_from_tokens(name, sha256_file(page), [(1, data["first"], data["second"])], kind="image", first_reader=data["readers"][0], second_reader=data["readers"][1], low_quality=data["low_quality"])
    template = directory / "template.xlsx"
    guess = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), experimental_image_sources=True)
    if guess.spec is None:
        return None
    form = read_form(template, guess.spec)
    return propose(form, source, plan_headers(form, source), experimental_image_sources=True), template, page


def with_page_image(export: Export, name: str, batch, page_path: Path):
    """Ảnh trang đã chuẩn hóa cho ảnh cắt của dòng vàng (Q6). Tọa độ `bbox` của lô tính trên đúng ảnh này. Chỉ dùng ca không cần xoay."""
    import cv2

    from vmai.ocr import prepare_image

    prepared = prepare_image(page_path)
    ok, encoded = cv2.imencode(".png", prepared.rgb[:, :, ::-1])
    if not ok:
        return batch, None
    out = export.files / name / "page_1.png"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(encoded.tobytes())
    page = Page(page=1, width=prepared.width, height=prepared.height, page_image_sha256=sha256_file(out))
    source = batch.sources[0].model_copy(update={"pages": [page]})
    return batch.model_copy(update={"sources": [source, *batch.sources[1:]]}), out


def pdf_batch():
    directory = GENERATED / "pdf" / "P01"
    source = read_pdf_table(directory / "source.pdf", experimental_pdf_sources=True)
    template = directory / "template.xlsx"
    guess = infer_spec(read_source(template), source, frozenset({"LEVEL_NAME_DIFFERS"}), experimental_pdf_sources=True)
    form = read_form(template, guess.spec)
    return propose(form, source, plan_headers(form, source), experimental_pdf_sources=True), template, directory / "source.pdf"


def pick_cover(candidates: dict[str, tuple[object, Path, Path]], seen: set[str]) -> list[str]:
    """Chọn tham lam: mỗi lần lấy lô thêm được nhiều mã và màu mới nhất, ít ô nhất khi hòa."""
    remaining = dict(candidates)
    picked = []

    def marks(name: str) -> set[str]:
        batch = remaining[name][0]
        return _codes(batch) | {f"color:{c.confidence.final_color.value}" for c in batch.changes}

    while remaining:
        best = max(sorted(remaining), key=lambda n: (len(marks(n) - seen), -len(remaining[n][0].changes)))
        if not marks(best) - seen:
            break
        seen |= marks(best)
        picked.append(best)
        del remaining[best]
    return picked


def guess_cases(export: Export) -> list[dict]:
    """Ca hỏi lại, và màn xác nhận. Mỗi mã câu hỏi một ca, cộng một ca đoán được."""
    out, seen = [], set()
    for directory in sorted(LAYOUTS.iterdir()):
        if not (directory / "source.xlsx").exists():
            continue
        try:
            source, template = read_source(directory / "source.xlsx"), read_source(directory / "template.xlsx")
            guess = infer_spec(template, source)
        except VmaiError:
            continue
        codes = {q.code for q in guess.questions} or {"(đoán được)"}
        if codes <= seen:
            continue
        seen |= codes
        out.append(
            {
                "case": directory.name,
                "template": export.rel(export.copy(directory.name, directory / "template.xlsx")),
                "source": export.rel(export.copy(directory.name, directory / "source.xlsx")),
                "status": guess.status,
                "questions": [q.model_dump() for q in guess.questions],
                "open_columns": [c.model_dump() for c in guess.open_columns],
                "confirmation_rows": [r.model_dump() for r in confirmation_rows(guess, source)],
                "source_options": [o.model_dump() for o in source_options(source)],
            }
        )
    return out


def error_cases(export: Export) -> list[dict]:
    out = []
    attempts = [
        ("PDF trang quét", GENERATED / "pdf" / "P02" / "source.pdf", lambda p: read_pdf_table(p, experimental_pdf_sources=True)),
        ("PDF có chữ ẩn", GENERATED / "pdf_traps" / "T1_chu_an_tr3.pdf", lambda p: read_pdf_table(p, experimental_pdf_sources=True)),
        ("PDF khi cờ tắt", GENERATED / "pdf" / "P01" / "source.pdf", read_pdf_table),
        ("Excel có mật khẩu", ROOT / "fixtures" / "templates" / "GIA-X19_co_mat_khau.xlsx", read_source),
    ]
    for i, (label, path, call) in enumerate(attempts, start=1):
        if not path.exists():
            continue
        try:
            call(path)
        except VmaiError as exc:
            copied = export.copy(f"loi_{i}", path)
            out.append({"case": label, "file": export.rel(copied), "code": exc.code, "message": str(exc)})
    return out


def _renamed_source(export: Export) -> Path:
    """Nguồn của X01, đổi tên cột kết quả thành tên app không biết, thêm một cột sai gấp 10 để có chỗ chọn nhầm."""
    wb = openpyxl.load_workbook(GENERATED / "xlsx" / "X01_m0" / "source.xlsx")
    ws = wb.worksheets[0]
    names = ("Kết quả", "KQ", "Ket qua")
    header_row = next(r for r in range(1, 10) if any(str(ws.cell(r, c).value or "").strip() in names for c in range(1, ws.max_column + 1)))
    result = next(c for c in range(1, ws.max_column + 1) if str(ws.cell(header_row, c).value or "").strip() in names)
    ws.cell(header_row, result, "Giá trị đo")
    extra = ws.max_column + 1
    ws.cell(header_row, extra, "Giá trị khác")
    for r in range(header_row + 1, ws.max_row + 1):
        value = ws.cell(r, result).value
        if isinstance(value, (int, float)):
            ws.cell(r, extra, value * 10)
    return _save_stable(wb, export.files / "us113_nguoi_chon" / "source.xlsx")


def us113(export: Export) -> tuple[dict, dict]:
    """US-1.13: đoán, màn xác nhận, hồ sơ, lô lần đầu và lần sau. Và đường người chọn cột, gồm chọn đúng và chọn nhầm."""
    accepted = frozenset({"LEVEL_NAME_DIFFERS"})
    base = GENERATED / "xlsx" / "X01_m0"
    folder = export.target / "us113"

    def save_profile(name: str, guess, form, source) -> dict:
        plan = plan_headers(form, source)
        profile = profile_from_guess(guess, form.fingerprint, plan.headers, confirmed_by=_PERSON, confirmed_at=_FIXED_TIME, profile_id=f"mau-{name}")
        _write_json(folder / f"{name}_ho_so.json", profile.model_dump_json(indent=1))
        return profile

    batches: dict[str, dict] = {}
    template = export.copy("us113_doan", base / "template.xlsx")
    source_path = export.copy("us113_doan", base / "source.xlsx")
    facts = read_source(source_path)
    guess = infer_spec(read_source(template), facts, accepted)
    _write_json(folder / "doan_spec_guess.json", guess.model_dump_json(indent=1))
    _write_json(folder / "doan_man_xac_nhan.json", json.dumps({"rows": [r.model_dump() for r in confirmation_rows(guess, facts)], "source_options": [o.model_dump() for o in source_options(facts)]}, ensure_ascii=False, indent=1))
    form = read_form(template, guess.spec)
    batches["us113_lan_dau"] = export.entry("us113_lan_dau", propose(form, facts, plan_headers(form, facts)), template, source_path, "Lần đầu, chưa có hồ sơ: mọi ô vàng")
    profile = save_profile("doan", guess, form, facts)
    batches["us113_lan_sau"] = export.entry("us113_lan_sau", propose(form, facts, plan_headers(form, facts, profile)), template, source_path, "Có hồ sơ. Cột ghép nhờ trả lời Cùng một thứ tối đa vàng (Q114); cột khác lên xanh theo Q17")

    template = export.copy("us113_nguoi_chon", base / "template.xlsx")
    chosen_path = _renamed_source(export)
    chosen_facts = read_source(chosen_path)
    asked = infer_spec(read_source(template), chosen_facts, accepted)
    _write_json(folder / "chon_cot_hoi_lai.json", asked.model_dump_json(indent=1))
    _write_json(folder / "chon_cot_man_xac_nhan.json", json.dumps({"rows": [r.model_dump() for r in confirmation_rows(asked, chosen_facts)], "source_options": [o.model_dump() for o in source_options(chosen_facts)]}, ensure_ascii=False, indent=1))
    right = infer_spec(read_source(template), chosen_facts, accepted, column_choices={c.letter: "Giá trị đo" for c in asked.open_columns})
    _write_json(folder / "chon_cot_dung_spec_guess.json", right.model_dump_json(indent=1))
    form = read_form(template, right.spec)
    batches["nguoi_chon_cot"] = export.entry("nguoi_chon_cot", propose(form, chosen_facts, plan_headers(form, chosen_facts)), template, chosen_path, "Người chọn đúng cột, chưa có hồ sơ: tối đa vàng")
    profile = save_profile("chon_cot_dung", right, form, chosen_facts)
    batches["nguoi_chon_cot_lan_sau"] = export.entry("nguoi_chon_cot_lan_sau", propose(form, chosen_facts, plan_headers(form, chosen_facts, profile)), template, chosen_path, "Người chọn đúng cột, có hồ sơ: vẫn tối đa vàng (Q115)")
    wrong = infer_spec(read_source(template), chosen_facts, accepted, column_choices={c.letter: "Giá trị khác" for c in asked.open_columns})
    form = read_form(template, wrong.spec)
    batches["nguoi_chon_nham_cot"] = export.entry("nguoi_chon_nham_cot", propose(form, chosen_facts, plan_headers(form, chosen_facts)), template, chosen_path, "Người chọn nhầm cột sai gấp 10: phải đỏ nhờ giới hạn")
    data = {
        "doan_spec_guess": "us113/doan_spec_guess.json",
        "doan_man_xac_nhan": "us113/doan_man_xac_nhan.json",
        "doan_ho_so": "us113/doan_ho_so.json",
        "chon_cot_hoi_lai": "us113/chon_cot_hoi_lai.json",
        "chon_cot_man_xac_nhan": "us113/chon_cot_man_xac_nhan.json",
        "chon_cot_dung_spec_guess": "us113/chon_cot_dung_spec_guess.json",
        "chon_cot_dung_ho_so": "us113/chon_cot_dung_ho_so.json",
        "column_choices_dung": {c.letter: "Giá trị đo" for c in asked.open_columns},
    }
    return data, batches


def _filled_template(batch, template: Path, target: Path) -> tuple[Path, list[str]]:
    """Biểu mẫu sau khi ghi các ô không đỏ, như app sẽ ghi sau khi duyệt. Chỉ để dựng fixture, `vmai` không ghi file."""
    wb = openpyxl.load_workbook(template)
    written = []
    for change in batch.changes:
        if change.confidence.final_color.value == "red":
            continue
        wb[change.target.sheet][change.target.a1_at_write] = float(change.value.normalized)
        written.append(change.target.a1_at_write)
    return _save_stable(wb, target), written


def stale_cases(export: Export) -> dict:
    """X14, X15 cho T-DEV-04. Có lô gốc, file trước và sau khi đổi, ô phải hạ đỏ và mã lý do mong đợi. `vmai` không phát các mã này."""
    base = GENERATED / "xlsx" / "X01_m0"
    answer = json.loads((base / "answer.json").read_text(encoding="utf-8"))
    template = export.copy("stale", FORM_WITH_FORMULAS, "template.xlsx")
    source_path = export.copy("stale", base / "source.xlsx")
    form, facts = read_form(template, SPEC), read_source(source_path)
    batch = propose(form, facts, plan_headers(form, facts, load_profile(answer, form.fingerprint)))
    root = export.files / "stale"
    entry = export.entry("stale_lo_goc", batch, template, source_path, "Lô gốc của X14, X14b, X15")
    filled, written = _filled_template(batch, template, root / "bieu_mau_da_dien.xlsx")
    sheet = batch.changes[0].target.sheet

    wb = openpyxl.load_workbook(filled)
    changed = written[:3]
    for a1 in changed:
        wb[sheet][a1] = float(wb[sheet][a1].value) + 1
    x14 = _save_stable(wb, root / "X14_sua_ngoai_app.xlsx")

    wb = openpyxl.load_workbook(filled)
    formula_cell = next((c.coordinate for row in wb[sheet].iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=") and "AVERAGE" in c.value), None)
    if formula_cell is not None:
        wb[sheet][formula_cell] = "=1+1"
    x14b = _save_stable(wb, root / "X14b_sua_cong_thuc.xlsx")

    wb = openpyxl.load_workbook(source_path)
    numeric = next(c for row in wb.worksheets[0].iter_rows(min_row=2) for c in row if isinstance(c.value, float))
    numeric.value = numeric.value + 1
    x15 = _save_stable(wb, root / "X15_nguon_da_doi.xlsx")
    return {
        "lo_goc": entry,
        "bieu_mau_da_dien": export.rel(filled),
        "bieu_mau_da_dien_sha256": sha256_file(filled),
        "o_da_ghi": len(written),
        "X14": {"file": export.rel(x14), "sha256": sha256_file(x14), "sheet": sheet, "cells": changed, "expected_reason": "STALE_VALUE_CHANGED", "label": "Số đã đổi sau khi điền"},
        "X14b": {"file": export.rel(x14b), "sha256": sha256_file(x14b), "sheet": sheet, "cells": [formula_cell] if formula_cell else [], "expected_reason": "STALE_FORMULA_CHANGED"},
        "X15": {"file": export.rel(x15), "sha256": sha256_file(x15), "original_sha256": sha256_file(source_path), "expected_reason": "STALE_SOURCE_CHANGED", "label": "File nguồn đã đổi", "note": "Mọi ô lấy từ nguồn này đều hạ đỏ"},
        "chua_co_ca": ["STALE_CELL_NOT_FOUND", "STALE_POSITIONAL_LINK", "STALE_PROFILE_REVOKED"],
    }


def _on_form_with_formulas(name: str) -> tuple[object, Path, Path]:
    directory = next(base / name for base in CASE_DIRS if (base / name).exists())
    answer = json.loads((directory / "answer.json").read_text(encoding="utf-8"))
    form, source = read_form(FORM_WITH_FORMULAS, SPEC), read_source(directory / "source.xlsx")
    return propose(form, source, plan_headers(form, source, load_profile(answer, form.fingerprint))), FORM_WITH_FORMULAS, directory / "source.xlsx"


def _has_red_in_stat_range(batch) -> bool:
    return any(c.confidence.final_color.value == "red" and re.fullmatch(r"B(?:[89]|[12]\d|3[0-7])", c.target.a1_at_write) for c in batch.changes)


def build(target: Path = TARGET) -> dict:
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True)
    export = Export(target)
    excel = excel_batches()

    def add(name: str, batch, template: Path, source: Path, note: str | None = None) -> dict:
        return export.entry(name, batch, export.copy(name, template), export.copy(name, source), note)

    entries: list[dict] = []
    seen: set[str] = set()
    required_names = list(REQUIRED)
    notes = dict(REQUIRED)
    # Đỏ trong B8:B37 trên biểu mẫu có công thức thật, để dựng đồ thị đầu vào: B39 phải hạ đỏ.
    red_source = next(n for n in sorted(excel) if n not in REQUIRED and _has_red_in_stat_range(excel[n][0]) and any(c.confidence.final_color.value != "red" for c in excel[n][0].changes))
    red_case = f"{red_source}_co_cong_thuc"
    excel[red_case] = _on_form_with_formulas(red_source)
    notes[red_case] = f"Nguồn {red_source} trên biểu mẫu có công thức. Có ô đỏ trong B8:B37, B39 phải hạ đỏ"
    required_names.append(red_case)
    for name in required_names:
        batch, template, source = excel[name]
        entries.append(add(name, batch, template, source, notes[name]))
        seen |= _codes(batch) | {f"color:{c.confidence.final_color.value}" for c in batch.changes}

    # Lô ảnh có `pages`, cho ảnh cắt của dòng vàng.
    made = image_batch("clean_s301")
    page_image = None
    if made is not None:
        batch, page_image = with_page_image(export, "anh_co_trang", made[0], made[2])
        entry = add("anh_co_trang", batch, made[1], made[2], "Ảnh có `pages`, tọa độ theo `page_image`. Ảnh đang sau cờ")
        entry["page_image"] = export.rel(page_image) if page_image else None
        entries.append(entry)

    candidates = {n: v for n, v in excel.items() if n not in required_names and n != red_source}
    for name in ("A08_total", "blur1.5_s301", "dark_s301"):
        made = image_batch(name)
        if made is not None:
            candidates[f"anh_{name}"] = made
    blind = GENERATED / "images_q90"
    if blind.exists():
        for directory in sorted(blind.iterdir()):
            if directory.is_dir() and "_clean_" in directory.name:
                made = image_batch(directory.name, blind, "ocr_cache_q90")
                if made is not None:
                    candidates[f"anh_{directory.name}"] = made
    candidates["pdf_P01"] = pdf_batch()
    for name in pick_cover(candidates, seen):
        batch, template, source = candidates[name]
        entries.append(add(name, batch, template, source))

    us113_data, chosen = us113(export)
    stale = stale_cases(export)
    entries.extend(chosen.values())
    entries.append(stale["lo_goc"])
    covered = set().union(*(set(e["reasons"]) for e in entries))
    index = {
        "note": "[MẪU GIẢ] Lô thật do propose sinh từ dữ liệu giả. Ảnh, PDF đang sau cờ, chỉ để dựng giao diện. Số đo chất lượng không lấy từ đây.",
        "red_in_stat_range": f"{red_case}.json",
        "batches": entries,
        "guesses": guess_cases(export),
        "errors": error_cases(export),
        "us113": us113_data,
        "stale": stale,
        "rule_codes_not_covered": sorted(set(_RULE_CODES) - covered),
    }
    _write_json(target / "index.json", json.dumps(index, ensure_ascii=False, indent=1))
    return index


def main() -> None:
    index = build()
    print(f"{len(index['batches'])} lô, {len(index['guesses'])} ca đoán, {len(index['errors'])} lỗi. Ra: {TARGET.relative_to(ROOT).as_posix()}")
    for entry in index["batches"]:
        print(" ", entry["file"], entry["colors"], entry["note"] or "")
    print("Câu hỏi có ca:", sorted({q["code"] for g in index["guesses"] for q in g["questions"]}))
    print("Lỗi có ca:", sorted({e["code"] for e in index["errors"]}))
    print("Ô đỏ trong cột thống kê:", index["red_in_stat_range"])
    print("Mã chưa có lô nào:", index["rule_codes_not_covered"])


if __name__ == "__main__":
    import _console

    _console.utf8()
    main()
