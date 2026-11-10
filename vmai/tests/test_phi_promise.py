"""Câu `PHI_COLUMN_SKIPPED` hứa hai điều: app không dùng cột người bệnh, và không gửi cột đó ra ngoài (T-AI-10, Q109).

Test đi trọn luồng của `vmai`: plan_headers -> map_headers qua provider giả có ghi lại yêu cầu -> propose.
Phạm vi: chứng minh trong đường gửi của `vmai`. Khi dịch vụ `llm` của app được nối (phần 4) phải kiểm lại ở mức app.
[MẪU GIẢ] Tên và mã người bệnh đều giả.
"""
from __future__ import annotations

import json
from pathlib import Path

import openpyxl
import pytest

from vmai import load_spec, plan_headers, propose, read_form, read_source
from vmai.llm.base import LLMResponse
from vmai.llm.factory import create_provider, load_profile
from vmai.llm_tasks import map_headers

import vmai

ROOT = Path(__file__).resolve().parents[2]
CASE = ROOT / "fixtures" / "eval" / "generated" / "xlsx" / "X01_m0"
SPEC = load_spec(Path(vmai.__file__).parent / "specs" / "GIA-QC-01.yaml")
PHI_HEADERS = ("Họ và tên", "Mã BN")



def _source_with_phi(tmp_path: Path) -> tuple[Path, list[str]]:
    """Nguồn X01_m0, đổi tên cột Kết quả để phải hỏi model, thêm hai cột người bệnh giả với giá trị dễ nhận."""
    tmp_path.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(CASE / "source.xlsx")
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ws.cell(row=1, column=headers.index("Kết quả") + 1, value="Trị số")
    name_col, id_col = ws.max_column + 1, ws.max_column + 2
    ws.cell(row=1, column=name_col, value=PHI_HEADERS[0])
    ws.cell(row=1, column=id_col, value=PHI_HEADERS[1])
    values = []
    for row in range(2, ws.max_row + 1):
        name, code = f"Nguyễn Giả Bệnh {row:04d}", f"BN-GIA-{row:06d}"
        ws.cell(row=row, column=name_col, value=name)
        ws.cell(row=row, column=id_col, value=code)
        values += [name, code]
    path = tmp_path / "nguon_co_nguoi_benh.xlsx"
    wb.save(path)
    return path, values


def _recording_provider():
    sent = []

    def handler(request):
        sent.append(request)
        asked = json.loads(request.messages[0].parts[0].text.split("Tiêu đề nguồn: ", 1)[1].replace("'", '"'))
        mappings = [{"source_header": h, "field": "result" if h == "Trị số" else None} for h in asked]
        return LLMResponse("fake", "fake-1", text=json.dumps({"mappings": mappings}))

    return create_provider(load_profile("fake"), handler=handler), sent


def _request_text(request) -> str:
    parts = [request.system] + [p.text for m in request.messages for p in m.parts if hasattr(p, "text")]
    return "\n".join(parts) + json.dumps(request.response_schema or {}, ensure_ascii=False)


def _run(tmp_path, forged: dict | None = None):
    path, values = _source_with_phi(tmp_path)
    form, source = read_form(CASE / "template.xlsx", SPEC), read_source(path)
    plan = plan_headers(form, source)
    provider, sent = _recording_provider()
    header_map, call = map_headers(provider, list(plan.llm_headers), plan.llm_fields)
    if forged:
        header_map = {**header_map, **forged}
    batch = propose(form, source, plan, header_map, (call,))
    return plan, sent, batch, values


def test_phi_columns_are_not_sent(tmp_path):
    plan, sent, _, values = _run(tmp_path)
    assert set(plan.phi_headers) == set(PHI_HEADERS)
    assert sent, "phải có ít nhất một yêu cầu tới provider, nếu không test không chứng minh gì"
    text = "\n".join(_request_text(r) for r in sent)
    assert "Trị số" in text  # tiêu đề cần hỏi đúng là đã đi qua đây
    for leaked in PHI_HEADERS + tuple(values):
        assert leaked not in text, leaked


def test_phi_columns_are_not_used(tmp_path):
    _, _, batch, values = _run(tmp_path)
    assert len(batch.changes) == 180  # luồng thật sự chạy, ô Kết quả vẫn được đề xuất
    assert "PHI_COLUMN_SKIPPED" in {n.reason for n in batch.notices}
    dumped = batch.model_dump_json()
    for value in values:
        assert value not in dumped, value
    for change in batch.changes:
        for evidence in change.evidence:
            assert not any(h in json.dumps(evidence.model_dump(), ensure_ascii=False) for h in PHI_HEADERS)


def test_forged_header_map_cannot_pull_a_phi_column_in(tmp_path):
    _, _, baseline, _ = _run(tmp_path / "a")
    _, _, forged, values = _run(tmp_path / "b", forged={PHI_HEADERS[0]: "result", PHI_HEADERS[1]: "code"})
    assert len(forged.changes) == len(baseline.changes)
    dumped = forged.model_dump_json()
    assert not any(v in dumped for v in values)


def test_source_file_name_never_goes_into_reason_details(tmp_path):
    """Q131: câu "Chép từ file {tên file}." do giao diện điền. `vmai` không chép tên file vào `detail`, vì tên file có thể
    chứa tên người. [MẪU GIẢ] Tên trong tên file là giả."""
    import shutil

    stem = "KQ_Nguyen_Gia_Thu"
    path = tmp_path / f"{stem}.xlsx"
    shutil.copyfile(CASE / "source.xlsx", path)
    form = read_form(CASE / "template.xlsx", SPEC)
    source = read_source(path)
    batch = propose(form, source, plan_headers(form, source))
    details = [r.detail or "" for c in batch.changes for r in c.confidence.reasons] + [n.detail or "" for n in batch.notices]
    assert any(r.code == "READ_XLSX_DIRECT" for c in batch.changes for r in c.confidence.reasons)
    assert not [d for d in details if stem.lower() in d.lower() or "nguyen_gia" in d.lower()]
