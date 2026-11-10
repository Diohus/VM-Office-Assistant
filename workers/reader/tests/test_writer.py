"""Bộ ghi theo Q48. So với Excel thật nằm ở tools/excel_check/q48_check.py."""
import hashlib
import io
import pathlib
import zipfile

import openpyxl
import pytest

from reader.writer import UUID_NAME, WriteRefused, apply

FIX = pathlib.Path(__file__).resolve().parents[3] / "fixtures" / "templates"
QC = (FIX / "GIA-QC-01_da_dien.xlsx").read_bytes()
UUID = "6d0c2f1a-6a3e-4a53-9e57-0f7b7f3c9a11"


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def book(data: bytes, **kw):
    return openpyxl.load_workbook(io.BytesIO(data), **kw)


def test_ghi_so_va_bo_moi_gia_tri_da_luu():
    out = apply(QC, [{"sheet": "KQ_QC_Thang", "cell": "B8", "value": 5.5}], UUID)
    ws = book(out["data"])["KQ_QC_Thang"]
    assert ws["B8"].value == 5.5
    assert str(ws["B39"].value).startswith("=")
    values = book(out["data"], data_only=True)
    for sheet in book(out["data"]).worksheets:
        for row in sheet.iter_rows():
            for c in row:
                if c.data_type == "f":
                    assert values[sheet.title][c.coordinate].value is None, c.coordinate
    assert b'fullCalcOnLoad="1"' in parts(out["data"])["xl/workbook.xml"]
    assert out["formulas_cleared"] >= 150


def test_so_trong_bieu_do_bi_bo():
    out = apply(QC, [{"sheet": "KQ_QC_Thang", "cell": "B8", "value": 500}], UUID)
    assert out["chart_caches_cleared"] > 0
    assert b"numCache" not in parts(out["data"])["xl/charts/chart1.xml"]


def test_phan_khong_lien_quan_giu_nguyen_tung_byte():
    before, after = parts(QC), parts(apply(QC, [{"sheet": "KQ_QC_Thang", "cell": "B8", "value": 5.5}], UUID)["data"])
    for name in ("xl/styles.xml", "xl/sharedStrings.xml", "xl/drawings/drawing1.xml", "xl/comments1.xml", "xl/theme/theme1.xml"):
        if name in before:
            assert hashlib.sha256(before[name]).digest() == hashlib.sha256(after[name]).digest(), name


def test_nhung_ma_dinh_danh_va_giu_khi_ghi_lan_hai():
    first = apply(QC, [{"sheet": "KQ_QC_Thang", "cell": "B8", "value": 5.5}], UUID)["data"]
    assert UUID.encode() in parts(first)["docProps/custom.xml"] and UUID_NAME.encode() in parts(first)["docProps/custom.xml"]
    second = apply(first, [{"sheet": "KQ_QC_Thang", "cell": "B9", "value": 5.6}], UUID)["data"]
    assert parts(second)["docProps/custom.xml"].count(UUID_NAME.encode()) == 1
    assert book(second).properties is not None


@pytest.mark.parametrize("text", ["=1+1", "+1", "-1", "@SUM(A1)", '=WEBSERVICE("http://example.invalid")'])
def test_chu_trong_giong_cong_thuc_van_la_chu(text):
    out = apply(QC, [{"sheet": "KQ_QC_Thang", "cell": "I8", "value": text}], UUID)
    cell = book(out["data"])["KQ_QC_Thang"]["I8"]
    assert cell.data_type in ("s", "inlineStr") and cell.value == text


def test_cong_thuc_goi_ra_ngoai_bi_tu_choi():
    with pytest.raises(WriteRefused) as err:
        apply(QC, [{"sheet": "KQ_QC_Thang", "cell": "I8", "formula": 'WEBSERVICE("http://example.invalid")'}], UUID)
    assert err.value.cells[0]["cell"] == "I8"


def test_o_khoa_bi_tu_choi_va_khong_ghi_o_nao():
    data = (FIX / "GIA-X19_sheet_khoa.xlsx").read_bytes()
    with pytest.raises(WriteRefused) as err:
        apply(data, [{"sheet": "GioiHan", "cell": "A1", "value": "x"}, {"sheet": "KQ_QC_Thang", "cell": "B8", "value": 1}], UUID)
    assert [c["cell"] for c in err.value.cells] == ["B8"]


@pytest.mark.parametrize("edit", [{"sheet": "Khong_co", "cell": "A1", "value": 1}, {"sheet": "KQ_QC_Thang", "cell": "1A", "value": 1}, {"sheet": "KQ_QC_Thang", "cell": "A1:B2", "value": 1}])
def test_dia_chi_sai_bi_tu_choi(edit):
    with pytest.raises(WriteRefused):
        apply(QC, [edit], UUID)


def test_ham_doi_moi_co_tien_to_con_chu_trong_ngoac_kep_thi_khong():
    out = apply(QC, [{"sheet": "KQ_QC_Thang", "cell": "I8", "formula": 'STDEV.S(B8:B37)&"STDEV.S("'}], UUID)
    value = book(out["data"])["KQ_QC_Thang"]["I8"].value
    assert value == '=_xlfn.STDEV.S(B8:B37)&"STDEV.S("'
