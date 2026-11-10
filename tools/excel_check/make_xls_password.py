"""Sinh fixture file .xls có mật khẩu giả, để kiểm app phân biệt được với .xls thường. Chạy: python make_xls_password.py"""
from __future__ import annotations

import pathlib
import shutil
import sys

from excel import excel

FIX = pathlib.Path(__file__).resolve().parents[2] / "fixtures" / "templates"
PASSWORD = "matkhau-gia"
XLS_FORMAT = 56


def main() -> int:
    with excel() as xl:
        wb, err = xl.open(xl.work.copy_in(FIX / "GIA-QC-01_da_dien.xlsx", "goc.xlsx"))
        if err:
            raise RuntimeError(err)
        saved = xl.work.path / "ra.xls"
        wb.SaveAs(str(saved), XLS_FORMAT, PASSWORD)
        wb.Close(False)
        shutil.copyfile(saved, FIX / "GIA-X19_xls_mat_khau.xls")
    print("Đã sinh GIA-X19_xls_mat_khau.xls")
    return 0


if __name__ == "__main__":
    sys.exit(main())
