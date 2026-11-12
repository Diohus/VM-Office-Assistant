"""Sinh fixture file .xls có mật khẩu giả, để kiểm app phân biệt được với .xls thường. Chạy: python make_xls_password.py

R-1016: sinh như clean_fixture_paths.py (thư mục trung tính, RemovePersonalInformation), để không có tên người trong phần mã hóa.
"""
from __future__ import annotations

import sys

from clean_fixture_paths import FIX, XL_XLS, save_with_password
from excel import excel

NAME = "GIA-X19_xls_mat_khau.xls"


def main() -> int:
    with excel() as xl:
        data = save_with_password(xl, (FIX / "GIA-QC-01_da_dien.xlsx").read_bytes(), NAME, XL_XLS)
    (FIX / NAME).write_bytes(data)
    print(f"Đã sinh {NAME}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
