"""Tiến trình con quá giờ, quá bộ nhớ, sập: worker phải trả lỗi dễ hiểu và sống tiếp. Chạy ngay trong container reader thật."""
from __future__ import annotations

import json
import subprocess

import pytest

from conftest import PROJECT

SCRIPT = """
import json, pathlib, tempfile
from reader import worker
d = pathlib.Path(tempfile.mkdtemp())
worker.SPOOL = d
worker.JOB_TIMEOUT_SECONDS = 2
worker.MEMORY_LIMIT_BYTES = 256 * 1024 * 1024
worker.CHILD_ARGS = ["-c", {code!r}]
(d / "x.xlsx").write_bytes(b"x")
(d / "x.job").write_text("{{}}")
worker._run_job(d / "x.job")
print(json.dumps({{"out": json.loads((d / "x.out.json").read_text()), "left": sorted(p.name for p in d.iterdir())}}))
"""


@pytest.mark.parametrize("name,code,reason", [
    ("qua_gio", "import time; time.sleep(60)", "timeout"),
    ("qua_bo_nho", "x = bytearray(2 * 1024 ** 3)", "crash"),
    ("sap", "import os; os._exit(9)", "crash"),
])
def test_tien_trinh_con_loi(name, code, reason):
    proc = subprocess.run(["docker", "exec", f"{PROJECT}-reader-1", "python", "-c", SCRIPT.format(code=code)], capture_output=True, text=True, encoding="utf8", timeout=60)
    assert proc.returncode == 0, proc.stderr[-800:]
    result = json.loads(proc.stdout.strip().splitlines()[-1])
    assert result["out"]["ok"] is False and result["out"]["detail"]["reason"] == reason and "không bị đổi" in result["out"]["message"]
    assert result["left"] == ["x.out.json"], "nội dung tải lên không được nằm lại"
