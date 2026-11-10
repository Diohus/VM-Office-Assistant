"""up.sh lần cài đầu: thiếu Python thì báo tiếng thường, không ghi compose.env; mật khẩu trống thì báo, không tự xóa.

Không cần Docker: chạy bản chép của up.sh trong thư mục tạm, python, docker, powershell là bản giả.
Chạy riêng được khi Docker tắt: python -m pytest tests/integration/test_up_script.py --noconftest
"""
from __future__ import annotations

import os
import pathlib
import shutil
import subprocess

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[2]
UP = ROOT / "tools" / "dev" / "up.sh"
SH = shutil.which("sh")

pytestmark = pytest.mark.skipif(SH is None, reason="Cần sh (Git Bash trên Windows)")


def stub(folder: pathlib.Path, name: str, body: str) -> None:
    path = folder / name
    path.write_text(f"#!/bin/sh\n{body}\n", encoding="utf8", newline="\n")
    path.chmod(0o755)


def run_up(tmp_path: pathlib.Path, python_body: str) -> tuple[subprocess.CompletedProcess, pathlib.Path]:
    repo = tmp_path / "app"
    (repo / "tools" / "dev").mkdir(parents=True)
    shutil.copyfile(UP, repo / "tools" / "dev" / "up.sh")
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    log = tmp_path / "docker.log"
    for name in ("python", "python3"):
        stub(bin_dir, name, python_body)
    stub(bin_dir, "docker", f'echo "$@" >> "{log.as_posix()}"')
    stub(bin_dir, "powershell", "exit 1")
    env = {**os.environ, "PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}"}
    result = subprocess.run([SH, "tools/dev/up.sh"], cwd=repo, env=env, capture_output=True, text=True, encoding="utf8", errors="replace", timeout=60)
    return result, log


def test_thieu_python_thi_bao_va_khong_tao_gi(tmp_path):
    # Giống lối tắt python của Microsoft Store: chạy được nhưng không phải Python thật.
    result, log = run_up(tmp_path, 'echo "Python was not found" >&2; exit 9009')
    assert result.returncode == 1
    assert "chưa có Python 3" in result.stdout and "python.org" in result.stdout
    assert not (tmp_path / "app" / ".local").exists(), "không ghi compose.env khi thiếu Python"
    assert not log.exists(), "không gọi docker"


def test_python_in_ra_chu_ngan_thi_coi_la_thieu(tmp_path):
    result, _ = run_up(tmp_path, "echo x")
    assert result.returncode == 1 and not (tmp_path / "app" / ".local" / "compose.env").exists()


def test_mat_khau_trong_thi_bao_va_khong_tu_xoa(tmp_path):
    local = tmp_path / "app" / ".local"
    local.mkdir(parents=True)
    (local / "compose.env").write_text("POSTGRES_PASSWORD=\n", encoding="utf8")
    result, log = run_up(tmp_path, "echo abcdefghijklmnopqrstuvwxyz012345")
    assert result.returncode == 1
    assert "mật khẩu kho dữ liệu bị trống" in result.stdout and "xóa riêng file .local/compose.env" in result.stdout
    assert (local / "compose.env").read_text(encoding="utf8") == "POSTGRES_PASSWORD=\n", "không tự sửa hay xóa"
    assert not log.exists()


def test_co_python_thi_ghi_mat_khau_va_bat_app(tmp_path):
    result, log = run_up(tmp_path, "echo abcdefghijklmnopqrstuvwxyz012345")
    assert result.returncode == 0, result.stdout + result.stderr
    env_file = tmp_path / "app" / ".local" / "compose.env"
    assert env_file.read_text(encoding="utf8") == "POSTGRES_PASSWORD=abcdefghijklmnopqrstuvwxyz012345\n"
    assert "--env-file .local/compose.env up -d --build --wait" in log.read_text(encoding="utf8")
    assert "App: http://127.0.0.1:" in result.stdout


def test_mat_khau_da_co_thi_giu_nguyen(tmp_path):
    local = tmp_path / "app" / ".local"
    local.mkdir(parents=True)
    (local / "compose.env").write_text("POSTGRES_PASSWORD=da_co_tu_truoc_khong_doi\n", encoding="utf8")
    result, _ = run_up(tmp_path, "exit 9009")
    assert result.returncode == 0, "đã có mật khẩu thì không cần Python"
    assert (local / "compose.env").read_text(encoding="utf8") == "POSTGRES_PASSWORD=da_co_tu_truoc_khong_doi\n"
