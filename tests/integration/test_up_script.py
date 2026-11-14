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


def run_up(tmp_path: pathlib.Path, python_body: str, **extra_env: str) -> tuple[subprocess.CompletedProcess, pathlib.Path]:
    repo = tmp_path / "app"
    (repo / "tools" / "dev").mkdir(parents=True, exist_ok=True)
    shutil.copyfile(UP, repo / "tools" / "dev" / "up.sh")
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir(exist_ok=True)
    log = tmp_path / "docker.log"
    for name in ("python", "python3"):
        stub(bin_dir, name, python_body)
    stub(bin_dir, "docker", f'echo "$@" >> "{log.as_posix()}"')
    stub(bin_dir, "powershell", "exit 1")
    # Không mang biến của người gọi (run.sh đặt VM_PROJECT, LLM_SECRET_FILE cho bộ test): mỗi ca tự nói rõ.
    env = {k: v for k, v in os.environ.items() if k not in ("VM_PROJECT", "LLM_SECRET_FILE", "WEB_PORT")}
    env.update(extra_env, PATH=f"{bin_dir}{os.pathsep}{os.environ['PATH']}")
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
    assert env_file.read_text(encoding="utf8") == "POSTGRES_PASSWORD=abcdefghijklmnopqrstuvwxyz012345\nLLM_SECRET_FILE=.local/llm_none.env\n"
    called = log.read_text(encoding="utf8")
    assert "--env-file .local/compose.env up -d --build --wait" in called
    # R-1020: thư mục dự án tuyệt đối, để buildx bake không đòi --allow=fs.read với Dockerfile.
    project_dir = called.split("--project-directory ", 1)[1].split(" ", 1)[0]
    assert pathlib.PurePosixPath(project_dir).is_absolute() or pathlib.PureWindowsPath(project_dir).is_absolute()
    assert "App: http://127.0.0.1:" in result.stdout


def test_mat_khau_da_co_thi_giu_nguyen(tmp_path):
    local = tmp_path / "app" / ".local"
    local.mkdir(parents=True)
    (local / "compose.env").write_text("POSTGRES_PASSWORD=da_co_tu_truoc_khong_doi\n", encoding="utf8")
    result, _ = run_up(tmp_path, "exit 9009")
    assert result.returncode == 0, "đã có mật khẩu thì không cần Python"
    assert (local / "compose.env").read_text(encoding="utf8").splitlines()[0] == "POSTGRES_PASSWORD=da_co_tu_truoc_khong_doi"


# Q135, PM chọn (A) 13/11: lối tắt Desktop chỉ đọc .local/compose.env, nên up.sh ghi lại ĐƯỜNG DẪN file key đã chọn (không đọc
# nội dung). Chỉ ghi cho app thật (vm-office) và khi up.sh tự chọn; bộ test không bao giờ ghi, vì compose.env dùng chung.
def _env_lines(tmp_path):
    return (tmp_path / "app" / ".local" / "compose.env").read_text(encoding="utf8").splitlines()


def test_app_that_co_env_thi_ghi_duong_dan_khong_ghi_noi_dung(tmp_path):
    (tmp_path / "app").mkdir()
    (tmp_path / "app" / ".env").write_text("OPENAI_API_KEY=sk-gia-khong-duoc-chep\n", encoding="utf8")
    result, _ = run_up(tmp_path, "echo abcdefghijklmnopqrstuvwxyz012345")
    assert result.returncode == 0, result.stdout + result.stderr
    lines = _env_lines(tmp_path)
    assert lines.count("LLM_SECRET_FILE=.env") == 1
    assert not any("sk-gia" in line for line in lines), "chỉ ghi đường dẫn, không chép nội dung .env"


def test_chay_lai_khong_nhan_doi_dong_va_doi_theo_lua_chon_moi(tmp_path):
    run_up(tmp_path, "echo abcdefghijklmnopqrstuvwxyz012345")
    assert "LLM_SECRET_FILE=.local/llm_none.env" in _env_lines(tmp_path)
    (tmp_path / "app" / ".env").write_text("OPENAI_API_KEY=x\n", encoding="utf8")
    run_up(tmp_path, "exit 9009")
    lines = _env_lines(tmp_path)
    assert [line for line in lines if line.startswith("LLM_SECRET_FILE=")] == ["LLM_SECRET_FILE=.env"]
    assert lines[0] == "POSTGRES_PASSWORD=abcdefghijklmnopqrstuvwxyz012345"


@pytest.mark.parametrize("extra", [{"VM_PROJECT": "vm-office-test"}, {"LLM_SECRET_FILE": "tests/integration/out/llm_canary.env"},
                                   {"VM_PROJECT": "vm-office-test", "LLM_SECRET_FILE": "tests/integration/out/llm_canary.env"}])
def test_bo_test_hoac_nguoi_goi_dat_khoa_thi_khong_ghi(tmp_path, extra):
    result, _ = run_up(tmp_path, "echo abcdefghijklmnopqrstuvwxyz012345", **extra)
    assert result.returncode == 0, result.stdout + result.stderr
    assert not any(line.startswith("LLM_SECRET_FILE=") for line in _env_lines(tmp_path))
