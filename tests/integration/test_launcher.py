"""T-DEV-03: lối tắt mở app. Chỉ thử trên bộ test 8089 và docker giả, không đụng 8088, không tắt Docker thật.

Không thay được phép thử khởi động lại máy thật, là tiêu chí của T-DEV-03.
"""
from __future__ import annotations

import os
import pathlib
import subprocess

import pytest

from conftest import PROJECT

ROOT = pathlib.Path(__file__).resolve().parents[2]
LAUNCHER = ROOT / "tools" / "dev" / "mo_app.cmd"
SHORTCUT = ROOT / "tools" / "dev" / "tao_loi_tat.cmd"
SAFE_ENV = {"VM_PROJECT": PROJECT, "WEB_PORT": "8089", "VM_NO_BROWSER": "1", "VM_NO_PAUSE": "1"}

pytestmark = pytest.mark.skipif(os.name != "nt", reason="Lối tắt chỉ dành cho Windows")


def run(**env) -> subprocess.CompletedProcess:
    return subprocess.run(["cmd.exe", "/c", str(LAUNCHER)], capture_output=True, text=True, encoding="utf8", errors="replace", timeout=200, env={**os.environ, **SAFE_ENV, **env})


def technical(text: str) -> list[str]:
    return [w for w in ("compose", "daemon", "npipe", "Traceback", "exit code", "127.0.0.1:8088") if w in text]


def test_docker_dang_chay_thi_mo_duoc_bo_rieng():
    proc = run()
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "http://127.0.0.1:8089" in proc.stdout and not technical(proc.stdout)


def _llm() -> tuple[str, str]:
    """Mã container llm của bộ test và file gắn vào /run/secrets/llm_env (đường dẫn, không đọc nội dung)."""
    fmt = '{{.Id}}|{{range .Mounts}}{{if eq .Destination "/run/secrets/llm_env"}}{{.Source}}{{end}}{{end}}'
    out = subprocess.run(["docker", "inspect", "-f", fmt, f"{PROJECT}-llm-1"], capture_output=True, text=True, timeout=60, check=True).stdout.strip()
    container, source = out.split("|", 1)
    return container, source


def test_loi_tat_khong_tao_lai_llm_va_giu_khoa(monkeypatch):
    """Q135 (PM 13/11): run.sh dựng bộ test bằng up.sh với khóa mồi. Lối tắt không có biến LLM_SECRET_FILE (như trên máy Owner),
    nên compose thấy đường dẫn file key khác. Đo 13/11: compose không tạo lại container khi chỉ đường dẫn secret khác, nên llm
    giữ nguyên container và vẫn gắn khóa mồi. Nếu một bản Docker sau này đổi cách đó, test này báo trước khi app thật mất key."""
    before = _llm()
    assert "llm_canary" in before[1], "bộ test phải chạy với khóa mồi (run.sh)"
    monkeypatch.delenv("LLM_SECRET_FILE", raising=False)
    proc = run()
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert _llm() == before, "lối tắt đã tạo lại llm hoặc đổi file key"


def test_chay_bo_test_khong_ghi_khoa_moi_vao_compose_env():
    """PM chọn (A): .local/compose.env dùng chung với app thật; bộ test không bao giờ ghi đường dẫn khóa mồi vào đó."""
    text = (ROOT / ".local" / "compose.env").read_text(encoding="utf8")
    assert "llm_canary" not in text and "tests/integration" not in text


def test_docker_khong_len_thi_bao_tieng_thuong(tmp_path):
    fake = tmp_path / "docker.cmd"
    fake.write_text("@exit /b 1\r\n", encoding="ascii")
    desktop = tmp_path / "desktop.cmd"
    desktop.write_text("@exit\r\n", encoding="ascii")  # start mở .cmd bằng cmd /K: exit /b để cửa sổ mở và giữ pipe
    proc = run(VM_DOCKER=str(fake), VM_DOCKER_DESKTOP=str(desktop), VM_WAIT_SECONDS="3")
    assert proc.returncode == 13
    assert "chưa bật xong" in proc.stdout and "không bị đổi" in proc.stdout and "nhấp lại" in proc.stdout and not technical(proc.stdout)


def test_docker_len_cham_thi_cho_roi_mo(tmp_path):
    real = subprocess.run(["where", "docker"], capture_output=True, text=True).stdout.splitlines()[0]
    flag = tmp_path / "da_bat"
    fake = tmp_path / "docker.cmd"
    fake.write_text(f'@if not exist "{flag}" exit /b 1\r\n@"{real}" %*\r\n', encoding="utf8")
    desktop = tmp_path / "desktop.cmd"
    desktop.write_text(f'@echo.> "{flag}"\r\n@exit\r\n', encoding="ascii")
    proc = run(VM_DOCKER=str(fake), VM_DOCKER_DESKTOP=str(desktop), VM_WAIT_SECONDS="30")
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "đang bật" in proc.stdout.lower() and "http://127.0.0.1:8089" in proc.stdout


def test_khong_tim_thay_docker_desktop_thi_bao(tmp_path):
    fake = tmp_path / "docker.cmd"
    fake.write_text("@exit /b 1\r\n", encoding="ascii")
    proc = run(VM_DOCKER=str(fake), VM_DOCKER_DESKTOP=str(tmp_path / "khong_co.exe"))
    assert proc.returncode == 12
    assert "chưa có Docker Desktop" in proc.stdout and not technical(proc.stdout)


def test_app_khong_bat_duoc_thi_bao(tmp_path):
    fake = tmp_path / "docker.cmd"
    fake.write_text('@if "%1"=="version" exit /b 0\r\n@exit /b 1\r\n', encoding="ascii")
    proc = run(VM_DOCKER=str(fake))
    assert proc.returncode == 14
    assert "chưa bật được" in proc.stdout and not technical(proc.stdout)


def test_khong_dung_lai_image_va_khong_doc_env():
    text = LAUNCHER.read_text(encoding="utf8")
    code = [line for line in text.splitlines() if not line.lower().startswith("rem")]
    assert " --no-build" in "\n".join(code) and "--build " not in "\n".join(code)
    assert not any(".env" in line.replace(".local/compose.env", "").replace(".local\\compose.env", "") for line in code)
    assert not any(word in text.lower() for word in ("executionpolicy", "reg add", "set-mppreference", "bcdedit", "schtasks"))


def shortcut(folder, pause: bool = False) -> subprocess.CompletedProcess:
    env = {**os.environ}
    env.pop("VM_NO_PAUSE", None)
    if not pause:
        env["VM_NO_PAUSE"] = "1"
    # Có pause thì gửi sẵn một phím, như người dùng bấm phím để đóng cửa sổ.
    return subprocess.run(["cmd.exe", "/c", str(SHORTCUT), str(folder)], input="x\r\n" if pause else None, stdin=None if pause else subprocess.DEVNULL,
                          capture_output=True, text=True, encoding="utf8", errors="replace", timeout=60, env=env)


def test_tao_bieu_tuong_trong_thu_muc_duoc_chi(tmp_path):
    proc = shortcut(tmp_path)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    links = list(tmp_path.glob("*.lnk"))
    assert len(links) == 1 and "VM-Office" in links[0].name


def test_tao_bieu_tuong_dung_lai_cho_doc_ket_qua(tmp_path):
    # T-DEV-06: nhấp đúp thì cửa sổ phải dừng để người dùng đọc "Đã tạo biểu tượng" (Owner gặp 10/11).
    proc = shortcut(tmp_path, pause=True)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    done = proc.stdout.index("Đã tạo biểu tượng")
    assert proc.stdout[done:].strip().count("\n") >= 1, "sau câu kết quả phải còn lời nhắc của pause"
    assert len(list(tmp_path.glob("*.lnk"))) == 1


def test_tao_bieu_tuong_thu_muc_khong_co_thi_bao_va_dung_lai(tmp_path):
    proc = shortcut(tmp_path / "khong_co", pause=True)
    assert proc.returncode == 1
    assert "Không tìm thấy thư mục" in proc.stdout and not list(tmp_path.glob("**/*.lnk"))
    assert proc.stdout.strip() != "Không tìm thấy thư mục để đặt biểu tượng.", "phải có lời nhắc của pause sau câu báo"


def test_file_cmd_dung_xuong_dong_cua_windows():
    for path in (LAUNCHER, SHORTCUT):
        data = path.read_bytes()
        assert data.count(b"\n") == data.count(b"\r\n"), f"{path.name}: nhãn goto của cmd có thể hỏng nếu xuống dòng kiểu Linux"
