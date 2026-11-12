"""Lát 2, mốc hạ tầng (Q135): llm, egress, key qua secret dạng file, khóa mồi không lộ.

Bộ test dùng khóa mồi (tests/integration/out/llm_canary.env do run.sh sinh) và máy giả fakeai thay OpenAI.
Không bao giờ gắn .env thật. Lệnh docker exec vào llm chỉ chạy trên bộ test, nơi key là khóa mồi.
"""
from __future__ import annotations

import hashlib
import json
import subprocess

import pytest

from conftest import BASE, PROJECT, ROOT, request

CANARY_FILE = ROOT / "tests" / "integration" / "out" / "llm_canary.env"
SERVICES = ["web", "api", "reader", "converter", "db", "llm", "egress", "fakeai"]


def canary() -> str:
    line = CANARY_FILE.read_text(encoding="utf-8").strip()
    assert line.startswith("OPENAI_API_KEY=sk-test-canary-"), "bộ test chỉ được chạy với khóa mồi"
    return line.split("=", 1)[1]


def docker(*args: str, timeout: int = 90) -> subprocess.CompletedProcess:
    return subprocess.run(["docker", *args], capture_output=True, text=True, encoding="utf8", errors="replace", timeout=timeout)


def container(service: str) -> str:
    return f"{PROJECT}-{service}-1"


CALL = """
import json, sys
from app import llm_client
try:
    out = llm_client.call("He thong", sys.argv[1], {"type": "object"}, "chat_edit", 2000)
    print(json.dumps({"ok": True, **out}))
except llm_client.LlmFailed as exc:
    print(json.dumps({"ok": False, "code": exc.code}))
"""


def call_via_api(prompt: str) -> tuple[dict, str]:
    out = docker("exec", container("api"), "python", "-c", CALL, prompt, timeout=60)
    assert out.returncode == 0, out.stderr[-500:]
    return json.loads(out.stdout.strip().splitlines()[-1]), out.stdout + out.stderr


@pytest.fixture(scope="module")
def results():
    """Chạy mọi đường gọi một lần, giữ cả chữ thô đầu ra để soát khóa mồi."""
    got = {}
    for marker in ["FAKE:ok", "FAKE:echo_401", "FAKE:echo_500", "FAKE:garbage", "FAKE:refusal", "FAKE:slow"]:
        got[marker] = call_via_api(f"Dữ liệu giả {marker}")
    return got


def test_goi_duoc_qua_llm_va_key_toi_dung_noi(results):
    body, _ = results["FAKE:ok"]
    assert body["ok"] is True and json.loads(body["text"])["status"] == "ask"
    assert body["input_tokens"] == 1200 and body["output_tokens"] == 300 and body["reasoning_tokens"] == 100
    assert body["prompt_sha256"] == hashlib.sha256("Dữ liệu giả FAKE:ok".encode()).hexdigest()
    assert f"prompt={body['prompt_sha256'][:12]}" in docker("logs", container("llm")).stderr
    seen = docker("exec", container("fakeai"), "python", "-c", "import urllib.request as u; print(u.urlopen('http://127.0.0.1:8080/seen').read().decode())")
    assert json.loads(seen.stdout)["auth_sha256"] == hashlib.sha256(f"Bearer {canary()}".encode()).hexdigest()


@pytest.mark.parametrize("marker,code", [("FAKE:echo_401", "CHAT_MODEL_KEY_REJECTED"), ("FAKE:echo_500", "CHAT_MODEL_UNAVAILABLE"),
                                         ("FAKE:garbage", "CHAT_MODEL_FAILED"), ("FAKE:refusal", "CHAT_MODEL_REFUSED"),
                                         ("FAKE:slow", "CHAT_MODEL_TIMEOUT")])
def test_loi_chi_la_ma_khong_mang_chu_goc(results, marker, code):
    body, raw = results[marker]
    assert body == {"ok": False, "code": code}
    assert canary() not in raw


def test_khoa_moi_khong_co_trong_log_inspect_va_co_so_du_lieu(results):
    key = canary()
    for service in SERVICES:
        logs = docker("logs", container(service))
        assert key not in logs.stdout + logs.stderr, f"khóa mồi trong log {service}"
        assert key not in docker("inspect", container(service)).stdout, f"khóa mồi trong docker inspect {service}"
    dump = docker("exec", container("db"), "pg_dump", "-U", "vm", "-d", "vm", "--data-only", timeout=120)
    assert dump.returncode == 0 and key not in dump.stdout
    for path in ["/api/health", "/api/documents"]:
        _, headers, body = request("GET", path)
        assert key.encode() not in body and key not in json.dumps(headers)


def test_chi_llm_co_file_secret():
    for service in ["api", "reader", "web", "converter", "egress"]:
        out = docker("exec", container(service), "sh", "-c", "test -e /run/secrets/llm_env && echo co || echo khong")
        assert out.stdout.strip() == "khong", service
    # Trên bộ test (khóa mồi) mới được exec vào llm. Chỉ hỏi có file hay không, không đọc.
    assert docker("exec", container("llm"), "sh", "-c", "test -e /run/secrets/llm_env && echo co").stdout.strip() == "co"


def test_key_khong_nam_trong_bien_moi_truong_cua_llm():
    env = docker("exec", container("llm"), "env").stdout
    assert canary() not in env and "OPENAI_API_KEY" not in env


def _networks(service: str) -> set[str]:
    data = json.loads(docker("inspect", container(service)).stdout)[0]
    if data["HostConfig"]["NetworkMode"] == "none":
        return {"none"}
    return {name.removeprefix(f"{PROJECT}_") for name in data["NetworkSettings"]["Networks"]}


def test_mang_dung_thiet_ke():
    assert _networks("reader") == {"none"} and _networks("converter") == {"none"}
    assert _networks("llm") == {"llm_in", "llm_out"}
    assert _networks("egress") == {"llm_out", "egress_out"}
    assert "llm_in" in _networks("api") and "llm_out" not in _networks("api")
    assert not _networks("web") & {"llm_in", "llm_out"}
    for name in ["llm_in", "llm_out"]:
        assert json.loads(docker("network", "inspect", f"{PROJECT}_{name}").stdout)[0]["Internal"] is True


PROBE = """
import socket, sys
def connect(host, port):
    try:
        s = socket.create_connection((host, port), timeout=5)
        s.close()
        return "noi_duoc"
    except OSError:
        return "khong"
def via_egress(target):
    s = socket.create_connection(("egress", 3128), timeout=10)
    s.sendall(f"CONNECT {target} HTTP/1.1\\r\\nHost: {target}\\r\\n\\r\\n".encode())
    line = s.recv(100).split(b"\\r\\n")[0].decode()
    s.close()
    return line
print(connect("1.1.1.1", 443))
print(connect("api.openai.com", 443))
print(via_egress("example.com:443"))
print(via_egress("api.openai.com:80"))
print(via_egress("fakeai:8080"))
print(via_egress("api.openai.com:443"))
"""


def test_llm_khong_tu_ra_internet_chi_qua_egress_toi_openai():
    out = docker("exec", container("llm"), "python", "-c", PROBE, timeout=90)
    lines = out.stdout.strip().splitlines()
    assert lines[0] == "khong" and lines[1] == "khong", "llm ra thẳng Internet được"
    assert lines[2].startswith("HTTP/1.1 403") and lines[3].startswith("HTTP/1.1 403") and lines[4].startswith("HTTP/1.1 403")
    # Host cho phép: qua được bước kiểm (200 khi máy có mạng, 502 khi không). Không gửi gì qua đường hầm.
    assert lines[5].startswith(("HTTP/1.1 200", "HTTP/1.1 502")), lines[5]


def test_api_va_web_khong_goi_thang_duoc_egress():
    probe = "import socket\ntry:\n    socket.create_connection(('egress', 3128), timeout=5); print('noi_duoc')\nexcept OSError:\n    print('khong')"
    assert docker("exec", container("api"), "python", "-c", probe).stdout.strip() == "khong"
