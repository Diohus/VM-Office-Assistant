"""egress, không mạng thật: chỉ CONNECT tới host:cổng cho phép, từ chối địa chỉ không công khai."""
from __future__ import annotations

import asyncio

import pytest

import egress as e


@pytest.mark.parametrize("head,want", [
    (b"CONNECT api.openai.com:443 HTTP/1.1\r\nHost: api.openai.com:443\r\n\r\n", ("api.openai.com", 443)),
    (b"CONNECT API.OpenAI.com.:443 HTTP/1.1\r\n\r\n", ("api.openai.com", 443)),
    (b"GET http://api.openai.com/ HTTP/1.1\r\n\r\n", None),
    (b"CONNECT api.openai.com HTTP/1.1\r\n\r\n", None),
    (b"CONNECT [::1]:443 HTTP/1.1\r\n\r\n", None),
    (b"CONNECT api.openai.com:443\r\n\r\n", None),
    ("CONNECT ápi.openai.com:443 HTTP/1.1\r\n\r\n".encode(), None),
])
def test_doc_dong_connect(head, want):
    assert e.parse_connect(head) == want


@pytest.mark.parametrize("host,port,ok", [("api.openai.com", 443, True), ("api.openai.com", 80, False), ("openai.com", 443, False),
                                          ("evil.api.openai.com", 443, False), ("api.openai.com.evil.com", 443, False), ("example.com", 443, False)])
def test_danh_sach_cho_phep(host, port, ok):
    assert e.allowed(host, port) is ok


@pytest.mark.parametrize("address,ok", [("8.8.8.8", True), ("10.0.0.1", False), ("127.0.0.1", False), ("172.18.0.2", False),
                                        ("169.254.169.254", False), ("::1", False), ("fd00::1", False)])
def test_chi_dia_chi_cong_khai(address, ok):
    assert e.public(address) is ok


def _talk(head: bytes) -> bytes:
    async def run():
        server = await asyncio.start_server(e.handle, "127.0.0.1", 0)
        port = server.sockets[0].getsockname()[1]
        reader, writer = await asyncio.open_connection("127.0.0.1", port)
        writer.write(head)
        await writer.drain()
        data = await asyncio.wait_for(reader.read(200), 5)
        writer.close()
        server.close()
        return data
    return asyncio.run(run())


def test_host_khac_bi_403():
    assert _talk(b"CONNECT example.com:443 HTTP/1.1\r\n\r\n").startswith(b"HTTP/1.1 403")


def test_phuong_thuc_khac_bi_405():
    assert _talk(b"GET http://api.openai.com/ HTTP/1.1\r\nHost: api.openai.com\r\n\r\n").startswith(b"HTTP/1.1 405")
