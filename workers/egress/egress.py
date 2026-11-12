"""Dịch vụ egress (Q135, R17): proxy chỉ nhận CONNECT tới đúng các host:cổng cho phép (mặc định api.openai.com:443).

- Không giải mã TLS: chỉ nối hai đầu, nên không thấy key hay nội dung.
- Từ chối mọi phương thức khác CONNECT, mọi host khác danh sách, và địa chỉ phân giải ra không phải địa chỉ công khai.
- Log chỉ ghi host, cổng, cho hay chặn, số byte. Chỉ thư viện chuẩn của Python.
"""
from __future__ import annotations

import asyncio
import ipaddress
import logging
import os
import socket

ALLOW = frozenset(x.strip().lower() for x in os.environ.get("EGRESS_ALLOW", "api.openai.com:443").split(",") if x.strip())
# Chỉ cho bộ test: cho phép địa chỉ nội bộ (máy giả trong mạng Docker). Bản chạy thật không đặt biến này.
ALLOW_PRIVATE = os.environ.get("EGRESS_ALLOW_PRIVATE") == "1"
PORT = int(os.environ.get("EGRESS_PORT", "3128"))
MAX_HEADER_BYTES = 8192
MAX_CONNECTIONS = 16
IDLE_SECONDS = 60
CONNECT_SECONDS = 10

log = logging.getLogger("egress")
_slots = asyncio.Semaphore(MAX_CONNECTIONS)


def parse_connect(head: bytes) -> tuple[str, int] | None:
    """Dòng đầu phải là `CONNECT host:port HTTP/1.x`. Trả (host, port) hoặc None."""
    try:
        line = head.split(b"\r\n", 1)[0].decode("ascii")
    except UnicodeDecodeError:
        return None
    parts = line.split(" ")
    if len(parts) != 3 or parts[0] != "CONNECT" or not parts[2].startswith("HTTP/1."):
        return None
    host, sep, port = parts[1].rpartition(":")
    if not sep or not port.isdigit() or not host or host.startswith("["):
        return None
    return host.lower().rstrip("."), int(port)


def allowed(host: str, port: int) -> bool:
    return f"{host}:{port}" in ALLOW


def public(address: str) -> bool:
    ip = ipaddress.ip_address(address)
    return ALLOW_PRIVATE or ip.is_global


async def _resolve(host: str, port: int) -> str | None:
    loop = asyncio.get_running_loop()
    infos = await loop.getaddrinfo(host, port, type=socket.SOCK_STREAM)
    for info in infos:
        address = info[4][0]
        if public(address):
            return address
    return None


async def _pipe(reader: asyncio.StreamReader, writer: asyncio.StreamWriter, counter: list[int]) -> None:
    try:
        while True:
            chunk = await asyncio.wait_for(reader.read(65536), IDLE_SECONDS)
            if not chunk:
                break
            counter[0] += len(chunk)
            writer.write(chunk)
            await writer.drain()
    except (asyncio.TimeoutError, ConnectionError, OSError):
        pass
    finally:
        try:
            writer.close()
        except Exception:  # noqa: BLE001
            pass


async def _refuse(writer: asyncio.StreamWriter, status: str) -> None:
    writer.write(f"HTTP/1.1 {status}\r\nContent-Length: 0\r\nConnection: close\r\n\r\n".encode())
    try:
        await writer.drain()
    finally:
        writer.close()


async def handle(reader: asyncio.StreamReader, writer: asyncio.StreamWriter) -> None:
    async with _slots:
        try:
            head = await asyncio.wait_for(reader.readuntil(b"\r\n\r\n"), CONNECT_SECONDS)
        except (asyncio.TimeoutError, asyncio.IncompleteReadError, asyncio.LimitOverrunError, ConnectionError):
            writer.close()
            return
        target = parse_connect(head)
        if target is None:
            log.info("refused method")
            await _refuse(writer, "405 Method Not Allowed")
            return
        host, port = target
        if not allowed(host, port):
            log.info("refused host=%s port=%s", host[:80], port)
            await _refuse(writer, "403 Forbidden")
            return
        try:
            address = await asyncio.wait_for(_resolve(host, port), CONNECT_SECONDS)
            if address is None:
                log.info("refused private host=%s", host)
                await _refuse(writer, "403 Forbidden")
                return
            up_reader, up_writer = await asyncio.wait_for(asyncio.open_connection(address, port), CONNECT_SECONDS)
        except (asyncio.TimeoutError, OSError):
            log.info("upstream failed host=%s", host)
            await _refuse(writer, "502 Bad Gateway")
            return
        writer.write(b"HTTP/1.1 200 Connection Established\r\n\r\n")
        await writer.drain()
        sent, received = [0], [0]
        await asyncio.gather(_pipe(reader, up_writer, sent), _pipe(up_reader, writer, received))
        log.info("tunnel host=%s sent=%s received=%s", host, sent[0], received[0])


async def serve() -> None:
    server = await asyncio.start_server(handle, "0.0.0.0", PORT, limit=MAX_HEADER_BYTES)
    log.info("egress ready allow=%s", ",".join(sorted(ALLOW)))
    async with server:
        await server.serve_forever()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(message)s")
    asyncio.run(serve())
