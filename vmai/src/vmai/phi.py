"""Dấu hiệu dữ liệu người bệnh, và cổng gửi ra ngoài (Q29).

Cổng một chiều: dò thấy thì dừng tới khi có người xử lý. Không thấy KHÔNG có nghĩa là được gửi.
Quyền gửi do chính sách của khách hàng quyết định, mặc định là không. Tên không có nhãn thì không dò được.
"""
from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

from .mapping import normalize_text

_PATTERNS = {
    "ho_ten": re.compile(r"\b(ho (va )?ten|ten (bn|benh nhan|nguoi benh)|benh nhan|nguoi benh)\b"),
    "ma_bn": re.compile(r"\b(ma (bn|benh nhan|nguoi benh)|pid|ma ho so|so ho so|sid)\b"),
    "ngay_sinh": re.compile(r"\b(ngay sinh|nam sinh|dob)\b"),
    "dinh_danh": re.compile(r"\b(cccd|cmnd|bhyt|the bao hiem|so dien thoai|dia chi)\b"),
}


# Cột về con người nói chung, kể cả nhân viên: AI không điền, giá trị không đi tiếp (câu 7 của khách hàng).
_PERSON_HEADER = re.compile(
    r"\b(nguoi|nhan vien|nv|ktv|ky thuat vien|bac si|bs|dieu duong|phu trach|ky ten|chu ky|ho (va )?ten|operator|staff|signed by)\b"
)
# Đủ để cảnh báo, không đủ để tin: tên không theo mẫu này sẽ lọt, nên lớp chính là không điền cột chữ theo mặc định.
_SURNAMES = frozenset(
    "nguyen tran le pham hoang huynh phan vu vo dang bui do ho ngo duong ly dinh truong mai cao lam".split()
)


def is_person_header(text: str) -> bool:
    return bool(_PERSON_HEADER.search(normalize_text(text)))


def looks_like_person_name(text: str) -> bool:
    words = text.split()
    if not 2 <= len(words) <= 5 or any(not w[0].isupper() or any(c.isdigit() for c in w) for w in words):
        return False
    return normalize_text(words[0]) in _SURNAMES


def phi_kind(text: str) -> str | None:
    """Loại dấu hiệu trong một chuỗi, ví dụ một tiêu đề cột."""
    norm = normalize_text(text)
    for kind, pattern in _PATTERNS.items():
        if pattern.search(norm):
            return kind
    return None


@dataclass(frozen=True)
class Token:
    text: str
    x: float = 0.0
    y: float = 0.0


@dataclass(frozen=True)
class PhiFinding:
    kind: str
    text: str


def scan_tokens(tokens: Iterable[Token]) -> list[PhiFinding]:
    """Dò trên chữ OCR nội bộ. Xét cả chuỗi ghép theo dòng vì engine có thể tách nhãn."""
    ordered = list(tokens)
    findings = [PhiFinding(kind, t.text) for t in ordered if (kind := phi_kind(t.text))]
    joined = " ".join(t.text for t in sorted(ordered, key=lambda t: (round(t.y / 20), t.x)))
    norm = normalize_text(joined)
    for kind, pattern in _PATTERNS.items():
        if pattern.search(norm) and not any(f.kind == kind for f in findings):
            findings.append(PhiFinding(kind, "(ghép token)"))
    return findings


class SendNotAllowed(Exception):
    """Chính sách chưa cho gửi loại dữ liệu này ra ngoài."""


class PhiDetected(Exception):
    """Có dấu hiệu dữ liệu người bệnh: dừng gửi tới khi có người xử lý."""

    def __init__(self, findings: list[PhiFinding]):
        super().__init__(f"Có dấu hiệu dữ liệu người bệnh: {sorted({f.kind for f in findings})}")
        self.findings = findings


@dataclass(frozen=True)
class ExternalSendPolicy:
    """Do khách hàng quyết định (câu 2). Mặc định không gửi gì."""

    allow_images: bool = False
    allow_documents: bool = False  # tài liệu hãng sample, Q66


def guard_external_image(tokens: Iterable[Token], policy: ExternalSendPolicy) -> None:
    """Gọi trước khi gửi ảnh ra ngoài. Qua được hàm này chỉ có nghĩa là chính sách cho phép và không dò thấy gì."""
    if not policy.allow_images:
        raise SendNotAllowed("Chính sách hiện tại không cho gửi ảnh ra ngoài")
    findings = scan_tokens(tokens)
    if findings:
        raise PhiDetected(findings)
