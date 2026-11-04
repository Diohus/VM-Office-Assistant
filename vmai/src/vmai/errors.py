"""Lỗi của vmai. `code` ổn định để ứng dụng dịch thành lời thường."""
from __future__ import annotations


class VmaiError(Exception):
    def __init__(self, code: str, detail: str):
        super().__init__(detail)
        self.code = code
        self.detail = detail


class SourceFileError(VmaiError):
    """File nguồn không đọc được hoặc không có bảng phù hợp."""


class TemplateSpecError(VmaiError):
    """Mô tả biểu mẫu sai, hoặc không khớp với file."""


class FeatureNotReleased(VmaiError):
    """Phần đọc ảnh và PDF thành bảng chưa có test riêng và chưa đo xong (Q81)."""

    def __init__(self, what: str):
        super().__init__("FEATURE_NOT_RELEASED", f"{what}: chưa phát hành.")


class MissingFields(VmaiError):
    """Nguồn thiếu cột bắt buộc. Cần hỏi người dùng; chưa có đề xuất nào được tạo."""

    def __init__(self, fields: list[str], headers: list[str]):
        super().__init__("MISSING_FIELDS", f"Chưa ghép được: {', '.join(fields)}")
        self.fields = fields
        self.headers = headers
