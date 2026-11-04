"""Mở Excel thật để đối chiếu, trên máy Windows. Không phải một phần của sản phẩm.

Hai ràng buộc từ sự cố ở M0 (2 tiến trình Excel ẩn còn sót, giữ khóa fixture):
- Excel luôn được tắt khi xong, kể cả khi script lỗi hoặc treo. Chỉ tắt đúng tiến trình do script này tạo.
- Chỉ mở file nằm trong thư mục làm việc tạm, không bao giờ mở file gốc trong fixtures.
"""
from __future__ import annotations

import contextlib
import pathlib
import shutil
import tempfile
import threading
import time

import pythoncom
import pywintypes
import win32api
import win32con
import win32com.client as win32
import win32event
import win32process

MSO_DISABLE_MACROS = 3
QUIT_WAIT_MS = 15000


def _terminate(pid: int) -> bool:
    """Tắt đúng một tiến trình. Không có gì tắt theo tên, để không đụng Excel khác trên máy."""
    try:
        handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid)
    except pywintypes.error:
        return False  # đã tự thoát
    try:
        win32api.TerminateProcess(handle, 1)
        win32event.WaitForSingleObject(handle, QUIT_WAIT_MS)
        return True
    finally:
        win32api.CloseHandle(handle)


def _wait_exit(pid: int, ms: int) -> bool:
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except pywintypes.error:
        return True
    try:
        return win32event.WaitForSingleObject(handle, ms) == win32event.WAIT_OBJECT_0
    finally:
        win32api.CloseHandle(handle)


class WorkDir:
    """Thư mục tạm chứa bản sao. Excel chỉ được mở file trong đây."""

    def __init__(self) -> None:
        self.path = pathlib.Path(tempfile.mkdtemp(prefix="vm_excel_check_"))

    def copy_in(self, src: pathlib.Path, name: str | None = None) -> pathlib.Path:
        dst = self.path / (name or src.name)
        shutil.copyfile(src, dst)
        return dst

    def write(self, name: str, data: bytes) -> pathlib.Path:
        dst = self.path / name
        dst.write_bytes(data)
        return dst

    def require_inside(self, path: pathlib.Path) -> pathlib.Path:
        path = pathlib.Path(path).resolve()
        if not path.is_relative_to(self.path.resolve()):
            raise ValueError(f"Chỉ được mở bản sao trong thư mục tạm, không mở: {path}")
        return path

    def remove(self) -> None:
        shutil.rmtree(self.path, ignore_errors=True)


class Excel:
    def __init__(self, xl, work: WorkDir, pid: int) -> None:
        self.app = xl
        self.work = work
        self.pid = pid

    def open(self, path: pathlib.Path):
        """Mở thường, không cập nhật liên kết. Trả (workbook, lỗi); Excel tự sửa file thì lỗi là 'repaired'."""
        path = self.work.require_inside(path)
        try:
            wb = self.app.Workbooks.Open(str(path), 0)
        except pywintypes.com_error as exc:
            return None, f"open-failed: {exc}"
        title = wb.Windows(1).Caption if wb.Windows.Count else wb.Name
        return wb, ("repaired" if "Repaired" in str(title) or "Repaired" in wb.Name else None)

    def open_protected(self, path: pathlib.Path):
        path = self.work.require_inside(path)
        return self.app.ProtectedViewWindows.Open(str(path))


@contextlib.contextmanager
def excel(limit_seconds: int = 300):
    """Một phiên Excel riêng, macro tắt. Quá `limit_seconds` thì tắt đúng phiên này, script báo lỗi ở lời gọi COM kế tiếp."""
    work = WorkDir()
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    done = threading.Event()

    def watchdog() -> None:
        if not done.wait(limit_seconds):
            print(f"Quá {limit_seconds} giây: tắt Excel PID {pid} do script này tạo", flush=True)
            _terminate(pid)

    threading.Thread(target=watchdog, daemon=True).start()
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    xl.AutomationSecurity = MSO_DISABLE_MACROS
    try:
        yield Excel(xl, work, pid)
    finally:
        done.set()
        with contextlib.suppress(Exception):
            for wb in list(xl.Workbooks):
                wb.Close(False)
        with contextlib.suppress(Exception):
            for w in list(xl.ProtectedViewWindows):
                w.Close()
        with contextlib.suppress(Exception):
            xl.Quit()
        del xl
        pythoncom.CoUninitialize()
        if not _wait_exit(pid, QUIT_WAIT_MS):
            _terminate(pid)
        time.sleep(0.2)
        work.remove()
