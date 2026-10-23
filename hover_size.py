"""Windows File Explorer hover-size tooltip.

Run with Start_Hover_Size.bat on Windows 10/11.
Requires uiautomation and pywin32. Does not modify or delete user files.
"""
from __future__ import annotations

import os
import queue
import stat
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

if sys.platform != "win32":
    raise SystemExit("This tool runs on Windows 10/11 only.")

import pythoncom
import uiautomation as auto
import win32api
import win32com.client
import win32con
import win32gui
import tkinter as tk

HOVER_DELAY = 0.50
POLL_INTERVAL = 0.10
CACHE_SECONDS = 120
MAX_CACHE_ENTRIES = 500
REPARSE_POINT = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
WINDOW_CLASSES = {"CabinetWClass", "ExploreWClass"}


def human_size(size: int) -> str:
    units = ("B", "KB", "MB", "GB", "TB", "PB")
    n = float(size)
    for unit in units:
        if n < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(n):,} B"
            return f"{n:,.2f} {unit}"
        n /= 1024
    return f"{size:,} B"


def short_name(path: str, max_len: int = 53) -> str:
    name = os.path.basename(path.rstrip("\\/")) or path
    return name if len(name) <= max_len else name[: max_len - 1] + "…"


@dataclass(frozen=True)
class FolderResult:
    size: int
    files: int
    folders: int
    skipped: int


class CancelScan(Exception):
    pass


def count_folder(path: str, still_current, progress) -> FolderResult:
    """Logical size of files under path, not Windows 'size on disk'.

    Never traverse symlinks/junctions/reparse-point directories; scan can be
    cancelled promptly when the pointer changes.
    """
    total = files = folders = skipped = examined = 0
    todo = [path]
    last_update = time.monotonic()
    while todo:
        if not still_current():
            raise CancelScan()
        directory = todo.pop()
        try:
            with os.scandir(directory) as entries:
                for item in entries:
                    examined += 1
                    if examined % 64 == 0 and not still_current():
                        raise CancelScan()
                    try:
                        info = item.stat(follow_symlinks=False)
                        if stat.S_ISDIR(info.st_mode):
                            folders += 1
                            # Junctions and other reparse points may loop outside
                            # the hovered folder. Count the directory but don't enter it.
                            if not (getattr(info, "st_file_attributes", 0) & REPARSE_POINT):
                                todo.append(item.path)
                            else:
                                skipped += 1
                        elif stat.S_ISREG(info.st_mode):
                            files += 1
                            total += info.st_size
                        # Link metadata is excluded from the size total.
                    except (OSError, PermissionError):
                        skipped += 1
                    now = time.monotonic()
                    if now - last_update > 0.40:
                        progress(FolderResult(total, files, folders, skipped))
                        last_update = now
        except CancelScan:
            raise
        except (OSError, PermissionError):
            skipped += 1
    return FolderResult(total, files, folders, skipped)


def get_explorer_item_name(x: int, y: int) -> Optional[str]:
    """Read the icon/row under the cursor using Windows UI Automation."""
    control = auto.ControlFromPoint(x, y)
    for _ in range(12):
        if not control:
            return None
        try:
            kind = control.ControlTypeName
            if kind in ("ListItemControl", "DataItemControl"):
                name = control.Name.strip()
                # Ignore empty items or placeholder/navigation rows.
                return name if name and name not in ("Items View",) else None
            if kind in ("WindowControl", "PaneControl") and control.Name == "Desktop":
                return None
            control = control.GetParentControl()
        except (RuntimeError, OSError, ValueError):
            return None
    return None


def resolve_displayed_file(folder_path: str, display_name: str) -> Optional[str]:
    """Resolve Explorer's display name to an actual path.

    Explorer can hide file extensions; match by stem only when unambiguous.
    """
    if not folder_path or folder_path.startswith("::") or not os.path.isdir(folder_path):
        return None
    name = display_name.strip()
    if not name or name in (".", "..") or any(ch in name for ch in "\\/\x00"):
        return None
    direct = os.path.join(folder_path, name)
    if os.path.lexists(direct):
        return direct
    matches = []
    try:
        with os.scandir(folder_path) as items:
            for item in items:
                # Explorer can display the name without a known extension.
                actual = item.name
                if actual.casefold() == name.casefold():
                    return item.path
                if os.path.splitext(actual)[0].casefold() == name.casefold():
                    matches.append(item.path)
                    if len(matches) > 1:
                        return None  # A hidden extension makes it ambiguous.
    except (PermissionError, OSError):
        return None
    return matches[0] if len(matches) == 1 else None


def hover_path_at(point: tuple[int, int], shell) -> Optional[str]:
    x, y = point
    try:
        hovered_hwnd = win32gui.WindowFromPoint(point)
        if not hovered_hwnd:
            return None
        explorer_hwnd = win32gui.GetAncestor(hovered_hwnd, 2)  # GA_ROOT
        if win32gui.GetClassName(explorer_hwnd) not in WINDOW_CLASSES:
            return None
        display_name = get_explorer_item_name(x, y)
        if not display_name:
            return None
        # The Explorer COM interface supplies the real directory (not the
        # address bar title). Several tab windows can sometimes share HWND.
        for window in shell.Windows():
            try:
                if int(window.HWND) != explorer_hwnd:
                    continue
                directory = str(window.Document.Folder.Self.Path)
                # Shell ParseName may correctly resolve localized display names.
                try:
                    item = window.Document.Folder.ParseName(display_name)
                    if item is not None:
                        target = str(item.Path)
                        if os.path.lexists(target):
                            return target
                except (AttributeError, OSError, pythoncom.com_error):
                    pass
                target = resolve_displayed_file(directory, display_name)
                if target:
                    return target
            except (AttributeError, OSError, pythoncom.com_error):
                continue
    except (OSError, RuntimeError, ValueError, pythoncom.com_error):
        pass
    return None


class HoverSize:
    def __init__(self):
        self.stop = threading.Event()
        self.events = queue.Queue()
        self.jobs = queue.Queue()
        self.lock = threading.Lock()
        self.token = 0
        self.target: Optional[str] = None
        self.visible_token = -1
        self.cache = {}  # path -> (timestamp, FolderResult)

        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title("Hover Folder Size")
        self.tip = tk.Toplevel(self.root)
        self.tip.withdraw()
        self.tip.overrideredirect(True)
        self.tip.attributes("-topmost", True)
        self.tip.configure(bg="#5E708A", padx=1, pady=1)
        self.label = tk.Label(
            self.tip, text="", justify="left", anchor="w",
            bg="#182233", fg="#F6F8FC",
            font=("Segoe UI", 10), padx=12, pady=9,
        )
        self.label.pack()
        self.tip.update_idletasks()
        try:
            # Do not activate or intercept clicks on Explorer items.
            hwnd = win32gui.GetParent(self.tip.winfo_id())
            if not hwnd:
                hwnd = self.tip.winfo_id()
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            ex_style |= (win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_NOACTIVATE |
                         win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)
        except (OSError, ValueError):
            pass  # Offset tooltip placement still avoids the pointer.
        self.root.after(65, self.drain_events)
        self.root.protocol("WM_DELETE_WINDOW", self.shutdown)

    def get_token(self) -> int:
        with self.lock:
            return self.token

    def new_request(self, path: Optional[str]) -> int:
        with self.lock:
            self.token += 1
            self.target = path
            return self.token

    def hover_worker(self):
        pythoncom.CoInitialize()
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            last = None
            stable_since = time.monotonic()
            checked = False
            while not self.stop.is_set():
                try:
                    cursor = win32gui.GetCursorPos()
                    now = time.monotonic()
                    if cursor != last:
                        last = cursor
                        stable_since = now
                        checked = False
                        token = self.new_request(None)
                        self.events.put(("hide", token, None, None))
                    elif not checked and now - stable_since >= HOVER_DELAY:
                        checked = True
                        path = hover_path_at(cursor, shell)
                        if path:
                            token = self.new_request(path)
                            self.events.put(("show", token, path, cursor))
                            if os.path.isdir(path):
                                self.jobs.put((token, path))
                    self.stop.wait(POLL_INTERVAL)
                except Exception:
                    # Windows UI Automation can temporarily fail as Explorer
                    # redraws/closes. Recover at the next pointer movement.
                    checked = True
                    self.stop.wait(0.25)
        finally:
            pythoncom.CoUninitialize()

    def scan_worker(self):
        while not self.stop.is_set():
            try:
                token, path = self.jobs.get(timeout=0.3)
            except queue.Empty:
                continue
            if token != self.get_token():
                continue
            cached = self.cache.get(path)
            if cached and time.monotonic() - cached[0] < CACHE_SECONDS:
                self.events.put(("result", token, path, cached[1]))
                continue
            try:
                result = count_folder(
                    path,
                    still_current=lambda: not self.stop.is_set() and token == self.get_token(),
                    progress=lambda partial: self.events.put(("progress", token, path, partial)),
                )
            except CancelScan:
                continue
            except Exception:
                self.events.put(("error", token, path, None))
                continue
            self.cache[path] = (time.monotonic(), result)
            if len(self.cache) > MAX_CACHE_ENTRIES:
                # Dictionary keys are ordered by insertion on current Python.
                self.cache.pop(next(iter(self.cache)))
            self.events.put(("result", token, path, result))

    def show_tooltip(self, point: tuple[int, int], text: str):
        self.label.configure(text=text)
        self.tip.update_idletasks()
        w, h = self.tip.winfo_reqwidth(), self.tip.winfo_reqheight()
        x, y = point
        try:
            work = win32api.GetMonitorInfo(win32api.MonitorFromPoint(point))["Work"]
            left, top, right, bottom = work
        except Exception:
            left, top, right, bottom = 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        pos_x = x + 22
        pos_y = y + 26
        if pos_x + w > right:
            pos_x = max(left, x - w - 14)
        if pos_y + h > bottom:
            pos_y = max(top, y - h - 16)
        self.tip.geometry(f"+{pos_x}+{pos_y}")
        self.tip.deiconify()
        self.tip.lift()

    def drain_events(self):
        try:
            while True:
                kind, token, path, payload = self.events.get_nowait()
                if token != self.get_token():
                    continue
                if kind == "hide":
                    self.tip.withdraw()
                    self.visible_token = -1
                elif kind == "show":
                    cursor = payload
                    self.visible_token = token
                    name = short_name(path)
                    if os.path.isdir(path):
                        cached = self.cache.get(path)
                        if cached and time.monotonic() - cached[0] < CACHE_SECONDS:
                            details = self.format_folder(cached[1], progress=False)
                        else:
                            details = "Calculating folder size…"
                    else:
                        try:
                            details = f"Size: {human_size(os.stat(path, follow_symlinks=False).st_size)}"
                        except OSError:
                            details = "File size unavailable"
                    self.show_tooltip(cursor, f"{name}\n{details}")
                    self.last_tip_point = cursor
                elif token == self.visible_token and kind in ("progress", "result", "error"):
                    if kind == "error":
                        body = "Cannot read this folder"
                    else:
                        body = self.format_folder(payload, progress=(kind == "progress"))
                    self.show_tooltip(self.last_tip_point, f"{short_name(path)}\n{body}")
        except queue.Empty:
            pass
        if not self.stop.is_set():
            self.root.after(65, self.drain_events)

    @staticmethod
    def format_folder(data: FolderResult, progress: bool) -> str:
        if progress:
            status = f"Scanning… {human_size(data.size)} so far"
        else:
            status = f"Folder size: {human_size(data.size)}"
        second = f"{data.files:,} files  •  {data.folders:,} folders"
        if data.skipped:
            second += f"  •  {data.skipped:,} inaccessible/links skipped"
        return f"{status}\n{second}"

    def shutdown(self):
        self.stop.set()
        self.tip.withdraw()
        self.root.quit()

    def run(self):
        threading.Thread(target=self.hover_worker, daemon=True, name="hover-detect").start()
        threading.Thread(target=self.scan_worker, daemon=True, name="folder-scan").start()
        print("Hover Folder Size is running.")
        print("Hover over an item inside a File Explorer folder for 0.5 seconds.")
        print("Close this command window or press Ctrl+C to exit.")
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.shutdown()
        finally:
            self.stop.set()


if __name__ == "__main__":
    HoverSize().run()
