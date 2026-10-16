"""Native Windows/macOS file dialogs, run in a short-lived child process.

Tk must own the main thread of its process, and the web server's request threads are not
main threads — so the dialog is shown by re-launching this program with `--dialog`.
The chosen path is passed back through a temporary file (a windowed .exe has no stdout),
which is deleted immediately.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile

MODES = ("open_source", "open_shared", "new_shared")


def _command() -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable]
    return [sys.executable, "-m", "personnel_evaluation_report"]


def ask(mode: str, initial_dir: str = "") -> str | None:
    """Show the dialog and return the selected path, or None if cancelled."""
    assert mode in MODES
    stub = os.environ.get("PER_REPORT_DEV_DIALOG_" + mode.upper())   # development/UI testing only
    if stub is not None:
        return stub or None
    fd, out = tempfile.mkstemp(prefix="per-report-dialog-", suffix=".json")
    os.close(fd)
    try:
        subprocess.run([*_command(), "--dialog", mode, initial_dir or "", out], check=False, timeout=3600)
        with open(out, encoding="utf-8") as fh:
            text = fh.read().strip()
        return json.loads(text).get("path") or None if text else None
    except (OSError, ValueError, subprocess.SubprocessError):
        return None
    finally:
        try:
            os.remove(out)
        except OSError:
            pass


def run_dialog(mode: str, initial_dir: str, out_file: str) -> None:
    """Entry point inside the child process."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    xlsx = [("فایل اکسل", "*.xlsx *.xlsm"), ("همه فایل‌ها", "*.*")]
    opts = {"parent": root, "initialdir": initial_dir or os.path.expanduser("~")}
    if mode == "open_source":
        path = filedialog.askopenfilename(title="انتخاب فایل اطلاعات پرسنلی", filetypes=xlsx, **opts)
    elif mode == "open_shared":
        path = filedialog.askopenfilename(title="انتخاب کارنامهٔ مشترک", filetypes=[("فایل اکسل", "*.xlsx")], **opts)
    else:
        path = filedialog.asksaveasfilename(title="ساخت کارنامهٔ مشترک تازه", defaultextension=".xlsx",
                                            initialfile="کارنامه پرسنلی.xlsx",
                                            filetypes=[("فایل اکسل", "*.xlsx")], **opts)
    root.destroy()
    with open(out_file, "w", encoding="utf-8") as fh:
        json.dump({"path": path or ""}, fh, ensure_ascii=False)
