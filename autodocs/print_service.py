from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Callable


WORD_EXTENSIONS = {".doc", ".docx", ".rtf"}


class PrintCancelled(RuntimeError):
    pass


def open_printer_preferences(printer_name: str) -> None:
    if not printer_name:
        raise RuntimeError("Chua chon may in.")
    subprocess.Popen(
        ["rundll32.exe", "printui.dll,PrintUIEntry", "/e", "/n", printer_name],
        close_fds=True,
    )


def print_files(
    files: list[str | Path],
    printer_name: str,
    copies: int,
    delay_seconds: float,
    progress: Callable[[int, int, str], None],
    should_stop: Callable[[], bool] | None = None,
) -> None:
    if not files:
        raise RuntimeError("Chua co file de in.")
    if not printer_name:
        raise RuntimeError("Chua chon may in.")
    if copies < 1:
        raise RuntimeError("So ban in phai lon hon 0.")

    try:
        import win32api
        import win32print
    except ImportError as exc:
        raise RuntimeError("Chua cai pywin32 nen khong the in hang loat tren Windows.") from exc

    paths = [Path(path).resolve() for path in files]
    missing = [str(path) for path in paths if not path.exists()]
    if missing:
        raise RuntimeError("Khong tim thay file:\n" + "\n".join(missing[:5]))

    current_default = win32print.GetDefaultPrinter()
    word = None
    total = len(paths)
    try:
        win32print.SetDefaultPrinter(printer_name)

        for index, path in enumerate(paths, start=1):
            if should_stop and should_stop():
                raise PrintCancelled("Da huy lenh in.")

            progress(index - 1, total, f"Dang in {path.name}")
            if path.suffix.lower() in WORD_EXTENSIONS:
                word = _print_word_file(path, copies, word)
            else:
                _print_with_default_app(path, printer_name, copies, win32api, should_stop)

            progress(index, total, f"Da gui {path.name}")
            if index < total:
                time.sleep(max(0.0, delay_seconds))
    finally:
        if word is not None:
            word.Quit()
        if current_default:
            win32print.SetDefaultPrinter(current_default)


def _print_word_file(path: Path, copies: int, word):
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Chua cai pywin32 hoac khong mo duoc Microsoft Word.") from exc

    if word is None:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

    doc = None
    try:
        doc = word.Documents.Open(str(path), ReadOnly=True)
        doc.PrintOut(Background=False, Copies=copies)
    finally:
        if doc is not None:
            doc.Close(False)
    return word


def _print_with_default_app(path: Path, printer_name: str, copies: int, win32api, should_stop) -> None:
    for _copy_index in range(copies):
        if should_stop and should_stop():
            raise PrintCancelled("Da huy lenh in.")
        try:
            win32api.ShellExecute(0, "printto", str(path), f'"{printer_name}"', str(path.parent), 0)
        except Exception:
            win32api.ShellExecute(0, "print", str(path), None, str(path.parent), 0)
        time.sleep(1.0)
