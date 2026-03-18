"""
Windows system app helpers.
Opens simple built-in applications like Notepad and Calculator.
"""
import ctypes
import os
import re
import subprocess
import time
from ctypes import wintypes

try:
    import pythoncom
    import win32com.client
except Exception:
    pythoncom = None
    win32com = None


_APP_ALIASES = {
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "calc": "calc.exe",
}


def _canonical_app_name(app_name: str) -> tuple[str, str | None]:
    cleaned = (app_name or "").strip().lower()
    command = _APP_ALIASES.get(cleaned)
    canonical = "calculator" if command == "calc.exe" else cleaned
    return canonical, command


def _open_if_needed(command: str):
    process = subprocess.Popen([command], cwd=os.getcwd())
    time.sleep(1.0)
    return process


def _find_window_titles_containing(fragment: str) -> list[str]:
    fragment = (fragment or "").strip().lower()
    if not fragment:
        return []

    titles = []
    user32 = ctypes.windll.user32

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def callback(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        title = buffer.value.strip()
        if title and fragment in title.lower():
            titles.append(title)
        return True

    user32.EnumWindows(callback, 0)
    return titles


def _activate_window(target_titles: list[str] | tuple[str, ...] | None = None, pid: int | None = None, timeout: float = 5.0) -> bool:
    if pythoncom is None or win32com is None:
        raise RuntimeError("pywin32 is required for system app typing automation.")

    target_titles = [title for title in (target_titles or []) if title]
    deadline = time.time() + max(timeout, 0.5)

    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        while time.time() < deadline:
            if pid is not None:
                try:
                    if shell.AppActivate(pid):
                        time.sleep(0.4)
                        return True
                except Exception:
                    pass

            for title in target_titles:
                try:
                    if shell.AppActivate(title):
                        time.sleep(0.4)
                        return True
                except Exception:
                    pass

                for partial_title in _find_window_titles_containing(title):
                    try:
                        if shell.AppActivate(partial_title):
                            time.sleep(0.4)
                            return True
                    except Exception:
                        continue

            time.sleep(0.35)
        return False
    finally:
        pythoncom.CoUninitialize()


def _send_keys_to_window(window_title: str | list[str] | tuple[str, ...], keys: str, pid: int | None = None, timeout: float = 5.0) -> bool:
    if pythoncom is None or win32com is None:
        raise RuntimeError("pywin32 is required for system app typing automation.")

    target_titles = [window_title] if isinstance(window_title, str) else list(window_title or [])

    pythoncom.CoInitialize()
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        if not _activate_window(target_titles=target_titles, pid=pid, timeout=timeout):
            return False
        time.sleep(0.3)
        shell.SendKeys(keys)
        return True
    finally:
        pythoncom.CoUninitialize()


def _normalize_calc_expression(expression: str) -> str:
    expr = (expression or "").strip().replace("x", "*").replace("X", "*").replace("Ã·", "/")
    return expr


def _safe_eval_expression(expression: str) -> str:
    expr = _normalize_calc_expression(expression)
    if not re.fullmatch(r"[0-9\.\+\-\*/\(\) ]+", expr):
        return ""
    try:
        value = eval(expr, {"__builtins__": {}}, {})
        return str(value)
    except Exception:
        return ""


def open_system_app(app_name: str) -> str:
    """Open a supported system app."""
    canonical, command = _canonical_app_name(app_name)
    if not command:
        supported = ", ".join(sorted(set(_APP_ALIASES.keys())))
        return f"❌ Unsupported system app '{app_name}'. Supported apps: {supported}"

    try:
        subprocess.Popen([command], cwd=os.getcwd())
        return f"✅ Opened {canonical}"
    except Exception as exc:
        return f"❌ Error opening {app_name}: {exc}"


def close_system_app(app_name: str) -> str:
    """Close a supported system app."""
    canonical, command = _canonical_app_name(app_name)
    if not command:
        supported = ", ".join(sorted(set(_APP_ALIASES.keys())))
        return f"❌ Unsupported system app '{app_name}'. Supported apps: {supported}"

    try:
        completed = subprocess.run(
            ["taskkill", "/IM", command, "/F"],
            capture_output=True,
            text=True,
            check=False,
        )
        if completed.returncode == 0:
            return f"✅ Closed {canonical}"
        stderr = (completed.stderr or completed.stdout or "").strip()
        return f"❌ Could not close {app_name}: {stderr or 'process not running'}"
    except Exception as exc:
        return f"❌ Error closing {app_name}: {exc}"


def write_in_system_app(app_name: str, text: str) -> str:
    """Type text into a supported system app window."""
    canonical, command = _canonical_app_name(app_name)
    if canonical != "notepad":
        return "❌ Writing is currently supported only for Notepad."
    if not text:
        return "❌ Please provide text to write."

    try:
        process = _open_if_needed(command)
        if not _send_keys_to_window(
            ["Untitled - Notepad", "Notepad"],
            text,
            pid=getattr(process, "pid", None),
            timeout=5.0,
        ):
            return "❌ Could not find the Notepad window to type into."
        return "✅ Wrote text in Notepad"
    except Exception as exc:
        return f"❌ Error writing in {canonical}: {exc}"


def calculate_in_calculator(expression: str) -> str:
    """Open Calculator, enter a simple arithmetic expression, and return the result."""
    canonical, command = _canonical_app_name("calculator")
    expr = _normalize_calc_expression(expression)
    if not expr:
        return "❌ Please provide an expression to calculate."

    try:
        process = _open_if_needed(command)
        activated = _send_keys_to_window(
            ["Calculator", "Calc"],
            "{ESC}" + expr + "=",
            pid=getattr(process, "pid", None),
            timeout=6.0,
        )
        if not activated:
            return "❌ Could not find the Calculator window to enter the expression."
        result = _safe_eval_expression(expr)
        if result:
            return f"✅ Calculated {expr} = {result} in Calculator"
        return f"✅ Entered {expr} in Calculator"
    except Exception as exc:
        return f"❌ Error using {canonical}: {exc}"
