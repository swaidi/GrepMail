"""
Platform helpers for GrepMail: settings location, auto start, opening files.
Everything that differs between Windows and Linux lives here.
"""
import os
import sys
import subprocess

IS_WINDOWS = sys.platform.startswith("win")
IS_LINUX = sys.platform.startswith("linux")
APP_NAME = "GrepMail"


def config_dir():
    """Per-user folder for settings (never the program folder)."""
    if IS_WINDOWS:
        base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get("XDG_CONFIG_HOME", os.path.expanduser("~/.config"))
    return os.path.join(base, APP_NAME)


def default_database_dir():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    return docs if os.path.isdir(docs) else os.path.expanduser("~")


def open_path(path):
    """Open a file or folder with the system's default program."""
    if IS_WINDOWS:
        os.startfile(path)  # noqa: S606 (Windows only)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


# ---------------------------------------------------------
# Auto start
# ---------------------------------------------------------
_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def _launch_command():
    """The command that starts this program."""
    if getattr(sys, "frozen", False):
        return '"%s"' % sys.executable
    script = os.path.abspath(sys.argv[0])
    exe = sys.executable
    if IS_WINDOWS:
        pyw = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.exists(pyw):
            exe = pyw
    return '"%s" "%s"' % (exe, script)


def _desktop_file():
    base = os.environ.get("XDG_CONFIG_HOME", os.path.expanduser("~/.config"))
    return os.path.join(base, "autostart", "grepmail.desktop")


def autostart_enabled():
    try:
        if IS_WINDOWS:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY) as k:
                winreg.QueryValueEx(k, APP_NAME)
            return True
        return os.path.exists(_desktop_file())
    except Exception:
        return False


def set_autostart(enabled):
    """Turn start-with-sign-in on or off for the current user only.
    Returns (ok, message)."""
    try:
        if IS_WINDOWS:
            import winreg
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _RUN_KEY) as k:
                if enabled:
                    winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, _launch_command())
                else:
                    try:
                        winreg.DeleteValue(k, APP_NAME)
                    except FileNotFoundError:
                        pass
            return True, "ok"
        path = _desktop_file()
        if enabled:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                f.write("[Desktop Entry]\nType=Application\nName=GrepMail\n"
                        "Comment=Saves tender attachments from your mail\n"
                        "Exec=%s\nIcon=grepmail\nTerminal=false\n"
                        "X-GNOME-Autostart-enabled=true\n" % _launch_command())
        elif os.path.exists(path):
            os.remove(path)
        return True, "ok"
    except Exception as e:
        return False, str(e)
