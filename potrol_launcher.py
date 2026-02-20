from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from socket import AF_INET, SOCK_STREAM, socket

from streamlit.web.cli import main as streamlit_cli_main

APP_TITLE = "POtrol"
APP_USER_MODEL_ID = "ChampagneMetals.POtrol"
HOST = "127.0.0.1"
DEFAULT_PORT = 8501
STARTUP_TIMEOUT_SECONDS = 45
THEME_ARGS = [
    "--theme.base=light",
    "--theme.primaryColor=#0b67c2",
    "--theme.backgroundColor=#edf2f8",
    "--theme.secondaryBackgroundColor=#ffffff",
    "--theme.textColor=#0f172a",
]
DESKTOP_MODE_ENV_VAR = "POTROL_DESKTOP_MODE"


def resolve_app_script() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / "potrol.py"
    return Path(__file__).with_name("potrol.py")


def resolve_icon_path() -> Path | None:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundled_icon = Path(sys._MEIPASS) / "assets" / "potrol-icon.ico"
        if bundled_icon.exists():
            return bundled_icon
    project_icon = Path(__file__).with_name("assets") / "potrol-icon.ico"
    if project_icon.exists():
        return project_icon
    return None


def parse_mode_args(argv: list[str]) -> tuple[bool, list[str]]:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--serve", action="store_true")
    parsed, passthrough = parser.parse_known_args(argv)
    return parsed.serve, passthrough


def is_port_open(host: str, port: int) -> bool:
    with socket(AF_INET, SOCK_STREAM) as probe:
        probe.settimeout(0.3)
        return probe.connect_ex((host, port)) == 0


def choose_port(preferred_port: int = DEFAULT_PORT) -> int:
    if not is_port_open(HOST, preferred_port):
        return preferred_port

    with socket(AF_INET, SOCK_STREAM) as probe:
        probe.bind((HOST, 0))
        return int(probe.getsockname()[1])


def wait_for_streamlit(port: int, server_proc: subprocess.Popen[bytes]) -> bool:
    health_url = f"http://{HOST}:{port}/_stcore/health"
    deadline = time.monotonic() + STARTUP_TIMEOUT_SECONDS

    while time.monotonic() < deadline:
        if server_proc.poll() is not None:
            return False
        try:
            with urllib.request.urlopen(health_url, timeout=1.0) as response:
                if response.status == 200:
                    return True
        except (urllib.error.URLError, TimeoutError):
            pass
        time.sleep(0.25)

    return False


def stop_process(process: subprocess.Popen[bytes] | None) -> None:
    if process is None or process.poll() is not None:
        return

    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)


def show_error(message: str) -> None:
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(None, message, APP_TITLE, 0x10)
            return
        except Exception:
            pass
    print(message, file=sys.stderr)


def set_windows_app_id() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


def apply_windows_taskbar_icon(window: object) -> None:
    if sys.platform != "win32":
        return

    icon_path = resolve_icon_path()
    if icon_path is None:
        return

    try:
        import ctypes
        from ctypes import wintypes

        native_window = getattr(window, "native", None)
        hwnd = None
        if native_window is not None:
            handle = getattr(native_window, "Handle", None)
            if handle is not None:
                hwnd = int(handle)
            elif isinstance(native_window, int):
                hwnd = native_window
        if not hwnd:
            return

        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x0010
        SM_CXICON = 11
        SM_CYICON = 12
        SM_CXSMICON = 49
        SM_CYSMICON = 50

        user32 = ctypes.windll.user32
        load_image = user32.LoadImageW
        load_image.argtypes = [
            wintypes.HINSTANCE,
            wintypes.LPCWSTR,
            wintypes.UINT,
            ctypes.c_int,
            ctypes.c_int,
            wintypes.UINT,
        ]
        load_image.restype = wintypes.HANDLE

        send_message = user32.SendMessageW
        send_message.argtypes = [
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        send_message.restype = wintypes.LPARAM

        big_w = int(user32.GetSystemMetrics(SM_CXICON))
        big_h = int(user32.GetSystemMetrics(SM_CYICON))
        small_w = int(user32.GetSystemMetrics(SM_CXSMICON))
        small_h = int(user32.GetSystemMetrics(SM_CYSMICON))

        hicon_big = load_image(None, str(icon_path), IMAGE_ICON, big_w, big_h, LR_LOADFROMFILE)
        hicon_small = load_image(None, str(icon_path), IMAGE_ICON, small_w, small_h, LR_LOADFROMFILE)

        if hicon_big:
            send_message(hwnd, WM_SETICON, ICON_BIG, hicon_big)
        if hicon_small:
            send_message(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
    except Exception:
        pass


def run_server_mode(streamlit_args: list[str]) -> int:
    app_script = resolve_app_script()
    cli_args = [
        "streamlit",
        "run",
        str(app_script),
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--server.headless=true",
        "--server.fileWatcherType=none",
    ]
    cli_args.extend(THEME_ARGS)
    cli_args.extend(streamlit_args)
    sys.argv = cli_args
    return streamlit_cli_main()


def build_server_args(passthrough_args: list[str], port: int) -> list[str]:
    blocked_prefixes = (
        "--server.port",
        "--server.address",
        "--server.headless",
        "--theme.",
    )
    filtered_args = [
        arg for arg in passthrough_args if not any(arg.startswith(prefix) for prefix in blocked_prefixes)
    ]
    return [
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        f"--server.address={HOST}",
        f"--server.port={port}",
        "--server.headless=true",
        "--server.fileWatcherType=none",
        *THEME_ARGS,
        *filtered_args,
    ]


def build_server_command(server_args: list[str]) -> list[str]:
    if getattr(sys, "frozen", False):
        command = [sys.executable, "--serve"]
    else:
        command = [sys.executable, str(Path(__file__).resolve()), "--serve"]
    command.extend(server_args)
    return command


def run_desktop_mode(passthrough_args: list[str]) -> int:
    port = choose_port()
    server_args = build_server_args(passthrough_args, port)
    server_cmd = build_server_command(server_args)
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    server_proc: subprocess.Popen[bytes] | None = None

    try:
        server_env = os.environ.copy()
        server_env[DESKTOP_MODE_ENV_VAR] = "1"
        server_proc = subprocess.Popen(
            server_cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creation_flags,
            env=server_env,
        )

        if not wait_for_streamlit(port, server_proc):
            show_error("POtrol server failed to start.")
            return 1

        import webview

        window = webview.create_window(
            APP_TITLE,
            f"http://{HOST}:{port}/",
            width=1400,
            height=900,
            min_size=(1000, 700),
        )
        webview.start(func=apply_windows_taskbar_icon, args=(window,), gui="edgechromium")
        return 0
    except Exception as exc:
        show_error(f"POtrol failed to open desktop window.\n\n{exc}")
        return 1
    finally:
        stop_process(server_proc)


def main() -> int:
    set_windows_app_id()
    serve_mode, passthrough_args = parse_mode_args(sys.argv[1:])
    if serve_mode:
        return run_server_mode(passthrough_args)
    return run_desktop_mode(passthrough_args)


if __name__ == "__main__":
    raise SystemExit(main())
