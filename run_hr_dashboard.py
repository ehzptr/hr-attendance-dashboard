from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

APP_NAME = "HR Attendance Dashboard"
DEFAULT_PORT = 8501
HOST = "127.0.0.1"


def project_dir() -> Path:
    return Path(__file__).resolve().parent


def find_free_port(start: int = DEFAULT_PORT) -> int:
    for port in range(start, start + 25):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind((HOST, port))
                return port
            except OSError:
                continue
    raise RuntimeError("Tidak menemukan port lokal yang tersedia untuk dashboard.")


def wait_for_server(port: int, timeout: float = 30.0) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            try:
                sock.connect((HOST, port))
                return True
            except OSError:
                time.sleep(0.25)
    return False


def main() -> int:
    app_path = project_dir() / "streamlit_app.py"
    if not app_path.exists():
        raise FileNotFoundError(f"File aplikasi tidak ditemukan: {app_path}")

    port = find_free_port()
    url = f"http://{HOST}:{port}"

    env = os.environ.copy()
    env.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")

    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(app_path),
        "--server.address",
        HOST,
        "--server.port",
        str(port),
        "--server.headless",
        "true",
        "--server.fileWatcherType",
        "none",
        "--browser.serverAddress",
        HOST,
        "--browser.serverPort",
        str(port),
        "--browser.gatherUsageStats",
        "false",
    ]

    creationflags = 0
    if os.name == "nt":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    process = subprocess.Popen(
        command,
        cwd=project_dir(),
        env=env,
        creationflags=creationflags,
    )

    try:
        if not wait_for_server(port):
            process.terminate()
            raise RuntimeError(
                "Dashboard gagal dijalankan. Pastikan dependency berhasil disiapkan oleh uv."
            )

        webbrowser.open(url)
        return process.wait()
    except KeyboardInterrupt:
        process.terminate()
        return process.wait()
    finally:
        if process.poll() is None:
            process.terminate()


if __name__ == "__main__":
    raise SystemExit(main())
