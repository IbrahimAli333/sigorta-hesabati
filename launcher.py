"""Desktop launcher for Sığorta Hesabatı Generatoru.

Starts the Flask app under waitress (single process), opens the user's default
browser to the local URL, and stays running until the user quits the process.
Intended to be wrapped by PyInstaller into a double-clickable .app / .exe.
"""
from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser


def _find_free_port(start: int = 5050, end: int = 5099) -> int:
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("Boş port tapılmadı (5050–5099). Başqa proqramları bağlayın.")


def main() -> None:
    # Desktop mode: no password gate (the user runs it on their own machine).
    os.environ.pop("APP_PASSWORD", None)

    # Import after env-cleanup so the app initializes without auth.
    from app import app  # noqa: E402

    port = _find_free_port()
    url = f"http://127.0.0.1:{port}"

    def _open_browser_when_ready() -> None:
        # Tiny delay so the server is listening before the browser hits it.
        time.sleep(1.2)
        try:
            webbrowser.open(url, new=2)
        except Exception:
            pass

    threading.Thread(target=_open_browser_when_ready, daemon=True).start()

    print("=" * 60)
    print("  Sığorta Hesabatı Generatoru")
    print(f"  Brauzerdə açılır: {url}")
    print()
    print("  Tətbiqi dayandırmaq üçün bu pəncərəni bağlayın")
    print("  (Mac: Cmd+Q, Windows: pəncərəni X ilə bağlayın).")
    print("=" * 60)

    try:
        from waitress import serve
        serve(app, host="127.0.0.1", port=port, threads=4, _quiet=True)
    except KeyboardInterrupt:
        pass
    except Exception as e:
        print(f"\nServer xətası: {e}")
        input("Davam etmək üçün Enter basın...")
        sys.exit(1)


if __name__ == "__main__":
    main()
