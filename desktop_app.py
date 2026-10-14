import os
from pathlib import Path
import shutil
import socket
import sys
from threading import Timer
import webbrowser


def _prepare_packaged_data() -> None:
    if not getattr(sys, "frozen", False):
        return

    resource_dir = Path(sys._MEIPASS)
    local_app_data = Path(
        os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
    )
    data_dir = local_app_data / "JLI Risk Governance" / "data"
    bundled_data_dir = resource_dir / "data"
    data_dir.mkdir(parents=True, exist_ok=True)

    for source in bundled_data_dir.iterdir():
        target = data_dir / source.name
        if source.is_file() and not target.exists():
            shutil.copy2(source, target)

    os.environ["JLI_DATA_DIR"] = str(data_dir)
    sys.path.insert(0, str(resource_dir))


def _available_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as listener:
        listener.bind(("127.0.0.1", 0))
        return int(listener.getsockname()[1])


def main() -> None:
    _prepare_packaged_data()

    import uvicorn
    from main import app

    port = _available_port()
    url = f"http://127.0.0.1:{port}"
    print("JLI Risk Governance is starting.")
    print(f"Dashboard: {url}")
    print("Keep this window open while using the dashboard.")
    Timer(1.5, webbrowser.open, args=(url,)).start()
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")


if __name__ == "__main__":
    main()
