# JLI Risk Governance

A local risk-governance dashboard for equity portfolio and trade-monitoring workflows. The application runs on your computer; it is not a hosted, multi-user service.

## Run locally

Use Python 3.10 or newer, then run these commands in PowerShell from the project directory:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python setup_db.py
python -m uvicorn main:app --reload
```

Open <http://127.0.0.1:8000> in your browser. The database setup command creates the initial database schema. Portfolio reports require the relevant database and workbook inputs; upload workbooks through the dashboard or place authorized input files in the local `data` directory. Application data is intentionally not included in this repository.

## Build the Windows executable

The Windows packaging script expects an initialized `data\risk.db`. To include the current trade-monitoring workbook and portfolio data in the package, place the authorized files in `data` before building. Then run:

```powershell
.\build_exe.ps1
```

The script installs the build dependencies into `.venv`, creates the executable under `dist`, and writes a ZIP package under `release`. These generated files and any bundled data are excluded from Git.

See [BUILD_WINDOWS_EXE.md](BUILD_WINDOWS_EXE.md) for details about using and distributing the packaged app.
