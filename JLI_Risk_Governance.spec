from pathlib import Path

from PyInstaller.building.build_main import Analysis, PYZ
from PyInstaller.building.api import COLLECT, EXE


ROOT = Path(SPECPATH)
datas = [(str(ROOT / "templates"), "templates")]
if (ROOT / "static").is_dir():
    datas.append((str(ROOT / "static"), "static"))
datas.append((str(ROOT / "data" / "risk.db"), "data"))
datas.extend(
    (str(workbook), "data")
    for workbook in sorted((ROOT / "data").glob("*.xlsx"))
)
if (ROOT / "images").is_dir():
    datas.append((str(ROOT / "images"), "images"))

analysis = Analysis(
    [str(ROOT / "desktop_app.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
python_archive = PYZ(analysis.pure)
executable = EXE(
    python_archive,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="JLI Risk Governance",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
)
COLLECT(
    executable,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    name="JLI Risk Governance",
)
