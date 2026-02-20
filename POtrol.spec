# -*- mode: python ; coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

project_root = Path(SPECPATH)

datas = [
    (str(project_root / "potrol.py"), "."),
    (str(project_root / "assets" / "potrol-logo.svg"), "assets"),
    (str(project_root / "assets" / "potrol-icon.svg"), "assets"),
    (str(project_root / "assets" / "potrol-icon.ico"), "assets"),
]
datas += copy_metadata("streamlit")
datas += collect_data_files("streamlit")
binaries = []
hiddenimports = []
hiddenimports += collect_submodules("streamlit.runtime.scriptrunner")
hiddenimports += collect_submodules("streamlit.runtime.scriptrunner_utils")
hiddenimports += ["numpy._core._exceptions", "tkinter", "tkinter.filedialog", "_tkinter"]

a = Analysis(
    ["potrol_launcher.py"],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="POtrol",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(project_root / "assets" / "potrol-icon.ico"),
)
