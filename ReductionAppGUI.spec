# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_submodules,
    copy_metadata,
)

block_cipher = None

project_root = Path(".").resolve()

datas = []
hiddenimports = []


def safe_copy_metadata(name: str, recursive: bool = True):
    try:
        return copy_metadata(name, recursive=recursive)
    except Exception:
        return []


# Library data
datas += collect_data_files("docx")
datas += collect_data_files("pymorphy2")
datas += collect_data_files("pymorphy2_dicts_ru")
datas += collect_data_files("openpyxl")

# Package metadata for frozen environments
# Important: import name "docx" has distribution name "python-docx"
datas += safe_copy_metadata("python-docx")
datas += safe_copy_metadata("pymorphy2")
datas += safe_copy_metadata("pymorphy2-dicts-ru")
datas += safe_copy_metadata("DAWG-Python")
datas += safe_copy_metadata("openpyxl")
datas += safe_copy_metadata("pandas")
datas += safe_copy_metadata("numpy")

hiddenimports += collect_submodules("pymorphy2")
hiddenimports += collect_submodules("pymorphy2_dicts_ru")
hiddenimports += collect_submodules("openpyxl")

# Project data
project_data_candidates = [
    "abbreviation_database",
]
for name in project_data_candidates:
    path = project_root / name
    if path.exists():
        datas.append((str(path), name))

a = Analysis(
    ["end_user_app.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ReductionAppGUI",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ReductionAppGUI",
)
