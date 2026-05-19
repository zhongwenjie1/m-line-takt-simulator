# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules

block_cipher = None

app_dir = Path.cwd().resolve()
project_root = app_dir.parent
main_py = app_dir / "main.py"
version_json = app_dir / "version.json"

if not main_py.exists():
    raise FileNotFoundError(f"main.py not found: {main_py}")

if not version_json.exists():
    raise FileNotFoundError(f"version.json not found: {version_json}")

datas = [
    (str(version_json), "."),
]
datas += collect_data_files("numpy")
datas += collect_data_files("pandas")

binaries = []
binaries += collect_dynamic_libs("numpy")
binaries += collect_dynamic_libs("pandas")

hiddenimports = [
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
]
hiddenimports += collect_submodules("numpy")
hiddenimports += collect_submodules("pandas")

a = Analysis(
    [str(main_py)],
    pathex=[str(project_root), str(app_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="生产指示小工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
