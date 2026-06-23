# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

block_cipher = None

pyqt6_datas, pyqt6_binaries, pyqt6_hiddenimports = collect_all("PyQt6")
extra_datas = [("WINDOWS-使用说明.txt", ".")]

a = Analysis(
    ["converter_gui.py"],
    pathex=[],
    binaries=pyqt6_binaries,
    datas=pyqt6_datas + extra_datas,
    hiddenimports=[
        "converter",
        "openpyxl",
        "openpyxl.cell._writer",
        "openpyxl.worksheet._reader",
        *pyqt6_hiddenimports,
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="trans-converter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="trans-converter",
)
