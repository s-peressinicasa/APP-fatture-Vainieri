# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# PyInstaller definisce sempre SPECPATH = cartella dove si trova questo .spec
PROJECT_ROOT = os.path.abspath(SPECPATH)

ENTRY_SCRIPT = os.path.join(PROJECT_ROOT, "src", "app", "main.py")
ICON_ICO = os.path.join(PROJECT_ROOT, "assets", "icon.ico")

hiddenimports = []
hiddenimports += collect_submodules("pandas")
hiddenimports += collect_submodules("openpyxl")
hiddenimports += collect_submodules("pdfplumber")

a = Analysis(
    [ENTRY_SCRIPT],
    pathex=[PROJECT_ROOT],
    binaries=[],
    datas=[
        (os.path.join(PROJECT_ROOT, "assets", "icon.ico"), "assets"),
        (os.path.join(PROJECT_ROOT, "assets", "icon.png"), "assets"),
    ],
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
    [],
    exclude_binaries=True,
    name="ControlloFattureVainieri",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=ICON_ICO,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="ControlloFattureVainieri",
)
