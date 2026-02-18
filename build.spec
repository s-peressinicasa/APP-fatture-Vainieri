# -*- mode: python ; coding: utf-8 -*-

import os

block_cipher = None

# PyInstaller definisce sempre SPECPATH = cartella dove si trova questo .spec
PROJECT_ROOT = os.path.abspath(SPECPATH)

ENTRY_SCRIPT = os.path.join(PROJECT_ROOT, "main.py")
ICON_ICO = os.path.join(PROJECT_ROOT, "assets", "icon.ico")

hiddenimports = []

a = Analysis(
    [ENTRY_SCRIPT],
    pathex=[PROJECT_ROOT, os.path.join(PROJECT_ROOT, "src")],
    binaries=[],
    datas=[
        (os.path.join(PROJECT_ROOT, "assets", "icon.ico"), "assets"),
        (os.path.join(PROJECT_ROOT, "assets", "icon.png"), "assets"),
        (os.path.join(PROJECT_ROOT, "src", "app", "engine", "prezzi_vainieri_2025.xlsx"), "app/engine"),
        (os.path.join(PROJECT_ROOT, "src", "app", "engine", "prezzi_vainieri_2026.xlsx"), "app/engine"),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pandas.tests",
        "numpy.tests",
        "openpyxl.tests",
        "pytest",
        
        # Web / QML / Quick
        "PySide6.QtQml",
        "PySide6.QtQuick",
        "PySide6.QtQuickWidgets",
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebEngineWidgets",
        "PySide6.QtWebEngineQuick",

        # Multimedia / device / extra
        "PySide6.QtMultimedia",
        "PySide6.QtMultimediaWidgets",
        "PySide6.QtBluetooth",
        "PySide6.QtNfc",
        "PySide6.QtSensors",
        "PySide6.QtPositioning",
        "PySide6.QtLocation",
        "PySide6.QtWebSockets",
        "PySide6.QtSerialPort",
        "PySide6.QtTextToSpeech",

        # 3D / remote objects (se presenti)
        "PySide6.Qt3DCore",
        "PySide6.Qt3DRender",
        "PySide6.Qt3DExtras",
        "PySide6.QtRemoteObjects",
    ],
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
