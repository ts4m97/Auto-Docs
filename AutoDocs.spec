# -*- mode: python ; coding: utf-8 -*-

datas = [
    ("assets", "assets"),
]

hiddenimports = [
    "pythoncom",
    "pywintypes",
    "win32api",
    "win32print",
    "win32timezone",
    "win32com",
    "win32com.client",
    "win32com.client.dynamic",
    "fitz",
    "docx",
    "openpyxl",
]


a = Analysis(
    ["run.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "pandas",
        "matplotlib",
        "numpy",
        "scipy",
        "tkinter",
        "pytest",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="AutoDocs",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="assets/autodocs.ico",
    version="packaging/windows_version_info.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="AutoDocs",
)
