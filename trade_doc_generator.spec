# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs, collect_submodules


project = Path(SPECPATH)

datas = [
    (str(project / "templates"), "templates"),
    (str(project / "resources" / "templates"), "resources/templates"),
]

optional_datas = [
    (project / "resources" / ".ai_key", "resources"),
    (project / "resources" / "LibreOfficePortable", "resources/LibreOfficePortable"),
    (project / "resources" / "libreoffice", "resources/libreoffice"),
]
for src, dest in optional_datas:
    if src.exists():
        datas.append((str(src), dest))

hiddenimports = []
hiddenimports += collect_submodules("rapidocr_onnxruntime")
hiddenimports += collect_submodules("onnxruntime")
hiddenimports += ["cv2", "fitz", "waitress"]

binaries = []
binaries += collect_dynamic_libs("onnxruntime")

datas += collect_data_files("rapidocr_onnxruntime")
datas += collect_data_files("onnxruntime")


a = Analysis(
    ["app.py"],
    pathex=[str(project)],
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
    [],
    exclude_binaries=True,
    name="TradeDocGenerator",
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="TradeDocGenerator",
)
