# PyInstaller spec for Sığorta Hesabatı Generatoru.
# Build:
#   pyinstaller sigorta.spec --clean --noconfirm
#
# Produces:
#   dist/Sigorta Hesabati.app            (macOS)
#   dist/Sigorta Hesabati/Sigorta...exe  (Windows)

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Bundle the Flask templates + static directory next to the executable.
datas = [
    ("templates", "templates"),
    ("static", "static"),
]
# openpyxl ships some data files (number formats etc.) that PyInstaller
# misses by default.
datas += collect_data_files("openpyxl")

hiddenimports = collect_submodules("openpyxl")

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "scipy", "PIL", "PyQt5", "PyQt6", "PySide2", "PySide6"],
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Sigorta Hesabati",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=sys.platform.startswith("win"),  # Windows: show console; macOS: windowed
    disable_windowed_traceback=False,
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
    name="Sigorta Hesabati",
)

# macOS .app bundle wrapper.
if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="Sigorta Hesabati.app",
        icon=None,
        bundle_identifier="az.saf.sigorta-hesabati",
        info_plist={
            "CFBundleName": "Sigorta Hesabati",
            "CFBundleDisplayName": "Sığorta Hesabatı",
            "CFBundleVersion": "1.0.0",
            "CFBundleShortVersionString": "1.0",
            "NSHighResolutionCapable": True,
            "LSBackgroundOnly": False,
        },
    )
