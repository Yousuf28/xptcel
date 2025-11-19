# xpt_open_in_excel.spec

block_cipher = None

from PyInstaller.utils.hooks import collect_submodules

# Collect all needed submodules in a clean way
hiddenimports = (
    collect_submodules("pandas")
    + collect_submodules("numpy")
    + collect_submodules("pyreadstat")
)

a = Analysis(
    ['xpt_open_in_excel.py'],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# EXE for Windows and for macOS bundle payload
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='xpt_open_in_excel',
    debug=False,
    strip=False,
    upx=False,
    console=True,      # change to False for release
)

# macOS .app bundle
app = BUNDLE(
    exe,
    name='xpt_open_in_excel.app',
    bundle_identifier="com.yourname.xptopen",
    icon=None,
    info_plist={
        "CFBundleName": "xpt_open_in_excel",
        "CFBundleDisplayName": "xpt_open_in_excel",
        "CFBundleIdentifier": "com.yourname.xptopen",
        "CFBundleVersion": "1.0",
        "CFBundleShortVersionString": "1.0",
        "NSHighResolutionCapable": True,

        # File association for .xpt files
        "CFBundleDocumentTypes": [
            {
                "CFBundleTypeName": "SAS XPORT File",
                "CFBundleTypeRole": "Editor",
                "CFBundleTypeExtensions": ["xpt"],
                "LSHandlerRank": "Owner",
            }
        ],
    }
)
