# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("resources/default_rules.txt", "resources"),
        ("resources/sun_app.ico", "resources"),
    ],
    hiddenimports=["keyring.backends.Windows"],
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
    a.datas,
    [],
    name="Kimi雅思单词本整理器_太阳版",
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
    icon=["resources/sun_app.ico"],
)
