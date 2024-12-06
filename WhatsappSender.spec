# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['WhatsappSender.py'],
    pathex=[],
    binaries=[],
    datas=[('red.json', '.'), ('whatsapp.json', '.'), ('parameters.json', '.'), ('auro.ico', '.')],
    hiddenimports=[],
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
    name='WhatsappSender',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch='x86_64',
    codesign_identity=None,
    entitlements_file=None,
    icon=['auro.ico'],
)
app = BUNDLE(
    exe,
    name='WhatsappSender.app',
    icon='auro.ico',
    bundle_identifier=None,
)
