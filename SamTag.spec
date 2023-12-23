# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['./sams_venv/lib/python3.11/site-packages'],
    binaries=[],
    datas=[('./src/*.pdf', 'src'), ('./src/fonts/*.ttf', 'src/fonts'), ('./src/app_icon/*', 'src/app_icon')],
    hiddenimports=['pkg_resources.py2_warn'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SamTag',
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
    icon=['src/app_icon/SamTag.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SamTag',
)
app = BUNDLE(
    coll,
    name='SamTag.app',
    icon='./src/app_icon/SamTag.ico',
    bundle_identifier=None,
)
