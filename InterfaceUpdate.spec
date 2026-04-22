# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Leitor.py'],
    pathex=[],
    binaries=[],
    datas=[('assets\\app_icon.ico', 'assets'), ('assets\\app_logo.png', 'assets'), ('assets\\sql\\correcao_grid_localizacao_produtos.sql', 'assets\\sql')],
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
    name='Interface Update',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    uac_admin=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets\\app_icon.ico'],
)
