# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['final_rev1_autosend.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\ph10010866\\AppData\\Local\\Programs\\Python\\Python310\\Lib\\site-packages\\openpyxl', 'openpyxl'), ('C:\\Users\\ph10010866\\AppData\\Local\\Programs\\Python\\Python310\\Lib\\site-packages\\win32com', 'win32com'), ('C:\\Users\\ph10010866\\AppData\\Local\\Programs\\Python\\Python310\\Lib\\site-packages\\mysql', 'mysql'), ('C:\\Users\\ph10010866\\AppData\\Local\\Programs\\Python\\Python310\\Lib\\site-packages\\pandas', 'pandas')],
    hiddenimports=['openpyxl', 'win32com', 'mysql', 'pandas'],
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
    a.binaries,
    a.datas,
    [],
    name='final_rev1_autosend',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app.ico'],
)
