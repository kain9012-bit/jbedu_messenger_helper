# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['jbedu_messenger_helper(v3.0).py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['win32timezone', 'openpyxl', 'PIL.Image', 'PIL.ImageTk', 'PIL.BmpImagePlugin', 'PIL.PngImagePlugin', 'PIL.JpegImagePlugin', 'PIL.GifImagePlugin'],
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
    name='전북특별자치도교육청_메신저 작업 도우미(v3.0)',
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
    icon='assets\\app-icon.ico',
)
