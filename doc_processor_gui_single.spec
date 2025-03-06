# -*- mode: python ; coding: utf-8 -*-
a = Analysis(
    ['doc_processor_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['_decimal', '_hashlib', '_lzma', '_bz2', '_multiprocessing', 'pyexpat', '_ssl', '_ctypes', '_queue',
             'unittest', 'email', 'html', 'http', 'xml', 'pydoc', 'doctest', 'argparse', 
             'pickle', 'calendar', 'ftplib', 'httplib2', 'pytz', 'asyncio', 'concurrent'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='项目开发类金额校验工具',
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
)