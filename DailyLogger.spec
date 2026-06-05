# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['c:\\Users\\LixiaoKuang\\OneDrive - VisionNav Robotics USA inc\\Desktop\\Code\\daily_logger.py'],
    pathex=[],
    datas=[
        ('virtual-journal-reader/dist', 'virtual-journal-reader/dist'),
        ('virtual-journal-reader/serve_reader.py', 'virtual-journal-reader'),
    ],
    binaries=[],
    hiddenimports=[
        'qrcode',
        'qrcode.image.pil',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['imageio_ffmpeg'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DailyLogger',
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
    name='DailyLogger',
)
