# -*- mode: python ; coding: utf-8 -*-
import os
import sys

# 獲取當前腳本的目錄
current_script_dir = os.path.dirname(os.path.realpath(sys.argv[0]))

block_cipher = None


a = Analysis(
    ['main.py'],
    pathex=[current_script_dir],
    binaries=[],
    datas=[('word_font\*','word_font'),
            ('fingerCache', 'fingerCache')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ShengHongFinger',
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
)

