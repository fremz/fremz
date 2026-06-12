# -*- mode: python ; coding: utf-8 -*-
# Fichier de configuration PyInstaller — RatiosFinanciers
# Usage : pyinstaller ratios.spec

from PyInstaller.utils.hooks import collect_all

datas_pdfplumber, binaries_pdfplumber, hiddenimports_pdfplumber = collect_all('pdfplumber')
datas_pdfminer,   binaries_pdfminer,   hiddenimports_pdfminer   = collect_all('pdfminer')
datas_cffi,       binaries_cffi,       hiddenimports_cffi        = collect_all('cffi')

a = Analysis(
    ['ratios_financiers.py'],
    pathex=[],
    binaries=binaries_pdfplumber + binaries_pdfminer + binaries_cffi,
    datas=datas_pdfplumber + datas_pdfminer + datas_cffi,
    hiddenimports=(
        hiddenimports_pdfplumber +
        hiddenimports_pdfminer   +
        hiddenimports_cffi       +
        [
            '_cffi_backend',
            'pdfminer.high_level',
            'pdfminer.layout',
            'pdfminer.converter',
            'pdfminer.pdfpage',
            'pdfminer.pdfinterp',
            'openpyxl',
            'pandas',
            'xlrd',
            'dotenv',
        ]
    ),
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
    name='RatiosFinanciers',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # Pas de fenêtre console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
