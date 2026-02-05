# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for Liturgie Samensteller."""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Collect all data files including PyQt6 plugins
datas = [
    ('src/i18n/*.json', 'src/i18n'),  # Translation files
] + collect_data_files('PyQt6', include_py_files=False)

# Collect all PyQt6 submodules to ensure nothing is missed
pyqt6_submodules = collect_submodules('PyQt6')

# Hidden imports that PyInstaller might miss
hiddenimports = pyqt6_submodules + [
    # PyQt6 core modules (explicit for safety)
    'PyQt6.QtCore',
    'PyQt6.QtGui',
    'PyQt6.QtWidgets',
    'PyQt6.QtPrintSupport',
    'PyQt6.sip',
    # python-pptx
    'pptx',
    'pptx.util',
    'pptx.enum.text',
    'pptx.enum.shapes',
    'pptx.enum.dml',
    'pptx.oxml',
    'pptx.oxml.ns',
    # lxml
    'lxml',
    'lxml._elementpath',
    'lxml.etree',
    # Excel
    'openpyxl',
    'openpyxl.cell',
    'openpyxl.workbook',
    # Windows COM
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',
    # Other
    'yt_dlp',
    'requests',
    'PIL',
    'PIL.Image',
]

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'cv2',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='LiturgieSamensteller',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon path here if you have one: icon='icon.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='LiturgieSamensteller',
)
