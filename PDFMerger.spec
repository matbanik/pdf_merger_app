# -*- mode: python ; coding: utf-8 -*-
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect all marker-pdf and surya data files
marker_datas = collect_data_files('marker')
surya_datas = collect_data_files('surya')

# Hidden imports for multiprocessing and marker dependencies
hidden_imports = [
    'multiprocessing',
    'multiprocessing.spawn',
    'multiprocessing.pool',
    'multiprocessing.managers',
    'torch.multiprocessing',
    'marker',
    'marker.converters',
    'marker.converters.pdf',
    'marker.models',
    'marker.output',
    'surya',
    'surya.settings',
    'surya.model',
    'cv2',
]

a = Analysis(
    ['pdf_merger_app.py'],
    pathex=[],
    binaries=[],
    datas=marker_datas + surya_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', '_pytest', 'pluggy', 'py', 'iniconfig'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PDFMerger',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='PDFMerger',
)
