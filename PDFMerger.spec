# -*- mode: python ; coding: utf-8 -*-
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect all marker-pdf and surya data files
marker_datas = collect_data_files('marker')
surya_datas = collect_data_files('surya')

# Collect data files for multi-format support
try:
    ebooklib_datas = collect_data_files('ebooklib')
except:
    ebooklib_datas = []

try:
    odf_datas = collect_data_files('odf')
except:
    odf_datas = []

# Combine all data files
all_datas = marker_datas + surya_datas + ebooklib_datas + odf_datas

# Hidden imports for multiprocessing and marker dependencies
hidden_imports = [
    # Multiprocessing
    'multiprocessing',
    'multiprocessing.spawn',
    'multiprocessing.pool',
    'multiprocessing.managers',
    'torch.multiprocessing',
    # Marker/Surya
    'marker',
    'marker.converters',
    'marker.converters.pdf',
    'marker.models',
    'marker.output',
    'surya',
    'surya.settings',
    'surya.model',
    'cv2',
    # Multi-format document support
    'docx',
    'docx.shared',
    'docx.oxml',
    'docx.text',
    'docx.document',
    'odf',
    'odf.opendocument',
    'odf.text',
    'odf.teletype',
    'ebooklib',
    'ebooklib.epub',
    'striprtf',
    'striprtf.striprtf',
    'bs4',
    'bs4.builder',
    'bs4.builder._htmlparser',
    'bs4.builder._lxml',
    # XML/HTML parsing
    'lxml',
    'lxml.etree',
    'lxml._elementpath',
    'html.parser',
    'html.entities',
]

a = Analysis(
    ['pdf_merger_app.py'],
    pathex=[],
    binaries=[],
    datas=all_datas,
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
