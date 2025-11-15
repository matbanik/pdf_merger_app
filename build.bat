@echo off
echo Building PDF Merger with Multiprocessing Fix
echo =============================================

rem Clean previous builds
echo Cleaning...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "*.spec" del "*.spec" 2>nul

rem Create a custom spec file with proper multiprocessing and multi-format support
echo Creating custom spec file with multiprocessing and multi-format support...
(
echo # -*- mode: python ; coding: utf-8 -*-
echo import sys
echo from PyInstaller.utils.hooks import collect_data_files, collect_submodules
echo.
echo # Collect all marker-pdf and surya data files
echo marker_datas = collect_data_files^('marker'^)
echo surya_datas = collect_data_files^('surya'^)
echo.
echo # Collect data files for multi-format support
echo try:
echo     ebooklib_datas = collect_data_files^('ebooklib'^)
echo except:
echo     ebooklib_datas = []
echo.
echo try:
echo     odf_datas = collect_data_files^('odf'^)
echo except:
echo     odf_datas = []
echo.
echo # Combine all data files
echo all_datas = marker_datas + surya_datas + ebooklib_datas + odf_datas
echo.
echo # Hidden imports for multiprocessing and marker dependencies
echo hidden_imports = [
echo     # Multiprocessing
echo     'multiprocessing',
echo     'multiprocessing.spawn',
echo     'multiprocessing.pool',
echo     'multiprocessing.managers',
echo     'torch.multiprocessing',
echo     # Marker/Surya
echo     'marker',
echo     'marker.converters',
echo     'marker.converters.pdf',
echo     'marker.models',
echo     'marker.output',
echo     'surya',
echo     'surya.settings',
echo     'surya.model',
echo     'cv2',
echo     # Multi-format document support
echo     'docx',
echo     'docx.shared',
echo     'docx.oxml',
echo     'docx.text',
echo     'docx.document',
echo     'odf',
echo     'odf.opendocument',
echo     'odf.text',
echo     'odf.teletype',
echo     'ebooklib',
echo     'ebooklib.epub',
echo     'striprtf',
echo     'striprtf.striprtf',
echo     'bs4',
echo     'bs4.builder',
echo     'bs4.builder._htmlparser',
echo     'bs4.builder._lxml',
echo     # XML/HTML parsing
echo     'lxml',
echo     'lxml.etree',
echo     'lxml._elementpath',
echo     'html.parser',
echo     'html.entities',
echo ]
echo.
echo a = Analysis^(
echo     ['pdf_merger_app.py'],
echo     pathex=[],
echo     binaries=[],
echo     datas=all_datas,
echo     hiddenimports=hidden_imports,
echo     hookspath=[],
echo     hooksconfig={},
echo     runtime_hooks=[],
echo     excludes=['pytest', '_pytest', 'pluggy', 'py', 'iniconfig'],
echo     win_no_prefer_redirects=False,
echo     win_private_assemblies=False,
echo     cipher=None,
echo     noarchive=False,
echo ^)
echo.
echo pyz = PYZ^(a.pure, a.zipped_data, cipher=None^)
echo.
echo exe = EXE^(
echo     pyz,
echo     a.scripts,
echo     [],
echo     exclude_binaries=True,
echo     name='PDFMerger',
echo     debug=False,
echo     bootloader_ignore_signals=False,
echo     strip=False,
echo     upx=False,
echo     console=False,
echo     disable_windowed_traceback=False,
echo     argv_emulation=False,
echo     target_arch=None,
echo     codesign_identity=None,
echo     entitlements_file=None,
echo ^)
echo.
echo coll = COLLECT^(
echo     exe,
echo     a.binaries,
echo     a.zipfiles,
echo     a.datas,
echo     strip=False,
echo     upx=False,
echo     upx_exclude=[],
echo     name='PDFMerger',
echo ^)
) > PDFMerger.spec

echo Building with multiprocessing support...
python -m PyInstaller --noconfirm PDFMerger.spec

echo.
if exist "dist\PDFMerger\PDFMerger.exe" (
    echo BUILD SUCCESS!
    echo Executable: dist\PDFMerger\PDFMerger.exe
    echo.
    echo IMPORTANT: Multiprocessing support enabled
    echo Note: Console window is hidden - check app's console output widget for progress
    echo Note: First run will download AI models ^(~1-2GB^)
    echo.
    echo Multi-format support included:
    echo   - Input formats: PDF, ODT, DOCX, TXT, RTF, EPUB, MD
    echo   - Output formats: PDF, ODT, DOCX, TXT, RTF, EPUB, MD
    echo.
    if exist "build" rmdir /s /q "build"
    if exist "PDFMerger.spec" del "PDFMerger.spec"
) else (
    echo BUILD FAILED!
    echo Check the error messages above.
)

pause