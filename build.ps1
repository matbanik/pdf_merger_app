# Build script for PDF Merger with Multi-Format Support
# This script builds a standalone executable using PyInstaller

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Building PDF Merger - Multi-Format" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Clean previous builds
Write-Host "[1/4] Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "build") {
    Remove-Item -Recurse -Force "build"
    Write-Host "  - Removed build directory" -ForegroundColor Gray
}
if (Test-Path "dist") {
    Remove-Item -Recurse -Force "dist"
    Write-Host "  - Removed dist directory" -ForegroundColor Gray
}
if (Test-Path "PDFMerger.spec") {
    Remove-Item -Force "PDFMerger.spec"
    Write-Host "  - Removed old spec file" -ForegroundColor Gray
}
Write-Host "[OK] Cleanup complete" -ForegroundColor Green
Write-Host ""

# Check if PDFMerger.spec exists (should exist in repo)
Write-Host "[2/4] Checking build configuration..." -ForegroundColor Yellow
if (-not (Test-Path "PDFMerger.spec")) {
    Write-Host "[ERROR] PDFMerger.spec not found!" -ForegroundColor Red
    Write-Host "The spec file should be in the project directory." -ForegroundColor Yellow
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
Write-Host "[OK] Build configuration found" -ForegroundColor Green
Write-Host ""

# Check if PyInstaller is installed
Write-Host "[3/4] Checking PyInstaller..." -ForegroundColor Yellow
try {
    $pyinstallerVersion = python -m PyInstaller --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[OK] PyInstaller version: $pyinstallerVersion" -ForegroundColor Green
    } else {
        throw "PyInstaller not found"
    }
} catch {
    Write-Host "[ERROR] PyInstaller is not installed!" -ForegroundColor Red
    Write-Host "Installing PyInstaller..." -ForegroundColor Yellow
    python -m pip install pyinstaller
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Failed to install PyInstaller" -ForegroundColor Red
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
    Write-Host "[OK] PyInstaller installed" -ForegroundColor Green
}
Write-Host ""

# Build the executable
Write-Host "[4/4] Building executable..." -ForegroundColor Yellow
Write-Host "This may take several minutes..." -ForegroundColor Gray
Write-Host ""

python -m PyInstaller --noconfirm PDFMerger.spec

Write-Host ""
# Check if build was successful
if (Test-Path "dist\PDFMerger\PDFMerger.exe") {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host " BUILD SUCCESS!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Executable location:" -ForegroundColor Cyan
    Write-Host "  dist\PDFMerger\PDFMerger.exe" -ForegroundColor White
    Write-Host ""
    Write-Host "Features included:" -ForegroundColor Cyan
    Write-Host "  - Multiprocessing support" -ForegroundColor White
    Write-Host "  - Multi-format input: PDF, ODT, DOCX, TXT, RTF, EPUB, MD" -ForegroundColor White
    Write-Host "  - Multi-format output: PDF, ODT, DOCX, TXT, RTF, EPUB, MD" -ForegroundColor White
    Write-Host "  - PII scrubbing" -ForegroundColor White
    Write-Host "  - Markdown conversion (with OCR support)" -ForegroundColor White
    Write-Host "  - PDF decryption (qpdf)" -ForegroundColor White
    Write-Host ""
    Write-Host "Important notes:" -ForegroundColor Yellow
    Write-Host "  - Console window is hidden (check app's console widget)" -ForegroundColor White
    Write-Host "  - First run will download AI models (~1-2GB)" -ForegroundColor White
    Write-Host "  - GPU acceleration available if CUDA is installed" -ForegroundColor White
    Write-Host ""

    # Clean up build artifacts
    if (Test-Path "build") {
        Remove-Item -Recurse -Force "build"
        Write-Host "Build artifacts cleaned up" -ForegroundColor Gray
    }

    # Ask if user wants to test the executable
    Write-Host ""
    $response = Read-Host "Do you want to run the executable now? (y/n)"
    if ($response -eq 'y' -or $response -eq 'Y') {
        Write-Host "Launching PDFMerger..." -ForegroundColor Cyan
        Start-Process "dist\PDFMerger\PDFMerger.exe"
    }
} else {
    Write-Host "========================================" -ForegroundColor Red
    Write-Host " BUILD FAILED!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Check the error messages above for details." -ForegroundColor Yellow
    Write-Host "Common issues:" -ForegroundColor Yellow
    Write-Host "  - Missing dependencies (run: pip install -r requirements.txt)" -ForegroundColor White
    Write-Host "  - Python version < 3.8 (requires Python 3.8+)" -ForegroundColor White
    Write-Host "  - Insufficient disk space" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
