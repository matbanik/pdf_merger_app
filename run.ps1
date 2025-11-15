# PDF Merger App - Automated Setup and Launch Script
# This script checks for Python, creates a virtual environment, installs dependencies, and launches the application

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Document Merger & PII Scrubber" -ForegroundColor Cyan
Write-Host " Multi-Format Support Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Function to check if Python is installed
function Test-PythonInstalled {
    try {
        $pythonVersion = python --version 2>&1
        if ($pythonVersion -match "Python (\d+)\.(\d+)\.(\d+)") {
            $major = [int]$matches[1]
            $minor = [int]$matches[2]
            Write-Host "[OK] Python found: $pythonVersion" -ForegroundColor Green
            if ($major -ge 3 -and $minor -ge 8) {
                return $true
            } else {
                Write-Host "[WARNING] Python 3.8 or higher is required. You have Python $major.$minor" -ForegroundColor Yellow
                return $false
            }
        }
        return $false
    }
    catch {
        return $false
    }
}

# Function to display Python installation instructions
function Show-PythonInstallInstructions {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host " Python is not installed or not found!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Python 3.8 or higher from:" -ForegroundColor Yellow
    Write-Host "https://www.python.org/downloads/" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Installation Instructions:" -ForegroundColor Yellow
    Write-Host "1. Download the latest Python installer for Windows" -ForegroundColor White
    Write-Host "2. Run the installer" -ForegroundColor White
    Write-Host "3. IMPORTANT: Check 'Add Python to PATH' during installation" -ForegroundColor Red
    Write-Host "4. Click 'Install Now'" -ForegroundColor White
    Write-Host "5. After installation, restart this script" -ForegroundColor White
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Check if Python is installed
if (-not (Test-PythonInstalled)) {
    Show-PythonInstallInstructions
}

# Set the virtual environment directory
$venvDir = "env"
$venvPython = Join-Path $venvDir "Scripts\python.exe"
$venvPip = Join-Path $venvDir "Scripts\pip.exe"

# Check if virtual environment exists
if (-not (Test-Path $venvDir)) {
    Write-Host "[INFO] Creating virtual environment..." -ForegroundColor Yellow
    try {
        python -m venv $venvDir
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create virtual environment"
        }
        Write-Host "[OK] Virtual environment created successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERROR] Failed to create virtual environment: $_" -ForegroundColor Red
        Write-Host "Please make sure Python venv module is installed" -ForegroundColor Yellow
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
} else {
    Write-Host "[OK] Virtual environment already exists" -ForegroundColor Green
}

# Activate virtual environment and install/upgrade dependencies
Write-Host "[INFO] Checking and installing dependencies..." -ForegroundColor Yellow

# Upgrade pip first
Write-Host "[INFO] Upgrading pip..." -ForegroundColor Cyan
try {
    & $venvPython -m pip install --upgrade pip --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[WARNING] Failed to upgrade pip, continuing anyway..." -ForegroundColor Yellow
    } else {
        Write-Host "[OK] Pip upgraded successfully" -ForegroundColor Green
    }
}
catch {
    Write-Host "[WARNING] Error upgrading pip: $_" -ForegroundColor Yellow
}

# Install requirements
if (Test-Path "requirements.txt") {
    Write-Host "[INFO] Installing requirements from requirements.txt..." -ForegroundColor Cyan
    Write-Host "(This may take several minutes for first-time installation...)" -ForegroundColor Yellow
    Write-Host ""

    try {
        & $venvPip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "[WARNING] Some packages may have failed to install" -ForegroundColor Yellow
            Write-Host "The application may still work, but some features might be unavailable" -ForegroundColor Yellow
            Write-Host ""
        } else {
            Write-Host ""
            Write-Host "[OK] All dependencies installed successfully" -ForegroundColor Green
        }
    }
    catch {
        Write-Host ""
        Write-Host "[WARNING] Error installing dependencies: $_" -ForegroundColor Yellow
        Write-Host "The application may still work, but some features might be unavailable" -ForegroundColor Yellow
        Write-Host ""
    }
} else {
    Write-Host "[WARNING] requirements.txt not found, skipping dependency installation" -ForegroundColor Yellow
}

# Launch the application
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Launching Application..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path "pdf_merger_app.py") {
    try {
        & $venvPython pdf_merger_app.py
    }
    catch {
        Write-Host ""
        Write-Host "[ERROR] Failed to launch application: $_" -ForegroundColor Red
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
} else {
    Write-Host "[ERROR] pdf_merger_app.py not found!" -ForegroundColor Red
    Write-Host "Make sure you're running this script from the application directory" -ForegroundColor Yellow
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host ""
Write-Host "Application closed." -ForegroundColor Cyan
