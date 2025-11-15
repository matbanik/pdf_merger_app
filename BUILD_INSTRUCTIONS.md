# Build Instructions for PDF Merger Executable

This document explains how to build a standalone executable for the Document Merger & PII Scrubber application.

## Prerequisites

### Required Software
- **Python 3.8 or higher** - [Download](https://www.python.org/downloads/)
  - Make sure to check "Add Python to PATH" during installation
- **Git** (optional, for cloning the repository)

### System Requirements
- **OS**: Windows 10/11 (64-bit recommended)
- **RAM**: 8GB minimum, 16GB recommended (for AI model processing)
- **Disk Space**:
  - ~500MB for build process
  - ~5GB for AI models (downloaded on first run)
- **GPU** (optional): NVIDIA GPU with CUDA support for faster processing

## Step-by-Step Build Process

### 1. Install Dependencies

Open PowerShell or Command Prompt in the project directory and run:

```powershell
# Using PowerShell (recommended)
.\run.ps1
```

This will automatically:
- Create a virtual environment
- Install all required Python packages
- Launch the application for testing

Or install manually:
```bash
pip install -r requirements.txt
```

### 2. Build the Executable

#### Option A: Using PowerShell (Recommended)
```powershell
.\build.ps1
```

#### Option B: Using Command Prompt
```cmd
build.bat
```

#### Option C: Manual Build
```bash
python -m PyInstaller --noconfirm PDFMerger.spec
```

### 3. Build Output

After successful build, you'll find:
- **Executable**: `dist\PDFMerger\PDFMerger.exe`
- **Supporting files**: All dependencies bundled in `dist\PDFMerger\` folder

## Build Configuration

The build is configured in `PDFMerger.spec` with the following features:

### Included Dependencies
- **Core PDF Processing**: PyMuPDF (fitz)
- **Multi-Format Support**:
  - python-docx (DOCX files)
  - odfpy (ODT files)
  - ebooklib (EPUB files)
  - striprtf (RTF files)
  - beautifulsoup4 (HTML/EPUB parsing)
- **AI/ML Libraries**:
  - PyTorch (CPU and GPU support)
  - marker-pdf (AI-powered PDF to Markdown)
  - surya (OCR support)
  - transformers (Hugging Face models)
- **GUI**: tkinter (included with Python)

### Hidden Imports
The spec file includes hidden imports for:
- Multiprocessing modules
- Document format handlers
- XML/HTML parsers
- AI model components

## Troubleshooting

### Build Fails with "Module not found"
**Solution**: Install all dependencies first
```bash
pip install -r requirements.txt
```

### Build Fails with "No module named PyInstaller"
**Solution**: Install PyInstaller
```bash
pip install pyinstaller
```

### Executable is Very Large (>500MB)
**Explanation**: This is normal. The executable includes:
- PyTorch (~400MB)
- All Python libraries
- AI model infrastructure
The actual AI models (~1-2GB) are downloaded separately on first run.

### GPU Support Not Working in Built Executable
**Solution**: Make sure:
1. CUDA toolkit is installed on target machine
2. PyTorch was installed with CUDA support:
   ```bash
   pip install torch torchvision --index-url https://download.pytorch.org/whl/cu118
   ```
3. Rebuild the executable after reinstalling PyTorch

### "DLL Load Failed" Error on Target Machine
**Causes**:
- Missing Visual C++ Redistributables
- Missing CUDA libraries (for GPU support)

**Solutions**:
1. Install [Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)
2. For GPU: Install [CUDA Toolkit](https://developer.nvidia.com/cuda-downloads)

### Build Works but Executable Crashes on Launch
**Debugging**:
1. Run in console mode to see errors:
   - Edit `PDFMerger.spec`: change `console=False` to `console=True`
   - Rebuild
2. Check Windows Event Viewer for crash details
3. Verify all dependencies in spec file match installed versions

## Distribution

### Single Folder Distribution (Default)
The `dist\PDFMerger\` folder contains:
- `PDFMerger.exe` - Main executable
- Supporting DLLs and libraries
- Python runtime

**To distribute**:
- Zip the entire `dist\PDFMerger\` folder
- Users extract and run `PDFMerger.exe`

### Single File Distribution (Optional)
To create a single executable file:

1. Edit `PDFMerger.spec`:
   ```python
   exe = EXE(
       pyz,
       a.scripts,
       a.binaries,      # Add this
       a.zipfiles,      # Add this
       a.datas,         # Add this
       [],
       name='PDFMerger',
       debug=False,
       bootloader_ignore_signals=False,
       strip=False,
       upx=True,        # Enable UPX compression
       upx_exclude=[],
       runtime_tmpdir=None,
       console=False,
   )
   ```

2. Rebuild:
   ```bash
   python -m PyInstaller --noconfirm PDFMerger.spec
   ```

**Note**: Single file builds are slower to start (extract to temp folder first).

## Build Optimization

### Reduce Executable Size
1. **Exclude unused AI models** (if not using advanced markdown):
   ```python
   excludes=['marker', 'surya', 'transformers', 'torch']
   ```

2. **Use UPX compression**:
   - Download UPX: https://upx.github.io/
   - Add to PATH
   - Set `upx=True` in spec file

3. **Remove debug symbols**:
   ```python
   strip=True,
   ```

### Improve Build Speed
1. Use `--onedir` instead of `--onefile` (default)
2. Disable UPX: `upx=False`
3. Skip unnecessary data collection

## Testing the Built Executable

### Basic Test
1. Run `dist\PDFMerger\PDFMerger.exe`
2. Add a test PDF file
3. Click "Start Merge"
4. Verify output file is created

### Full Test Checklist
- [ ] Application launches without errors
- [ ] Can add files (PDF, DOCX, ODT, TXT, RTF, EPUB, MD)
- [ ] Can merge files to all output formats
- [ ] PII scrubbing works
- [ ] PDF decryption works (with qpdf)
- [ ] Markdown conversion works (simple mode)
- [ ] Markdown conversion works (advanced mode - if AI models downloaded)
- [ ] GPU acceleration works (if CUDA available)
- [ ] Settings are saved and loaded correctly

## Common Build Scenarios

### Building for Distribution Without AI Models
If you want a smaller executable without AI features:

1. Edit `requirements.txt` - remove:
   ```
   marker-pdf>=0.2.0
   torch>=2.0.0
   torchvision>=0.15.0
   transformers>=4.30.0
   ```

2. Edit `PDFMerger.spec` - remove from hidden_imports:
   ```python
   'marker', 'surya', 'torch.multiprocessing', 'cv2',
   ```

3. Remove from data collection:
   ```python
   marker_datas = []
   surya_datas = []
   ```

4. Rebuild

**Result**: ~100MB executable (instead of ~500MB) but no advanced markdown conversion.

### Building for Network Deployment
For running on shared network drives:

1. Use `--onefile` mode
2. Set specific temporary directory:
   ```python
   runtime_tmpdir='_internal'
   ```
3. Test on target network before deployment

## Support

If you encounter issues not covered here:
1. Check the application's console output for errors
2. Review the PyInstaller build log
3. Create an issue on GitHub with:
   - Python version (`python --version`)
   - PyInstaller version (`pyinstaller --version`)
   - Full error message
   - Build log output
