# Installation Guide

## Quick Start (Recommended)

```bash
# Install core dependencies
pip install -r requirements.txt

# Run the application
python -m src.macpst.main
```

## What You Get

### ✅ **Core Installation (Always Works)**
- Complete GUI application
- PST file format detection and validation
- Basic PST parsing (extracts some messages)
- Full conversion to EML, MBOX, PDF formats
- Message filtering and duplicate detection
- Batch processing capabilities

### 🚀 **Enhanced Installation (Included)**
Full PST parsing is now included in the main requirements:

```bash
# This includes libpff-python for complete PST support
pip install -r requirements.txt

# Alternative if pip fails:
conda install -c conda-forge libpff-python
```

## Installation Options

### Option 1: Simple Installation
```bash
pip install -r requirements.txt
```
**Result**: App works with basic PST parsing

### Option 2: Full Installation
```bash
./install_deps.sh
```
**Result**: Attempts to install enhanced PST support

### Option 3: Manual Step-by-Step
```bash
# Core dependencies (required)
pip install reportlab python-dateutil cryptography lxml pillow chardet

# Full PST parsing support
pip install libpff-python
```

## Troubleshooting

### "libpff-python installation failed"
**The app will still work with basic PST parsing!** Try alternative installation or use the fallback mode.

### "Import error for pypff"
The code looks for the `pypff` module from libpff-python. If installation failed, basic parsing will be used automatically.

### "GUI won't start"
Make sure you have Python 3.8+ with tkinter:
```bash
python -c "import tkinter; print('Tkinter OK')"
```

## What Works Without libpff-python

- ✅ PST file detection and validation
- ✅ Basic message extraction (sample messages)
- ✅ Complete GUI interface
- ✅ All output format conversions (EML, MBOX, PDF)
- ✅ Message filtering and processing
- ✅ Batch operations

## What Requires libpff-python

- 🔍 Full PST message extraction
- 📁 Complete folder structure reading
- 📎 Attachment extraction
- 📧 Complete email metadata

## Running the Application

```bash
# Start the GUI
python -m src.macpst.main

# Or if you installed as a package
macpst-converter
```

The app will automatically detect what PST parsing capabilities are available and adjust accordingly.