# Windows Utilities

A collection of Python utilities for Windows system administration and file management. This toolkit includes scripts for batch font installation and automated zip file extraction.

## Overview

This repository contains two main utilities:

1. **Font Installer** - Automatically installs font files from multiple folders with admin elevation support
2. **Zip Extractor** - Batch extracts zip files into organized subfolders

Both tools include comprehensive error handling, detailed logging, and user-friendly progress reporting.

## Features

### Font Installer
- **Batch Installation**: Install fonts from all subfolders automatically
- **Multiple Installation Methods**: Uses Windows Shell COM, direct file operations, registry, and Windows API
- **Smart Detection**: Checks if fonts are already installed
- **Flexible Overwrite Options**: Ask, skip, or force overwrite existing fonts
- **Dry Run Mode**: Preview what would be installed without actually installing
- **Admin Privilege Detection**: Warns if not running with administrator rights
- **Supported Formats**: .ttf, .otf, .ttc, .fon, .fnt
- **Detailed Logging**: Shows real-time progress and comprehensive summary

### Zip Extractor
- **Batch Processing**: Process all zip files in a folder automatically
- **Smart Folder Creation**: Creates subfolders based on zip file names
- **Validation**: Checks zip file integrity before extraction
- **Error Handling**: Handles corrupted files and permission issues gracefully
- **Progress Logging**: Shows real-time progress and detailed summary
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Quick Start

### Prerequisites

- Python 3.6 or higher
- Windows operating system (for Font Installer)
- Administrator privileges (recommended for Font Installer)

### Installation

1. **Clone or download** this repository:
   ```bash
   git clone <repository-url>
   cd WindowsUtils
   ```

2. **Install dependencies** (for Font Installer):
   ```bash
   pip install -r requirements.txt
   ```

   The Zip Extractor uses only Python standard library and requires no additional packages.

## Font Installer Usage

### Basic Usage

1. **Install fonts from current directory (with prompts):**
   ```bash
   python font_installer.py
   ```

2. **Install fonts from specific folder:**
   ```bash
   python font_installer.py "C:\Downloads\Fonts"
   ```

3. **Dry run (preview without installing):**
   ```bash
   python font_installer.py --dry-run
   ```

4. **Force overwrite existing fonts:**
   ```bash
   python font_installer.py --force
   ```

5. **Skip existing fonts automatically:**
   ```bash
   python font_installer.py --overwrite no
   ```

### Command Line Options

| Option | Description | Example |
|--------|-------------|---------|
| `folder_path` | Path to folder with font subfolders | `python font_installer.py "C:\Fonts"` |
| `--dry-run` | Preview fonts without installing | `python font_installer.py --dry-run` |
| `--overwrite {yes,no,ask}` | Handle existing fonts | `python font_installer.py --overwrite yes` |
| `--force` | Force overwrite (same as `--overwrite yes`) | `python font_installer.py --force` |
| `--verbose, -v` | Enable verbose logging | `python font_installer.py -v` |
| `--no-admin-check` | Skip administrator privilege check | `python font_installer.py --no-admin-check` |

### Font Installer Examples

**Example 1: Install fonts with prompts for existing fonts**
```bash
python font_installer.py "C:\Downloads\FontCollection"
```

**Example 2: Preview what would be installed**
```bash
python font_installer.py "C:\Downloads\FontCollection" --dry-run
```

**Example 3: Install fonts, skipping existing ones**
```bash
python font_installer.py "C:\Downloads\FontCollection" --overwrite no
```

**Example 4: Force install all fonts, replacing existing**
```bash
python font_installer.py "C:\Downloads\FontCollection" --force
```

### Font Installer Output Example

```
Font Installer for Windows
========================================
Processing folder: C:\Users\YourName\Downloads\Fonts
----------------------------------------
2024-01-15 14:30:25 - INFO - Found 15 font file(s) in 3 folder(s) to install
2024-01-15 14:30:25 - INFO - Processing font: Roboto-Regular.ttf
2024-01-15 14:30:26 - INFO - Successfully installed: Roboto-Regular.ttf
2024-01-15 14:30:26 - INFO - Processing font: Roboto-Bold.ttf
2024-01-15 14:30:26 - INFO - Skipped: Roboto-Bold.ttf - Font already installed (skipped)

============================================================
FONT INSTALLATION SUMMARY
============================================================
Folders with fonts processed: 3
Total font files processed: 15
Successfully installed: 12
Failed installations: 0
Skipped (already installed): 3

Detailed Results:
----------------------------------------

Folder: C:\Users\YourName\Downloads\Fonts\Roboto
  ✓ Roboto-Regular.ttf
  ⏭️  Roboto-Bold.ttf - Font already installed (skipped)
  ✓ Roboto-Italic.ttf

Folder: C:\Users\YourName\Downloads\Fonts\OpenSans
  ✓ OpenSans-Regular.ttf
  ✓ OpenSans-Bold.ttf
============================================================

🎉 Font installation completed!
📝 Note: You may need to restart applications to see new fonts.
```

## Zip Extractor Usage

### Basic Usage

1. **Extract zip files in current directory:**
   ```bash
   python zip_extractor.py
   ```

2. **Extract zip files in specific folder:**
   ```bash
   python zip_extractor.py "C:\path\to\your\zip\folder"
   ```

3. **Linux/macOS example:**
   ```bash
   python zip_extractor.py "/home/user/Downloads"
   ```

### Zip Extractor Examples

**Example 1: Current directory**
```bash
cd C:\Downloads
python zip_extractor.py
```

**Example 2: Specific folder**
```bash
python zip_extractor.py "C:\Users\YourName\Documents\Archives"
```

### Example Structure

**Before extraction:**
```
my_folder/
├── documents.zip
├── photos.zip
└── projects.zip
```

**After running the script:**
```
my_folder/
├── documents.zip
├── documents/          # ← Contents of documents.zip
│   ├── file1.pdf
│   └── file2.docx
├── photos.zip
├── photos/            # ← Contents of photos.zip
│   ├── image1.jpg
│   └── image2.png
├── projects.zip
└── projects/          # ← Contents of projects.zip
    ├── project1/
    └── project2/
```

### Zip Extractor Output Example

```
Zip File Extractor
Processing folder: C:\Users\YourName\Downloads
--------------------------------------------------
2024-01-15 14:30:25 - INFO - Found 3 zip file(s) to process
2024-01-15 14:30:25 - INFO - Processing: documents.zip
2024-01-15 14:30:25 - INFO - Created/verified folder: C:\Users\YourName\Downloads\documents
2024-01-15 14:30:26 - INFO - Successfully extracted 5 file(s) from documents.zip
2024-01-15 14:30:26 - INFO - Processing: photos.zip
2024-01-15 14:30:26 - INFO - Created/verified folder: C:\Users\YourName\Downloads\photos
2024-01-15 14:30:27 - INFO - Successfully extracted 12 file(s) from photos.zip

==================================================
EXTRACTION SUMMARY
==================================================
Total zip files processed: 3
Successfully extracted: 2
Failed extractions: 1

Detailed Results:
------------------------------
✓ documents.zip → C:\Users\YourName\Downloads\documents (5 files)
✓ photos.zip → C:\Users\YourName\Downloads\photos (12 files)
✗ corrupted.zip - Error: Invalid zip file
==================================================
```

## Error Handling

Both utilities include comprehensive error handling:

### Font Installer
- Invalid or corrupted font files
- Permission issues (warns if not running as admin)
- Registry access errors
- File system errors
- Font already installed detection

### Zip Extractor
- Invalid or corrupted zip files
- Permission denied errors
- File system errors
- Missing directory errors

## Dependencies

- **Font Installer**: Requires `pywin32` (see [requirements.txt](requirements.txt))
- **Zip Extractor**: No external dependencies (uses Python standard library)

Install dependencies:
```bash
pip install -r requirements.txt
```

## Technical Details

### Font Installer Methods

The [`font_installer.py`](font_installer.py) uses multiple installation methods in order of preference:

1. **Windows Shell COM**: Uses [`install_font_shell_com`](font_installer.py) with automatic elevation handling
2. **Direct File Operations**: Uses [`install_font_powershell`](font_installer.py) to avoid system dialogs
3. **Copy & Registry**: Uses [`install_font_copy`](font_installer.py) for manual installation
4. **Windows API**: Falls back to [`ctypes.windll.gdi32.AddFontResourceW`](font_installer.py) if needed

### Key Functions

#### Font Installer
- [`is_admin()`](font_installer.py): Checks for administrator privileges
- [`is_font_installed()`](font_installer.py): Detects if font is already installed
- [`install_font()`](font_installer.py): Main installation function with multiple methods
- [`find_font_files()`](font_installer.py): Recursively finds all font files
- [`install_fonts_from_folder()`](font_installer.py): Batch processes fonts from folder

#### Zip Extractor
- [`extract_zip_files()`](zip_extractor.py): Main extraction function
- [`print_summary()`](zip_extractor.py): Displays detailed extraction results

## Logging

Both utilities use Python's `logging` module for progress tracking:

- **INFO**: Standard operation messages
- **DEBUG**: Verbose mode (use `--verbose` flag for Font Installer)
- **ERROR**: Error messages and failures
- **WARNING**: Non-critical issues

## Supported File Formats

### Font Installer
- `.ttf` - TrueType Font
- `.otf` - OpenType Font
- `.ttc` - TrueType Collection
- `.fon` - Bitmap Font
- `.fnt` - Bitmapped Font

### Zip Extractor
- `.zip` - Standard ZIP archives
