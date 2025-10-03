# Windows Utilities

A collection of utilities for Windows system administration, troubleshooting, and file management. This toolkit includes scripts for batch font installation, automated zip file extraction, and Bluetooth audio diagnostics and repair.

## Overview

This repository contains four main utilities:

1. **Font Installer** (Python) - Automatically installs font files from multiple folders with admin elevation support
2. **Zip Extractor** (Python) - Batch extracts zip files into organized subfolders
3. **Bluetooth Audio Diagnostics** (PowerShell) - Comprehensive diagnostics for Bluetooth audio device issues
4. **Bluetooth Audio Fix** (PowerShell) - Automated fixes for common Bluetooth audio problems

All tools include comprehensive error handling, detailed logging, and user-friendly progress reporting.

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

### Bluetooth Audio Diagnostics
- **Comprehensive Device Scanning**: Lists all audio devices with detailed status
- **Driver Status Checks**: Verifies Bluetooth audio driver and A2DP profile status
- **Volume Level Detection**: Identifies volume-related issues
- **PnP Device Analysis**: Checks for device problem codes and errors
- **Service Status**: Verifies Bluetooth and audio service health
- **Device Switching Tests**: Tests ability to switch between audio devices
- **Registry Validation**: Checks audio configuration in Windows registry

### Bluetooth Audio Fix
- **A2DP Driver Reset**: Disables and re-enables Bluetooth audio drivers
- **Audio Endpoint Refresh**: Resets audio endpoint configurations
- **Service Restart**: Restarts Bluetooth and audio services
- **Default Device Configuration**: Forces correct default device selection
- **Automated Verification**: Checks if fixes resolved the issues
- **Guided Troubleshooting**: Provides next steps if automated fixes don't work

## Quick Start

### Prerequisites

- Python 3.6 or higher (for Python utilities)
- Windows operating system
- PowerShell 5.1 or higher (for Bluetooth scripts)
- Administrator privileges (required for Font Installer and Bluetooth scripts)

### Installation

1. **Clone or download** this repository:
   ```bash
   git clone <repository-url>
   cd WindowsUtils
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

   For Bluetooth audio scripts, install the AudioDeviceCmdlets module:
   ```powershell
   Install-Module -Name AudioDeviceCmdlets
   ```

   Note: The Zip Extractor uses only Python standard library and requires no additional packages.

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

## Bluetooth Audio Diagnostics Usage

### Running Diagnostics

1. **Run diagnostics on Bluetooth audio device:**
   ```powershell
   .\bluetooth_audio_diagnostics.ps1
   ```

   **Note**: Must be run as Administrator in PowerShell.

### What It Checks

The diagnostics script performs 11 comprehensive checks:

1. Current default audio device details
2. All playback devices with their status
3. Specific device volume levels
4. PnP device status and problem codes
5. Bluetooth audio driver status
6. A2DP profile availability
7. Audio device driver information
8. System audio configuration (registry)
9. Device ID retrieval for manual checks
10. Device switching capability test
11. Error codes and device problems

### Diagnostics Output Example

```powershell
========================================
Deep Audio Diagnostics
========================================

[1] Current Default Device - Full Details:
   Device: Soundbar (S6520)
   Volume: 45%
   Muted: False

[2] All Playback Devices with Status:
Index Default Name              DeviceState
----- ------- ----              -----------
1     True    Soundbar (S6520)  Active
2     False   Speakers          Active

[10] Testing Device Switching:
   Current device: Soundbar (S6520)
   Attempting to switch to device Index 2...
   Switched to: Speakers
   Switching back to S6520...
   Current device: Soundbar (S6520)
   Device switching works!

========================================
Diagnostics Complete!
========================================
```

## Bluetooth Audio Fix Usage

### Running Fixes

1. **Apply automated fixes:**
   ```powershell
   .\bluetooth_audio_fix.ps1
   ```

   **Note**: Must be run as Administrator in PowerShell.

### What It Fixes

The fix script applies 5 different fixes automatically:

1. **A2DP Driver Reset**: Disables and re-enables the Bluetooth A2DP audio driver
2. **Audio Endpoint Reset**: Resets the audio endpoint configuration
3. **Bluetooth Services Restart**: Restarts Bluetooth Support and Audio Gateway services
4. **Audio Services Restart**: Restarts Windows Audio and Audio Endpoint Builder services
5. **Default Device Reset**: Forces the device to be set as default again

### Fix Script Output Example

```powershell
========================================
Advanced Bluetooth Audio Fix
========================================

[Fix 1] Resetting Bluetooth Audio Driver (A2DP)...
   Found S6520 A2DP device. Resetting...
   Disabled A2DP driver.
   Re-enabled A2DP driver.

[Fix 2] Resetting Audio Endpoint...
   Found S6520 audio endpoint. Resetting...
   Disabled audio endpoint.
   Re-enabled audio endpoint.

[Fix 3] Restarting Bluetooth Services...
   Bluetooth Support Service restarted.
   Bluetooth Audio Gateway Service restarted.

[Fix 4] Restarting Audio Services...
   Audio services restarted.

[Fix 5] Setting S6520 as default device...
   S6520 set as default.

[Verification] Checking S6520 status now:
   Current Device: Soundbar (S6520)
   Volume: 50%
   Muted: False

   ✓ Volume is now readable! Audio stream should work.

========================================
Fixes Applied!
========================================

Next Steps:
1. Test audio NOW (play YouTube video, etc.)
2. If still not working, the Bluetooth connection may need to be removed:
   - Go to Settings > Bluetooth & devices
   - Find S6520 and click 'Remove device'
   - Turn off/on the speaker and re-pair it
3. Check if your soundbar has multiple audio modes (stereo/surround)
   - Some modes don't work properly with Windows Bluetooth
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

### Bluetooth Audio Scripts
- Missing AudioDeviceCmdlets module
- Insufficient permissions (requires Administrator)
- Device not found or disconnected
- Service restart failures
- Driver enable/disable errors

## Dependencies

- **Font Installer**: Requires `pywin32` (see requirements.txt)
- **Zip Extractor**: No external dependencies (uses Python standard library)
- **Bluetooth Audio Scripts**: Requires `AudioDeviceCmdlets` PowerShell module

Install Python dependencies:
```bash
pip install -r requirements.txt
```

Install PowerShell module (as Administrator):
```powershell
Install-Module -Name AudioDeviceCmdlets
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

#### Bluetooth Audio Scripts
- **Diagnostics**: 11 comprehensive checks covering devices, drivers, services, and configuration
- **Fix Script**: 5 automated fixes for common Bluetooth audio issues
- Both scripts include color-coded output for easy reading and status verification

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

### Bluetooth Audio Scripts
- Works with all Bluetooth audio devices
- Specifically tested with A2DP audio profile
- Supports stereo and headset profiles

## Troubleshooting

### Font Installer Issues

**Problem**: "This script is not running with administrator privileges"
- **Solution**: Right-click Command Prompt/PowerShell and select "Run as Administrator"

**Problem**: "All installation methods failed"
- **Solution**: Ensure you have admin rights and the font file is not corrupted

**Problem**: Font shows as installed but doesn't appear in applications
- **Solution**: Restart the application or log out and log back in

### Zip Extractor Issues

**Problem**: "Invalid zip file" error
- **Solution**: The zip file may be corrupted. Try downloading it again

**Problem**: "Permission denied" error
- **Solution**: Check that you have write permissions for the target folder

**Problem**: Extracted folder is empty
- **Solution**: The zip file may be password-protected or empty

### Bluetooth Audio Script Issues

**Problem**: "Module AudioDeviceCmdlets not found"
- **Solution**: Install the module: `Install-Module -Name AudioDeviceCmdlets` (as Administrator)

**Problem**: "Access Denied" or permission errors
- **Solution**: Run PowerShell as Administrator (right-click PowerShell → Run as Administrator)

**Problem**: Diagnostics show "Volume: " (blank)
- **Solution**: Run the fix script - this indicates audio stream isn't working properly

**Problem**: Device shows in list but no audio plays
- **Solution**: Try the fix script, then check Volume Mixer to ensure app-specific volume isn't at 0

**Problem**: Fixes don't work
- **Solution**: 
  1. Remove the Bluetooth device from Settings
  2. Turn off the Bluetooth speaker/headphones
  3. Re-pair the device
  4. Run the diagnostics script again

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

## License

This project is open source. Please check the repository for license information.

## Authors

LiteObject

## Acknowledgments

- Python utilities built with Python's standard library and pywin32
- PowerShell scripts utilize AudioDeviceCmdlets module
- Designed for Windows system administrators and power users

## Support

If you encounter issues or have questions:
1. Check the Troubleshooting section above
2. Run the diagnostics scripts for detailed information
3. Open an issue on GitHub with diagnostic output

## Utility Summary

| Utility | Language | Purpose | Admin Required |
|---------|----------|---------|----------------|
| Font Installer | Python | Batch install fonts from folders | Recommended |
| Zip Extractor | Python | Batch extract zip files | No |
| Bluetooth Audio Diagnostics | PowerShell | Diagnose Bluetooth audio issues | Yes |
| Bluetooth Audio Fix | PowerShell | Fix common Bluetooth audio problems | Yes |
