# Zip File Extractor

A Python script that automatically extracts all zip files in a folder into respective subfolders matching the zip file names.

## 📋 Overview

This script loops through all zip files in a specified directory and extracts each one into its own subfolder. For example:
- `documents.zip` → extracted to `documents/` folder
- `photos.zip` → extracted to `photos/` folder
- `projects.zip` → extracted to `projects/` folder

If the subfolders don't exist, they are created automatically.

## ✨ Features

- 🔄 **Batch Processing**: Processes all zip files in a folder automatically
- 📁 **Smart Folder Creation**: Creates subfolders based on zip file names
- 🛡️ **Error Handling**: Handles corrupted files, permission issues, and other errors gracefully
- 📊 **Progress Logging**: Shows real-time progress and detailed summary
- 🎯 **Flexible Usage**: Can specify custom folder or use current directory
- ✅ **Validation**: Checks zip file integrity before extraction

## 🚀 Quick Start

### Prerequisites

- Python 3.6 or higher
- Windows, macOS, or Linux

### Installation

1. **Clone or download** this repository:
   ```bash
   git clone <repository-url>
   cd windows_util
   ```

2. **No additional packages required** - uses Python standard library only!

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

## 💡 Usage Examples

### Example 1: Current Directory
```bash
# Navigate to folder with zip files
cd C:\Downloads
python C:\path\to\zip_extractor.py
```

### Example 2: Specific Folder
```bash
python zip_extractor.py "C:\Users\YourName\Documents\Archives"
```

### Example 3: With Python Virtual Environment
```bash
# Activate virtual environment (if using one)
.venv\Scripts\Activate.ps1  # Windows PowerShell
source .venv/bin/activate   # Linux/macOS

# Run the script
python zip_extractor.py
```

## 📁 Example Structure

**Before extraction:**
```
my_folder/
├── documents.zip
├── photos.zip
├── projects.zip
└── zip_extractor.py
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
├── projects/          # ← Contents of projects.zip
│   ├── project1/
│   └── project2/
└── zip_extractor.py
```

## 🔧 Advanced Options

### Command Line Arguments

| Argument | Description | Example |
|----------|-------------|---------|
| `folder_path` | Target folder containing zip files | `python zip_extractor.py "C:\Archives"` |
| (none) | Uses current working directory | `python zip_extractor.py` |

### Logging Levels

The script uses Python's logging module. To see more detailed output, you can modify the logging level in the script:

```python
# In zip_extractor.py, change this line:
logging.basicConfig(level=logging.DEBUG)  # More verbose
logging.basicConfig(level=logging.WARNING)  # Less verbose
```

## 📊 Output Examples

### Successful Execution
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
