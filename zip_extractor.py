#!/usr/bin/env python3
"""
Zip File Extractor

This script loops through all zip files in a specified folder and extracts them
into respective subfolders matching the zip file names (without extension).
If the subfolders don't exist, they are created automatically.

Usage:
    python zip_extractor.py [folder_path]

If no folder_path is provided, it will process zip files in the current directory.
"""

import os
import zipfile
import sys
from pathlib import Path
import logging

def setup_logging():
    """Set up logging to display progress and errors."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

def extract_zip_files(folder_path):
    """
    Extract all zip files in the specified folder to matching subfolders.
    
    Args:
        folder_path (str): Path to the folder containing zip files
    
    Returns:
        dict: Summary of extraction results
    """
    folder_path = Path(folder_path).resolve()
    
    if not folder_path.exists():
        logging.error(f"Folder does not exist: {folder_path}")
        return {"error": "Folder not found"}
    
    if not folder_path.is_dir():
        logging.error(f"Path is not a directory: {folder_path}")
        return {"error": "Not a directory"}
    
    # Find all zip files in the folder
    zip_files = list(folder_path.glob("*.zip"))
    
    if not zip_files:
        logging.info(f"No zip files found in: {folder_path}")
        return {"processed": 0, "successful": 0, "failed": 0, "details": []}
    
    logging.info(f"Found {len(zip_files)} zip file(s) to process")
    
    results = {
        "processed": len(zip_files),
        "successful": 0,
        "failed": 0,
        "details": []
    }
    
    for zip_file in zip_files:
        try:
            # Create subfolder name (zip filename without extension)
            subfolder_name = zip_file.stem
            subfolder_path = folder_path / subfolder_name
            
            logging.info(f"Processing: {zip_file.name}")
            
            # Create subfolder if it doesn't exist
            subfolder_path.mkdir(exist_ok=True)
            logging.info(f"Created/verified folder: {subfolder_path}")
            
            # Extract zip file
            with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                # Check if zip file is valid
                zip_ref.testzip()
                
                # Extract all contents
                zip_ref.extractall(subfolder_path)
                
                # Count extracted files
                extracted_files = len(zip_ref.namelist())
                
            logging.info(f"Successfully extracted {extracted_files} file(s) from {zip_file.name}")
            
            results["successful"] += 1
            results["details"].append({
                "zip_file": zip_file.name,
                "status": "success",
                "extracted_to": str(subfolder_path),
                "files_extracted": extracted_files
            })
            
        except zipfile.BadZipFile:
            error_msg = f"Invalid zip file: {zip_file.name}"
            logging.error(error_msg)
            results["failed"] += 1
            results["details"].append({
                "zip_file": zip_file.name,
                "status": "failed",
                "error": "Invalid zip file"
            })
            
        except PermissionError:
            error_msg = f"Permission denied accessing: {zip_file.name}"
            logging.error(error_msg)
            results["failed"] += 1
            results["details"].append({
                "zip_file": zip_file.name,
                "status": "failed",
                "error": "Permission denied"
            })
            
        except Exception as e:
            error_msg = f"Unexpected error with {zip_file.name}: {str(e)}"
            logging.error(error_msg)
            results["failed"] += 1
            results["details"].append({
                "zip_file": zip_file.name,
                "status": "failed",
                "error": str(e)
            })
    
    return results

def print_summary(results):
    """Print a summary of the extraction results."""
    if "error" in results:
        print(f"\nError: {results['error']}")
        return
    
    print("\n" + "="*50)
    print("EXTRACTION SUMMARY")
    print("="*50)
    print(f"Total zip files processed: {results['processed']}")
    print(f"Successfully extracted: {results['successful']}")
    print(f"Failed extractions: {results['failed']}")
    
    if results['details']:
        print("\nDetailed Results:")
        print("-" * 30)
        for detail in results['details']:
            if detail['status'] == 'success':
                print(f"✓ {detail['zip_file']} → {detail['extracted_to']} ({detail['files_extracted']} files)")
            else:
                print(f"✗ {detail['zip_file']} - Error: {detail['error']}")
    
    print("="*50)

def main():
    """Main function to handle command line arguments and run the extraction."""
    setup_logging()
    
    # Determine folder path
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        folder_path = os.getcwd()
    
    print(f"Zip File Extractor")
    print(f"Processing folder: {Path(folder_path).resolve()}")
    print("-" * 50)
    
    # Run extraction
    results = extract_zip_files(folder_path)
    
    # Print summary
    print_summary(results)

if __name__ == "__main__":
    main()