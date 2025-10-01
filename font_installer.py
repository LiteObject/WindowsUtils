#!/usr/bin/env python3
"""
Font Installer for Windows

This script loops through all subfolders in a specified directory and installs
font files (.ttf, .otf, .woff, .woff2) found in those folders.

Usage:
    python font_installer.py [folder_path]

If no folder_path is provided, it will process subfolders in the current
directory.

Note: This script requires administrative privileges on Windows to install fonts
to the system font directory.
"""

import os
import shutil
import winreg
from pathlib import Path
import logging
import ctypes
import argparse
import time

# Font file extensions supported
FONT_EXTENSIONS = {".ttf", ".otf", ".ttc", ".fon", ".fnt"}
FONT_REGISTRY_PATH = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except AttributeError:
        return False


def is_font_installed(font_path):
    """Check if a font is already installed in the system."""
    try:
        font_filename = os.path.basename(font_path)
        fonts_dir = os.path.join(os.environ["WINDIR"], "Fonts")
        installed_font_path = os.path.join(fonts_dir, font_filename)

        # Check if font file exists in Fonts directory
        if os.path.exists(installed_font_path):
            return True

        # Also check registry for font entries
        try:
            fonts_key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE, FONT_REGISTRY_PATH, 0, winreg.KEY_READ
            )
            with fonts_key:
                i = 0
                while True:
                    try:
                        _, value, _ = winreg.EnumValue(fonts_key, i)
                        if value.lower() == font_filename.lower():
                            return True
                        i += 1
                    except OSError:
                        break
        except OSError:
            pass

        return False

    except OSError as e:
        logging.debug("Error checking if font is installed %s: %s", font_path, str(e))
        return False


def get_font_name_from_file(font_path):
    """
    Extract font name from font file for registry entry.
    This is a simplified approach - for production use, consider using
    a font parsing library like fonttools.
    """
    font_name = Path(font_path).stem
    extension = Path(font_path).suffix.lower()

    # Basic font name formatting
    if extension in [".ttf", ".otf"]:
        return f"{font_name} (TrueType)"
    if extension == ".ttc":
        return f"{font_name} (TrueType Collection)"
    return font_name


def install_font_registry(font_path, font_name, force_overwrite=False):
    """Install font by adding registry entry."""
    try:
        # Open the Fonts registry key
        fonts_key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, FONT_REGISTRY_PATH, 0, winreg.KEY_WRITE
        )
        with fonts_key:

            # Check if font is already registered
            if not force_overwrite:
                try:
                    existing_value = winreg.QueryValueEx(fonts_key, font_name)[0]
                    if existing_value:
                        logging.debug(
                            "Font already registered in registry: %s", font_name
                        )
                        return False
                except OSError:
                    # Font not found in registry, proceed with installation
                    pass

            # Add or update font in registry
            winreg.SetValueEx(
                fonts_key, font_name, 0, winreg.REG_SZ, os.path.basename(font_path)
            )
            logging.debug("Font registered in registry: %s", font_name)
        return True
    except OSError as e:
        logging.error("Registry installation failed for %s: %s", font_path, str(e))
        return False


def install_font_shell_com(font_path, force_overwrite=False):
    """Install font using Windows Shell COM object."""
    try:
        # pylint: disable=import-outside-toplevel
        import win32com.client

        # Check if font already exists and we shouldn't overwrite
        fonts_dir = os.path.join(os.environ["WINDIR"], "Fonts")
        destination = os.path.join(fonts_dir, os.path.basename(font_path))

        if os.path.exists(destination) and not force_overwrite:
            return False

        # Use Shell.Application to copy font with elevation
        shell = win32com.client.Dispatch("Shell.Application")
        fonts_folder = shell.Namespace(0x14)  # Fonts folder

        # Copy the font file - this will handle elevation automatically
        # 4=no progress dialog, 16=yes to all
        fonts_folder.CopyHere(font_path, 4 + 16)

        return True

    except ImportError:
        logging.debug("win32com not available, falling back to other methods")
        return False
    except OSError as e:
        logging.error("Shell COM installation failed for %s: %s", font_path, str(e))
        return False


def install_font_powershell(font_path, force_overwrite=False):
    """Install font using direct file operations (avoiding system dialogs)."""
    try:
        fonts_dir = os.path.join(os.environ["WINDIR"], "Fonts")
        destination = os.path.join(fonts_dir, os.path.basename(font_path))

        # Check if font already exists
        if os.path.exists(destination) and not force_overwrite:
            return False

        # Remove existing file if it exists to avoid dialog
        if os.path.exists(destination):
            try:
                os.remove(destination)
                logging.debug("Removed existing font: %s", destination)
            except OSError as e:
                logging.error("Failed to remove existing font: %s", str(e))
                return False

        # Copy the font file
        shutil.copy2(font_path, destination)

        # Register in registry
        font_name = get_font_name_from_file(font_path)
        if install_font_registry(destination, font_name, force_overwrite):
            # Notify system of font change without using Windows API
            # that triggers dialog
            try:
                # Use a simpler notification method
                # Small delay to ensure file operations complete
                time.sleep(0.1)
                logging.debug("Font installed successfully: %s", font_name)
                return True
            except OSError as e:
                logging.debug("Font installed but notification failed: %s", str(e))
                return True

        return False

    except OSError as e:
        logging.error("Direct installation failed for %s: %s", font_path, str(e))
        return False


def install_font_copy(font_path, force_overwrite=False):
    """Install font by copying to Windows Fonts directory."""
    try:
        # Get Windows Fonts directory
        fonts_dir = os.path.join(os.environ["WINDIR"], "Fonts")
        destination = os.path.join(fonts_dir, os.path.basename(font_path))

        # If file exists and we're not forcing overwrite, check if we
        # should skip
        if os.path.exists(destination) and not force_overwrite:
            return False

        # Remove existing file if it exists (to avoid system dialog)
        if os.path.exists(destination):
            try:
                os.remove(destination)
                logging.debug("Removed existing font file: %s", destination)
            except OSError as e:
                logging.error("Failed to remove existing font: %s", str(e))
                return False

        # Copy font file to Fonts directory
        shutil.copy2(font_path, destination)

        # Add registry entry
        font_name = get_font_name_from_file(font_path)
        return install_font_registry(destination, font_name, force_overwrite)

    except OSError as e:
        logging.error("Copy installation failed for %s: %s", font_path, str(e))
        return False


def install_font(font_path, overwrite_mode="ask"):
    """
    Install a single font file using multiple methods.
    Tries different installation methods in order of preference.

    Args:
        font_path (str): Path to the font file to install
        overwrite_mode (str): How to handle existing fonts ('yes', 'no', 'ask')

    Returns:
        tuple: (success: bool, message: str)
    """
    font_path = str(Path(font_path).resolve())
    force_overwrite = overwrite_mode == "yes"

    # Check if font is already installed
    if is_font_installed(font_path):
        font_name = os.path.basename(font_path)

        if overwrite_mode == "no":
            return False, f"Font already installed (skipped): {font_name}"
        if overwrite_mode == "ask":
            response = (
                input(
                    f"\nFont '{font_name}' is already installed. " f"Overwrite? (y/n): "
                )
                .lower()
                .strip()
            )
            if response not in ["y", "yes"]:
                return False, f"Font installation skipped by user: {font_name}"
            force_overwrite = True
        # If overwrite_mode == 'yes', continue with installation

        logging.info("Overwriting existing font: %s", font_name)

    # Method 1: Try Windows Shell COM object (handles elevation automatically)
    if install_font_shell_com(font_path, force_overwrite):
        return True, "Installed successfully using Windows Shell COM"

    # Method 2: Try direct file operations (avoids system dialogs)
    if install_font_powershell(font_path, force_overwrite):
        return True, "Installed successfully using direct file operations"

    # Method 3: Try copy and registry method
    if install_font_copy(font_path, force_overwrite):
        return True, "Installed successfully using copy method"

    # Method 4: Use Windows API (fallback) - may still show dialog
    try:
        # Only use this as last resort and with caution
        if force_overwrite:
            # Use Windows API to install font
            gdi32 = ctypes.windll.gdi32
            result = gdi32.AddFontResourceW(font_path)
            if result > 0:
                # Notify system of font change
                user32 = ctypes.windll.user32
                user32.SendMessageW(0xFFFF, 0x001D, 0, 0)  # WM_FONTCHANGE
                return True, "Installed successfully using Windows API"
    except OSError as e:
        logging.error("API installation failed for %s: %s", font_path, str(e))

    return False, "All installation methods failed"


def find_font_files(folder_path):
    """Find all font files in the given folder and its subfolders."""
    folder_path = Path(folder_path)
    font_files = []
    folders_with_fonts = set()

    for root, _, files in os.walk(folder_path):
        root_path = Path(root)
        fonts_in_folder = []

        for file in files:
            if Path(file).suffix.lower() in FONT_EXTENSIONS:
                font_file = root_path / file
                font_files.append(font_file)
                fonts_in_folder.append(font_file)

        # Track folders that contain fonts
        if fonts_in_folder:
            folders_with_fonts.add(root_path)
            logging.debug("Found %d font(s) in: %s", len(fonts_in_folder), root_path)
        else:
            # Only log skipped folders if they are subfolders (not the root)
            if root_path != folder_path:
                logging.debug("Skipping folder (no fonts found): %s", root_path)

    return font_files, folders_with_fonts


def preview_fonts_from_folder(folder_path):
    """
    Preview all font files that would be installed from subfolders
    (dry run mode).

    Args:
        folder_path (str): Path to the folder containing subfolders with fonts

    Returns:
        dict: Summary of fonts that would be installed
    """
    folder_path = Path(folder_path).resolve()

    if not folder_path.exists():
        logging.error("Folder does not exist: %s", folder_path)
        return {"error": "Folder not found"}

    if not folder_path.is_dir():
        logging.error("Path is not a directory: %s", folder_path)
        return {"error": "Not a directory"}

    # Find all font files in subfolders
    font_files, folders_with_fonts = find_font_files(folder_path)

    if not font_files:
        logging.info("No font files found in: %s", folder_path)
        return {
            "processed": 0,
            "successful": 0,
            "failed": 0,
            "details": [],
            "folders_processed": 0,
        }

    logging.info(
        "Found %d font file(s) in %d folder(s) that would be installed",
        len(font_files),
        len(folders_with_fonts),
    )

    results = {
        "processed": len(font_files),
        "successful": len(font_files),  # In dry run, assume all would succeed
        "failed": 0,
        "details": [],
        "folders_processed": len(folders_with_fonts),
        "folders_with_fonts": folders_with_fonts,
    }

    for font_file in font_files:
        logging.info("Would install font: %s", font_file.name)
        results["details"].append(
            {
                "font_file": font_file.name,
                "path": str(font_file.parent),
                "status": "would_install",
            }
        )

    return results


def install_fonts_from_folder(folder_path, overwrite_mode="ask"):
    """
    Install all font files found in subfolders of the specified directory.

    Args:
        folder_path (str): Path to the folder containing subfolders with fonts
        overwrite_mode (str): How to handle existing fonts ('yes', 'no', 'ask')

    Returns:
        dict: Summary of installation results
    """
    folder_path = Path(folder_path).resolve()

    if not folder_path.exists():
        logging.error("Folder does not exist: %s", folder_path)
        return {"error": "Folder not found"}

    if not folder_path.is_dir():
        logging.error("Path is not a directory: %s", folder_path)
        return {"error": "Not a directory"}

    # Find all font files in subfolders
    font_files, folders_with_fonts = find_font_files(folder_path)

    if not font_files:
        logging.info("No font files found in: %s", folder_path)
        return {
            "processed": 0,
            "successful": 0,
            "failed": 0,
            "skipped": 0,
            "details": [],
            "folders_processed": 0,
        }

    logging.info(
        "Found %d font file(s) in %d folder(s) to install",
        len(font_files),
        len(folders_with_fonts),
    )

    results = {
        "processed": len(font_files),
        "successful": 0,
        "failed": 0,
        "skipped": 0,
        "details": [],
        "folders_processed": len(folders_with_fonts),
        "folders_with_fonts": folders_with_fonts,
    }

    for font_file in font_files:
        try:
            logging.info("Processing font: %s", font_file.name)

            success, message = install_font(font_file, overwrite_mode)

            if success:
                logging.info("Successfully installed: %s", font_file.name)
                results["successful"] += 1
                results["details"].append(
                    {
                        "font_file": font_file.name,
                        "path": str(font_file.parent),
                        "status": "success",
                        "message": message,
                    }
                )
            else:
                if (
                    "skipped" in message.lower()
                    or "already installed" in message.lower()
                ):
                    logging.info("Skipped: %s - %s", font_file.name, message)
                    results["skipped"] += 1
                    results["details"].append(
                        {
                            "font_file": font_file.name,
                            "path": str(font_file.parent),
                            "status": "skipped",
                            "message": message,
                        }
                    )
                else:
                    logging.error("Failed to install: %s - %s", font_file.name, message)
                    results["failed"] += 1
                    results["details"].append(
                        {
                            "font_file": font_file.name,
                            "path": str(font_file.parent),
                            "status": "failed",
                            "error": message,
                        }
                    )

        except OSError as e:
            error_msg = f"Unexpected error installing {font_file.name}: {e}"
            logging.error(error_msg)
            results["failed"] += 1
            results["details"].append(
                {
                    "font_file": font_file.name,
                    "path": str(font_file.parent),
                    "status": "failed",
                    "error": str(e),
                }
            )

    return results


def print_summary(results, dry_run=False):
    """Print a summary of the installation results."""
    if "error" in results:
        print(f"\nError: {results['error']}")
        return

    mode_text = "DRY RUN - FONT PREVIEW" if dry_run else "FONT INSTALLATION SUMMARY"

    print("\n" + "=" * 60)
    print(mode_text)
    print("=" * 60)

    # Show folder processing info
    folders_processed = results.get("folders_processed", 0)
    if folders_processed > 0:
        print(f"Folders with fonts processed: {folders_processed}")

    print(f"Total font files processed: {results['processed']}")

    if dry_run:
        print(f"Fonts that would be installed: {results['successful']}")
    else:
        print(f"Successfully installed: {results['successful']}")
        print(f"Failed installations: {results['failed']}")

        skipped = results.get("skipped", 0)
        if skipped > 0:
            print(f"Skipped (already installed): {skipped}")

    if results["details"]:
        detail_header = "Fonts Found:" if dry_run else "Detailed Results:"
        print(f"\n{detail_header}")
        print("-" * 40)

        # Group by folder
        folders = {}
        for detail in results["details"]:
            folder = detail["path"]
            if folder not in folders:
                folders[folder] = []
            folders[folder].append(detail)

        for folder, files in folders.items():
            print(f"\nFolder: {folder}")
            for detail in files:
                if dry_run or detail["status"] == "success":
                    icon = "📋" if dry_run else "✓"
                    print(f"  {icon} {detail['font_file']}")
                elif detail["status"] == "skipped":
                    msg = detail.get("message", "Skipped")
                    print(f"  ⏭️  {detail['font_file']} - {msg}")
                else:
                    error = detail.get("error", "Unknown error")
                    print(f"  ✗ {detail['font_file']} - Error: {error}")

    # Show summary of skipped folders
    if results["processed"] == 0:
        print("\nNo folders with font files were found to process.")
        print(
            "   Folders without .ttf, .otf, .ttc, .fon, or .fnt files "
            "are automatically skipped."
        )

    print("=" * 60)


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Install font files from all subfolders in a specified directory."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python font_installer.py
    # Install fonts from current directory (prompt for existing)
  python font_installer.py "C:\\Downloads\\Fonts"
    # Install fonts from specific folder
  python font_installer.py --dry-run
    # Show what would be installed without installing
  python font_installer.py --overwrite yes
    # Overwrite existing fonts without prompting
  python font_installer.py --overwrite no
    # Skip existing fonts automatically
  python font_installer.py --force
    # Same as --overwrite yes (force overwrite)

Supported font formats: .ttf, .otf, .ttc, .fon, .fnt

Note: Running as Administrator is recommended for proper font installation.
        """,
    )

    parser.add_argument(
        "folder_path",
        nargs="?",
        default=os.getcwd(),
        help=(
            "Path to folder containing subfolders with font files "
            "(default: current directory)"
        ),
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help=("Show what fonts would be installed without actually " "installing them"),
    )

    parser.add_argument(
        "--no-admin-check",
        action="store_true",
        help="Skip administrator privilege check and warning",
    )

    parser.add_argument(
        "--verbose", "-v", action="store_true", help="Enable verbose logging output"
    )

    parser.add_argument(
        "--overwrite",
        choices=["yes", "no", "ask"],
        default="ask",
        help=(
            'Handle existing fonts: "yes" (overwrite), "no" (skip), '
            '"ask" (prompt user) - default: ask'
        ),
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help=(
            "Force overwrite existing fonts without prompting "
            "(same as --overwrite yes)"
        ),
    )

    return parser.parse_args()


def main():
    """Main function to handle command line arguments and run installation."""
    args = parse_arguments()

    # Set up logging based on verbosity
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    print("Font Installer for Windows")
    print("=" * 40)

    # Check for admin privileges (unless skipped)
    if not args.no_admin_check and not is_admin():
        print(
            "⚠️  WARNING: This script is not running with administrator " "privileges."
        )
        print("   Some font installations may fail without admin rights.")
        print("   Consider running as administrator for best results.")
        print("   Use --no-admin-check to skip this warning.")
        print()

        response = input("Continue anyway? (y/n): ").lower().strip()
        if response not in ["y", "yes"]:
            print("Exiting...")
            return

    folder_path = args.folder_path

    # Determine overwrite mode
    if args.force:
        overwrite_mode = "yes"
    else:
        overwrite_mode = args.overwrite

    if args.dry_run:
        print("🔍 DRY RUN MODE - No fonts will be actually installed")
        print()
    else:
        if overwrite_mode == "yes":
            print(
                "🔄 OVERWRITE MODE: Existing fonts will be replaced "
                "without prompting"
            )
        elif overwrite_mode == "no":
            print("⏭️  SKIP MODE: Existing fonts will be skipped")
        else:
            print("❓ ASK MODE: You will be prompted for each existing font")
        print()

    print(f"Processing folder: {Path(folder_path).resolve()}")
    print("-" * 40)

    # Run installation (or dry run)
    if args.dry_run:
        results = preview_fonts_from_folder(folder_path)
    else:
        results = install_fonts_from_folder(folder_path, overwrite_mode)

    # Print summary
    print_summary(results, dry_run=args.dry_run)

    if not args.dry_run and int(results.get("successful", 0)) > 0:
        print("\n🎉 Font installation completed!")
        print("📝 Note: You may need to restart applications to see new fonts.")


if __name__ == "__main__":
    main()
