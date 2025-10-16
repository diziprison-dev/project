#!/usr/bin/env python3
"""
Advanced setup script for Product Price Calculator
"""

import subprocess
import sys
import os
from pathlib import Path

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"\n{description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✓ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ {description} failed:")
        print(f"Error: {e.stderr}")
        return False

def check_python():
    """Check if Python is installed and accessible"""
    try:
        result = subprocess.run([sys.executable, "--version"], capture_output=True, text=True)
        print(f"✓ Python found: {result.stdout.strip()}")
        return True
    except Exception as e:
        print(f"✗ Python not found: {e}")
        return False

def check_excel_file():
    """Check if Excel file exists"""
    excel_file = Path("sturmmm.xlsx")
    if excel_file.exists():
        print(f"✓ Excel file found: {excel_file}")
        return True
    else:
        print(f"⚠ Excel file not found: {excel_file}")
        print("Please make sure sturmmm.xlsx is in the same folder as this script.")
        return False

def install_requirements():
    """Install required packages"""
    packages = [
        ("pandas>=1.5.0", "pandas"),
        ("openpyxl>=3.0.0", "openpyxl"),
        ("pyinstaller>=5.0.0", "pyinstaller")
    ]
    
    for package, name in packages:
        if not run_command(f"pip install {package}", f"Installing {name}"):
            return False
    return True

def test_imports():
    """Test if all required modules can be imported"""
    modules = ["pandas", "openpyxl", "tkinter"]
    
    for module in modules:
        try:
            __import__(module)
            print(f"✓ {module} import successful")
        except ImportError as e:
            print(f"✗ {module} import failed: {e}")
            return False
    return True

def main():
    print("=" * 60)
    print("Ürün Fiyat Hesaplayıcı - Advanced Setup")
    print("=" * 60)
    
    # Check Python
    if not check_python():
        print("\n✗ Python is not installed or not accessible.")
        print("Please install Python from: https://www.python.org/downloads/")
        print("Make sure to check 'Add Python to PATH' during installation.")
        input("\nPress Enter to exit...")
        return False
    
    # Check Excel file
    check_excel_file()
    
    # Install requirements
    if not install_requirements():
        print("\n✗ Package installation failed!")
        input("\nPress Enter to exit...")
        return False
    
    # Test imports
    if not test_imports():
        print("\n✗ Import test failed!")
        input("\nPress Enter to exit...")
        return False
    
    print("\n" + "=" * 60)
    print("✓ SETUP COMPLETED SUCCESSFULLY!")
    print("=" * 60)
    print("\nYou can now run the application:")
    print("1. Double-click 'run_app.bat'")
    print("2. Or run: python main.py")
    print("\nTo create Windows executable:")
    print("1. Run: python build_advanced.py")
    print("\nIMPORTANT: Make sure 'sturmmm.xlsx' is in the same folder!")
    print("=" * 60)
    
    input("\nPress Enter to exit...")
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
