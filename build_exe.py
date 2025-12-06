import os
import shutil
import PyInstaller.__main__
import sys
from pathlib import Path

def main():
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        os.system("pip install pyinstaller")
    
    # Define paths
    base_dir = Path(__file__).parent.absolute()
    build_dir = base_dir / 'build'
    dist_dir = base_dir / 'dist'
    
    # Clean previous builds
    for folder in [build_dir, dist_dir]:
        if folder.exists():
            print(f"Removing existing {folder}...")
            shutil.rmtree(folder, ignore_errors=True)
    
    # Create a temporary directory for additional files
    temp_dir = base_dir / 'temp_build'
    if temp_dir.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
    temp_dir.mkdir(exist_ok=True)
    
    # Copy necessary data files
    data_files = ['arial.ttf', 'arialbd.ttf']
    for file in data_files:
        src = base_dir / file
        if src.exists():
            shutil.copy2(src, temp_dir / file)
    
    # Prepare PyInstaller command
    pyinstaller_args = [
        '--name=VulnerabilityReportConverter',
        '--onefile',
        '--windowed',  # Don't show console window
        '--add-data', f'arial.ttf;.',
        '--add-data', f'arialbd.ttf;.',
        '--icon=NONE',  # You can add an .ico file here if you have one
        '--clean',
        '--noconfirm',
        'gui.py'  # Main entry point
    ]
    
    # Run PyInstaller
    print("Building executable with PyInstaller...")
    PyInstaller.__main__.run(pyinstaller_args)
    
    # Clean up
    shutil.rmtree(temp_dir, ignore_errors=True)
    
    print("\nBuild complete!")
    print(f"Executable location: {dist_dir / 'VulnerabilityReportConverter.exe'}")
    print("\nTo distribute the application, you only need to share the .exe file.")
    print("The application will automatically include all necessary dependencies.")

if __name__ == "__main__":
    main()
