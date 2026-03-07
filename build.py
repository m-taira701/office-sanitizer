import os
import sys
import shutil
import subprocess
from pathlib import Path

def main():
    """
    Build script for Office Sanitizer using PyInstaller.
    Creates a single-file executable.
    """
    print("--- Building Office Sanitizer with PyInstaller ---")

    # Project root
    root_dir = Path(__file__).parent.resolve()
    
    # Entry point
    script_path = root_dir / "main.py"
    if not script_path.exists():
        print(f"Error: {script_path} not found.")
        sys.exit(1)

    # Output name
    app_name = "OfficeSanitizer.exe" if sys.platform == "win32" else "OfficeSanitizer.bin"
    
    # PyInstaller arguments
    args = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed", # Disable console
        "--name", app_name.replace(".exe", "").replace(".bin", ""),
    ]

    # Add metadata and icon for Windows
    if sys.platform == "win32":
        icon_path = root_dir / "resources" / "icon.ico"
        if icon_path.exists():
            args.extend(["--icon", str(icon_path)])
            print(f"Using icon: {icon_path}")
            
        # Optional: Adding further metadata if needed, usually version info is in a separate text file for PyInstaller
        # Pyinstaller handles metadata differently, often requiring a version file.
        # For now, we will rely on PyInstaller's standard packing.

    # Add main script
    args.append(str(script_path))

    print(f"Running command: {' '.join(args)}")
    
    try:
        subprocess.check_call(args, cwd=root_dir)
        print("\n--- Build Successful! ---")
        
        # Determine output location
        # PyInstaller puts the output in the 'dist' folder by default
        dist_dir = root_dir / "dist"
        
        # PyInstaller uses the --name for the output file
        expected_output_name = f"{app_name.replace('.exe', '').replace('.bin', '')}"
        if sys.platform == "win32":
             expected_output_name += ".exe"
             
        exe_path = dist_dir / expected_output_name
        
        if exe_path.exists() and app_name != expected_output_name:
            # Rename if necessary for CI consistency (e.g. Mac wants .bin)
            target_path = dist_dir / app_name
            if target_path.exists():
                os.remove(target_path)
            shutil.move(str(exe_path), str(target_path))
            print(f"Moved/Renamed to: {target_path}")
        elif exe_path.exists():
             print(f"Executable ready at: {exe_path}")
        else:
             print("Warning: Expected output not found in dist folder.")

        # Clean up PyInstaller temp files
        build_dir = root_dir / "build"
        spec_file = root_dir / f"{app_name.replace('.exe', '').replace('.bin', '')}.spec"
        if build_dir.exists():
             shutil.rmtree(build_dir)
        if spec_file.exists():
             os.remove(spec_file)

    except subprocess.CalledProcessError as e:
        print(f"\n--- Build Failed! ---")
        print(f"Error code: {e.returncode}")
        sys.exit(1)

if __name__ == "__main__":
    main()
