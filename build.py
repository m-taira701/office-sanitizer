import os
import sys
import shutil
import subprocess
from pathlib import Path

def main():
    """
    Build script for Office Sanitizer using Nuitka.
    Creates a single-file executable.
    """
    print("--- Building Office Sanitizer with Nuitka ---")

    # Project root
    root_dir = Path(__file__).parent.resolve()
    
    # Entry point
    script_path = root_dir / "main.py"
    if not script_path.exists():
        print(f"Error: {script_path} not found.")
        sys.exit(1)

    # Output name
    app_name = "OfficeSanitizer.exe" if sys.platform == "win32" else "OfficeSanitizer.bin"
    
    # Nuitka arguments
    # Note: --onefile requires zstandard
    args = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--onefile",
        "--enable-plugin=pyside6",
        "--include-package=office_sanitizer",
        "--output-filename=" + app_name,
        "--assume-yes-for-downloads", # Allow downloading ccache/dependencies if needed
    ]

    # Windows specific
    if sys.platform == "win32":
        args.append("--windows-disable-console")
        
        # Add metadata to reduce false positives
        args.extend([
            "--windows-company-name=OfficeSanitizerProject",
            "--windows-product-name=Office Sanitizer",
            "--windows-file-version=1.0.0.0",
            "--windows-product-version=1.0.0.0",
            "--windows-file-description=Office Sanitizer Application",
            "--copyright=Copyright (C) 2026 OfficeSanitizer Project",
        ])

        icon_path = root_dir / "resources" / "icon.ico"
        if icon_path.exists():
            args.append(f"--windows-icon-from-ico={icon_path}")
            print(f"Using icon: {icon_path}")
    
    # Mac specific
    elif sys.platform == "darwin":
        args.append("--macos-disable-console")
        pass

    # Add main script
    args.append(str(script_path))

    print(f"Running command: {' '.join(args)}")
    
    try:
        subprocess.check_call(args, cwd=root_dir)
        print("\n--- Build Successful! ---")
        
        # Determine output location
        # Nuitka onefile puts the exe in the cwd by default
        exe_path = root_dir / app_name
        if exe_path.exists():
            print(f"Executable: {exe_path}")
            # Move to dist/ for consistency 
            dist_dir = root_dir / "dist"
            dist_dir.mkdir(exist_ok=True)
            target_path = dist_dir / app_name
            
            # Remove existing if any
            if target_path.exists():
                os.remove(target_path)
                
            shutil.move(str(exe_path), str(target_path))
            print(f"Moved to: {target_path}")

    except subprocess.CalledProcessError as e:
        print(f"\n--- Build Failed! ---")
        print(f"Error code: {e.returncode}")
        sys.exit(1)

if __name__ == "__main__":
    main()
