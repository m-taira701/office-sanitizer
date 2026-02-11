
import sys

# DEBUG: Imports for Nuitka analysis
from office_sanitizer.cli import main as cli_main
from office_sanitizer.gui.main_window import main as gui_main

def main():
    if len(sys.argv) > 1:
        cli_main()
    else:
        gui_main()

if __name__ == "__main__":
    main()
