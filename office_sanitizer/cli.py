import argparse
import sys
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from office_sanitizer.excel import sanitize_excel
from office_sanitizer.pptx import sanitize_pptx

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    args = parser.parse_args()

    p = Path(args.path)
    if p.is_dir():
        sanitize_excel(args.path)
        sanitize_pptx(args.path)
        return

    if p.is_file():
        ext = p.suffix.lower()
        if ext == ".xlsx":
            sanitize_excel(args.path)
            return
        if ext == ".pptx":
            sanitize_pptx(args.path)
            return

    raise ValueError(f"Unsupported path: {args.path}")

if __name__ == "__main__":
    main()
