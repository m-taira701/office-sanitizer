# office_sanitizer/cli.py
import argparse
import sys
from pathlib import Path

from .core.docx_sanitizer import DocxSanitizer, WordSanitizeOptions
from .core.xlsx_sanitizer import XlsxSanitizer, ExcelSanitizeOptions
from .core.pptx_sanitizer import PptxSanitizer, PowerPointSanitizeOptions


def main():
    parser = argparse.ArgumentParser(description="Office Sanitizer CLI")
    parser.add_argument("path", help="Path to file or directory")
    parser.add_argument("--recursive", "-r", action="store_true", help="Recursive search for directory")
    parser.add_argument("--inplace", "-i", action="store_true", help="Overwrite original files")
    parser.add_argument("--no-metadata", action="store_true", help="Do NOT remove metadata (default: remove)")
    parser.add_argument("--no-comments", action="store_true", help="Do NOT remove comments (default: remove)")
    parser.add_argument("--no-revisions", action="store_true", help="Do NOT remove revisions (Word only) (default: remove)")
    
    args = parser.parse_args()
    
    target_path = Path(args.path)
    if not target_path.exists():
        print(f"Error: Path not found: {target_path}")
        sys.exit(1)

    # Common options
    common_opts = {
        "remove_metadata": not args.no_metadata,
        "recursive": args.recursive,
        "in_place": args.inplace,
    }

    # Determine file type(s) or iterate
    # For CLI, we can just instantiate all sanitizers if directory, or specific if file.
    # But core Logic handles recursion.
    # We need to dispatch based on extension if file, or run all if dir.
    
    sanitizers = [
        (DocxSanitizer(), "docx", WordSanitizeOptions(**common_opts, remove_comments=not args.no_comments, remove_revisions=not args.no_revisions)),
        (XlsxSanitizer(), "xlsx", ExcelSanitizeOptions(**common_opts, remove_comments=not args.no_comments)),
        (PptxSanitizer(), "pptx", PowerPointSanitizeOptions(**common_opts, remove_comments=not args.no_comments)),
    ]

    print(f"Processing: {target_path}")
    
    if target_path.is_file():
        ext = target_path.suffix.lower().lstrip(".")
        processed = False
        for sanitizer, s_ext, opts in sanitizers:
            if ext == s_ext:
                sanitizer.sanitize(target_path, opts)
                processed = True
                break
        if not processed:
            print(f"Skipping unsupported file: {target_path.name}")
    else:
        # Directory mode: run all sanitizers
        for sanitizer, _, opts in sanitizers:
            sanitizer.sanitize(target_path, opts)
            
    print("Done.")

if __name__ == "__main__":
    main()
