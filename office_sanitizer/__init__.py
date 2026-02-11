from .core.xlsx_sanitizer import XlsxSanitizer, ExcelSanitizeOptions
from .core.docx_sanitizer import DocxSanitizer, WordSanitizeOptions
from .core.pptx_sanitizer import PptxSanitizer, PowerPointSanitizeOptions

# Legacy compatibility (optional, but good for keeping API stable)
# sanitize_excel = XlsxSanitizer().sanitize ... 
# But for now, let's just expose the classes

__all__ = [
    "XlsxSanitizer",
    "ExcelSanitizeOptions",
    "DocxSanitizer",
    "WordSanitizeOptions",
    "PptxSanitizer",
    "PowerPointSanitizeOptions",
]
