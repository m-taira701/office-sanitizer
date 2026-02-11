
import sys
import unittest
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

class TestImports(unittest.TestCase):
    def test_core_imports(self):
        try:
            from office_sanitizer.core.base import Sanitizer
            from office_sanitizer.core.docx_sanitizer import DocxSanitizer
            from office_sanitizer.core.xlsx_sanitizer import XlsxSanitizer
            from office_sanitizer.core.pptx_sanitizer import PptxSanitizer
            from office_sanitizer.core.utils import iter_office_files
        except ImportError as e:
            self.fail(f"Failed to import core modules: {e}")

    def test_gui_imports(self):
        try:
            from office_sanitizer.gui.main_window import MainWindow
            from office_sanitizer.gui.drop_widget import DropWidget
            from office_sanitizer.gui.worker import SanitizeWorker
        except ImportError as e:
            self.fail(f"Failed to import GUI modules: {e}")

if __name__ == '__main__':
    unittest.main()
