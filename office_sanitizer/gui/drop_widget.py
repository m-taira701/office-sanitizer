# office_sanitizer/gui/drop_widget.py
from PySide6.QtCore import Qt, Signal, QMimeData
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor
from PySide6.QtWidgets import QFrame, QLabel, QVBoxLayout

class DropWidget(QFrame):
    """
    ドラッグ＆ドロップを受け付けるウィジェット。
    モダンなデザインで、ファイルをドロップする領域を表示します。
    """
    files_dropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setup_ui()

    def setup_ui(self):
        # スタイルの設定
        self.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.setStyleSheet("""
            DropWidget {
                border: 2px dashed #5c5c5c;
                border-radius: 10px;
                background-color: #2b2b2b;
            }
            DropWidget:hover {
                border-color: #3daee9;
                background-color: #323232;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.label = QLabel("ここにOfficeファイルをドラッグ＆ドロップしてください\n(.docx, .xlsx, .pptx)")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("color: #aaaaaa; font-size: 16px; font-weight: bold;")
        
        layout.addWidget(self.label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                DropWidget {
                    border: 2px dashed #3daee9;
                    background-color: #3b3b3b;
                }
            """)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            DropWidget {
                border: 2px dashed #5c5c5c;
                border-radius: 10px;
                background-color: #2b2b2b;
            }
        """)

    def dropEvent(self, event: QDropEvent):
        files = []
        for url in event.mimeData().urls():
            files.append(url.toLocalFile())
        
        if files:
            self.files_dropped.emit(files)
        
        # リセットスタイル
        self.setStyleSheet("""
            DropWidget {
                border: 2px dashed #5c5c5c;
                border-radius: 10px;
                background-color: #2b2b2b;
            }
        """)
        event.acceptProposedAction()
