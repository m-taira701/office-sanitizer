# office_sanitizer/gui/main_window.py
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QThreadPool
from PySide6.QtGui import QAction, QIcon, QPalette, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QCheckBox, QGroupBox, QTextEdit, QListWidget,
    QSplitter, QLabel, QFileDialog, QMessageBox, QProgressBar
)

from .drop_widget import DropWidget
from .worker import SanitizeWorker

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Office Sanitizer")
        self.resize(1000, 700)
        
        self.threadpool = QThreadPool()
        self.file_list = []
        
        self.setup_ui()
        self.apply_dark_theme()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        
        # Splitter: Left (Files) | Right (Settings)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # --- Left Panel ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        # Drop Area
        self.drop_widget = DropWidget()
        self.drop_widget.files_dropped.connect(self.add_files)
        left_layout.addWidget(self.drop_widget, stretch=1)
        
        # File List
        left_layout.addWidget(QLabel("処理対象ファイル:"))
        self.list_widget = QListWidget()
        left_layout.addWidget(self.list_widget, stretch=2)
        
        btn_layout = QHBoxLayout()
        self.clear_btn = QPushButton("リストクリア")
        self.clear_btn.clicked.connect(self.clear_files)
        btn_layout.addWidget(self.clear_btn)
        
        self.add_btn = QPushButton("ファイル追加...")
        self.add_btn.clicked.connect(self.open_file_dialog)
        btn_layout.addWidget(self.add_btn)
        
        left_layout.addLayout(btn_layout)
        
        splitter.addWidget(left_widget)
        
        # --- Right Panel (Settings) ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 0, 0, 0)
        
        # Settings Group
        settings_group = QGroupBox("サニタイズ設定")
        settings_layout = QVBoxLayout()
        
        self.check_metadata = QCheckBox("メタデータ削除 (作成者、プロパティ等)")
        self.check_metadata.setChecked(True)
        self.check_metadata.setToolTip("docProps/core.xml 等を削除します")
        settings_layout.addWidget(self.check_metadata)
        
        self.check_comments = QCheckBox("コメント削除")
        self.check_comments.setChecked(True)
        self.check_comments.setToolTip("Word/Excel/PPT内のコメントを削除します")
        settings_layout.addWidget(self.check_comments)
        
        self.check_revisions = QCheckBox("変更履歴削除 (Wordのみ)")
        self.check_revisions.setChecked(True)
        settings_layout.addWidget(self.check_revisions)
        
        self.check_inplace = QCheckBox("上書き保存 (注意)")
        self.check_inplace.setChecked(False)
        self.check_inplace.setToolTip("チェックすると元のファイルを上書きします。オフなら別名で保存します。")
        settings_layout.addWidget(self.check_inplace)
        
        settings_group.setLayout(settings_layout)
        right_layout.addWidget(settings_group)
        
        # Action Buttons
        self.run_btn = QPushButton("サニタイズ実行")
        self.run_btn.setFixedHeight(50)
        self.run_btn.setStyleSheet("font-weight: bold; font-size: 14px; background-color: #3daee9; color: white;")
        self.run_btn.clicked.connect(self.run_sanitization)
        right_layout.addWidget(self.run_btn)
        
        right_layout.addStretch()
        
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 1)
        
        main_layout.addWidget(splitter, stretch=3)
        
        # --- Bottom Panel (Logs) ---
        log_group = QGroupBox("処理ログ")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("background-color: #1e1e1e; color: #dcdcdc; font-family: Consolas, monospace;")
        log_layout.addWidget(self.log_text)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        log_layout.addWidget(self.progress_bar)
        
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group, stretch=1)

    def apply_dark_theme(self):
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
        palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
        QApplication.setPalette(palette)
        
    def add_files(self, files: list[str]):
        for f in files:
            path = Path(f)
            if path not in self.file_list:
                self.file_list.append(path)
                self.list_widget.addItem(str(path))
    
    def clear_files(self):
        self.file_list.clear()
        self.list_widget.clear()
        
    def open_file_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "ファイルを選択", 
            "", 
            "Office Files (*.docx *.xlsx *.pptx)"
        )
        if files:
            self.add_files(files)
            
    def run_sanitization(self):
        if not self.file_list:
            QMessageBox.warning(self, "警告", "ファイルが選択されていません。")
            return
            
        self.log_message("--- サニタイズ処理開始 ---")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, len(self.file_list))
        self.progress_bar.setValue(0)
        self.run_btn.setEnabled(False)
        
        self.completed_count = 0
        
        options = {
            "remove_metadata": self.check_metadata.isChecked(),
            "remove_comments": self.check_comments.isChecked(),
            "remove_revisions": self.check_revisions.isChecked(),
            "in_place": self.check_inplace.isChecked(),
        }
        
        for f in self.file_list:
            worker = SanitizeWorker(f, options)
            worker.signals.log.connect(self.log_message)
            worker.signals.finished.connect(self.on_worker_finished)
            self.threadpool.start(worker)
            
    def log_message(self, msg: str):
        self.log_text.append(msg)
        # スクロール
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())
        
    def on_worker_finished(self):
        self.completed_count += 1
        self.progress_bar.setValue(self.completed_count)
        
        if self.completed_count >= len(self.file_list):
            self.log_message("--- 全処理完了 ---")
            self.run_btn.setEnabled(True)
            self.progress_bar.setVisible(False)
            QMessageBox.information(self, "完了", "すべての処理が完了しました。")

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
