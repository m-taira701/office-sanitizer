# office_sanitizer/gui/worker.py
import traceback
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, QRunnable, Signal, Slot

from ..core.base import CommonSanitizeOptions
from ..core.docx_sanitizer import DocxSanitizer, WordSanitizeOptions
from ..core.xlsx_sanitizer import XlsxSanitizer, ExcelSanitizeOptions
from ..core.pptx_sanitizer import PptxSanitizer, PowerPointSanitizeOptions


class WorkerSignals(QObject):
    """
    WorkerからGUIへの通信用シグナル
    """
    started = Signal()
    finished = Signal()
    error = Signal(str)
    result = Signal(object)
    log = Signal(str)  # ログメッセージ用


class SanitizeWorker(QRunnable):
    """
    別スレッドでサニタイズ処理を実行するWorker
    """
    def __init__(self, file_path: Path, options: dict[str, Any]):
        super().__init__()
        self.file_path = file_path
        self.options = options
        self.signals = WorkerSignals()

    @Slot()
    def run(self):
        self.signals.started.emit()
        try:
            ext = self.file_path.suffix.lower()
            
            # オプションの構築と実行
            # メタデータ削除などの共通オプション
            common_opts = {
                "remove_metadata": self.options.get("remove_metadata", True),
                "recursive": False, # 単一ファイル処理を基本とする
                "in_place": False, # GUIからは別名保存を基本とするか、設定に従う。今回は安全のためデフォルトは別名
                "output_dir": self.options.get("output_dir", None)
            }
            # in_place設定があれば上書き
            if self.options.get("in_place", False):
                common_opts["in_place"] = True

            self.signals.log.emit(f"処理開始: {self.file_path.name}")

            if ext == ".docx":
                opts = WordSanitizeOptions(
                    **common_opts,
                    remove_comments=self.options.get("remove_comments", True),
                    remove_revisions=self.options.get("remove_revisions", True)
                )
                sanitizer = DocxSanitizer()
                sanitizer.sanitize(self.file_path, opts)

            elif ext == ".xlsx":
                opts = ExcelSanitizeOptions(
                    **common_opts,
                    remove_comments=self.options.get("remove_comments", True),
                    # Excel固有オプションがあればここで追加
                )
                sanitizer = XlsxSanitizer()
                sanitizer.sanitize(self.file_path, opts)

            elif ext == ".pptx":
                opts = PowerPointSanitizeOptions(
                    **common_opts,
                    remove_comments=self.options.get("remove_comments", True)
                )
                sanitizer = PptxSanitizer()
                sanitizer.sanitize(self.file_path, opts)

            else:
                self.signals.log.emit(f"スキップ: 未対応の拡張子 {ext}")
                self.signals.finished.emit()
                return

            self.signals.log.emit(f"完了: {self.file_path.name}")

        except Exception as e:
            err_msg = traceback.format_exc()
            self.signals.error.emit(str(e))
            self.signals.log.emit(f"エラー: {str(e)}")
        finally:
            self.signals.finished.emit()
