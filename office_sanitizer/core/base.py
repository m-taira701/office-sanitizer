# office_sanitizer/core/base.py
from __future__ import annotations

import os
from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

@dataclass(frozen=True)
class CommonSanitizeOptions:
    """
    共通のサニタイズオプション
    
    Attributes:
        remove_metadata (bool): プロパティ（作成者、作成日時など）を削除するかどうか。
        recursive (bool): ディレクトリ指定時に再帰的に処理するかどうか。
        in_place (bool): 上書き保存するかどうか。Falseの場合は別名保存。
        output_dir (Optional[Path]): 出力先ディレクトリ。指定がある場合、ファイル名を維持してそこに保存。
    """
    remove_metadata: bool = True
    recursive: bool = True
    in_place: bool = True
    output_dir: Optional[Path] = None


class Sanitizer(ABC):
    """
    Sanitizer Base Class
    """

    @abstractmethod
    def sanitize(self, path: str | os.PathLike, options: CommonSanitizeOptions) -> None:
        """
        指定されたファイルまたはディレクトリをサニタイズします。
        """
        pass
