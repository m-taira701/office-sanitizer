# office_sanitizer/core/pptx_sanitizer.py
from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from pathlib import Path

from .base import Sanitizer, CommonSanitizeOptions
from .utils import (
    iter_office_files,
    resolve_output_path,
    process_with_temp_copy,
    zip_rewrite,
    zip_sanitize_docprops,
)

logger = logging.getLogger(__name__)

@dataclass(frozen=True)
class PowerPointSanitizeOptions(CommonSanitizeOptions):
    """
    PowerPointサニタイズオプション
    
    Attributes:
        remove_comments (bool): コメントを削除するかどうか。
    """
    remove_comments: bool = True


class PptxSanitizer(Sanitizer):
    def sanitize(self, path: str | os.PathLike, options: PowerPointSanitizeOptions = PowerPointSanitizeOptions()) -> None:
        """
        PowerPointファイル(.pptx)をサニタイズします。
        
        Args:
            path: ファイルまたはディレクトリのパス
            options: サニタイズオプション (PowerPointSanitizeOptions)
        """
        p = Path(path)

        targets: list[Path]
        if p.is_dir():
            targets = list(iter_office_files(p, "*.pptx", recursive=options.recursive, skip_prefix="~$"))
        elif p.is_file():
            targets = [p]
        else:
            raise FileNotFoundError(f"Path not found: {p}")

        for f in targets:
            try:
                self._sanitize_one_pptx(f, options)
                logger.info(f"サニタイズ完了: {f.name}")
            except Exception as e:
                logger.error(f"サニタイズ失敗: {f.name} - {str(e)}")

    def _sanitize_one_pptx(self, src: Path, options: PowerPointSanitizeOptions) -> None:
        if src.suffix.lower() != ".pptx":
            return

        dst = resolve_output_path(src, options.in_place, options.output_dir, ".sanitized.pptx")

        def processor(tmp_pptx_path: Path) -> None:
            if options.remove_comments:
                _zip_sanitize_comments(tmp_pptx_path)

            if options.remove_metadata:
                zip_sanitize_docprops(tmp_pptx_path)

        process_with_temp_copy(src, dst, processor)


# --- XML / ZIP Utilities for PowerPoint ---

_PPT_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _minimal_comments_xml() -> bytes:
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:cmLst xmlns:p="{_PPT_NS}" xmlns:a="{_DRAW_NS}"/>'
    )
    return xml.encode("utf-8")


def _minimal_comment_authors_xml() -> bytes:
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:cmAuthorLst xmlns:p="{_PPT_NS}"/>'
    )
    return xml.encode("utf-8")


def _zip_sanitize_comments(pptx_path: Path) -> None:
    def rewrite(name: str, data: bytes) -> bytes | None:
        if name.startswith("ppt/comments") and name.endswith(".xml"):
            return _minimal_comments_xml()
        if name.startswith("ppt/slideComments") and name.endswith(".xml"):
            return _minimal_comments_xml()
        if name.startswith("ppt/") and name.endswith("commentAuthors.xml"):
            return _minimal_comment_authors_xml()
        return None

    zip_rewrite(pptx_path, rewrite)
