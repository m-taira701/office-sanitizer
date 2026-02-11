# office_sanitizer/core/docx_sanitizer.py
from __future__ import annotations

import logging
import os
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
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
class WordSanitizeOptions(CommonSanitizeOptions):
    """
    Wordサニタイズオプション
    
    Attributes:
        remove_comments (bool): コメントを削除するかどうか。
        remove_revisions (bool): 変更履歴を削除するかどうか。
    """
    remove_comments: bool = True
    remove_revisions: bool = True


class DocxSanitizer(Sanitizer):
    def sanitize(self, path: str | os.PathLike, options: WordSanitizeOptions = WordSanitizeOptions()) -> None:
        """
        Wordファイル(.docx)をサニタイズします。
        
        Args:
            path: ファイルまたはディレクトリのパス
            options: サニタイズオプション (WordSanitizeOptions)
        """
        p = Path(path)

        targets: list[Path]
        if p.is_dir():
            targets = list(iter_office_files(p, "*.docx", recursive=options.recursive, skip_prefix="~$"))
        elif p.is_file():
            targets = [p]
        else:
            raise FileNotFoundError(f"Path not found: {p}")

        for f in targets:
            try:
                self._sanitize_one_docx(f, options)
                logger.info(f"サニタイズ完了: {f.name}")
            except Exception as e:
                logger.error(f"サニタイズ失敗: {f.name} - {str(e)}")

    def _sanitize_one_docx(self, src: Path, options: WordSanitizeOptions) -> None:
        if src.suffix.lower() != ".docx":
            return

        dst = resolve_output_path(src, options.in_place, options.output_dir, ".sanitized.docx")

        def processor(tmp_docx_path: Path) -> None:
            if options.remove_comments:
                _zip_sanitize_comments(tmp_docx_path)
            if options.remove_revisions:
                _zip_sanitize_revisions(tmp_docx_path)

            if options.remove_metadata:
                zip_sanitize_docprops(tmp_docx_path)

        process_with_temp_copy(src, dst, processor)


# --- XML / ZIP Utilities for Word ---

_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _minimal_comments_xml() -> bytes:
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:comments xmlns:w="{_WORD_NS}"/>'
    )
    return xml.encode("utf-8")


def _minimalize_root_xml(xml_bytes: bytes) -> bytes | None:
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return None

    new_root = ET.Element(root.tag, attrib=root.attrib)
    return ET.tostring(new_root, encoding="utf-8", xml_declaration=True)


def _strip_comments_from_xml(xml_bytes: bytes) -> bytes | None:
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return None

    targets = {"commentRangeStart", "commentRangeEnd", "commentReference", "commentId"}

    for parent in root.iter():
        # iterate over a static list since we may remove children
        for child in list(parent):
            if _local_name(child.tag) in targets:
                parent.remove(child)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def _zip_sanitize_comments(docx_path: Path) -> None:
    def rewrite(name: str, data: bytes) -> bytes | None:
        if name == "word/comments.xml":
            return _minimal_comments_xml()
        if name in {
            "word/commentsExtended.xml",
            "word/commentsIds.xml",
            "word/people.xml",
            "word/commentsEx.xml",
        }:
            return _minimalize_root_xml(data)
        if name.startswith("word/") and name.endswith(".xml"):
            stripped = _strip_comments_from_xml(data)
            return stripped
        return None

    zip_rewrite(docx_path, rewrite)


def _strip_revisions_from_xml(xml_bytes: bytes) -> bytes | None:
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return None

    targets = {
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "moveFromRangeStart",
        "moveToRangeEnd",
        "pPrChange",
        "rPrChange",
        "tblPrChange",
        "tblGridChange",
        "trPrChange",
        "tcPrChange",
        "sectPrChange",
        "numberingChange",
    }

    for parent in root.iter():
        for child in list(parent):
            if _local_name(child.tag) in targets:
                parent.remove(child)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _zip_sanitize_revisions(docx_path: Path) -> None:
    def rewrite(name: str, data: bytes) -> bytes | None:
        if name.startswith("word/") and name.endswith(".xml"):
            stripped = _strip_revisions_from_xml(data)
            return stripped
        return None

    zip_rewrite(docx_path, rewrite)
