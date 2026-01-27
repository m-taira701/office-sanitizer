# office_sanitizer/excel.py
from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.views import Selection


@dataclass(frozen=True)
class ExcelSanitizeOptions:
    # TODO 外部変数化
    zoom: int = 100
    focus_cell: str = "A1"
    set_first_sheet_active: bool = True
    remove_comments: bool = True  # cell comments / threaded comments
    remove_metadata: bool = True  # core/custom/app props
    recursive: bool = True
    in_place: bool = True
    output_dir: Optional[Path] = None  # if set, writes output here (keeps filename)


def sanitize_excel(path: str | os.PathLike, options: ExcelSanitizeOptions = ExcelSanitizeOptions()) -> None:
    """
    Sanitize .xlsx files:
      - Zoom -> options.zoom
      - Active cell -> options.focus_cell (A1)
      - Active sheet -> first sheet (optional)
      - Remove comments (optional)
      - Remove metadata (docProps/core.xml, docProps/custom.xml, docProps/app.xml) (optional)

    Accepts a file or directory. Directory mode finds *.xlsx (skips temp "~$" files).
    """
    p = Path(path)

    targets: list[Path]
    if p.is_dir():
        targets = list(_iter_xlsx_files(p, recursive=options.recursive))
    elif p.is_file():
        targets = [p]
    else:
        raise FileNotFoundError(f"Path not found: {p}")

    for f in targets:
        _sanitize_one_xlsx(f, options)


def _iter_xlsx_files(root: Path, recursive: bool) -> Iterable[Path]:
    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    for f in root.glob(pattern):
        # Skip Office temp files like "~$foo.xlsx"
        if f.name.startswith("~$"):
            continue
        # Skip non-files (just in case)
        if not f.is_file():
            continue
        yield f


def _sanitize_one_xlsx(src: Path, options: ExcelSanitizeOptions) -> None:
    if src.suffix.lower() != ".xlsx":
        return

    if options.output_dir is not None:
        options.output_dir.mkdir(parents=True, exist_ok=True)
        dst = options.output_dir / src.name
    else:
        dst = src if options.in_place else src.with_name(src.stem + ".sanitized.xlsx")

    # Work in a temp dir to avoid corrupting the original if something fails mid-way
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        tmp_workbook_path = td_path / "workbook.xlsx"

        # Copy source to temp first (important for in_place writes)
        shutil.copy2(src, tmp_workbook_path)

        # 1) openpyxl pass: zoom, selection, comments, basic properties
        _openpyxl_sanitize(tmp_workbook_path, options)

        # 2) zip-level pass: aggressively blank docProps/* metadata
        if options.remove_metadata:
            _zip_sanitize_docprops(tmp_workbook_path)

        # Move into place
        if dst.resolve() == src.resolve():
            # in-place overwrite
            shutil.copy2(tmp_workbook_path, dst)
        else:
            shutil.copy2(tmp_workbook_path, dst)


def _openpyxl_sanitize(xlsx_path: Path, options: ExcelSanitizeOptions) -> None:
    # data_only=False to preserve formulas etc.
    wb = load_workbook(filename=str(xlsx_path), data_only=False, keep_vba=False)

    # Active sheet -> first sheet
    if options.set_first_sheet_active and wb.sheetnames:
        wb.active = 0

    # Remove comments (cell comments + (best-effort) threaded comments)
    if options.remove_comments:
        for ws in wb.worksheets:
            # Cell comments
            for row in ws.iter_rows():
                for cell in row:
                    if cell.comment is not None:
                        cell.comment = None

            # Threaded comments exist in newer Excel formats; openpyxl support is partial.
            # Best effort: drop known attributes if present.
            if hasattr(ws, "_comments"):
                try:
                    ws._comments = []  # type: ignore[attr-defined]
                except Exception:
                    pass

    # Zoom + focus cell
    focus = options.focus_cell
    for ws in wb.worksheets:
        sv = ws.sheet_view

        # Zoom
        try:
            sv.zoomScale = int(options.zoom)
        except Exception:
            pass

        # Top-left & selection
        try:
            sv.topLeftCell = focus
        except Exception:
            pass

        # Ensure there is a Selection object; openpyxl keeps a list
        try:
            if not sv.selection:
                sv.selection = [Selection(activeCell=focus, sqref=focus)]
            else:
                sv.selection[0].activeCell = focus
                sv.selection[0].sqref = focus
        except Exception:
            pass

    # Workbook properties: blank common fields (zip pass will hard-blank docProps later)
    if options.remove_metadata:
        props = wb.properties
        # OpenXML core props equivalents (best-effort)
        for attr in (
            "creator",
            "lastModifiedBy",
            "title",
            "subject",
            "description",
            "keywords",
            "category",
            "contentStatus",
            "identifier",
            "language",
            "version",
        ):
            if hasattr(props, attr):
                try:
                    setattr(props, attr, None)
                except Exception:
                    pass

        # created/modified timestamps
        for attr in ("created", "modified"):
            if hasattr(props, attr):
                try:
                    setattr(props, attr, None)
                except Exception:
                    pass

    wb.save(filename=str(xlsx_path))
    wb.close()


# --- Zip/XML sanitization (docProps) -----------------------------------------

_CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

_CUSTOM_NS = {
    "cp": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}

_APP_NS = {
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}


def _zip_sanitize_docprops(xlsx_path: Path) -> None:
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        entries = {info.filename: zin.read(info.filename) for info in zin.infolist()}
        compression = zin.compression

    if "docProps/core.xml" in entries:
        entries["docProps/core.xml"] = _minimal_core_xml()
    if "docProps/custom.xml" in entries:
        entries["docProps/custom.xml"] = _minimal_custom_xml()
    if "docProps/app.xml" in entries:
        entries["docProps/app.xml"] = _minimal_app_xml()

    tmp = xlsx_path.with_suffix(".xlsx.tmp")
    try:
        with zipfile.ZipFile(tmp, "w", compression=compression) as zout:
            for name, data in entries.items():
                zout.writestr(name, data)
        tmp.replace(xlsx_path)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass


def _minimal_core_xml() -> bytes:
    # Minimal core-properties container with namespaces declared.
    # Leaving it empty avoids leaking creator/modifiedBy/timestamps/etc.
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<cp:coreProperties'
        f' xmlns:cp="{_CORE_NS["cp"]}"'
        f' xmlns:dc="{_CORE_NS["dc"]}"'
        f' xmlns:dcterms="{_CORE_NS["dcterms"]}"'
        f' xmlns:dcmitype="{_CORE_NS["dcmitype"]}"'
        f' xmlns:xsi="{_CORE_NS["xsi"]}">'
        f"</cp:coreProperties>"
    )
    return xml.encode("utf-8")


def _minimal_custom_xml() -> bytes:
    # Minimal custom-properties container.
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="{_CUSTOM_NS["cp"]}" xmlns:vt="{_CUSTOM_NS["vt"]}">'
        f"</Properties>"
    )
    return xml.encode("utf-8")


def _minimal_app_xml() -> bytes:
    # Extended properties can include Company/Manager etc.
    # Provide minimal container; Office regenerates some fields but avoids embedding org data.
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="{_APP_NS["ep"]}" xmlns:vt="{_APP_NS["vt"]}">'
        f"</Properties>"
    )
    return xml.encode("utf-8")
