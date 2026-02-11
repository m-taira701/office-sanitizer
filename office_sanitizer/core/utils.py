# office_sanitizer/core/utils.py
from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Callable, Iterable, Optional

# --- Constants for XML Namespaces ---

CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

CUSTOM_NS = {
    "cp": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}

APP_NS = {
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}


# --- Resource Path Helper ---

def get_resource_path(relative_path: str) -> Path:
    """
    Get absolute path to resource.
    Nuitka preserves __file__ even in --onefile mode (pointing to the temp dir),
    so we can rely on relative paths from this file.
    """
    # office_sanitizer/core/utils.py -> office_sanitizer/core/ -> office_sanitizer/ -> root/
    base_path = Path(__file__).parent.parent.parent
    return base_path / relative_path



# --- File Iteration & Path Resolution ---

def iter_office_files(root: Path, pattern: str, recursive: bool, skip_prefix: Optional[str] = None) -> Iterable[Path]:
    """Iterate over files matching a pattern, optionally recursively."""
    glob_pattern = f"**/{pattern}" if recursive else pattern
    for f in root.glob(glob_pattern):
        if skip_prefix and f.name.startswith(skip_prefix):
            continue
        if not f.is_file():
            continue
        yield f


def resolve_output_path(src: Path, in_place: bool, output_dir: Optional[Path], suffix: str) -> Path:
    """Determine the destination path based on options."""
    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir / src.name
    if in_place:
        return src
    return src.with_name(src.stem + suffix)


def process_with_temp_copy(src: Path, dst: Path, processor: Callable[[Path], None]) -> None:
    """
    Copy source to a temp file, process it, then move/copy to destination.
    Safe against corruption/interruption.
    """
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        tmp_path = td_path / src.name

        # Copy source to temp first (important for in_place writes)
        shutil.copy2(src, tmp_path)

        # Mutate temp file
        processor(tmp_path)

        # Move into place
        if dst.resolve() == src.resolve():
            # If in-place, copy back over original
            shutil.copy2(tmp_path, dst)
        else:
            # If new destination
            shutil.copy2(tmp_path, dst)


# --- ZIP Manipulation ---

class _DeleteEntry:
    """Sentinel object to indicate an entry should be deleted."""
    pass


DELETE_ENTRY = _DeleteEntry()


def zip_rewrite(path: Path, rewrite_fn: Callable[[str, bytes], bytes | None | _DeleteEntry]) -> None:
    """
    Rewrite a zip file in-place by iterating over entries and applying a rewrite function.
    """
    with zipfile.ZipFile(path, "r") as zin:
        entries = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]
        compression = zin.compression

    out_entries: list[tuple[str, bytes]] = []
    for name, data in entries:
        new_data = rewrite_fn(name, data)
        if isinstance(new_data, _DeleteEntry):
            continue
        if new_data is None:
            new_data = data
        out_entries.append((name, new_data))

    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for name, data in out_entries:
                zout.writestr(name, data)
        tmp.replace(path)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass


# --- Metadata Sanitization Helpers ---

def minimal_core_xml() -> bytes:
    """Minimal core-properties container."""
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<cp:coreProperties'
        f' xmlns:cp="{CORE_NS["cp"]}"'
        f' xmlns:dc="{CORE_NS["dc"]}"'
        f' xmlns:dcterms="{CORE_NS["dcterms"]}"'
        f' xmlns:dcmitype="{CORE_NS["dcmitype"]}"'
        f' xmlns:xsi="{CORE_NS["xsi"]}">'
        f"</cp:coreProperties>"
    )
    return xml.encode("utf-8")


def minimal_custom_xml() -> bytes:
    """Minimal custom-properties container."""
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="{CUSTOM_NS["cp"]}" xmlns:vt="{CUSTOM_NS["vt"]}">'
        f"</Properties>"
    )
    return xml.encode("utf-8")


def minimal_app_xml() -> bytes:
    """Minimal extended-properties container."""
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="{APP_NS["ep"]}" xmlns:vt="{APP_NS["vt"]}">'
        f"</Properties>"
    )
    return xml.encode("utf-8")


def zip_sanitize_docprops(zip_path: Path) -> None:
    """Rewrite docProps/*.xml with minimal versions."""
    def rewrite(name: str, data: bytes) -> bytes | None:
        if name == "docProps/core.xml":
            return minimal_core_xml()
        if name == "docProps/custom.xml":
            return minimal_custom_xml()
        if name == "docProps/app.xml":
            return minimal_app_xml()
        return None

    zip_rewrite(zip_path, rewrite)
