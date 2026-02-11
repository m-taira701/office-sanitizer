# office_sanitizer/common.py
from __future__ import annotations

import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Optional


@dataclass(frozen=True)
class CommonSanitizeOptions:
    remove_metadata: bool = True  # core/custom/app props
    recursive: bool = True
    in_place: bool = True
    output_dir: Optional[Path] = None  # if set, writes output here (keeps filename)


def iter_office_files(root: Path, pattern: str, recursive: bool, skip_prefix: Optional[str] = None) -> Iterable[Path]:
    glob_pattern = f"**/{pattern}" if recursive else pattern
    for f in root.glob(glob_pattern):
        if skip_prefix and f.name.startswith(skip_prefix):
            continue
        if not f.is_file():
            continue
        yield f


def resolve_output_path(src: Path, in_place: bool, output_dir: Optional[Path], suffix: str) -> Path:
    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir / src.name
    if in_place:
        return src
    return src.with_name(src.stem + suffix)


def process_with_temp_copy(src: Path, dst: Path, processor: Callable[[Path], None]) -> None:
    # Work in a temp dir to avoid corrupting the original if something fails mid-way
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        tmp_path = td_path / src.name

        # Copy source to temp first (important for in_place writes)
        shutil.copy2(src, tmp_path)

        # Mutate temp file
        processor(tmp_path)

        # Move into place
        if dst.resolve() == src.resolve():
            shutil.copy2(tmp_path, dst)
        else:
            shutil.copy2(tmp_path, dst)


class _DeleteEntry:
    pass


DELETE_ENTRY = _DeleteEntry()


def zip_rewrite(path: Path, rewrite_fn: Callable[[str, bytes], bytes | None | _DeleteEntry]) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        entries = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]
        compression = zin.compression

    out_entries: list[tuple[str, bytes]] = []
    for name, data in entries:
        new_data = rewrite_fn(name, data)
        if new_data is DELETE_ENTRY:
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


def minimal_core_xml() -> bytes:
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


def minimal_custom_xml() -> bytes:
    # Minimal custom-properties container.
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="{_CUSTOM_NS["cp"]}" xmlns:vt="{_CUSTOM_NS["vt"]}">'
        f"</Properties>"
    )
    return xml.encode("utf-8")


def minimal_app_xml() -> bytes:
    # Extended properties can include Company/Manager etc.
    # Provide minimal container; Office regenerates some fields but avoids embedding org data.
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Properties xmlns="{_APP_NS["ep"]}" xmlns:vt="{_APP_NS["vt"]}">'
        f"</Properties>"
    )
    return xml.encode("utf-8")


def zip_sanitize_docprops(xlsx_path: Path) -> None:
    def rewrite(name: str, data: bytes) -> bytes | None:
        if name == "docProps/core.xml":
            return minimal_core_xml()
        if name == "docProps/custom.xml":
            return minimal_custom_xml()
        if name == "docProps/app.xml":
            return minimal_app_xml()
        return None

    zip_rewrite(xlsx_path, rewrite)
