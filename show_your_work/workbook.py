"""Opening a workbook safely, and holding the two views a check may need.

openpyxl can show a cell's formula or its last cached result, never both in one
load. Checks need both, so the file is opened twice and the two views are kept
side by side.
"""

from __future__ import annotations

import contextlib
import zipfile
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


class UnreadableWorkbook(Exception):
    """The file could not be opened as a modern Excel workbook."""


# Only formats openpyxl actually parses. .xls is a different, older binary format.
READABLE_SUFFIXES = {".xlsx", ".xlsm"}

# An .xlsx is a zip, and openpyxl parses its parts into memory with the standard
# library's XML parser. Two shapes of file abuse that: a part that unpacks to far
# more than any real workbook, and XML that declares entities to expand on parse.
# Both are refused here, before openpyxl opens the file, so a hostile file becomes a
# plain refusal (exit 2) rather than an out-of-memory kill or a silent empty report.
#
# The limits are set well above any ordinary workbook and well below a bomb. A dense
# 50,000-row sheet unpacks to about 15 MB with a best single-part ratio near 6x; a
# pathologically repetitive but legitimate one reaches about 16x. A 200 KB bomb whose
# sheet part unpacks to 200 MB drove memory past 700 MB at a ratio near 1000x.
MAX_TOTAL_UNCOMPRESSED = 512 * 1024 * 1024  # 512 MiB summed across every part
MAX_PART_RATIO = 100  # any single part's uncompressed:compressed size

# A ratio is only evidence of a bomb on a part big enough to be one. Small parts
# compress extremely well for ordinary reasons — a theme, a styles table, a sheet of
# repeated values — and judging those on ratio alone would refuse real workbooks for
# no gain, because a part this size cannot exhaust anything. Many small parts are
# still caught, by the total above.
MIN_PART_SIZE_TO_JUDGE = 4 * 1024 * 1024

# The parts openpyxl parses as XML. A real .xlsx never declares a document type or an
# entity in any of them; both are the machinery of an entity-expansion payload, so
# their mere presence is refused rather than measured.
_XML_PART_SUFFIXES = (".xml", ".rels")
_FORBIDDEN_XML_MARKERS = (b"<!DOCTYPE", b"<!ENTITY")
# A DOCTYPE and its entity declarations live in the XML prolog, before the root
# element, so scanning the first slice of each part is enough to find them and keeps
# the scan itself bounded even on a part that is individually large.
_XML_SCAN_BYTES = 1024 * 1024


def _mib(n: int) -> str:
    return f"{n / (1024 * 1024):.0f} MB"


def _preflight_container(path: Path) -> None:
    """Refuse a file shaped like a zip bomb or an XML entity-expansion payload.

    Runs before openpyxl opens the file. A file that is not a valid zip is left for
    openpyxl, whose existing error path names the corrupt-or-password case. Any refusal
    here is an UnreadableWorkbook with a plain reason, so the caller exits 2.
    """
    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        # Not a readable zip. Let load_workbook raise, so the existing message
        # (corrupt, mislabelled, or open-password) is the one the reader sees.
        return
    except Exception:
        return

    with archive:
        infos = archive.infolist()

        total = sum(info.file_size for info in infos)
        if total > MAX_TOTAL_UNCOMPRESSED:
            raise UnreadableWorkbook(
                f"This file unpacks to {_mib(total)}, far more than any ordinary "
                f"workbook (the limit is {_mib(MAX_TOTAL_UNCOMPRESSED)}). It is refused "
                "unread, because opening it could exhaust memory. If it is genuine, it "
                "is unusually large; confirm it in Excel before trusting it."
            )

        for info in infos:
            if info.compress_size <= 0 or info.file_size < MIN_PART_SIZE_TO_JUDGE:
                continue
            ratio = info.file_size / info.compress_size
            if ratio > MAX_PART_RATIO:
                raise UnreadableWorkbook(
                    f"One part of this file expands {ratio:.0f}x when unpacked "
                    f"({_mib(info.file_size)} from {info.compress_size} bytes), the "
                    "signature of a decompression bomb rather than a spreadsheet. It "
                    "is refused unread."
                )

        for info in infos:
            if not info.filename.lower().endswith(_XML_PART_SUFFIXES):
                continue
            try:
                with archive.open(info) as part:
                    head = part.read(_XML_SCAN_BYTES)
            except Exception:
                # A part that will not even open cleanly is openpyxl's problem to
                # report, not this pre-flight's.
                continue
            if any(marker in head for marker in _FORBIDDEN_XML_MARKERS):
                raise UnreadableWorkbook(
                    "This workbook's XML declares an entity or document type, which a "
                    "spreadsheet never needs and which is the mechanism of an XML "
                    "entity-expansion attack. It is refused unread."
                )


@dataclass
class LoadedWorkbook:
    """Both views of one file, plus what loading it revealed."""

    path: Path
    formulas: Workbook
    values: Workbook | None
    has_cached_values: bool
    macro_enabled: bool

    @property
    def sheet_titles(self) -> list[str]:
        return list(self.formulas.sheetnames)

    def close(self) -> None:
        for wb in (self.formulas, self.values):
            if wb is not None:
                # Closing must never be the thing that ends a run.
                with contextlib.suppress(Exception):
                    wb.close()


def _describe_open_failure(path: Path, exc: Exception) -> str:
    if path.suffix.lower() == ".xls":
        return (
            "This is the older .xls format, which this tool cannot read. "
            "Re-save it as .xlsx and run again."
        )
    if isinstance(exc, zipfile.BadZipFile):
        return (
            "The file is not a valid .xlsx container. It is either corrupt, mislabelled, "
            "or protected with an open password. These look identical from the outside, "
            "so this tool cannot tell you which."
        )
    if isinstance(exc, KeyError):
        return (
            "The file is a zip archive but does not contain a workbook. It is most "
            "likely corrupt or was renamed from another format."
        )
    return f"The file could not be opened as a workbook: {exc}"


def load(path_str: str) -> LoadedWorkbook:
    """Open a workbook for auditing, or raise UnreadableWorkbook with a plain reason."""
    path = Path(path_str)

    if not path.exists():
        raise UnreadableWorkbook(f"No such file: {path}")
    if path.is_dir():
        raise UnreadableWorkbook(f"{path} is a directory, not a workbook.")
    if path.suffix.lower() == ".xls":
        raise UnreadableWorkbook(
            "This is the older .xls format, which this tool cannot read. "
            "Re-save it as .xlsx and run again."
        )
    if path.suffix.lower() not in READABLE_SUFFIXES:
        raise UnreadableWorkbook(
            f"{path.suffix or 'This file'} is not a format this tool reads. "
            f"Supported: {', '.join(sorted(READABLE_SUFFIXES))}."
        )

    # Inspect the zip container before openpyxl parses anything into memory.
    _preflight_container(path)

    try:
        formulas = load_workbook(path, data_only=False, keep_vba=False)
    except Exception as exc:  # openpyxl raises a wide range of parse errors
        raise UnreadableWorkbook(_describe_open_failure(path, exc)) from exc

    values: Workbook | None
    try:
        values = load_workbook(path, data_only=True, keep_vba=False)
    except Exception:
        # The formula view already opened, so a failure here costs a check, not the run.
        values = None

    has_cached = _detect_cached_values(formulas, values)

    return LoadedWorkbook(
        path=path,
        formulas=formulas,
        values=values,
        has_cached_values=has_cached,
        macro_enabled=_contains_macros(path),
    )


def _contains_macros(path: Path) -> bool:
    """Look for a VBA part inside the container rather than trusting the extension.

    A macro workbook renamed to .xlsx still carries its code, so the file name is not
    evidence either way.
    """
    try:
        with zipfile.ZipFile(path) as archive:
            return any("vbaproject" in name.lower() for name in archive.namelist())
    except Exception:
        return False


def _detect_cached_values(formulas: Workbook, values: Workbook | None) -> bool:
    """True when Excel has saved results alongside the formulas.

    A workbook written by a library and never opened in Excel carries formulas with
    no stored results. Every value-based check is blind on such a file, and saying so
    is more useful than reporting a clean pass.
    """
    if values is None:
        return False

    for sheet in formulas.worksheets:
        if sheet.title not in values.sheetnames:
            continue
        value_sheet = values[sheet.title]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type != "f":
                    continue
                cached = value_sheet.cell(cell.row, cell.column).value
                if cached is not None:
                    return True
        # Keep scanning other sheets; a single formula sheet may legitimately be empty.
    return False
