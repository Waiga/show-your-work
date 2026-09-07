"""Show Your Work: report checkable reasons a spreadsheet's numbers may not be trustworthy.

The tool reports what it can verify from the file itself. It does not judge whether a
number is correct, and it reports what it could not check as "not checked" rather than
as a pass.
"""

# Read from the installed package rather than repeated here. A hand-written copy
# drifts the moment a release is cut, and then the tool misreports itself — which
# is exactly the kind of unexplained number it exists to find.
try:
    from importlib.metadata import PackageNotFoundError
    from importlib.metadata import version as _installed_version

    try:
        __version__ = _installed_version("unexplained-cells")
    except PackageNotFoundError:  # running from a source tree, not installed
        __version__ = "unknown (not installed)"
except ImportError:  # pragma: no cover - Python 3.7 and earlier
    __version__ = "unknown"

from show_your_work.findings import Finding, Level, Report, Unchecked

__all__ = ["Finding", "Level", "Report", "Unchecked", "__version__"]
