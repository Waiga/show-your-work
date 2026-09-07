"""Show Your Work: report checkable reasons a spreadsheet's numbers may not be trustworthy.

The tool reports what it can verify from the file itself. It does not judge whether a
number is correct, and it reports what it could not check as "not checked" rather than
as a pass.
"""

__version__ = "0.1.0"

from show_your_work.findings import Finding, Level, Report, Unchecked

__all__ = ["Finding", "Level", "Report", "Unchecked", "__version__"]
