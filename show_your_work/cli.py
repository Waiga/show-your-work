"""Command line entry point."""

from __future__ import annotations

import argparse
import sys

from show_your_work import __version__
from show_your_work import report as report_module
from show_your_work.checks import ALL_CHECKS
from show_your_work.findings import Level, Report
from show_your_work.workbook import UnreadableWorkbook, load

EXIT_CLEAN = 0
EXIT_FINDINGS = 1
EXIT_UNREADABLE = 2

THRESHOLDS = {"high": Level.HIGH, "medium": Level.MEDIUM, "low": Level.LOW, "never": None}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="show-your-work",
        description=(
            "Report checkable reasons a spreadsheet's numbers may not be trustworthy. "
            "Runs entirely on your machine and never uploads the file."
        ),
        epilog=(
            "Exit codes: 0 nothing at or above the fail level, "
            "1 findings at or above it, 2 the file could not be read."
        ),
    )
    parser.add_argument(
        "workbook",
        nargs="+",
        help="path to an .xlsx or .xlsm file (more than one may be given)",
    )
    parser.add_argument(
        "--format", choices=("text", "json"), default="text", help="output format"
    )
    parser.add_argument(
        "--show-values",
        action="store_true",
        help="include cell contents in the report (withheld by default)",
    )
    parser.add_argument(
        "--verbose", action="store_true", help="explain why each finding matters"
    )
    parser.add_argument(
        "--fail-on",
        choices=tuple(THRESHOLDS),
        default="high",
        help="lowest level that makes the command exit non-zero (default: high)",
    )
    parser.add_argument("--version", action="version", version=f"show-your-work {__version__}")
    return parser


def run(path: str) -> Report:
    """Open a workbook, run every check, and return the report."""
    book = load(path)
    try:
        report = Report(path=str(book.path), sheets_read=book.sheet_titles)
        for check in ALL_CHECKS:
            check(book, report)
        return report
    finally:
        book.close()


def _exit_for(report: Report, threshold: Level | None) -> int:
    """The exit code one report earns against the fail level."""
    if threshold is None:
        return EXIT_CLEAN
    worst = report.worst_level()
    if worst is not None and worst.rank >= threshold.rank:
        return EXIT_FINDINGS
    return EXIT_CLEAN


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    threshold = THRESHOLDS[args.fail_on]

    # A pre-commit hook or a shell glob hands over several workbooks at once. Each is
    # opened, checked and reported on its own; one unreadable file does not stop the
    # rest. The command's exit code is the worst any single file earned: 2 if any could
    # not be read, otherwise 1 if any had findings at or above the fail level, else 0.
    reports: list[Report] = []
    worst_exit = EXIT_CLEAN
    for path in args.workbook:
        try:
            report = run(path)
        except UnreadableWorkbook as exc:
            print(f"show-your-work: {path}: {exc}", file=sys.stderr)
            worst_exit = EXIT_UNREADABLE
            continue
        reports.append(report)
        if worst_exit != EXIT_UNREADABLE:
            worst_exit = max(worst_exit, _exit_for(report, threshold))

    if args.format == "json":
        # One file stays a single object, so a caller can parse it directly; several
        # files become an array, so the whole of stdout remains valid JSON.
        if len(args.workbook) == 1 and reports:
            print(report_module.to_json(reports[0], show_values=args.show_values))
        else:
            objects = [
                report_module.to_json(r, show_values=args.show_values, indent=None)
                for r in reports
            ]
            print("[" + ", ".join(objects) + "]")
    else:
        for i, report in enumerate(reports):
            if i:
                print()
            print(
                report_module.to_text(
                    report, show_values=args.show_values, verbose=args.verbose
                )
            )

    return worst_exit


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
