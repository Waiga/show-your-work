"""Turning a Report into something a person or a script can read.

Cell contents stay out of the output unless the caller asks for them. A spreadsheet
under review usually holds someone's real numbers, and a report is the thing most
likely to get forwarded.
"""

from __future__ import annotations

import json

from show_your_work.findings import Level, Report

LEVEL_LABEL = {Level.HIGH: "HIGH  ", Level.MEDIUM: "MEDIUM", Level.LOW: "LOW   "}

NOT_CHECKED_ALWAYS = [
    (
        "Whether any number is correct",
        (
            "The tool reports that a number is unexplained, never that it is wrong. "
            "A hardcoded figure may well be the right one."
        ),
    ),
    (
        "Whether the formula logic matches the intent",
        (
            "A formula can be consistent, uniform and completely wrong about the "
            "business. Only a reader who knows the model can judge that."
        ),
    ),
    (
        "Anything outside the cells",
        (
            "Charts, pivot table caches, VBA code, embedded objects, data validation "
            "and conditional formatting rules are not inspected."
        ),
    ),
]


def to_json(report: Report, show_values: bool = False) -> str:
    payload = {
        "file": report.path,
        "sheets_read": report.sheets_read,
        "counts": report.counts(),
        "findings": [f.as_dict(show_values) for f in report.sorted_findings()],
        "not_checked": [u.as_dict() for u in report.unchecked]
        + [{"topic": t, "reason": r} for t, r in NOT_CHECKED_ALWAYS],
        "values_included": show_values,
    }
    return json.dumps(payload, indent=2)


def _rule(char: str = "-", width: int = 72) -> str:
    return char * width


def to_text(report: Report, show_values: bool = False, verbose: bool = False) -> str:
    lines: list[str] = []
    counts = report.counts()
    total = sum(counts.values())

    lines.append(f"Show Your Work — {report.path}")
    lines.append(_rule("="))
    lines.append(
        f"Read {len(report.sheets_read)} sheet(s). "
        f"{total} finding(s): {counts['high']} high, "
        f"{counts['medium']} medium, {counts['low']} low."
    )
    lines.append("")

    if not report.findings:
        lines.append("Nothing checkable stood out. See 'Not checked' below before")
        lines.append("treating that as a clean bill of health.")
        lines.append("")
    else:
        current_level = None
        for finding in report.sorted_findings():
            if finding.level is not current_level:
                current_level = finding.level
                lines.append(f"{finding.level.value.upper()}")
                lines.append(_rule())
            lines.append(
                f"{LEVEL_LABEL[finding.level]}  {finding.sheet}!{finding.location}"
                f"  [{finding.check}]"
            )
            lines.append(f"        {finding.summary}")
            if verbose and finding.detail:
                for chunk in _wrap(finding.detail, 62):
                    lines.append(f"        {chunk}")
            if show_values and finding.sample is not None:
                lines.append(f"        value: {finding.sample}")
            lines.append("")

    lines.append("NOT CHECKED")
    lines.append(_rule())
    for item in report.unchecked:
        lines.append(f"  - {item.topic}")
        for chunk in _wrap(item.reason, 66):
            lines.append(f"      {chunk}")
    for topic, reason in NOT_CHECKED_ALWAYS:
        lines.append(f"  - {topic}")
        for chunk in _wrap(reason, 66):
            lines.append(f"      {chunk}")
    lines.append("")

    if not show_values:
        lines.append("Cell contents were withheld. Re-run with --show-values to include them.")

    return "\n".join(lines)


def _wrap(text: str, width: int) -> list[str]:
    words, out, line = text.split(), [], ""
    for word in words:
        if line and len(line) + 1 + len(word) > width:
            out.append(line)
            line = word
        else:
            line = f"{line} {word}".strip()
    if line:
        out.append(line)
    return out
