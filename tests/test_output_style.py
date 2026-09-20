"""No punctuation dash ever reaches the reader.

The em dash, the en dash and the spaced hyphen are banned as punctuation in
anything published under Waiga Arya's name, and a tool's own output is published
the moment somebody runs it. The header line used to read
``Show Your Work — path``; it now reads ``Show Your Work: path``.

Two tests guard it, because either one alone leaves a hole.

``test_no_punctuation_dash_reaches_any_rendered_output`` runs every format the
README documents and reads what a person would actually see. It cannot see a
message no fixture happens to trigger.

``test_no_punctuation_dash_sits_in_a_string_that_could_be_printed`` reads the
source instead, so a dash typed into a check message nothing here exercises
still fails. It walks the syntax tree rather than the text, so comments never
reach it and docstrings are skipped by name: neither is printed.

What is deliberately allowed: a line of repeated hyphens, which is a rule under
a heading, and the ``- `` that opens an item in the "Not checked" list. Those are
layout, the same way a bullet is in Markdown. A hyphen inside a compound word or
an identifier, such as ``--show-values`` or ``non-zero``, is spelling.
"""

from __future__ import annotations

import ast
import json
import pathlib
import re

import pytest

from show_your_work import report as report_module
from show_your_work.cli import build_parser, main

# Every character a reader would see as a dash, plus the ASCII sequences that
# print as one. " -- " is the one that hides from a search for the em dash while
# reading identically on the page.
DASH_FORMS = {
    "em dash": "—",
    "en dash": "–",
    "figure dash": "‒",
    "horizontal bar": "―",
    "minus sign": "−",
    "non-breaking hyphen": "‑",
    "spaced double hyphen": " -- ",
    "spaced hyphen": " - ",
}

_BULLET = re.compile(r"^\s*[-*]\s")
_RULE = set("-=_")


def punctuation_dashes(text: str) -> list[str]:
    """Every dash in ``text`` that a reader would read as punctuation.

    A line made only of rule characters is skipped, and a leading bullet marker
    is removed before the line is scanned. Both are layout, not punctuation.
    """
    found = []
    for number, line in enumerate(text.splitlines(), start=1):
        stripped = line.strip()
        if stripped and set(stripped) <= _RULE:
            continue
        scannable = _BULLET.sub("", line)
        for name, form in DASH_FORMS.items():
            if form in scannable:
                found.append(f"line {number}: {name} in {line.strip()!r}")
    return found


def _strings_from_json(blob: str) -> str:
    """Every string in a JSON document, keys included, one per line.

    ``json.dumps`` escapes a non-ASCII character by default, so an em dash
    reaches standard output as the six characters ``\\u2014`` and a search of the
    raw text would not find it. Parsing first is the only honest way to read it.
    """
    out: list[str] = []

    def walk(node):
        if isinstance(node, str):
            out.append(node)
        elif isinstance(node, dict):
            for key, value in node.items():
                out.append(key)
                walk(value)
        elif isinstance(node, list):
            for value in node:
                walk(value)

    walk(json.loads(blob))
    return "\n".join(out)


# Enough workbooks between them to make most check messages render at least once.
FIXTURES = (
    "overwritten",
    "short_total",
    "odd_formula",
    "looping",
    "hidden_things",
    "text_numbers",
    "linked",
    "clean_model",
)

# Every documented invocation from the README's "Use" section.
TEXT_FLAGS = ([], ["--verbose"], ["--show-values"], ["--verbose", "--show-values"])


@pytest.mark.parametrize("fixture", FIXTURES)
def test_no_punctuation_dash_reaches_any_rendered_output(fixture, request, capsys):
    path = request.getfixturevalue(fixture)

    for flags in TEXT_FLAGS:
        main([path, *flags])
        rendered = capsys.readouterr().out
        assert rendered.strip(), f"no output to check for {flags}"
        assert not punctuation_dashes(rendered), (
            f"text output with {flags or 'no flags'} carries a punctuation dash: "
            f"{punctuation_dashes(rendered)}"
        )

    for flags in ([], ["--show-values"]):
        main([path, "--format", "json", *flags])
        payload = capsys.readouterr().out
        assert not punctuation_dashes(_strings_from_json(payload)), (
            f"json output with {flags or 'no flags'} carries a punctuation dash"
        )


def test_no_punctuation_dash_in_a_run_over_several_files(overwritten, hidden_things, capsys):
    """A pre-commit hook hands over every staged workbook in one run."""
    main([overwritten, hidden_things])
    assert not punctuation_dashes(capsys.readouterr().out)

    main([overwritten, hidden_things, "--format", "json"])
    assert not punctuation_dashes(_strings_from_json(capsys.readouterr().out))


def test_no_punctuation_dash_in_the_help_or_the_version(capsys):
    assert not punctuation_dashes(build_parser().format_help())
    assert not punctuation_dashes(build_parser().format_usage())

    with pytest.raises(SystemExit):
        main(["--version"])
    assert not punctuation_dashes(capsys.readouterr().out)


def test_no_punctuation_dash_in_a_refusal_message(tmp_path, capsys):
    """The message a reader sees when the file cannot be read is output too."""
    main([str(tmp_path / "missing.xlsx")])
    assert not punctuation_dashes(capsys.readouterr().err)

    legacy = tmp_path / "old.xls"
    legacy.write_bytes(b"not really a workbook")
    main([str(legacy)])
    assert not punctuation_dashes(capsys.readouterr().err)


def test_the_header_line_names_the_file_without_a_dash(overwritten, capsys):
    """The one line this change was opened for."""
    main([overwritten])
    first = capsys.readouterr().out.splitlines()[0]
    assert first.startswith("Show Your Work: ")
    assert "—" not in first


def test_the_always_printed_not_checked_block_is_clean():
    """Printed on every run, findings or none, so it is checked on its own."""
    for topic, reason in report_module.NOT_CHECKED_ALWAYS:
        assert not punctuation_dashes(topic)
        assert not punctuation_dashes(reason)


# --- the source side: a message no fixture here happens to trigger ---


def _printable_strings(tree: ast.AST) -> list[str]:
    """Every string literal in a module except the docstrings.

    A comment never reaches the reader and never reaches this list either, because
    the syntax tree does not carry one. A docstring is skipped by identifying the
    node it belongs to, which is stable across Python versions in a way that
    tokenising an f-string is not.
    """
    docstrings = set()
    for node in ast.walk(tree):
        if isinstance(
            node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)
        ):
            body = getattr(node, "body", None)
            if (
                body
                and isinstance(body[0], ast.Expr)
                and isinstance(body[0].value, ast.Constant)
                and isinstance(body[0].value.value, str)
            ):
                docstrings.add(id(body[0].value))

    return [
        node.value
        for node in ast.walk(tree)
        if isinstance(node, ast.Constant)
        and isinstance(node.value, str)
        and id(node) not in docstrings
    ]


def _package_modules() -> list[pathlib.Path]:
    package = pathlib.Path(report_module.__file__).parent
    return sorted(package.rglob("*.py"))


def test_the_package_has_modules_to_read():
    """Guards the test below: an empty list would pass it without reading anything."""
    assert len(_package_modules()) >= 5


@pytest.mark.parametrize("module", _package_modules(), ids=lambda p: p.name)
def test_no_punctuation_dash_sits_in_a_string_that_could_be_printed(module):
    offences = []
    tree = ast.parse(module.read_text(encoding="utf-8"))
    for literal in _printable_strings(tree):
        offences += [f"{literal!r}: {hit}" for hit in punctuation_dashes(literal)]
    assert not offences, f"{module.name} holds a printable string with a dash: {offences}"
