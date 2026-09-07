"""The command line contract: output, exit codes, and what is withheld."""

from __future__ import annotations

import json

from show_your_work.cli import EXIT_CLEAN, EXIT_FINDINGS, EXIT_UNREADABLE, main


def test_clean_file_exits_zero(clean_model, capsys):
    assert main([clean_model]) == EXIT_CLEAN


def test_high_finding_exits_one(overwritten, capsys):
    assert main([overwritten]) == EXIT_FINDINGS


def test_fail_on_never_always_exits_zero(overwritten, capsys):
    assert main([overwritten, "--fail-on", "never"]) == EXIT_CLEAN


def test_fail_on_low_catches_a_low_finding(hidden_things, capsys):
    assert main([hidden_things, "--fail-on", "low"]) == EXIT_FINDINGS


def test_missing_file_exits_two(tmp_path, capsys):
    assert main([str(tmp_path / "nope.xlsx")]) == EXIT_UNREADABLE
    assert "No such file" in capsys.readouterr().err


def test_wrong_format_is_refused_by_name(tmp_path, capsys):
    legacy = tmp_path / "old.xls"
    legacy.write_bytes(b"not really a workbook")
    assert main([str(legacy)]) == EXIT_UNREADABLE
    assert "older .xls format" in capsys.readouterr().err


def test_corrupt_file_names_the_password_possibility(tmp_path, capsys):
    broken = tmp_path / "broken.xlsx"
    broken.write_bytes(b"PK\x03\x04 not a real archive")
    assert main([str(broken)]) == EXIT_UNREADABLE
    assert "open password" in capsys.readouterr().err


# --- privacy ---


def test_cell_contents_are_withheld_by_default(overwritten, capsys):
    main([overwritten])
    out = capsys.readouterr().out
    assert "4200" not in out
    assert "--show-values" in out


def test_show_values_includes_them(overwritten, capsys):
    main([overwritten, "--show-values"])
    assert "4200" in capsys.readouterr().out


def test_json_withholds_samples_by_default(overwritten, capsys):
    main([overwritten, "--format", "json"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["values_included"] is False
    assert all("sample" not in f for f in payload["findings"])


def test_json_is_wellformed_and_carries_not_checked(overwritten, capsys):
    main([overwritten, "--format", "json", "--show-values"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["counts"]["high"] >= 1
    assert payload["not_checked"], "the report must always say what it did not check"
    assert any(f.get("sample") for f in payload["findings"])


def test_text_report_always_lists_what_was_not_checked(clean_model, capsys):
    main([clean_model])
    out = capsys.readouterr().out
    assert "NOT CHECKED" in out
    assert "Whether any number is correct" in out


def test_clean_report_does_not_claim_the_file_is_fine(clean_model, capsys):
    main([clean_model])
    out = capsys.readouterr().out
    assert "clean bill of health" in out  # phrased as a caveat, not a pass
