"""The vocabulary of a report: findings, levels, and honest gaps."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class Level(str, Enum):
    """How much a finding should worry the reader.

    HIGH   the number a reader would quote may be wrong or unexplainable
    MEDIUM the sheet hides something, or depends on something outside itself
    LOW    a nuisance that makes review harder without changing a total
    """

    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"

    @property
    def rank(self) -> int:
        return {"high": 3, "medium": 2, "low": 1}[self.value]


@dataclass(frozen=True)
class Finding:
    """One checkable observation, anchored to a place in the workbook.

    ``detail`` must describe structure, never cell contents. Anything that could
    reproduce the data itself belongs in ``sample``, which is withheld unless the
    caller explicitly asks to see values.
    """

    check: str
    level: Level
    sheet: str
    location: str
    summary: str
    detail: str = ""
    sample: str | None = None

    def as_dict(self, show_values: bool = False) -> dict:
        out = {
            "check": self.check,
            "level": self.level.value,
            "sheet": self.sheet,
            "location": self.location,
            "summary": self.summary,
        }
        if self.detail:
            out["detail"] = self.detail
        if show_values and self.sample is not None:
            out["sample"] = self.sample
        return out


@dataclass(frozen=True)
class Unchecked:
    """Something the tool deliberately did not or could not verify."""

    topic: str
    reason: str

    def as_dict(self) -> dict:
        return {"topic": self.topic, "reason": self.reason}


@dataclass
class Report:
    """Everything one run learned about one file."""

    path: str
    findings: list[Finding] = field(default_factory=list)
    unchecked: list[Unchecked] = field(default_factory=list)
    sheets_read: list[str] = field(default_factory=list)

    def add(self, finding: Finding) -> None:
        self.findings.append(finding)

    def note_unchecked(self, topic: str, reason: str) -> None:
        self.unchecked.append(Unchecked(topic, reason))

    def sorted_findings(self) -> list[Finding]:
        return sorted(
            self.findings,
            key=lambda f: (-f.level.rank, f.sheet, f.check, f.location),
        )

    def counts(self) -> dict[str, int]:
        counts = {"high": 0, "medium": 0, "low": 0}
        for f in self.findings:
            counts[f.level.value] += 1
        return counts

    def worst_level(self) -> Level | None:
        if not self.findings:
            return None
        return max((f.level for f in self.findings), key=lambda level: level.rank)
