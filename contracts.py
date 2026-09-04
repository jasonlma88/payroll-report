#!/usr/bin/env python3
"""
Output contracts.

"The output must be compatible with QuickBooks" is not something you can test
by diffing against one good file. It is a *specification*, and it belongs in
the code as data so it can be checked against any output, including the
thousands of random ones fuzz.py generates.

A Contract is a name, an exporter, and a list of rules. A rule reads a report
(and the exported bytes) and yields human-readable violations. Adding a new
downstream system means adding a Contract, not editing the engine.

    python3 contracts.py            # list the contracts and their rules
    python3 contracts.py --export quickbooks_iif in.xlsx out.iif
"""

from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path

from payroll_core import BUCKETS, OUTPUT_COLUMNS

# --------------------------------------------------------------------------
# Framework
# --------------------------------------------------------------------------


# A rule is one of two very different things, and conflating them makes the
# whole exercise useless:
#
#   ERROR    — we produced something wrong. Always our bug. Must never happen.
#   ADVISORY — the report is correct, but this particular input cannot be
#              imported into that system. The user needs to hear it; it is not
#              a defect in the code and must not fail a build.
ERROR = "error"
ADVISORY = "advisory"


@dataclass
class Violation:
    contract: str
    rule: str
    severity: str
    message: str

    def __str__(self) -> str:
        return f"[{self.severity}] {self.contract}/{self.rule}: {self.message}"


@dataclass
class Contract:
    """What some downstream consumer requires of a payroll report."""

    name: str
    description: str
    rules: list = field(default_factory=list)      # list[(name, callable, severity)]
    exporter: object = None                        # (result) -> str, or None

    def validate(self, result) -> list:
        """Return a list of Violations; empty means the report conforms."""
        text = None
        if self.exporter is not None:
            try:
                text = self.exporter(result)
            except Exception as exc:               # an exporter that dies is our bug
                return [Violation(self.name, "exporter", ERROR,
                                  f"raised {type(exc).__name__}: {exc}")]

        violations = []
        for rule_name, rule, severity in self.rules:
            try:
                for problem in rule(result, text) or []:
                    violations.append(Violation(self.name, rule_name, severity, problem))
            except Exception as exc:
                violations.append(Violation(self.name, rule_name, ERROR,
                                            f"rule raised {type(exc).__name__}: {exc}"))
        return violations


def hours(row, column) -> float:
    """A bucket cell as a number — blank cells read as 0."""
    value = row[column]
    return 0.0 if value == "" or value is None else float(value)


# --------------------------------------------------------------------------
# Contract 1 — invariants of the transformation itself
#
# These hold no matter who consumes the report. If one of these breaks, the
# numbers are wrong, and nothing downstream can save you.
# --------------------------------------------------------------------------


def _rule_buckets_partition_total(result, _text):
    for row in result.rows:
        parts = sum(hours(row, b) for b in BUCKETS)
        if abs(parts - row["Personal Total"]) > 1e-9:
            yield (
                f"{row['Employee']}: buckets sum to {parts:g} "
                f"but Personal Total is {row['Personal Total']:g}"
            )


def _rule_no_duplicate_employees(result, _text):
    seen = set()
    for row in result.rows:
        if row["Employee"] in seen:
            yield f"{row['Employee']} appears more than once"
        seen.add(row["Employee"])


def _rule_totals_are_finite_numbers(result, _text):
    for row in result.rows:
        for column in BUCKETS + ["Personal Total"]:
            value = row[column]
            if value == "":
                continue
            if not isinstance(value, (int, float)) or value != value or value in (
                float("inf"),
                float("-inf"),
            ):
                yield f"{row['Employee']}.{column} is not a finite number: {value!r}"


def _rule_no_negative_totals(result, _text):
    for row in result.rows:
        if row["Personal Total"] < 0:
            yield f"{row['Employee']} has a negative total ({row['Personal Total']:g})"


def _rule_employee_names_nonempty(result, _text):
    for row in result.rows:
        if not str(row["Employee"]).strip():
            yield "an employee row has a blank name"


def _rule_blank_means_zero(result, _text):
    """A blank bucket must mean zero, never a dropped number."""
    for row in result.rows:
        for bucket in BUCKETS:
            if row[bucket] == "":
                continue
            if row[bucket] == 0:
                yield f"{row['Employee']}.{bucket} is a literal 0 rather than blank"


# Two weeks of wall-clock time is 336 hours. Anything past that is a typo,
# a units mix-up, or a duplicated import — never a real timesheet.
PLAUSIBLE_PERIOD_HOURS = 400


def _rule_plausible_hours(result, _text):
    for row in result.rows:
        if row["Personal Total"] > PLAUSIBLE_PERIOD_HOURS:
            yield (
                f"{row['Employee']} totals {row['Personal Total']:g} hours, "
                f"more than the {PLAUSIBLE_PERIOD_HOURS} h that fit in a pay period "
                "— check for a typo or a doubled import"
            )


CORE = Contract(
    name="core",
    description="Invariants every consumer depends on, whatever the format.",
    rules=[
        ("buckets-partition-total", _rule_buckets_partition_total, ERROR),
        ("no-duplicate-employees", _rule_no_duplicate_employees, ERROR),
        ("finite-numbers", _rule_totals_are_finite_numbers, ERROR),
        ("no-negative-totals", _rule_no_negative_totals, ADVISORY),
        ("names-non-empty", _rule_employee_names_nonempty, ERROR),
        ("blank-means-zero", _rule_blank_means_zero, ERROR),
        ("plausible-hours", _rule_plausible_hours, ADVISORY),
    ],
)


# --------------------------------------------------------------------------
# Contract 2 — QuickBooks Desktop, IIF TIMEACT records
#
# A worked example of a real external format, with the constraints that
# actually bite: tab-delimited so no tabs or newlines in any field, names
# capped at 41 characters, and durations as HH:MM so anything finer than a
# minute cannot be represented.
# --------------------------------------------------------------------------

QB_NAME_LIMIT = 41
QB_PAYROLL_ITEMS = {
    "Sick Leave": "Sick Hourly",
    "Vacation": "Vacation Hourly",
    "Work Hours": "Regular Pay",
}
_QB_FORBIDDEN = re.compile(r"[\t\r\n]")


def quickbooks_duration(value: float) -> str:
    """Hours as QuickBooks' HH:MM. 61.25 -> '61:15'."""
    total_minutes = int(round(value * 60))
    sign = "-" if total_minutes < 0 else ""
    total_minutes = abs(total_minutes)
    return f"{sign}{total_minutes // 60}:{total_minutes % 60:02d}"


def qb_field(value) -> str:
    """A value safe to place in a tab-delimited IIF record."""
    return _QB_FORBIDDEN.sub(" ", str(value)).strip()


def export_quickbooks_iif(result) -> str:
    """Render the report as an IIF file of TIMEACT records."""
    end_date = "" 
    if result.period:
        end_date = datetime.strptime(result.period.split("_to_")[1], "%Y-%m-%d").strftime(
            "%m/%d/%Y"
        )

    out = io.StringIO()
    out.write("!TIMERHDR\tVER\tREL\tCOMPANYNAME\tIMPORTEDBEFORE\tFROMTIMER\n")
    out.write("TIMERHDR\t8\t0\t\tN\tN\n")
    out.write("!TIMEACT\tDATE\tJOB\tEMP\tITEM\tPITEM\tDURATION\tXFERTOPAYROLL\n")
    for row in result.rows:
        for bucket in BUCKETS:
            value = hours(row, bucket)
            # Anything under half a minute renders as "0:00", which QuickBooks
            # rejects outright. Dropping it is closer to the truth than a
            # record claiming zero time was worked.
            if not value or quickbooks_duration(value) in ("0:00", "-0:00"):
                continue
            out.write(
                "TIMEACT\t{date}\t\t{emp}\t\t{item}\t{dur}\tY\n".format(
                    date=end_date,
                    emp=qb_field(row["Employee"]),
                    item=QB_PAYROLL_ITEMS[bucket],
                    dur=quickbooks_duration(value),
                )
            )
    return out.getvalue()


def _rule_qb_no_delimiters_in_fields(_result, text):
    """The exporter must neutralise tabs and newlines, not pass them through."""
    if text is None:
        return
    for number, line in enumerate(text.splitlines(), 1):
        if not line.startswith("TIMEACT\t"):
            continue
        for position, value in enumerate(line.split("\t")):
            if _QB_FORBIDDEN.search(value):
                yield f"line {number} field {position} still contains a delimiter: {value!r}"


def _rule_qb_name_length(result, _text):
    for row in result.rows:
        name = str(row["Employee"])
        if len(name) > QB_NAME_LIMIT:
            yield f"{name!r} is {len(name)} characters; QuickBooks caps names at {QB_NAME_LIMIT}"


def _rule_qb_duration_is_representable(result, _text):
    """QuickBooks stores HH:MM, so sub-minute hours would be silently rounded."""
    for row in result.rows:
        for bucket in BUCKETS:
            value = hours(row, bucket)
            if not value:
                continue
            if abs(value * 60 - round(value * 60)) > 1e-6:
                yield (
                    f"{row['Employee']}.{bucket} = {value!r} h is not a whole number "
                    f"of minutes; QuickBooks would round it to {quickbooks_duration(value)}"
                )


def _rule_qb_every_row_is_wellformed(result, text):
    if text is None:
        return
    expected = None
    for number, line in enumerate(text.splitlines(), 1):
        if line.startswith("!TIMEACT"):
            expected = len(line.split("\t"))
        elif line.startswith("TIMEACT"):
            if expected is None:
                yield f"line {number}: a TIMEACT record before its !TIMEACT header"
                continue
            fields = line.split("\t")
            if len(fields) != expected:
                yield (
                    f"line {number}: {len(fields)} fields, header declares {expected}"
                )


def _rule_qb_hours_round_trip(result, text):
    """Re-read the exported file and confirm the hours survived the trip."""
    if text is None:
        return
    recovered: dict[str, float] = {}
    for line in text.splitlines():
        if not line.startswith("TIMEACT\t"):
            continue
        fields = line.split("\t")
        if len(fields) < 7:
            continue
        emp, duration = fields[3], fields[6]
        try:
            sign = -1 if duration.startswith("-") else 1
            hh, mm = duration.lstrip("-").split(":")
            recovered[emp] = recovered.get(emp, 0.0) + sign * (int(hh) + int(mm) / 60)
        except ValueError:
            yield f"duration {duration!r} for {emp} is not HH:MM"
            return
    for row in result.rows:
        want = row["Personal Total"]
        got = recovered.get(qb_field(row["Employee"]), 0.0)
        if abs(want - got) > 1 / 120:          # half a minute
            yield f"{row['Employee']}: exported {got:g} h but the report says {want:g} h"


def _rule_qb_no_zero_records(result, text):
    if text is None:
        return
    for number, line in enumerate(text.splitlines(), 1):
        if line.startswith("TIMEACT\t") and line.split("\t")[6] == "0:00":
            yield f"line {number}: a zero-duration record QuickBooks will reject"


QUICKBOOKS_IIF = Contract(
    name="quickbooks-iif",
    description="QuickBooks Desktop timer import (tab-delimited IIF, TIMEACT records).",
    exporter=export_quickbooks_iif,
    rules=[
        ("no-delimiters-in-fields", _rule_qb_no_delimiters_in_fields, ERROR),
        ("name-length", _rule_qb_name_length, ADVISORY),
        ("duration-representable", _rule_qb_duration_is_representable, ADVISORY),
        ("records-well-formed", _rule_qb_every_row_is_wellformed, ERROR),
        ("hours-round-trip", _rule_qb_hours_round_trip, ERROR),
        ("no-zero-records", _rule_qb_no_zero_records, ERROR),
    ],
)


# --------------------------------------------------------------------------
# Contract 3 — a plain CSV, the lowest common denominator for any other system
# --------------------------------------------------------------------------


def export_csv(result) -> str:
    out = io.StringIO()
    writer = csv.writer(out, lineterminator="\n")
    writer.writerow(OUTPUT_COLUMNS)
    for row in result.rows:
        writer.writerow([row[c] for c in OUTPUT_COLUMNS])
    return out.getvalue()


def _rule_csv_round_trips(result, text):
    if text is None:
        return
    rows = list(csv.reader(io.StringIO(text)))
    if rows[0] != OUTPUT_COLUMNS:
        yield f"header is {rows[0]}, expected {OUTPUT_COLUMNS}"
        return
    if len(rows) - 1 != len(result.rows):
        yield f"{len(rows) - 1} data rows, expected {len(result.rows)}"
        return
    for parsed, original in zip(rows[1:], result.rows):
        if parsed[0] != str(original["Employee"]):
            yield f"employee {parsed[0]!r} != {original['Employee']!r} after a round trip"
        try:
            if abs(float(parsed[-1]) - original["Personal Total"]) > 1e-9:
                yield f"{parsed[0]}: total {parsed[-1]} != {original['Personal Total']}"
        except ValueError:
            yield f"{parsed[0]}: total {parsed[-1]!r} is not a number after a round trip"


def _rule_csv_no_formula_injection(result, _text):
    """A leading =, +, - or @ makes Excel/Sheets execute the cell on open."""
    for row in result.rows:
        name = str(row["Employee"])
        if name[:1] in ("=", "+", "-", "@"):
            yield f"{name!r} starts with {name[0]!r} and would be read as a formula"


GENERIC_CSV = Contract(
    name="generic-csv",
    description="Plain CSV for any system that is not QuickBooks.",
    exporter=export_csv,
    rules=[
        ("round-trips", _rule_csv_round_trips, ERROR),
        ("no-formula-injection", _rule_csv_no_formula_injection, ADVISORY),
    ],
)


ALL_CONTRACTS = [CORE, QUICKBOOKS_IIF, GENERIC_CSV]
BY_NAME = {c.name.replace("-", "_"): c for c in ALL_CONTRACTS}


def validate_all(result, contracts=None) -> list:
    violations = []
    for contract in contracts or ALL_CONTRACTS:
        violations.extend(contract.validate(result))
    return violations


def errors(violations) -> list:
    return [v for v in violations if v.severity == ERROR]


def advisories(violations) -> list:
    return [v for v in violations if v.severity == ADVISORY]


def main() -> int:
    import sys

    args = sys.argv[1:]
    if args[:1] == ["--export"]:
        name, source, dest = args[1], args[2], args[3]
        from payroll_core import process_timesheet

        contract = BY_NAME[name.replace("-", "_")]
        result = process_timesheet(source)
        Path(dest).write_text(contract.exporter(result), encoding="utf-8")
        problems = contract.validate(result)
        print(f"Wrote {dest} ({contract.name})")
        for problem in problems:
            print(f"  ! {problem}")
        return 1 if problems else 0

    for contract in ALL_CONTRACTS:
        print(f"{contract.name}  --  {contract.description}")
        for rule_name, _, severity in contract.rules:
            print(f"    - [{severity:9}] {rule_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
