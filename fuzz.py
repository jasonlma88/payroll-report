#!/usr/bin/env python3
"""
Property-based robustness testing.

selftest.py checks known cases. This checks *properties* — statements that must
hold for every possible timesheet, on thousands of randomly generated ones,
including deliberately hostile ones.

    python3 fuzz.py                 # 300 cases
    python3 fuzz.py -n 5000         # longer run
    python3 fuzz.py --seed 12345    # reproduce a specific failure

A failing case is shrunk to the smallest input that still fails and written to
fuzz-failure-<seed>.xlsx so it can be opened and looked at.
"""

from __future__ import annotations

import argparse
import random
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from contracts import ALL_CONTRACTS, advisories, errors, hours, validate_all  # noqa: E402
from payroll_core import (                                      # noqa: E402
    BUCKETS,
    PayrollError,
    force_text,
    process_timesheet,
    save_report,
)

HEADER = ["Person", "Description", "Project", "Hours"]

# Names chosen to break things: unicode, RTL, tabs, quotes, formula prefixes,
# very long, and the ordinary case so we still cover normality.
NASTY_NAMES = [
    "Ada Lovelace",
    "Robin (Robbie) Vale",
    "O'Brien, Seán",
    "Müller-Şahin",
    "毛大勤",
    "Иванов Иван",
    "  padded  ",
    "name\twith\ttab",
    "name\nwith\nnewline",
    '=cmd|"/c calc"!A1',
    "+1234567890",
    "-minus-lead",
    "@at-lead",
    "A" * 60,
    "x",
    "Employee; DROP TABLE",
    "emoji 🎉 name",
    "Person",                      # collides with the header keyword
    "Worked Time",                 # collides with the subtotal keyword
]

PROJECTS = [
    ("CD-SICK Sick Leave", "Sick Leave"),
    ("cd-sick lower case", "Sick Leave"),
    ("CD-VAC Paid Vacation", "Vacation"),
    ("CD-HOL Company Holiday", "Work Hours"),
    ("ACME-1 Regular Project", "Work Hours"),
    ("SICKLE-CELL-9 Sickle Cell Study", "Work Hours"),   # must NOT be sick leave
    ("VACUUM-1 Vacuum Systems", "Work Hours"),           # must NOT be vacation
    ("Contoso-CTX-301/302 Long Name", "Work Hours"),
    ("proj\twith\ttab", "Work Hours"),
]

# Adversarial but *plausible* values. Absurd magnitudes (1e12 hours) only
# generated advisory noise and exercised IEEE754, not this program.
HOUR_VALUES = [
    0, 0.25, 0.5, 1, 7.5, 8, 40, 61.25, 80, 168,
    "8", "8.5", " 7 ", "1,024.5",          # hours arriving as text
    -4,                                     # a correction entry
    "", None, "abc", True,                  # blank, missing, junk, boolean
]


ADVISORY_COUNTS: dict[str, int] = {}


class Failure(Exception):
    def __init__(self, prop: str, detail: str):
        super().__init__(f"{prop}: {detail}")
        self.prop = prop
        self.detail = detail


# --------------------------------------------------------------------------
# Generating timesheets
# --------------------------------------------------------------------------


def make_case(rng: random.Random) -> dict:
    """A random timesheet, described as data so it can be shrunk."""
    employees = []
    for _ in range(rng.randint(1, 8)):
        entries = [
            (rng.choice(PROJECTS)[0], rng.choice(HOUR_VALUES))
            for _ in range(rng.randint(0, 6))
        ]
        employees.append({"name": rng.choice(NASTY_NAMES), "entries": entries})
    return {
        "employees": employees,
        "title_rows": rng.randint(0, 4),
        "period": rng.choice([None, "8/9/26 - 8/22/26", "12/1/2026 - 12/31/2026"]),
        "blank_between": rng.random() < 0.8,
        "state_subtotal": rng.random() < 0.7,
        "trailing_blanks": rng.randint(0, 3),
    }


def numeric(value) -> float:
    """What the engine should read a cell as, for the oracle below."""
    if value is None or isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except ValueError:
        return 0.0


def write_case(case: dict, path: Path) -> Path:
    book = Workbook()
    sheet = book.active
    sheet.title = "Payroll Report (jobs)"

    for i in range(case["title_rows"]):
        sheet.append([f"Some title {i}", None, None, None])
    if case["period"]:
        sheet.append([case["period"], None, None, None])
    sheet.append(HEADER)

    for employee in case["employees"]:
        total = sum(numeric(h) for _, h in employee["entries"])
        sheet.append([employee["name"], None, None, total if case["state_subtotal"] else None])
        force_text(sheet.cell(row=sheet.max_row, column=1))
        sheet.append([None, "Worked Time", None, total])
        for project, value in employee["entries"]:
            sheet.append([None, None, project, value])
            force_text(sheet.cell(row=sheet.max_row, column=3))
        if case["blank_between"]:
            sheet.append([None, None, None, None])
    for _ in range(case["trailing_blanks"]):
        sheet.append([None, None, None, None])

    book.save(path)
    return path


# --------------------------------------------------------------------------
# The properties
# --------------------------------------------------------------------------


def expected_totals(case: dict) -> dict:
    """An independent oracle: what the buckets should be, computed a different way."""
    wanted: dict[str, dict] = {}
    for employee in case["employees"]:
        name = str(employee["name"]).strip()
        if not name or name.lower() == "person":
            continue
        entry = wanted.setdefault(name, {b: 0.0 for b in BUCKETS})
        for project, value in employee["entries"]:
            lowered = project.lower()
            if "cd-sick" in lowered or "sick leave" in lowered:
                bucket = "Sick Leave"
            elif "cd-vac" in lowered or "paid vacation" in lowered:
                bucket = "Vacation"
            else:
                bucket = "Work Hours"
            entry[bucket] += numeric(value)
    return wanted


def check_case(case: dict, tmp: Path, index: int) -> None:
    """Run every property against one generated timesheet."""
    source = write_case(case, tmp / f"case{index}.xlsx")

    try:
        result = process_timesheet(source)
    except PayrollError:
        # A clean refusal is always an acceptable outcome; a crash is not.
        return
    except Exception as exc:
        raise Failure("no-uncaught-exceptions", f"{type(exc).__name__}: {exc}") from exc

    # 1. Conservation — nothing gained, nothing lost, checked against the oracle.
    wanted = expected_totals(case)
    got = {str(r["Employee"]): r for r in result.rows}
    for name, buckets in wanted.items():
        if name not in got:
            raise Failure("no-employee-dropped", f"{name!r} is missing from the report")
        for bucket, value in buckets.items():
            actual = hours(got[name], bucket)
            if abs(actual - value) > 1e-6 * max(1.0, abs(value)):
                raise Failure(
                    "hours-conserved", f"{name!r}.{bucket}: got {actual!r}, expected {value!r}"
                )
    for name in got:
        if name not in wanted:
            raise Failure("no-employee-invented", f"{name!r} is in the report but not the input")

    # 2. Every declared contract, including QuickBooks. Advisories describe
    #    input that some system cannot accept; only errors mean we are broken.
    found = validate_all(result)
    for violation in errors(found):
        raise Failure(f"contract:{violation.contract}/{violation.rule}", violation.message)
    for violation in advisories(found):
        ADVISORY_COUNTS[f"{violation.contract}/{violation.rule}"] = (
            ADVISORY_COUNTS.get(f"{violation.contract}/{violation.rule}", 0) + 1
        )

    # 3. Determinism — the same input twice must give the same answer.
    again = process_timesheet(source)
    if [dict(r) for r in again.rows] != [dict(r) for r in result.rows]:
        raise Failure("deterministic", "two runs on the same file disagreed")

    # 4. Excel round trip — what we wrote back must read as what we computed.
    written = save_report(result, tmp / f"out{index}.xlsx", overwrite=True)
    reread = _read_report(written)
    if len(reread) != len(result.rows):
        raise Failure("excel-round-trip", f"wrote {len(result.rows)} rows, read back {len(reread)}")
    for original, back in zip(result.rows, reread):
        if str(original["Employee"]).strip() != str(back[0] or "").strip():
            raise Failure("excel-round-trip", f"{original['Employee']!r} came back as {back[0]!r}")
        # Excel stores 15 significant digits, so the comparison has to be
        # relative. An absolute epsilon fails on values no payroll will ever see.
        want, back_value = original["Personal Total"], (back[4] or 0)
        if abs(want - back_value) > 1e-12 * max(1.0, abs(want)):
            raise Failure(
                "excel-round-trip",
                f"{original['Employee']!r}: {want} came back as {back_value}",
            )

    # 5. Warnings must be honest: they fire exactly when the file's stated
    #    subtotal disagrees with the line items.
    if case["state_subtotal"]:
        mismatches = [w for w in result.warnings if "the line items add up to" in w]
        if mismatches:
            raise Failure(
                "no-false-alarm", f"subtotals were consistent but got {mismatches[0]!r}"
            )


def _read_report(path: Path) -> list:
    from openpyxl import load_workbook

    book = load_workbook(path, data_only=True, read_only=True)
    try:
        return [tuple(r) for r in book.active.iter_rows(min_row=2, values_only=True)]
    finally:
        book.close()


# --------------------------------------------------------------------------
# Shrinking
# --------------------------------------------------------------------------


def shrink(case: dict, tmp: Path, prop: str) -> dict:
    """Cut the failing case down to something a human can read.

    Only shrinks toward the *same* property — otherwise the minimal case that
    comes out describes a different bug than the one that was found.
    """

    def still_fails(candidate: dict) -> bool:
        try:
            check_case(candidate, tmp, 0)
            return False
        except Failure as failure:
            return failure.prop == prop
        except Exception:
            return prop == "crash"

    current = case
    changed = True
    while changed:
        changed = False
        # Drop employees.
        for i in range(len(current["employees"])):
            if len(current["employees"]) <= 1:
                break
            candidate = dict(current)
            candidate["employees"] = current["employees"][:i] + current["employees"][i + 1:]
            if still_fails(candidate):
                current, changed = candidate, True
                break
        if changed:
            continue
        # Drop entries within an employee.
        for e, employee in enumerate(current["employees"]):
            for i in range(len(employee["entries"])):
                candidate = dict(current)
                trimmed = dict(employee)
                trimmed["entries"] = employee["entries"][:i] + employee["entries"][i + 1:]
                candidate["employees"] = (
                    current["employees"][:e] + [trimmed] + current["employees"][e + 1:]
                )
                if still_fails(candidate):
                    current, changed = candidate, True
                    break
            if changed:
                break
        if changed:
            continue
        # Simplify the surrounding noise.
        for key, simple in (("title_rows", 0), ("trailing_blanks", 0), ("period", None)):
            if current[key] != simple:
                candidate = dict(current)
                candidate[key] = simple
                if still_fails(candidate):
                    current, changed = candidate, True
                    break
    return current


def describe(case: dict) -> str:
    lines = [f"  period={case['period']!r} title_rows={case['title_rows']} "
             f"subtotal_stated={case['state_subtotal']}"]
    for employee in case["employees"]:
        lines.append(f"  employee {employee['name']!r}")
        for project, value in employee["entries"]:
            lines.append(f"      {project!r}  hours={value!r}")
    return "\n".join(lines)


# --------------------------------------------------------------------------


def _report_advisories() -> None:
    if not ADVISORY_COUNTS:
        return
    print("\nAdvisories (correct output, but this input cannot be imported "
          "into that system):")
    for rule, count in sorted(ADVISORY_COUNTS.items(), key=lambda kv: -kv[1]):
        print(f"  {count:5}x  {rule}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("-n", "--cases", type=int, default=300)
    parser.add_argument("--seed", type=int, default=None)
    parser.add_argument("--keep-going", action="store_true",
                        help="report every distinct property that fails, not just the first")
    args = parser.parse_args()

    base_seed = args.seed if args.seed is not None else random.randrange(1 << 30)
    print(f"Fuzzing {args.cases} cases against {len(ALL_CONTRACTS)} contracts "
          f"(seed {base_seed})")

    failures: dict[str, tuple] = {}
    with tempfile.TemporaryDirectory() as raw:
        tmp = Path(raw)
        for i in range(args.cases):
            seed = base_seed + i
            case = make_case(random.Random(seed))
            try:
                check_case(case, tmp, i)
            except Failure as failure:
                if failure.prop in failures:
                    continue
                failures[failure.prop] = (seed, failure, shrink(case, tmp, failure.prop))
                if not args.keep_going:
                    break
            except Exception as exc:                    # a crash is itself a failure
                failures.setdefault("crash", (seed, Failure("crash", repr(exc)), case))
                if not args.keep_going:
                    break

        if not failures:
            print(f"\nAll {args.cases} cases passed every property "
                  f"and every contract rule.")
            _report_advisories()
            return 0

        print(f"\n{len(failures)} propert{'y' if len(failures) == 1 else 'ies'} failed:\n")
        for prop, (seed, failure, minimal) in failures.items():
            print(f"[{prop}]  seed {seed}")
            print(f"  {failure.detail}")
            print("  smallest input that still fails:")
            print(describe(minimal))
            out = Path(f"fuzz-failure-{seed}.xlsx")
            write_case(minimal, out)
            print(f"  written to {out}\n")
        _report_advisories()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
