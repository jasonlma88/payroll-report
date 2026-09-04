#!/usr/bin/env python3
"""
Self-test for the payroll engine.

Runs on any platform with no test framework installed:

    python3 selftest.py

The frozen application runs this too, via  "Payroll Report" --selftest,
so a non-technical user can confirm their download works before trusting it.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from payroll_core import (  # noqa: E402
    PayrollError,
    force_text,
    process_timesheet,
    save_report,
)

HERE = Path(__file__).resolve().parent
SAMPLE_IN = HERE / "example_timesheet.xlsx"
SAMPLE_OUT = HERE / "example_report.xlsx"

_results: list[tuple[bool, str, str]] = []
_skipped: list[str] = []


# Where output goes. The frozen Windows build has no console, so the GUI
# swaps this for a function that appends to a window.
emit = print


def set_output(fn) -> None:
    global emit
    emit = fn


def check(name: str, condition: bool, detail: str = "") -> None:
    _results.append((bool(condition), name, detail))


def build_sheet(rows: list, path: Path, title: str = "Payroll Report (jobs)") -> Path:
    book = Workbook()
    sheet = book.active
    sheet.title = title
    for row in rows:
        sheet.append(row)
        for column in range(1, len(row) + 1):
            force_text(sheet.cell(row=sheet.max_row, column=column))
    book.save(path)
    return path


def cells(path: Path) -> list:
    sheet = load_workbook(path, data_only=True).active
    return [tuple(r) for r in sheet.iter_rows(values_only=True)]


HEADER = ["Person", "Description", "Project", "Hours"]


def basic_rows() -> list:
    return [
        ["Payroll Report", None, None, None],
        ["8/9/26 - 8/22/26", None, None, None],
        [None, None, None, None],
        HEADER,
        ["Ada Lovelace", None, None, 80],
        [None, "Worked Time", None, 80],
        [None, None, "CD-SICK Sick Leave", 8],
        [None, None, "CD-VAC Paid Vacation", 8],
        [None, None, "ACME-1 Project One", 64],
        [None, None, None, None],
        ["Grace Hopper", None, None, 40],
        [None, "Worked Time", None, 40],
        [None, None, "CD-HOL Company Holiday", 8],
        [None, None, "ACME-2 Project Two", 32],
    ]


def run(tmp: Path) -> None:
    # --- the real sample, if it shipped alongside -------------------------
    if SAMPLE_IN.exists() and SAMPLE_OUT.exists():
        result = process_timesheet(SAMPLE_IN)
        written = save_report(result, tmp / "sample.xlsx", overwrite=True)
        check("sample reproduces the reference report byte for byte",
              cells(written) == cells(SAMPLE_OUT))
        check("sample finds every employee", result.employee_count == 14,
              f"got {result.employee_count}")
        check("sample totals 899.25 hours", abs(result.total_hours - 899.25) < 0.01,
              f"got {result.total_hours}")
        check("sample raises no warnings", not result.warnings, str(result.warnings))
    else:
        # The frozen application ships without the example workbooks; the
        # synthetic cases below cover the same ground.
        _skipped.append("example-workbook checks (no sample files alongside)")

    # --- bucketing --------------------------------------------------------
    result = process_timesheet(build_sheet(basic_rows(), tmp / "basic.xlsx"))
    ada, grace = result.rows
    check("sick leave bucketed", ada["Sick Leave"] == 8, str(ada))
    check("vacation bucketed", ada["Vacation"] == 8, str(ada))
    check("work hours bucketed", ada["Work Hours"] == 64, str(ada))
    check("personal total sums the buckets", ada["Personal Total"] == 80, str(ada))
    check("company holiday counts as work", grace["Work Hours"] == 40, str(grace))
    check("zero buckets render blank", grace["Sick Leave"] == "", str(grace))
    check("pay period parsed", result.period == "2026-08-09_to_2026-08-22", str(result.period))
    check("suggested filename uses the period",
          result.suggested_filename == "Payroll_2026-08-09_to_2026-08-22.xlsx")

    # --- a project named 'sick' must not become sick leave ----------------
    rows = basic_rows()
    rows[6] = [None, None, "SICKLE-CELL-1 Sickle Cell Study", 8]
    result = process_timesheet(build_sheet(rows, tmp / "sickle.xlsx"))
    check("a client project containing 'sick' stays work hours",
          result.rows[0]["Work Hours"] == 72 and result.rows[0]["Sick Leave"] == "",
          str(result.rows[0]))

    # --- header not on row 4, and columns reordered -----------------------
    rows = [["stray note", None, None, None]] + basic_rows()
    result = process_timesheet(build_sheet(rows, tmp / "shifted.xlsx"))
    check("header found when it is not on row 4", result.header_row == 5,
          f"got {result.header_row}")

    reordered = [
        [r[3], r[0], r[2], r[1]] if r != HEADER else ["Hours", "Person", "Project", "Description"]
        for r in basic_rows()
    ]
    result = process_timesheet(build_sheet(reordered, tmp / "reordered.xlsx"))
    check("columns matched by name, not position",
          result.rows[0]["Personal Total"] == 80, str(result.rows[0]))

    # --- the cross-check must actually fire -------------------------------
    rows = basic_rows()
    rows[4] = ["Ada Lovelace", None, None, 999]        # subtotal disagrees
    result = process_timesheet(build_sheet(rows, tmp / "mismatch.xlsx"))
    check("mismatched subtotal is reported",
          any("Ada Lovelace" in w and "999" in w for w in result.warnings),
          str(result.warnings))

    rows = basic_rows()
    rows[6] = [None, None, "CD-SICK Sick Leave", "not a number"]
    result = process_timesheet(build_sheet(rows, tmp / "badhours.xlsx"))
    check("unreadable hours are reported, not silently dropped",
          any("no readable Hours" in w for w in result.warnings), str(result.warnings))

    # --- rejections -------------------------------------------------------
    for name, path, needle in [
        ("a finished report is rejected as input",
         build_sheet([list(r) for r in cells(SAMPLE_OUT)], tmp / "finished.xlsx")
         if SAMPLE_OUT.exists() else None, "Could not find a header row"),
        ("a missing file is reported", tmp / "nope.xlsx", "File not found"),
        ("an .xls file is reported", tmp / "old.xls", "old .xls format"),
        ("a lock file is rejected", tmp / "~$lock.xlsx", "lock file"),
    ]:
        if path is None:
            continue
        try:
            process_timesheet(path)
            check(name, False, "no error raised")
        except PayrollError as exc:
            check(name, needle in str(exc), str(exc).replace("\n", " ")[:90])

    # --- output contracts -------------------------------------------------
    from contracts import ALL_CONTRACTS, advisories, errors, export_quickbooks_iif, \
        quickbooks_duration, validate_all

    result = process_timesheet(build_sheet(basic_rows(), tmp / "contracts.xlsx"))
    found = validate_all(result)
    check("a clean report violates no contract rule", not errors(found),
          "; ".join(str(v) for v in errors(found)))
    check("a clean report raises no advisories either", not advisories(found),
          "; ".join(str(v) for v in advisories(found)))

    check("quarter hours convert to QuickBooks HH:MM",
          quickbooks_duration(61.25) == "61:15", quickbooks_duration(61.25))
    check("whole hours convert to QuickBooks HH:MM",
          quickbooks_duration(8) == "8:00", quickbooks_duration(8))
    iif = export_quickbooks_iif(result)
    widths = {len(line.split("\t")) for line in iif.splitlines() if line.startswith("TIMEACT\t")}
    check("every IIF record has the declared field count", widths == {8}, str(widths))

    # A name that would break the tab-delimited format must be neutralised.
    rows = basic_rows()
    rows[4] = ["Bad\tName\nHere", None, None, 80]
    hostile = process_timesheet(build_sheet(rows, tmp / "hostile.xlsx"))
    check("a name containing tabs cannot break the IIF export",
          not errors(validate_all(hostile, ALL_CONTRACTS)),
          "; ".join(str(v) for v in errors(validate_all(hostile, ALL_CONTRACTS))))

    # A name Excel would execute as a formula must survive as text.
    rows = basic_rows()
    rows[4] = ["=SUM(A1:A9)", None, None, 80]
    formula = process_timesheet(build_sheet(rows, tmp / "formula.xlsx"))
    back = cells(save_report(formula, tmp / "formula-out.xlsx", overwrite=True))
    check("a name starting with '=' survives as text, not a formula",
          back[1][0] == "=SUM(A1:A9)", repr(back[1][0]))

    # --- writing ----------------------------------------------------------
    result = process_timesheet(build_sheet(basic_rows(), tmp / "basic2.xlsx"))
    out = tmp / "outdir"
    first = save_report(result, out)
    second = save_report(result, out)
    check("output folder is created on demand", first.parent == out)
    check("an existing report is never overwritten",
          first != second and "(2)" in second.name, f"{first.name} / {second.name}")
    forced = save_report(result, first, overwrite=True)
    check("--force overwrites in place", forced == first)

    grid = cells(first)
    check("output header is correct",
          grid[0] == ("Employee", "Sick Leave", "Vacation", "Work Hours", "Personal Total"),
          str(grid[0]))
    check("blank buckets are empty cells, not empty strings",
          grid[2][1] is None, repr(grid[2][1]))


def main() -> int:
    _results.clear()
    _skipped.clear()
    with tempfile.TemporaryDirectory() as raw:
        try:
            run(Path(raw))
        except Exception:
            import traceback
            emit(traceback.format_exc())
            check("self-test completed without crashing", False)

    passed = sum(1 for ok, _, _ in _results if ok)
    for ok, name, detail in _results:
        mark = "PASS" if ok else "FAIL"
        emit(f"  [{mark}] {name}" + (f"  -- {detail}" if detail and not ok else ""))
    for name in _skipped:
        emit(f"  [SKIP] {name}")

    emit(f"\n{passed}/{len(_results)} checks passed")
    return 0 if passed == len(_results) else 1


if __name__ == "__main__":
    sys.exit(main())
