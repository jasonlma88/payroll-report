#!/usr/bin/env python3
"""
Payroll Report Generator (command line).

Turns a timesheet export into a per-employee payroll report.
The processing itself lives in payroll_core.py; payroll_gui.py is the
point-and-click version of this same thing.

Requirements:
    pip install openpyxl

Usage:
    python3 payroll_processor.py <timesheet.xlsx> [-o OUTPUT_DIR_OR_FILE]
    python3 payroll_processor.py                  # newest Excel file here, with a prompt
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from contracts import advisories, errors, validate_all  # noqa: E402
from payroll_core import PayrollError, process_timesheet, safe_path, save_report  # noqa: E402


def find_latest_excel_file(directory: Path) -> Path | None:
    """Newest .xlsx/.xls in `directory`, ignoring Excel's ~$ lock files."""
    candidates = [
        p
        for pattern in ("*.xlsx", "*.xls")
        for p in directory.glob(pattern)
        if not p.name.startswith("~$")
    ]
    return max(candidates, key=lambda p: p.stat().st_mtime) if candidates else None


def print_summary(result) -> None:
    counts = result.entry_counts

    print(f"\n{'=' * 60}")
    print("PAYROLL REPORT SUMMARY")
    print(f"{'=' * 60}")
    if result.period:
        print(f"Pay period:      {result.period.replace('_to_', ' to ')}")
    print(f"Total employees: {result.employee_count}")
    print(f"Total hours:     {result.total_hours:g}")
    print(
        f"Entries:         {counts.get('Work Hours', 0)} work, "
        f"{counts.get('Sick Leave', 0)} sick, {counts.get('Vacation', 0)} vacation"
    )

    for column, label in (("Vacation", "Vacation time"), ("Sick Leave", "Sick leave")):
        rows = result.with_bucket(column)
        if rows:
            print(f"\n{label}:")
            for row in rows:
                print(f"  {row['Employee']}: {row[column]:g} hours")

    checks = validate_all(result)
    for violation in errors(checks):
        print(f"\nBUG: {violation}")
        print("Please report this — the report may be wrong.")
    notes = advisories(checks)
    if notes:
        print(f"\n{len(notes)} compatibility note(s):")
        for note in notes:
            print(f"  - {note.contract}: {note.message}")

    if result.warnings:
        print(f"\n{len(result.warnings)} thing(s) to check:")
        for warning in result.warnings:
            print(f"  ! {warning}")
    else:
        print("\nEvery employee's line items match the totals in the source file.")
    print(f"{'=' * 60}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a payroll report from a timesheet export.")
    parser.add_argument("input", nargs="?", help="Timesheet .xlsx (default: newest Excel file here)")
    parser.add_argument("-o", "--output", default=".", help="Output folder or file path (default: .)")
    parser.add_argument("-f", "--force", action="store_true", help="Overwrite an existing report")
    parser.add_argument("-y", "--yes", action="store_true", help="Skip the auto-detect confirmation")
    args = parser.parse_args()

    if args.input:
        source = safe_path(args.input)
        if not source.exists():
            sys.exit(f"Error: file not found: {source}")
    else:
        here = Path.cwd()
        print(f"Searching for Excel files in {here}...")
        source = find_latest_excel_file(here)
        if source is None:
            sys.exit("Error: no .xls/.xlsx files found in the current directory.")
        modified = datetime.fromtimestamp(source.stat().st_mtime)
        print(f"Found: {source.name}  (modified {modified:%Y-%m-%d %H:%M:%S})")
        if not args.yes and input("Use this file? (y/n): ").strip().lower() not in ("y", "yes"):
            sys.exit("Cancelled.")

    try:
        print(f"Reading {source}...")
        result = process_timesheet(source)
        written = save_report(result, args.output, overwrite=args.force)
    except PayrollError as exc:
        sys.exit(f"Error: {exc}")

    print_summary(result)
    print(f"\nSaved: {written}")


if __name__ == "__main__":
    main()
