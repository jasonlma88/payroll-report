# Payroll Report

Turns the timesheet export (the "Payroll Report (jobs)" sheet) into a one-row-per-employee
payroll report: **Employee · Sick Leave · Vacation · Work Hours · Personal Total**.

---

## For the person running payroll

Download the file for your machine from the
[Releases page](https://github.com/jasonlma88/payroll-report/releases), unzip it, and
double-click it. There is nothing to install — no Python, no pip, no dependencies.

| Your machine | Download |
| --- | --- |
| Mac (Apple Silicon — M1/M2/M3/M4) | `Payroll-Report-macOS-arm64.zip` |
| Windows 10 / 11 | `Payroll-Report-Windows.zip` |
| Linux | `Payroll-Report-Linux.tar.gz` |

Then:

1. **Choose…** the timesheet file
2. **Choose…** the folder to save the report in
3. **Create Report**

It remembers both choices for next time. **Show in Finder / Show in Explorer** opens the
result. **Self-test** confirms the app is working correctly — press it if a report ever
looks wrong.

### The one-time security prompt

Neither build is signed with a paid developer certificate, so each OS asks once:

- **macOS** — the app "cannot be opened because Apple cannot check it for malicious software."
  On **macOS 15 (Sequoia) and later** the old right-click → Open shortcut no longer works.
  Double-click the app once and let it be blocked, then go to
  **System Settings → Privacy & Security**, scroll to the **Security** section, and click
  **Open Anyway** next to the message about Payroll Report. Authenticate, then confirm.
  On **macOS 14 and earlier**, right-click (or Control-click) the app → **Open** → **Open**.
- **Windows** — "Windows protected your PC."
  **More info** → **Run anyway**. Only needed the first time.

Either way it is once per machine, not once per report.

**Avoiding the prompt entirely.** Both warnings are triggered by a "downloaded from the
internet" flag (`com.apple.quarantine` on macOS, Mark-of-the-Web on Windows) that the
*browser* attaches — not by anything in the app. A copy that never came from a browser
never gets flagged. So if you put the app on the company file share and people copy it
from there, neither prompt appears at all. Note that AirDrop and email attachments **do**
set the flag; a mounted network share, a USB drive, or `scp` do not.

**Removing it properly** costs money: an Apple Developer ID ($99/yr) to sign and notarize
the Mac build, and a Windows code-signing certificate for the .exe — and an ordinary (OV)
Windows certificate still shows SmartScreen until it builds download reputation, so only
an EV certificate ($400+/yr, hardware token) silences it immediately. For an internal tool
used by a handful of people, the file-share route above is the better trade.

### Keeping it handy

- **macOS** — drag the app into `/Applications`, then drag it to the Dock.
- **Windows** — right-click the `.exe` → **Pin to Start**, or **Send to → Desktop (create shortcut)**.
- **Linux** — `chmod +x payroll-report`, then drop a `.desktop` file in `~/.local/share/applications/`.

---

## For whoever maintains it

### Run from source

```bash
pip install -r requirements.txt      # just openpyxl
python3 payroll_gui.py               # the window
python3 payroll_processor.py PayrollReportExample.xlsx -o ~/Desktop
python3 selftest.py                  # 26 checks
```

Flags on the CLI: `-o` output folder or file · `-f` overwrite instead of adding "(2)" ·
`-y` skip the confirm prompt when auto-picking the newest file.

### Build the standalone app

```bash
python3 -m venv .buildenv && .buildenv/bin/pip install openpyxl pyinstaller
.buildenv/bin/python build.py        # -> dist/
```

`build.py` runs the self-test first and refuses to build if anything fails.

**PyInstaller cannot cross-compile** — a Windows `.exe` has to be built on Windows.
`.github/workflows/build.yml` builds all four targets on GitHub's runners; push a `v*`
tag and it attaches them to a release:

```bash
git tag v1.0.0 && git push origin v1.0.0
```

That is the practical way to ship Windows and Linux builds from a Mac. Each build takes about 90 seconds. There is deliberately **no Intel Mac build**: the
`macos-13` runner pool is tiny and the job repeatedly never got scheduled at all, holding
every release open. Two commented-out lines in the matrix restore it if anyone needs one.

### Files

| File | What it is |
| --- | --- |
| `payroll_core.py` | The engine: read, categorize, verify, write. No UI, no printing. |
| `payroll_gui.py` | The window. Also the frozen app's entry point. |
| `payroll_processor.py` | The command-line front end. |
| `selftest.py` | 26 checks, runnable from the command line or the Self-test button. |
| `build.py` | Builds the standalone app for the current platform. |
| `make_example.py` | Builds the synthetic example workbooks. |
| `contracts.py` | Output contracts — what downstream systems require. |
| `fuzz.py` | Property-based fuzzing against those contracts. |
| `example_timesheet.xlsx` | Example **input** — fully synthetic. |
| `example_report.xlsx` | Example **output** — what the tool produces from it. |

**No real payroll data is in this repo.** The example workbooks are generated by
`make_example.py`; every person and client in them is invented, and the standard
fictional company names (Contoso, Northwind, Globex, Initech, Fabrikam, Aperture,
Umbra, Wonka) are used deliberately so they cannot be mistaken for a real client.
Real exports are gitignored by name.

The synthetic data is not arbitrary — it reproduces every structural feature that
makes a real export tricky: quarter-hours, an employee with no hours at all,
multi-project employees, sick/vacation/holiday codes, slashes and trailing spaces
in project codes, and names with parentheses, apostrophes, accents and non-Latin
scripts.

### How hours are bucketed

`CATEGORY_RULES` at the top of `payroll_core.py`:

- **Sick Leave** — project code contains `CD-SICK` or `sick leave`
- **Vacation** — contains `CD-VAC` or `paid vacation`
- **Work Hours** — everything else, including Company Holiday and all client projects

Add a charge code or a whole new bucket by editing that list. Nothing else changes.

### Output contracts

"The output must be compatible with QuickBooks" is a *specification*, and diffing
against one known-good file cannot test it. So the contracts live in `contracts.py`
as data:

```bash
python3 contracts.py                       # list contracts and their rules
python3 contracts.py --export quickbooks_iif timesheet.xlsx payroll.iif
```

| Contract | What it requires |
| --- | --- |
| `core` | Invariants any consumer depends on: buckets partition the total, no duplicate or invented employees, finite numbers, non-empty names. |
| `quickbooks-iif` | QuickBooks Desktop timer import. Tab-delimited, so no tabs or newlines in any field; names ≤ 41 characters; durations as `HH:MM`, so `61.25 h` must become `61:15`. |
| `generic-csv` | Round-trips, and no leading `=`/`+`/`-`/`@` that a spreadsheet would execute. |

Every rule is one of two things, and conflating them makes the exercise useless:

- **error** — we produced something wrong. Always our bug. Must never happen.
- **advisory** — the report is correct, but *this input* cannot be imported into
  that system. The user needs to hear it; it must not fail a build.

Both the CLI and the window run these on every real report, so a name too long for
QuickBooks or an implausible 900-hour total is reported when it happens, not months later.

Adding a downstream system means adding a `Contract` — not editing the engine.

### Property-based fuzzing

`selftest.py` checks known cases. `fuzz.py` checks *properties* — statements that must
hold for every possible timesheet — against thousands of randomly generated hostile ones:
unicode and RTL names, embedded tabs and newlines, formula prefixes, 60-character names,
names colliding with the format's own keywords (`Person`, `Worked Time`), negative hours,
`1e-7` and `1e12` hours, hours as text, blank and missing values.

```bash
python3 fuzz.py                  # 300 cases
python3 fuzz.py -n 5000          # longer
python3 fuzz.py --seed 12345     # reproduce a specific failure
```

The properties: hours are conserved against an independently-computed oracle, no employee
is dropped or invented, every contract holds, the same input twice gives the same answer,
the Excel round-trip preserves what we computed, and warnings fire exactly when the
source file's own subtotals disagree.

A failure is **shrunk** to the smallest input that still fails *that same property*, then
written to `fuzz-failure-<seed>.xlsx` so you can open it.

`build.py` and CI both run the fuzzer and refuse to ship if it fails.

### What it checks for you

- Finds the header row and the pay period itself, so an extra title row won't break it.
- Matches columns by name, so a reordered export still works.
- **Cross-checks every employee's line items against the subtotal the source file states
  for them**, and reports any that disagree. This is the check that catches a silent
  misparse, which is the failure mode that actually costs money.
- Refuses Excel `~$` lock files, `.xls`, non-Excel files, and finished reports fed back in
  as input — each with a specific message saying what to do instead.
- Never overwrites: an existing report becomes `… (2).xlsx`.
- Says "it is probably open in Excel" instead of raising `PermissionError`.

### Dependencies

`openpyxl` only. It deliberately does **not** use pandas: pandas and numpy added ~300 MB
to the frozen application and nothing this tool needs. The build is 45 MB.
