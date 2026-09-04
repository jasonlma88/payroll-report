#!/usr/bin/env python3
"""
Build a standalone Payroll Report application for the machine you run this on.

    python3 build.py

The result lands in dist/ and needs no Python, no pip and no dependencies on
the machine it is copied to:

    macOS    dist/Payroll Report.app         (drag to Applications)
    Windows  dist/Payroll Report.exe         (single file, double-click)
    Linux    dist/payroll-report             (single file, chmod +x)

PyInstaller cannot cross-compile: each platform's build must run on that
platform. .github/workflows/build.yml does all three on CI if you would rather
not find a Windows machine.
"""

from __future__ import annotations

import platform
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ENTRY = HERE / "payroll_gui.py"
SYSTEM = platform.system()

# Everything pandas used to drag in, plus the bits of the stdlib that no GUI
# needs. Excluding them roughly halves the download.
EXCLUDES = [
    "pandas", "numpy", "matplotlib", "scipy", "PIL", "IPython", "jedi",
    "pytest", "setuptools", "pip", "lib2to3", "pydoc_data", "sqlite3",
    "test", "unittest", "distutils", "email", "http", "xmlrpc",
]


def run(cmd: list) -> None:
    print("+", " ".join(str(c) for c in cmd), flush=True)
    subprocess.run(cmd, check=True, cwd=HERE)


def main() -> int:
    if not ENTRY.exists():
        sys.exit(f"Cannot find {ENTRY}")

    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        sys.exit(
            "PyInstaller is not installed for this interpreter.\n"
            f"Run:  {sys.executable} -m pip install pyinstaller openpyxl"
        )

    try:
        import tkinter  # noqa: F401
    except ImportError:
        sys.exit(
            "This Python has no tkinter, so it cannot build the GUI.\n"
            "macOS/Windows: use the python.org installer.\n"
            "Debian/Ubuntu: sudo apt install python3-tk\n"
            "Fedora: sudo dnf install python3-tkinter"
        )

    print("Running the self-test before building...", flush=True)
    if subprocess.run([sys.executable, str(HERE / "selftest.py")], cwd=HERE).returncode != 0:
        sys.exit("Self-test failed — not building. Fix the failures above first.")

    print("\nFuzzing against the output contracts...", flush=True)
    fuzz = subprocess.run(
        [sys.executable, str(HERE / "fuzz.py"), "-n", "400"], cwd=HERE
    )
    if fuzz.returncode != 0:
        sys.exit("Fuzzing failed — not building. Fix the failures above first.")

    for stale in ("build", "dist"):
        shutil.rmtree(HERE / stale, ignore_errors=True)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--windowed",
        "--name", "Payroll Report" if SYSTEM in ("Darwin", "Windows") else "payroll-report",
        # payroll_core is imported at runtime from the frozen bundle, so name it.
        "--hidden-import", "payroll_core",
        "--hidden-import", "selftest",
        "--hidden-import", "contracts",
        "--paths", str(HERE),
    ]
    for module in EXCLUDES:
        cmd += ["--exclude-module", module]

    if SYSTEM == "Darwin":
        # onedir: launches faster than onefile and Gatekeeper is happier with it.
        cmd += ["--osx-bundle-identifier", "io.github.jasonlma88.payrollreport"]
    else:
        # A single file is the only thing worth emailing to someone.
        cmd += ["--onefile"]

    icon = next((HERE / n for n in ("icon.icns", "icon.ico", "icon.png") if (HERE / n).exists()), None)
    if icon:
        cmd += ["--icon", str(icon)]

    cmd.append(str(ENTRY))
    run(cmd)

    dist = HERE / "dist"
    if SYSTEM == "Darwin":
        app = dist / "Payroll Report.app"
        # Ad-hoc sign so macOS 15+ will launch it without "damaged" complaints.
        # This is not notarization: on another Mac it still needs the
        # right-click > Open gesture once. See README.
        subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(app)], check=False)
        subprocess.run(["xattr", "-cr", str(app)], check=False)
        built = app
    elif SYSTEM == "Windows":
        built = dist / "Payroll Report.exe"
    else:
        built = dist / "payroll-report"
        if built.exists():
            built.chmod(0o755)

    if not built.exists():
        sys.exit(f"Build finished but {built} is missing.")

    size = sum(f.stat().st_size for f in built.rglob("*")) if built.is_dir() else built.stat().st_size
    print(f"\nBuilt: {built}")
    print(f"Size:  {size / 1e6:.0f} MB")
    print("\nThis runs on a machine with no Python installed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
