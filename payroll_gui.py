#!/usr/bin/env python3
"""
Payroll Report — minimal desktop UI.

Pick a timesheet file, pick a folder, press Create Report.

Run with:  python3 payroll_gui.py
"""

from __future__ import annotations

import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, ttk

sys.path.insert(0, str(Path(__file__).resolve().parent))
from contracts import advisories, errors, validate_all  # noqa: E402
from payroll_core import PayrollError, process_timesheet, safe_path, save_report  # noqa: E402

APP_TITLE = "Payroll Report"
REVEAL_LABEL = (
    "Show in Finder" if sys.platform == "darwin"
    else "Show in Explorer" if sys.platform.startswith("win")
    else "Open Folder"
)
SETTINGS = Path.home() / ".payroll_report_prefs"


def load_prefs() -> dict:
    prefs = {}
    try:
        for line in SETTINGS.read_text().splitlines():
            key, _, value = line.partition("=")
            if key and value:
                prefs[key.strip()] = value.strip()
    except OSError:
        pass
    return prefs


def save_prefs(prefs: dict) -> None:
    try:
        SETTINGS.write_text("\n".join(f"{k}={v}" for k, v in prefs.items()))
    except OSError:
        pass


def reveal(path: Path) -> None:
    """Show a file or folder in the OS file browser."""
    # A windowed frozen build has no console; without this flag Windows opens
    # one for a fraction of a second every time.
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", "-R", str(path)], check=False)
        elif sys.platform.startswith("win"):
            # explorer wants "/select,<path>" as a single argument, and it
            # exits non-zero even when it works.
            subprocess.run(f'explorer /select,"{path}"', shell=True,
                           check=False, creationflags=flags)
        else:
            # Try the file manager first, fall back to opening the folder.
            for cmd in (["xdg-open", str(path.parent)], ["gio", "open", str(path.parent)]):
                try:
                    subprocess.run(cmd, check=False)
                    return
                except FileNotFoundError:
                    continue
    except OSError:
        pass


class PayrollApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.prefs = load_prefs()
        self.events: queue.Queue = queue.Queue()
        self.last_output: Path | None = None
        self.running = False

        root.title(APP_TITLE)
        root.minsize(560, 480)
        root.geometry("640x560")

        self.input_var = tk.StringVar(value=self.prefs.get("input", ""))
        self.output_var = tk.StringVar(
            value=self.prefs.get("output", str(Path.home() / "Desktop"))
        )
        self.status_var = tk.StringVar(value="Choose a timesheet file to begin.")

        self._build(root)
        self._refresh_run_state()
        root.after(100, self._drain_events)

    # -- layout ------------------------------------------------------------
    def _build(self, root: tk.Tk) -> None:
        # Derive every font from the system default so the window looks native
        # on each platform instead of guessing point sizes.
        base = tkfont.nametofont("TkDefaultFont")
        bold = base.copy()
        bold.configure(weight="bold")
        small = base.copy()
        small.configure(size=max(base.cget("size") - 1, 9))

        outer = ttk.Frame(root, padding=(18, 16))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)

        # --- the two choices ------------------------------------------------
        ttk.Label(outer, text="Timesheet file", font=bold).grid(
            row=0, column=0, sticky="w", pady=(0, 2)
        )
        ttk.Entry(outer, textvariable=self.input_var).grid(
            row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8)
        )
        ttk.Button(outer, text="Choose…", command=self.pick_input).grid(row=1, column=2)

        ttk.Label(outer, text="Save the report into", font=bold).grid(
            row=2, column=0, sticky="w", pady=(14, 2)
        )
        ttk.Entry(outer, textvariable=self.output_var).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8)
        )
        ttk.Button(outer, text="Choose…", command=self.pick_output).grid(row=3, column=2)

        # --- the one action --------------------------------------------------
        actions = ttk.Frame(outer)
        actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(18, 4))
        actions.columnconfigure(0, weight=1)

        self.run_btn = ttk.Button(
            actions, text="Create Report", command=self.run, default="active"
        )
        self.run_btn.grid(row=0, column=1)
        root.bind("<Return>", lambda _e: self.run())
        root.bind("<KP_Enter>", lambda _e: self.run())
        self.reveal_btn = ttk.Button(
            actions, text=REVEAL_LABEL, command=self.show_output, state="disabled"
        )
        self.reveal_btn.grid(row=0, column=2, padx=(8, 0))

        ttk.Separator(outer).grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 10))

        self.status_label = ttk.Label(
            outer, textvariable=self.status_var, justify="left", font=bold
        )
        self.status_label.grid(row=6, column=0, columnspan=3, sticky="w")
        # Wrap to the real width instead of a hardcoded guess.
        outer.bind(
            "<Configure>",
            lambda e: self.status_label.configure(wraplength=max(e.width - 40, 200)),
        )

        # --- results ----------------------------------------------------------
        log_frame = ttk.Frame(outer)
        log_frame.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        outer.rowconfigure(7, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        try:
            bg = ttk.Style().lookup("TEntry", "fieldbackground") or "#ffffff"
        except tk.TclError:
            bg = "#ffffff"
        try:
            fg = ttk.Style().lookup("TLabel", "foreground") or "#000000"
        except tk.TclError:
            fg = "#000000"

        self.log = tk.Text(
            log_frame, height=10, wrap="word", state="disabled", relief="flat",
            background=bg, foreground=fg, borderwidth=0, highlightthickness=0,
            padx=10, pady=8, font=base,
        )
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, command=self.log.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)
        # Light-theme colours are too dark to read on a dark background, so
        # pick the palette from the pane's actual background.
        if is_dark(bg, root):
            warn, error, ok, dim = "#e8b339", "#ff8b7d", "#5fd08a", "#9aa0ad"
        else:
            warn, error, ok, dim = "#9a6700", "#b42318", "#116329", "#6b6f7a"
        self.log.tag_configure("warn", foreground=warn)
        self.log.tag_configure("error", foreground=error)
        self.log.tag_configure("ok", foreground=ok)
        self.log.tag_configure("dim", foreground=dim)

        # On macOS, a file dropped on the Dock icon arrives through this.
        if sys.platform == "darwin":
            try:
                root.createcommand("::tk::mac::OpenDocument", self._open_document)
            except tk.TclError:
                pass

    # -- helpers -----------------------------------------------------------
    def _open_document(self, *paths) -> None:
        if paths:
            self.input_var.set(paths[0])
            self._refresh_run_state()

    def say(self, text: str, tag: str | None = None) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _refresh_run_state(self) -> None:
        ready = bool(self.input_var.get().strip()) and bool(self.output_var.get().strip())
        self.run_btn.configure(state=("normal" if ready and not self.running else "disabled"))

    # -- actions -----------------------------------------------------------
    def pick_input(self) -> None:
        start = self.input_var.get() or self.prefs.get("input", "")
        initial = str(Path(start).parent) if start else str(Path.home())
        chosen = filedialog.askopenfilename(
            title="Choose the timesheet file",
            initialdir=initial,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if chosen:
            self.input_var.set(chosen)
            self.status_var.set(f"Ready: {Path(chosen).name}")
            self._refresh_run_state()

    def pick_output(self) -> None:
        chosen = filedialog.askdirectory(
            title="Choose where to save the report",
            initialdir=self.output_var.get() or str(Path.home()),
            mustexist=False,
        )
        if chosen:
            self.output_var.set(chosen)
            self._refresh_run_state()

    def show_output(self) -> None:
        if self.last_output and self.last_output.exists():
            reveal(self.last_output)

    def run(self) -> None:
        if self.running or str(self.run_btn["state"]) == "disabled":
            return
        source = safe_path(self.input_var.get().strip())
        target = safe_path(self.output_var.get().strip())

        if not source.is_file():
            messagebox.showerror(APP_TITLE, f"Timesheet file not found:\n{source}")
            return
        if target.exists() and not target.is_dir():
            messagebox.showerror(APP_TITLE, f"Not a folder:\n{target}")
            return

        save_prefs({"input": str(source), "output": str(target)})

        self.running = True
        self.run_btn.configure(state="disabled")
        self.reveal_btn.configure(state="disabled")
        self.clear_log()
        self.status_var.set("Working…")
        self.say(f"Reading {source.name}")

        threading.Thread(target=self._work, args=(source, target), daemon=True).start()
        self._refresh_run_state()

    def _work(self, source: Path, target: Path) -> None:
        """Runs off the UI thread; talks back through self.events."""
        try:
            result = process_timesheet(source)
            # The output field is always a folder, even if the user typed a new name.
            target.mkdir(parents=True, exist_ok=True)
            written = save_report(result, target)
            self.events.put(("done", result, written))
        except PayrollError as exc:
            self.events.put(("error", str(exc), None))
        except Exception:
            self.events.put(("error", "Unexpected problem:\n\n" + traceback.format_exc(), None))

    def _drain_events(self) -> None:
        try:
            while True:
                kind, a, b = self.events.get_nowait()
                if kind == "done":
                    self._on_done(a, b)
                else:
                    self._on_error(a)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_events)

    def _on_done(self, result, written: Path) -> None:
        self.running = False
        self.last_output = written
        counts = result.entry_counts

        self.say(f"Sheet '{result.sheet_name}', header on row {result.header_row}")
        if result.period:
            self.say(f"Pay period: {result.period.replace('_to_', ' to ')}")
        self.say(
            f"{result.employee_count} employees · "
            f"{counts.get('Work Hours', 0)} work entries · "
            f"{counts.get('Sick Leave', 0)} sick · "
            f"{counts.get('Vacation', 0)} vacation"
        )
        self.say(f"Total hours: {result.total_hours:g}")

        checks = validate_all(result)
        for violation in errors(checks):
            self.say(f"BUG — please report this: {violation}", "error")

        if result.warnings:
            self.say("")
            self.say(f"{len(result.warnings)} thing(s) to check:", "warn")
            for warning in result.warnings:
                self.say(f"  • {warning}", "warn")
        else:
            self.say("Every employee's line items match the totals in the source file.", "ok")

        notes = advisories(checks)
        if notes:
            self.say("")
            self.say(f"{len(notes)} compatibility note(s):", "warn")
            for note in notes:
                self.say(f"  • {note.contract}: {note.message}", "warn")

        self.say("")
        self.say(f"Saved: {written}", "ok")
        self.status_var.set(f"Done — {written.name}")
        self.reveal_btn.configure(state="normal")
        self._refresh_run_state()

    def _on_error(self, message: str) -> None:
        self.running = False
        self.say(message, "error")
        self.status_var.set("Could not create the report.")
        messagebox.showerror(APP_TITLE, message)
        self._refresh_run_state()


def pick_theme(root: tk.Tk) -> None:
    """Use each platform's best-looking ttk theme.

    macOS and Windows default to a native theme already. Linux defaults to
    'clam' only sometimes; the alternatives ('alt', 'default', 'classic') look
    like Motif from 1995, so choose explicitly.
    """
    style = ttk.Style(root)
    available = set(style.theme_names())
    for preferred in (
        ["aqua"] if sys.platform == "darwin"
        else ["vista", "winnative"] if sys.platform.startswith("win")
        else ["clam"]
    ):
        if preferred in available:
            try:
                style.theme_use(preferred)
            except tk.TclError:
                continue
            break


def is_dark(colour: str, root: tk.Tk) -> bool:
    """True if a Tk colour is dark enough to need light text on top."""
    try:
        r, g, b = root.winfo_rgb(colour)                # 16-bit per channel
    except tk.TclError:
        return False
    return (0.299 * r + 0.587 * g + 0.114 * b) / 65535 < 0.5


def enable_hidpi() -> None:
    """Stop the window looking blurry on high-density Windows displays."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)      # Windows 8.1+
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()       # Windows 7/Vista
        except Exception:
            pass


def smoke_test() -> int:
    """One end-to-end conversion, so CI can prove the built binary actually runs.

    This is a health check, not the test suite: the tests live in selftest.py
    and fuzz.py and run before the binary is ever built.
    """
    import tempfile

    from openpyxl import Workbook

    with tempfile.TemporaryDirectory() as raw:
        source = Path(raw) / "smoke.xlsx"
        book = Workbook()
        sheet = book.active
        for row in (
            ["Payroll Report", None, None, None],
            ["8/9/26 - 8/22/26", None, None, None],
            [None, None, None, None],
            ["Person", "Description", "Project", "Hours"],
            ["Test Person", None, None, 80],
            [None, "Worked Time", None, 80],
            [None, None, "CD-SICK Sick Leave", 8],
            [None, None, "ACME-1 Project", 72],
        ):
            sheet.append(row)
        book.save(source)

        result = process_timesheet(source)
        written = save_report(result, Path(raw) / "out.xlsx", overwrite=True)
        row = result.rows[0]
        ok = (
            len(result.rows) == 1
            and row["Sick Leave"] == 8
            and row["Work Hours"] == 72
            and row["Personal Total"] == 80
            and written.exists()
            and not errors(validate_all(result))
        )
    print("smoke test: " + ("OK" if ok else "FAILED"))
    return 0 if ok else 1


def main() -> None:
    if "--smoke" in sys.argv[1:]:
        sys.exit(smoke_test())

    enable_hidpi()
    root = tk.Tk()
    pick_theme(root)
    # No manual "tk scaling" call: macOS and Windows both report the right
    # scaling already, and overriding it made every widget ~40% oversized.
    if sys.platform == "darwin":
        try:
            root.call("::tk::unsupported::MacWindowStyle", "style", root, "document")
        except tk.TclError:
            pass
    PayrollApp(root)
    # Make sure the window comes to the front when launched from an .app bundle.
    root.lift()
    root.attributes("-topmost", True)
    root.after(300, lambda: root.attributes("-topmost", False))
    root.mainloop()


if __name__ == "__main__":
    main()
