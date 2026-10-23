"""
Desktop GUI for the DIFOT report pipeline, for a non-technical user who
won't open a terminal. Wraps pipeline.run_pipeline() -- the exact same
logic the CLI uses -- behind a simple point-and-click window:

  1. Browse for the raw SAP export file.
  2. Pick a vendor from a dropdown (read from that file).
  3. Click Generate.

The actual report generation runs on a background thread so the window
never freezes, and progress/errors are relayed back to the main thread
through a queue (Tkinter widgets aren't thread-safe to touch directly
from a worker thread).
"""

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .io import find_source_file, load_source, set_source_file
from .pipeline import PipelineError, run_pipeline
from .transform import list_vendors

WINDOW_TITLE = "DIFOT Report Generator"


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(WINDOW_TITLE)
        self.geometry("640x480")
        self.minsize(560, 420)

        self._queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._source_df = None  # loaded lazily once a source file is picked
        self._running = False

        self._build_widgets()
        self._try_load_existing_source()
        self.after(100, self._poll_queue)

    # -----------------------------------------------------------
    # LAYOUT
    # -----------------------------------------------------------
    def _build_widgets(self) -> None:
        pad = {"padx": 10, "pady": 6}

        source_frame = ttk.LabelFrame(self, text="1. Source file (raw SAP export)")
        source_frame.pack(fill="x", **pad)

        self.source_label = ttk.Label(source_frame, text="No file selected", foreground="#888")
        self.source_label.pack(side="left", padx=10, pady=10, fill="x", expand=True)

        ttk.Button(source_frame, text="Browse...", command=self._on_browse).pack(
            side="right", padx=10, pady=10
        )

        vendor_frame = ttk.LabelFrame(self, text="2. Vendor")
        vendor_frame.pack(fill="x", **pad)

        self.vendor_var = tk.StringVar()
        self.vendor_combo = ttk.Combobox(
            vendor_frame, textvariable=self.vendor_var, state="disabled", values=[]
        )
        self.vendor_combo.pack(fill="x", padx=10, pady=10)

        self.generate_button = ttk.Button(
            self, text="3. Generate Report", command=self._on_generate, state="disabled"
        )
        self.generate_button.pack(pady=10)

        self.progress = ttk.Progressbar(self, mode="indeterminate")

        log_frame = ttk.LabelFrame(self, text="Status")
        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = tk.Text(log_frame, height=10, state="disabled", wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

        self.open_folder_button = ttk.Button(
            self, text="Open output folder", command=self._on_open_folder, state="disabled"
        )
        self.open_folder_button.pack(pady=(0, 10))
        self._last_output_path: Path | None = None

    # -----------------------------------------------------------
    # STARTUP: if a source file is already sitting in data/source
    # (e.g. from a previous run), load its vendors automatically
    # -----------------------------------------------------------
    def _try_load_existing_source(self) -> None:
        try:
            existing = find_source_file()
        except (FileNotFoundError, ValueError):
            return
        self._load_source_in_background(existing)

    # -----------------------------------------------------------
    # SOURCE FILE
    # -----------------------------------------------------------
    def _on_browse(self) -> None:
        chosen = filedialog.askopenfilename(
            title="Select the raw SAP export file",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not chosen:
            return
        self._load_source_in_background(Path(chosen))

    def _load_source_in_background(self, path: Path) -> None:
        self.source_label.config(text=f"Loading {path.name}...", foreground="#888")
        self.vendor_combo.config(state="disabled", values=[])
        self.generate_button.config(state="disabled")
        self._log(f"Loading source file: {path.name}")

        def work() -> None:
            try:
                # Only copy into data/source if it isn't already there
                # (the auto-detected startup case) -- set_source_file()
                # would otherwise just re-copy the file onto itself.
                try:
                    current = find_source_file()
                except (FileNotFoundError, ValueError):
                    current = None
                if current != path:
                    set_source_file(path)
                df = load_source()
                vendors = list_vendors(df)
            except Exception as e:
                self._queue.put(("source_error", str(e)))
                return
            self._queue.put(("source_loaded", (path.name, df, vendors)))

        threading.Thread(target=work, daemon=True).start()

    # -----------------------------------------------------------
    # GENERATE
    # -----------------------------------------------------------
    def _on_generate(self) -> None:
        vendor = self.vendor_var.get().strip()
        if not vendor:
            messagebox.showwarning(WINDOW_TITLE, "Please select a vendor first.")
            return

        self._running = True
        self.generate_button.config(state="disabled")
        self.vendor_combo.config(state="disabled")
        self.open_folder_button.config(state="disabled")
        self.progress.pack(fill="x", padx=10, pady=(0, 6))
        self.progress.start(12)
        self._log(f"\nGenerating report for '{vendor}'...")

        def work() -> None:
            try:
                result = run_pipeline(vendor, progress=lambda msg: self._queue.put(("progress", msg)))
            except PipelineError as e:
                self._queue.put(("run_error", str(e)))
                return
            except Exception as e:  # unexpected bug -- still show it, don't crash silently
                self._queue.put(("run_error", f"Unexpected error: {e}"))
                return
            self._queue.put(("run_done", result))

        threading.Thread(target=work, daemon=True).start()

    # -----------------------------------------------------------
    # QUEUE POLLING -- the only place worker-thread results are
    # applied to widgets, since Tkinter isn't thread-safe
    # -----------------------------------------------------------
    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()

                if kind == "source_loaded":
                    name, df, vendors = payload
                    self._source_df = df
                    self.source_label.config(text=f"{name}  ({len(df)} rows)", foreground="black")
                    self.vendor_combo.config(state="readonly", values=vendors)
                    self.generate_button.config(state="normal" if vendors else "disabled")
                    self._log(f"Loaded {len(df)} rows, {len(vendors)} vendor(s) found.")

                elif kind == "source_error":
                    self.source_label.config(text="Failed to load file", foreground="#c0392b")
                    self._log(f"Error loading source file: {payload}")
                    messagebox.showerror(WINDOW_TITLE, str(payload))

                elif kind == "progress":
                    self._log(str(payload))

                elif kind == "run_done":
                    self._finish_run()
                    result = payload
                    self._last_output_path = result.output_path
                    self.open_folder_button.config(state="normal")
                    self._log(f"\nDone. Saved: {result.output_path.name}")
                    self._log(f"Semana {result.semana} | {result.filtered_rows} rows for {result.vendor_name}")
                    messagebox.showinfo(
                        WINDOW_TITLE,
                        f"Report generated:\n\n{result.output_path.name}",
                    )

                elif kind == "run_error":
                    self._finish_run()
                    self._log(f"\nError: {payload}")
                    messagebox.showerror(WINDOW_TITLE, str(payload))

        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _finish_run(self) -> None:
        self._running = False
        self.progress.stop()
        self.progress.pack_forget()
        self.generate_button.config(state="normal")
        self.vendor_combo.config(state="readonly")

    # -----------------------------------------------------------
    # HELPERS
    # -----------------------------------------------------------
    def _log(self, message: str) -> None:
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _on_open_folder(self) -> None:
        if self._last_output_path and self._last_output_path.exists():
            os.startfile(self._last_output_path.parent)  # Windows-only, matches this app's target


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
