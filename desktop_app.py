#!/usr/bin/env python3
"""Simple desktop app for Trial Balance generation."""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from convert_trial_balance import RunResult, load_config, run_conversion


class TrialBalanceApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Trial Balance Generator")
        self.root.geometry("860x520")
        self.root.minsize(820, 460)

        if getattr(sys, "frozen", False):
            # Running from packaged executable (PyInstaller).
            self.base_dir = Path(sys.executable).resolve().parent
        else:
            self.base_dir = Path(__file__).resolve().parent
        self.config_path = self.base_dir / "config.yaml"
        self.last_result: RunResult | None = None

        self.input_var = tk.StringVar(value=str(self.base_dir / "input" / "zostava.xlsx"))
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Input file (zostava.xlsx):").grid(
            row=0, column=0, sticky="w"
        )
        entry = ttk.Entry(frame, textvariable=self.input_var, width=90)
        entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="Browse", command=self._browse_input).grid(
            row=1, column=1, sticky="ew"
        )

        button_row = ttk.Frame(frame)
        button_row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 6))
        self.run_button = ttk.Button(button_row, text="RUN", command=self._run)
        self.run_button.pack(side=tk.LEFT)
        self.open_output_button = ttk.Button(
            button_row, text="Open output folder", command=self._open_output_folder
        )
        self.open_output_button.pack(side=tk.LEFT, padx=(8, 0))
        self.open_report_button = ttk.Button(
            button_row, text="Open report", command=self._open_report
        )
        self.open_report_button.pack(side=tk.LEFT, padx=(8, 0))
        self.open_run_report_button = ttk.Button(
            button_row, text="Open run report", command=self._open_run_report
        )
        self.open_run_report_button.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(frame, textvariable=self.status_var).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 4)
        )

        self.log = tk.Text(frame, height=22, wrap=tk.WORD)
        self.log.grid(row=4, column=0, columnspan=2, sticky="nsew")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=4, column=2, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(4, weight=1)

    def _append_log(self, text: str) -> None:
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def _browse_input(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select zostava.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self.input_var.set(selected)

    def _set_busy(self, busy: bool) -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        self.run_button.configure(state=state)
        self.open_output_button.configure(state=state)
        self.open_report_button.configure(state=state)
        self.open_run_report_button.configure(state=state)

    def _resolve_path(self, value: str) -> Path:
        path = Path(value).expanduser()
        if not path.is_absolute():
            path = self.base_dir / path
        return path.resolve()

    def _run(self) -> None:
        input_raw = self.input_var.get().strip()
        if not input_raw:
            messagebox.showerror("Error", "Input file path is required.")
            return

        input_path = self._resolve_path(input_raw)
        if not input_path.exists():
            messagebox.showerror("Error", f"Input file not found:\n{input_path}")
            return
        if input_path.suffix.lower() != ".xlsx":
            messagebox.showerror("Error", "Input file must be an .xlsx file.")
            return

        self.log.delete("1.0", tk.END)
        self._set_busy(True)
        self.status_var.set("Running...")
        self._append_log(f"RUN started for: {input_path}")

        thread = threading.Thread(target=self._run_worker, args=(input_path,), daemon=True)
        thread.start()

    def _run_worker(self, input_path: Path) -> None:
        try:
            cfg = load_config(self.config_path)
            template_path = self._resolve_path(str(cfg.template["path"]))
            result = run_conversion(cfg, input_path=input_path, template_path=template_path)
            self.root.after(0, self._run_success, result)
        except Exception as exc:  # pragma: no cover - UI error path
            details = "".join(traceback.format_exception(exc))
            self.root.after(0, self._run_error, details)

    def _run_success(self, result: RunResult) -> None:
        self.last_result = result
        self._append_log(f"Saved: {result.output_path}")
        if result.report_path:
            self._append_log(f"Report saved: {result.report_path}")
        self._append_log(f"Run report: {result.summary_path}")
        self._append_log(
            f"Period: {result.period.year}/{result.period.month:02d}"
        )
        self._append_log(
            "Selected latest month: "
            f"{result.summary.selected_year}/{result.summary.selected_month:02d} "
            f"(rows {result.summary.rows_before_filter} -> "
            f"{result.summary.rows_after_filter})"
        )
        md_dal_status = (
            "OK"
            if abs(result.summary.template_diff) <= result.summary.tolerance
            else "MISMATCH"
        )
        self._append_log(
            "MD/DAL check (template accounts): "
            f"{md_dal_status} "
            f"(debit={result.summary.template_debit:.2f}, "
            f"credit={result.summary.template_credit:.2f}, "
            f"diff={result.summary.template_diff:.2f})"
        )
        self._append_log(
            f"Missing accounts in template: {len(result.summary.missing_accounts)}"
        )
        if result.summary.missing_accounts:
            self._append_log("Top missing accounts by diff:")
            for item in result.summary.missing_accounts[:10]:
                self._append_log(
                    f"  {item.account}: debit={item.debit:.2f}, "
                    f"credit={item.credit:.2f}, diff={item.diff:.2f}"
                )
        self.status_var.set("Done")
        self._set_busy(False)
        messagebox.showinfo("Success", "Files were generated successfully.")

    def _run_error(self, details: str) -> None:
        self._append_log("ERROR")
        self._append_log(details.rstrip())
        self.status_var.set("Failed")
        self._set_busy(False)
        messagebox.showerror(
            "Run failed",
            "Generation failed. Check the log area for details.",
        )

    def _open_path(self, path: Path) -> None:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
            return
        if os.name == "posix":
            cmd = ["open", str(path)] if "darwin" in os.uname().sysname.lower() else ["xdg-open", str(path)]
            subprocess.Popen(cmd)
            return
        raise RuntimeError(f"Unsupported OS for open action: {os.name}")

    def _open_output_folder(self) -> None:
        target = self.base_dir / "output"
        target.mkdir(parents=True, exist_ok=True)
        try:
            self._open_path(target)
        except Exception as exc:  # pragma: no cover - UI error path
            messagebox.showerror("Error", f"Cannot open folder:\n{exc}")

    def _open_report(self) -> None:
        if not self.last_result or not self.last_result.report_path:
            messagebox.showerror("Error", "No report generated yet. Run first.")
            return
        try:
            self._open_path(self.last_result.report_path)
        except Exception as exc:  # pragma: no cover - UI error path
            messagebox.showerror("Error", f"Cannot open report:\n{exc}")

    def _open_run_report(self) -> None:
        if not self.last_result:
            messagebox.showerror("Error", "No run report generated yet. Run first.")
            return
        try:
            self._open_path(self.last_result.summary_path)
        except Exception as exc:  # pragma: no cover - UI error path
            messagebox.showerror("Error", f"Cannot open run report:\n{exc}")


def main() -> None:
    root = tk.Tk()
    app = TrialBalanceApp(root)
    app._append_log("Ready")
    root.mainloop()


if __name__ == "__main__":
    main()
