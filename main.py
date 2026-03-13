"""
MIS Report Automation Tool
──────────────────────────
Select any sales CSV → click Generate → get a formatted Excel MIS report.
"""

import os
import threading
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from analyzer import (kpi_summary, monthly_revenue, region_performance,
                      top_customers, top_products)
from data_cleaner import cleaning_summary, load_and_clean
from report_generator import generate_report, WHITE

# ── Colours & fonts ────────────────────────────────────────────────────────
BG          = "#1F3864"
CARD_BG     = "#FFFFFF"
ACCENT      = "#2ECC71"
TEXT_DARK   = "#1F3864"
TEXT_LIGHT  = "#FFFFFF"
BTN_HOVER   = "#27AE60"
WHITE       = "#FFFFFF"
FONT_TITLE  = ("Calibri", 18, "bold")
FONT_LABEL  = ("Calibri", 11)
FONT_SMALL  = ("Calibri", 9)
FONT_BTN    = ("Calibri", 11, "bold")


class MISApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MIS Report Automation Tool")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._center_window(600, 480)

        self.csv_path    = tk.StringVar(value="No file selected")
        self.output_dir  = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        self.status_text = tk.StringVar(value="Ready.")
        self.progress    = tk.DoubleVar(value=0)

        self._build_ui()

    # ── Window helpers ─────────────────────────────────────────────────────
    def _center_window(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── UI construction ────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=BG, pady=20)
        header.pack(fill="x")
        tk.Label(header, text="📊  MIS Report Automation",
                 font=FONT_TITLE, bg=BG, fg=TEXT_LIGHT).pack()
        tk.Label(header, text="Upload a sales CSV and generate a formatted Excel report instantly.",
                 font=FONT_SMALL, bg=BG, fg="#A9C4E4").pack(pady=(4, 0))

        # ── Card ──────────────────────────────────────────────────────────
        card = tk.Frame(self, bg=CARD_BG, bd=0, padx=30, pady=24)
        card.pack(fill="both", expand=True, padx=24, pady=(0, 20))

        # CSV input
        self._section_label(card, "Step 1 — Select Input CSV")
        csv_row = tk.Frame(card, bg=CARD_BG)
        csv_row.pack(fill="x", pady=(4, 12))
        tk.Entry(csv_row, textvariable=self.csv_path, font=FONT_SMALL,
                 fg="#555", state="readonly", width=48,
                 relief="solid", bd=1).pack(side="left", fill="x", expand=True)
        self._btn(csv_row, "Browse", self._browse_csv, small=True).pack(
            side="left", padx=(8, 0))

        # Output folder
        self._section_label(card, "Step 2 — Select Output Folder")
        out_row = tk.Frame(card, bg=CARD_BG)
        out_row.pack(fill="x", pady=(4, 12))
        tk.Entry(out_row, textvariable=self.output_dir, font=FONT_SMALL,
                 fg="#555", state="readonly", width=48,
                 relief="solid", bd=1).pack(side="left", fill="x", expand=True)
        self._btn(out_row, "Browse", self._browse_output, small=True).pack(
            side="left", padx=(8, 0))

        # Generate button
        self._section_label(card, "Step 3 — Generate Report")
        self.gen_btn = self._btn(card, "⚡  Generate MIS Report",
                                 self._start_generation)
        self.gen_btn.pack(fill="x", pady=(6, 16))

        # Progress
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("green.Horizontal.TProgressbar",
                        troughcolor="#E0E0E0", background=ACCENT,
                        thickness=12)
        self.pbar = ttk.Progressbar(card, variable=self.progress, maximum=100,
                                    style="green.Horizontal.TProgressbar")
        self.pbar.pack(fill="x")

        # Status
        tk.Label(card, textvariable=self.status_text,
                 font=FONT_SMALL, bg=CARD_BG, fg="#555",
                 anchor="w").pack(fill="x", pady=(6, 0))

    def _section_label(self, parent, text):
        tk.Label(parent, text=text, font=("Calibri", 10, "bold"),
                 bg=CARD_BG, fg=TEXT_DARK, anchor="w").pack(fill="x")

    def _btn(self, parent, text, command, small=False):
        bg = "#2F5496" if not small else "#E8EEF7"
        fg = WHITE if not small else TEXT_DARK
        font = FONT_BTN if not small else FONT_SMALL
        btn = tk.Button(parent, text=text, command=command,
                        bg=bg, fg=fg, font=font,
                        relief="flat", cursor="hand2",
                        padx=10, pady=6 if not small else 4,
                        activebackground=BTN_HOVER, activeforeground=WHITE)
        return btn

    # ── File dialogs ───────────────────────────────────────────────────────
    def _browse_csv(self):
        path = filedialog.askopenfilename(
            title="Select Sales CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.csv_path.set(path)
            self._set_status("CSV selected. Ready to generate.")

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_dir.set(path)

    # ── Generation pipeline ────────────────────────────────────────────────
    def _start_generation(self):
        if self.csv_path.get() == "No file selected":
            messagebox.showwarning("No File", "Please select a CSV file first.")
            return
        if not os.path.exists(self.output_dir.get()):
            messagebox.showwarning("Invalid Folder", "Output folder does not exist.")
            return

        self.gen_btn.configure(state="disabled")
        self.progress.set(0)
        # Run in background thread so UI stays responsive
        thread = threading.Thread(target=self._run_pipeline, daemon=True)
        thread.start()

    def _run_pipeline(self):
        try:
            self._set_status("Loading CSV...", 5)

            raw_df = pd.read_csv(self.csv_path.get(),encoding="latin-1",on_bad_lines='skip')
            self._set_status("Cleaning data...", 20)

            clean_df = load_and_clean(self.csv_path.get())
            summary  = cleaning_summary(raw_df, clean_df)
            self._set_status(
                f"Cleaned: {summary['rows_before']} → {summary['rows_after']} rows. Analysing...",
                40
            )

            kpis        = kpi_summary(clean_df)
            monthly     = monthly_revenue(clean_df)
            customers   = top_customers(clean_df)
            products    = top_products(clean_df)
            region      = region_performance(clean_df)
            self._set_status("Building Excel report...", 70)

            output_path = generate_report(
                clean_df     = clean_df,
                kpis         = kpis,
                monthly      = monthly,
                top_customers = customers,
                top_products  = products,
                region       = region,
                source_path  = self.csv_path.get(),
                output_path  = self.output_dir.get(),
            )
            self._set_status(f"✅  Report saved: {os.path.basename(output_path)}", 100)
            messagebox.showinfo("Done!",
                                f"Report generated successfully!\n\n{output_path}")

        except Exception as e:
            self._set_status(f"❌  Error: {e}", 0)
            messagebox.showerror("Error", traceback.format_exc())

        finally:
            self.gen_btn.configure(state="normal")

    def _set_status(self, message: str, progress: float = None):
        self.status_text.set(message)
        if progress is not None:
            self.progress.set(progress)
        self.update_idletasks()


# ── Entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = MISApp()
    app.mainloop()