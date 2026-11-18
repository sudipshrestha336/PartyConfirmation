"""
Balance Confirmation Bulk Generator — GUI (PDF Output Version)
==============================================================
Tkinter-based interface to select Word template, Excel data file,
output folder, and run the merge process, outputting high-quality PDF files.
"""

import re
import sys
import os
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Ensure required libraries are installed
try:
    from docx import Document
    import openpyxl
    from docx2pdf import convert
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx", "openpyxl", "docx2pdf"])
    from docx import Document
    import openpyxl
    from docx2pdf import convert


# ─────────────────────────────────────────────
#  AMOUNT FORMATTING CONFIG
# ─────────────────────────────────────────────
AMOUNT_COLUMNS  = ["Opening_Balance", "Taxable_Purchase", "VAT_Purchase", "Non_Taxable_Purchase", "TDS", "Closing_Balance", "Taxable_Sales","Non_Taxable_Sales","VAT_Sales","Advance_Tax_Received"]
AMOUNT_DECIMALS = 2
# ─────────────────────────────────────────────


# ── Core merge logic ──────────────────────────────────────────────────────────

def format_value(value, header: str = "") -> str:
    if value is None:
        return ""
    if header.upper() in [c.upper() for c in AMOUNT_COLUMNS]:
        try:
            return f"{float(value):,.{AMOUNT_DECIMALS}f}"
        except (ValueError, TypeError):
            pass
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    if isinstance(value, (int, float)):
        return str(value)
    return str(value).strip()


def load_excel_data(excel_path: str, sheet_name=None):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel file is empty.")
    headers = [str(h).strip() if h is not None else f"COL{i}" for i, h in enumerate(rows[0])]
    data = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        record = {headers[i]: format_value(v, headers[i]) for i, v in enumerate(row)}
        data.append(record)
    return headers, data


def apply_placeholders(text: str, record: dict) -> str:
    def replacer(match):
        key = match.group(1).strip()
        return record.get(key, match.group(0))
    return re.sub(r"\{\{([^}]+)\}\}", replacer, text)


def replace_in_paragraph(paragraph, record):
    full_text = "".join(run.text for run in paragraph.runs)
    if "{{" not in full_text:
        return
    new_text = apply_placeholders(full_text, record)
    if new_text == full_text:
        return
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""


def replace_in_table(table, record):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_in_paragraph(para, record)
            for nested_table in cell.tables:
                replace_in_table(nested_table, record)


def fill_template(template_path, record, output_docx_path):
    doc = Document(template_path)
    for para in doc.paragraphs:
        replace_in_paragraph(para, record)
    for table in doc.tables:
        replace_in_table(table, record)
    for section in doc.sections:
        for hdr in [section.header, section.first_page_header, section.even_page_header]:
            if hdr:
                for para in hdr.paragraphs:
                    replace_in_paragraph(para, record)
                for table in hdr.tables:
                    replace_in_table(table, record)
        for ftr in [section.footer, section.first_page_footer, section.even_page_footer]:
            if ftr:
                for para in ftr.paragraphs:
                    replace_in_paragraph(para, record)
                for table in ftr.tables:
                    replace_in_table(table, record)
    doc.save(output_docx_path)


def make_safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Balance Confirmation Bulk Generator (PDF) - Fonepay Payment Service Limited")
        self.resizable(False, False)
        self.configure(bg="#f0f0f0")

        # ── Variables ──
        self.template_var   = tk.StringVar()
        self.excel_var      = tk.StringVar()
        self.output_var     = tk.StringVar()
        self.pattern_var    = tk.StringVar(value="Confirmation_{{PARTY_NAME}}.pdf")
        self.sheet_var      = tk.StringVar(value="")
        self.columns_var    = tk.StringVar()   # read-only, shows detected columns

        self._build_ui()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        PAD = dict(padx=12, pady=6)

        # Title bar
        title_frame = tk.Frame(self, bg="#1a5276", pady=10)
        title_frame.grid(row=0, column=0, sticky="ew")
        tk.Label(title_frame, text="⚙  Balance Confirmation Bulk Generator (PDF Export)",
                 font=("Segoe UI", 13, "bold"), fg="white", bg="#1a5276").pack()

        # Main frame
        frm = tk.Frame(self, bg="#f0f0f0", padx=20, pady=16)
        frm.grid(row=1, column=0, sticky="nsew")

        def section_label(text, row):
            tk.Label(frm, text=text, font=("Segoe UI", 9, "bold"),
                     fg="#1a5276", bg="#f0f0f0").grid(row=row, column=0, columnspan=3,
                     sticky="w", pady=(12, 2))

        def file_row(label, var, cmd, row, filetypes=None):
            tk.Label(frm, text=label, bg="#f0f0f0",
                     font=("Segoe UI", 9)).grid(row=row, column=0, sticky="w", **PAD)
            entry = tk.Entry(frm, textvariable=var, width=48,
                             font=("Segoe UI", 9), state="readonly",
                             readonlybackground="white", relief="solid")
            entry.grid(row=row, column=1, sticky="ew", padx=4)
            tk.Button(frm, text="Browse…", command=cmd, width=9,
                      font=("Segoe UI", 9), bg="#2980b9", fg="white",
                      relief="flat", cursor="hand2").grid(row=row, column=2, padx=(4, 0))

        # ── Section 1: Files ──
        section_label("1.  Select Files", 0)
        file_row("Word Template (.docx)", self.template_var, self._browse_template, 1)
        file_row("Excel Data (.xlsx)",    self.excel_var,    self._browse_excel,    2)

        # ── Section 2: Output ──
        section_label("2.  Output Settings", 3)
        file_row("Output Folder",         self.output_var,   self._browse_output,   4)

        tk.Label(frm, text="File Name Pattern", bg="#f0f0f0",
                 font=("Segoe UI", 9)).grid(row=5, column=0, sticky="w", **PAD)
        tk.Entry(frm, textvariable=self.pattern_var, width=48,
                 font=("Segoe UI", 9), relief="solid").grid(row=5, column=1,
                 sticky="ew", padx=4, pady=6)

        tk.Label(frm, text="Sheet Name (optional)", bg="#f0f0f0",
                 font=("Segoe UI", 9)).grid(row=6, column=0, sticky="w", **PAD)
        tk.Entry(frm, textvariable=self.sheet_var, width=48,
                 font=("Segoe UI", 9), relief="solid").grid(row=6, column=1,
                 sticky="ew", padx=4, pady=6)
        tk.Label(frm, text="leave blank = first sheet", bg="#f0f0f0",
                 font=("Segoe UI", 8), fg="gray").grid(row=6, column=2, sticky="w")

        # ── Section 3: Preview columns ──
        section_label("3.  Excel Columns Detected", 7)
        col_entry = tk.Entry(frm, textvariable=self.columns_var, width=48,
                             font=("Segoe UI", 9), state="readonly",
                             readonlybackground="#eaf4fb", relief="solid", fg="#1a5276")
        col_entry.grid(row=8, column=0, columnspan=3, sticky="ew", padx=4, pady=4)

        # ── Run button ──
        tk.Button(frm, text="▶   Generate PDFs", command=self._run,
                  font=("Segoe UI", 11, "bold"), bg="#1e8449", fg="white",
                  relief="flat", cursor="hand2", pady=8).grid(
                  row=9, column=0, columnspan=3, sticky="ew", pady=(16, 6))

        # ── Progress bar ──
        self.progress = ttk.Progressbar(frm, mode="determinate", length=400)
        self.progress.grid(row=10, column=0, columnspan=3, sticky="ew", pady=4)

        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(frm, textvariable=self.status_var, bg="#f0f0f0",
                 font=("Segoe UI", 9), fg="#555").grid(
                 row=11, column=0, columnspan=3, sticky="w")

        # ── Log box ──
        section_label("4.  Log", 12)
        self.log = scrolledtext.ScrolledText(frm, width=64, height=12,
                                             font=("Consolas", 9),
                                             state="disabled", relief="solid",
                                             bg="#1c1c1c", fg="#e0e0e0",
                                             insertbackground="white")
        self.log.grid(row=13, column=0, columnspan=3, sticky="ew", pady=4)

        # tag colors
        self.log.tag_config("ok",    foreground="#2ecc71")
        self.log.tag_config("err",   foreground="#e74c3c")
        self.log.tag_config("info",  foreground="#3498db")
        self.log.tag_config("head",  foreground="#f39c12", font=("Consolas", 9, "bold"))

    # ── Browse helpers ─────────────────────────────────────────────────────────

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Select Word Template",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")])
        if path:
            self.template_var.set(path)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel Data File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")])
        if path:
            self.excel_var.set(path)
            self._preview_columns(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            self.output_var.set(path)

    def _preview_columns(self, excel_path):
        try:
            sheet = self.sheet_var.get().strip() or None
            headers, _ = load_excel_data(excel_path, sheet)
            self.columns_var.set("  |  ".join(headers))
        except Exception as e:
            self.columns_var.set(f"Error reading columns: {e}")

    # ── Logging ────────────────────────────────────────────────────────────────

    def _log(self, msg, tag=""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── Run ────────────────────────────────────────────────────────────────────

    def _run(self):
        template = self.template_var.get().strip()
        excel    = self.excel_var.get().strip()
        out_dir  = self.output_var.get().strip()
        pattern  = self.pattern_var.get().strip()
        sheet    = self.sheet_var.get().strip() or None

        # Validate
        if not template:
            messagebox.showwarning("Missing", "Please select a Word template."); return
        if not excel:
            messagebox.showwarning("Missing", "Please select an Excel file."); return
        if not out_dir:
            messagebox.showwarning("Missing", "Please select an output folder."); return
        if not pattern:
            messagebox.showwarning("Missing", "Please enter a file name pattern."); return

        # Force output file extension configuration to match PDF requirements
        if not pattern.lower().endswith(".pdf"):
            if pattern.lower().endswith(".docx"):
                pattern = pattern[:-5] + ".pdf"
            else:
                pattern += ".pdf"
            self.pattern_var.set(pattern)

        # Run in background thread so UI stays responsive
        threading.Thread(
            target=self._run_merge,
            args=(template, excel, out_dir, pattern, sheet),
            daemon=True
        ).start()

    def _run_merge(self, template, excel, out_dir, pattern, sheet):
        self._log("\n=== Starting Merge & PDF Generation ===", "head")
        self.status_var.set("Loading Excel data…")

        try:
            headers, records = load_excel_data(excel, sheet)
            self._log(f"  Loaded {len(records)} parties  |  Columns: {', '.join(headers)}", "info")
        except Exception as e:
            self._log(f"❌  Failed to read Excel: {e}", "err")
            self.status_var.set("Error reading Excel.")
            return

        Path(out_dir).mkdir(parents=True, exist_ok=True)
        total   = len(records)
        success = 0
        errors  = []

        self.progress["maximum"] = total
        self.progress["value"]   = 0

        for i, record in enumerate(records, 1):
            try:
                # 1. Determine safe filenames
                pdf_filename = apply_placeholders(pattern, record)
                pdf_filename = make_safe_filename(pdf_filename)
                
                # Make dynamic unique temporary Word file names to compile into PDF
                docx_filename = pdf_filename.replace(".pdf", ".docx").replace(".PDF", ".docx")
                if not docx_filename.endswith(".docx"):
                    docx_filename += ".docx"
                
                temp_docx_path = Path(out_dir) / f"~temp_{i}_{docx_filename}"
                final_pdf_path = Path(out_dir) / pdf_filename

                # 2. Fill the text modifications into a temporary Word docx
                fill_template(template, record, str(temp_docx_path))
                
                # 3. Use system layout conversion engine to spit out a clean PDF
                self.status_var.set(f"Converting {i}/{total} to PDF…")
                convert(str(temp_docx_path), str(final_pdf_path))
                
                # 4. Wipe temporary file traces out of the folder
                if temp_docx_path.exists():
                    os.remove(temp_docx_path)

                self._log(f"  [{i:>4}/{total}] ✅  {pdf_filename}", "ok")
                success += 1
            except Exception as e:
                msg = f"Row {i}: {e}"
                errors.append(msg)
                self._log(f"  [{i:>4}/{total}] ❌  {msg}", "err")
                try:
                    if 'temp_docx_path' in locals() and temp_docx_path.exists():
                        os.remove(temp_docx_path)
                except:
                    pass

            self.progress["value"] = i
            self.status_var.set(f"Processed {i} of {total}…")
            self.update_idletasks()

        # Summary Window Update
        self._log(f"\n{'='*44}", "head")
        self._log(f"  ✅  {success} PDF files created in:\n      {out_dir}", "ok")
        if errors:
            self._log(f"  ❌  {len(errors)} errors:", "err")
            for err in errors:
                self._log(f"      - {err}", "err")
        self._log(f"{'='*44}\n", "head")
        self.status_var.set(f"Done! {success} PDFs created. {len(errors)} errors.")
        self.progress["value"] = total

        if not errors:
            messagebox.showinfo("Complete",
                f"✅  {success} PDFs generated successfully!\n\nSaved to:\n{out_dir}")
        else:
            messagebox.showwarning("Complete with errors",
                f"{success} PDFs generated.\n{len(errors)} errors — see log for details.")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
