
# Party Confirmation Tool (Balance Confirmation Bulk Generator)

A desktop tool that generates party balance confirmation letters in bulk. It reads party details and financial records from an Excel file, fills a Word template for each party, and exports ready-to-send PDF files.

Built with Python and Tkinter to replace the manual copy-paste of balances into individual confirmation letters during year-end audits.

## Features

- Simple GUI to select the Word template, Excel data file, and output folder
- Generates one PDF confirmation per party (one Excel row = one PDF)
- Placeholder-based templates using `{{COLUMN_NAME}}`, so any Excel column can be used
- Works in the document body, tables (including nested tables), headers, and footers
- Automatic number formatting for amount columns (thousands separator, 2 decimals)
- Customizable output file names, for example `Confirmation_{{PARTY_NAME}}.pdf`
- Live progress bar and colour-coded log, with per-row error reporting
- Runs in a background thread, so the interface stays responsive

## Requirements

- Python 3.8 or later
- **Microsoft Word installed** (Windows or macOS), because PDF conversion uses `docx2pdf`, which relies on Word
- Python packages:
  - `python-docx`
  - `openpyxl`
  - `docx2pdf`

Tkinter ships with most Python installations. The script tries to install missing packages automatically on first run. You can also install them manually:

```bash
pip install python-docx openpyxl docx2pdf
```

## Usage

1. Run the tool:
   ```bash
   python party_confirmation_tool.py
   ```
2. Select the **Word Template** (`.docx`).
3. Select the **Excel Data** file (`.xlsx` or `.xlsm`). The detected column headers are shown in the window.
4. Select the **Output Folder**.
5. (Optional) Change the **File Name Pattern** and enter a **Sheet Name**. If left blank, the first sheet is used.
6. Click **Generate PDFs** and follow the progress in the log.

## Preparing the Excel file

- The first row must contain column headers.
- Each following row is one party. Fully empty rows are skipped.
- Header names are used as placeholder names, so they must match the template exactly.

Example:

| PARTY_NAME | Opening_Balance | Taxable_Purchase | VAT_Purchase | TDS | Closing_Balance |
|---|---|---|---|---|---|
| ABC Traders Pvt. Ltd. | 150000 | 400000 | 52000 | 8000 | 594000 |

### Amount formatting

Columns listed in `AMOUNT_COLUMNS` (top of the script) are formatted as `1,234,567.89`:

```
Opening_Balance, Taxable_Purchase, VAT_Purchase, Non_Taxable_Purchase, TDS,
Closing_Balance, Taxable_Sales, Non_Taxable_Sales, VAT_Sales, Advance_Tax_Received
```

To add or remove columns, or to change the number of decimals, edit `AMOUNT_COLUMNS` and `AMOUNT_DECIMALS`.

## Preparing the Word template

Write placeholders in double curly braces wherever the data should appear:

```
Dear {{PARTY_NAME}},

As per our records, your closing balance as of year-end is NPR {{Closing_Balance}}.
```

- Placeholder names must match the Excel column headers exactly.
- Unmatched placeholders are left unchanged, which makes missing columns easy to spot.
- Tables, headers, and footers are supported.
- Type each placeholder in one go. Avoid changing the formatting partway through a placeholder.

## File name pattern

Placeholders also work in output file names:

| Pattern | Result |
|---|---|
| `Confirmation_{{PARTY_NAME}}.pdf` | `Confirmation_ABC Traders Pvt. Ltd..pdf` |
| `{{PARTY_CODE}}_Confirmation.pdf` | `P001_Confirmation.pdf` |

Characters not allowed in file names (`\ / * ? : " < > |`) are replaced with `_`. The `.pdf` extension is added automatically if missing.

## How it works

1. Reads the Excel rows into records (header → value).
2. For each record, fills the Word template and saves a temporary `.docx`.
3. Converts the `.docx` to PDF using `docx2pdf`.
4. Deletes the temporary file and logs the result.

## Limitations

- Requires Microsoft Word for PDF conversion (not supported on Linux).
- If a paragraph contains a placeholder, the paragraph's text is merged into its first run, so mixed formatting inside that paragraph (for example, one bold word) may be lost. Keep placeholder paragraphs uniformly formatted.
- Only the first row is treated as the header row.
- Formulas in Excel are read as their last saved values, so save the workbook before running.

## Data privacy

Party names and balances are confidential. Do not commit real Excel files, templates with company details, or generated PDFs to this repository. Use anonymized sample data for testing and add data folders to `.gitignore`.

## Tech stack

Python · Tkinter · python-docx · openpyxl · docx2pdf

## License

Add a license of your choice (for example, MIT) before publishing.
