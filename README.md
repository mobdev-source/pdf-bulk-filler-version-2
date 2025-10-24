# pdf-bulk-filler

Desktop MVP for mapping CSV/Excel data to fillable PDF forms with a drag-and-drop workflow.

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -e .[dev]
```

Launch the UI with:

```bash
python -m pdf_bulk_filler.main
```

Use `python -m pdf_bulk_filler.main --no-gui` for automated smoke checks, or `--data/--template/--mapping` to preload resources when starting the app.

## Key Features

- Split-screen PySide6 UI that keeps spreadsheet columns visible beside the PDF preview, complete with per-page navigation and zoom shortcuts.
- Filter the available columns instantly with the built-in search box, then drag headers onto highlighted PDF form fields to create reusable mappings.
- Excel imports prompt for worksheet selection when several sheets are present, and saved mappings remember the chosen sheet. When headers or data start deeper in the file, use **Adjust Data Range** to set header/data rows and the first column before mapping fields.
- Persist mappings to JSON and reload them later; sample datasets and templates live under `assets/data/` and `assets/templates/` (e.g., `sample_contacts.xlsx`, `Fillable_CIS-Individual-BPI.pdf`) for quick demos.
- Generate individual or combined PDFs locally via a background worker; progress updates stream to the status bar, with optional read-only output for immutable delivery.

To build standalone executables, wire PyInstaller against the console script `pdf-bulk-filler`.

## Development Tips

- Run the test suite with `pytest -q` before committing changes.
- The CLI accepts `--no-gui` for headless validation and `--data/--template/--mapping` arguments to preload resources during manual testing.
- Keep new fixtures under `tests/fixtures/` and leverage the sample workbooks/PDFs in `assets/` for regression coverage.


