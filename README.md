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

- Split-screen interface built with PySide6 showing spreadsheet columns alongside a rendered PDF page with per-page navigation, zoom controls, and menu shortcuts.
- Drag column headers onto highlighted PDF form fields to build reusable mappings, then prune assignments directly from the mapping table.
- Excel imports prompt for worksheet selection when multiple sheets exist, and the chosen sheet stays attached to saved mappings. If headers begin deeper in the file, use 'Adjust Data Range' to set header/data rows and the first column before mapping fields.
- Persist mappings to JSON and regenerate them later; sample assets live under `assets/` (`sample_contacts.csv` + `sample_invoice.pdf`) for quick demos.
- Batch-generate filled PDFs locally using a responsive background worker; progress streams to the status bar and output can be flattened for immutable delivery.

To build standalone executables, wire PyInstaller against the console script `pdf-bulk-filler`.


