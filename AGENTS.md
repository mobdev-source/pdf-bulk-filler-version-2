# Repository Guidelines

## Project Structure & Module Organization
Code follows the src layout. Core package lives in `src/pdf_bulk_filler`, with `main.py` hosting the CLI stub and `__init__.py` exposing public entry points. Tests sit in `tests/` and should mirror package modules one-to-one. Sample data and templates live in `assets/data/` and `assets/templates/`; treat those as canonical fixtures and add new ones sparingly. UI components live under `src/pdf_bulk_filler/ui/`—keep widgets dumb, push orchestration into `MainWindow`, and avoid embedding business logic in view code.

## Build, Test, and Development Commands
Create a fresh environment with:
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -e .[dev]
```
Run the CLI locally via `python -m pdf_bulk_filler.main --help`; use `--no-gui` for headless smoke tests. Package metadata builds with `python -m build` once the codebase is ready for distribution. Continuous validation relies on `pytest`, so `pytest -q` should pass before every push. The fixture workbook `assets/data/sample_contacts.xlsx` exists to exercise worksheet selection; preserve it when reshaping demos.

## Coding Style & Naming Conventions
Use PEP 8 defaults: four-space indentation, snake_case for functions and variables, PascalCase for classes, and UPPER_CASE for constants. Type hints are required on public functions; prefer `Path` over bare strings for I/O. Keep functions focused, document assumptions with concise docstrings, and raise explicit errors instead of failing silently. PySide6 work should use layout managers, emit status-bar feedback for long operations, mirror toolbar actions with keyboard shortcuts, and pipe worksheet/data-range selections through the shared dialogs so mappings retain header/data offsets.

## Testing Guidelines
Write tests with `pytest` and place them under `tests/`, naming files `test_<module>.py`. Mirror CLI behavior with focused unit tests that exercise argument parsing and data handling. Use `capsys` to assert console output, and add regression fixtures in `tests/fixtures/` for complex PDFs. `tests/test_pdf_engine.py` covers flattening against the sample invoice, while `tests/test_data_loader.py` asserts Excel metadata—extend them when data-loading semantics evolve. GUI-heavy code should stay thin enough that underlying services remain unit-testable.

## Commit & Pull Request Guidelines
Adopt Conventional Commit prefixes (`feat:`, `fix:`, `docs:`, etc.) to keep history searchable. Limit the subject to 72 characters and expand context in the body when altering behavior. Pull requests should describe the scenario, note testing evidence (`pytest`, manual runs), and attach sample input/output PDFs when relevant.

## Security & Configuration Tips
Never commit real customer data or unredacted PDFs. Store API keys and service credentials in environment variables or `.env` files that stay untracked. When integrating external storage, capture configuration steps in `docs/` so automation agents can replay them deterministically.



