# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Invoice PDF auto-rename and verification tool. Renames PDFs based on contract number lookup from Excel, and verifies invoice data against Excel records using offline OCR.

## Key Commands

```bash
uv sync                                # Install dependencies
uv sync --extra dev                    # Install with dev dependencies

uv run main.py generate-test-data --output ./test_data   # Generate mock data
uv run main.py rename --dir ./pdfs --excel ./合同号索引表.xlsx   # Rename PDFs
uv run main.py verify --dir ./pdfs --excel ./发票验证表.xlsx     # Verify invoices
uv run main.py -v rename --dir ./pdfs --excel ./合同号索引表.xlsx  # Verbose logging

uv run pytest tests/ -v                # Run all tests
uv run pytest tests/test_contract.py -v  # Run single test file
uv run pytest -k test_name             # Run single test by name
```

## Architecture

- **ocr.py** — PDF text extraction: PyMuPDF text layer first, PaddleOCR fallback for scanned PDFs. Invoice field parsing via regex.
- **contract.py** — Contract index Excel loading. 甲方合同号 extraction from `YW-xxx` names (SZ prefix → second dash, otherwise → first dash).
- **rename.py** — PDF classification (Type 1: `2100002601080A_xxx.pdf`, Type 2: `dzfp_xxx.pdf`) and batch rename. Output to `{dir}_Renamed/` folder (copies files, originals untouched).
- **verify.py** — Excel vs OCR comparison (columns AF/BZ/CA/CB/CC/CD, rows 4+). Fills empty invoice_no/invoice_date from OCR. Copies filled PDFs to `{dir}_filled/` renamed as `{invoice_no}.pdf`. Saves to `{name}_Verified.xlsx`.
- **report.py** — Generates text reports after rename/verify steps, saved as `rename_report.txt` and `verify_report.txt`.
- **excel_utils.py** — Excel column mapping and read-only access with openpyxl.
- **gui.pyw** — tkinter GUI launcher. Double-click `run.bat` to open. File browser dialogs for folder/file selection.

## Conventions

- **Offline only** — No network calls. PaddleOCR runs on local CPU.
- **Python 3.12** (pinned: PaddlePaddle only supports 3.12), managed by `uv`.
- **TDD** — Write failing test first, then implement.
- **After each commit**: Review if CLAUDE.md and README.md need updates. Update them in the same commit if the change affects architecture, commands, or usage.
