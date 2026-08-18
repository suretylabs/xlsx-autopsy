# Changelog

## 0.2.0 — 2026-08-17

Launch surface.

- Positional workbook path: `uv run xlsx-autopsy report.xlsx`
- `-V` / `--version`
- Hatchling version is sourced from `xlsx_autopsy.__version__`
- SECURITY.md, Dependabot, PyPI trusted-publish workflow
- Published to PyPI as [`xlsx-autopsy`](https://pypi.org/project/xlsx-autopsy/) `0.2.0`
- GitHub issue template that forbids real workbooks
- README positions this against Excel, openpyxl, and xlwings

## 0.1.2 — 2026-08-17

uv-native public surface.

- Install and run docs are `uv sync` / `uv run` / `uvx`. There is no pip path.
- `[tool.uv] package = true` so uv treats this as a project package.

## 0.1.1 — 2026-08-17

Fail-closed public slice.

- CLI exits `1` on a missing workbook, a corrupt zip, or a formula-worker failure.
- Connection redaction covers brace-wrapped and quoted OLEDB values, including secrets that contain semicolons.
- Formula scout includes a sheet when the scout window errors or the part is larger than the window. Sheets past `max_sheets` are recorded, not dropped.
- Formula workers discard their CSV on error. The parent never COPY-loads a missing file as success.
- Each run wipes DuckDB, the blueprint, and `parquet/` unless `--keep-outputs` is set.
- Punctuation-only sheet names (`!!!`) become `sheet_<sha256-8>` instead of an empty table name.
- Pin `polars>=1.0,<2`. Drop unused `python-calamine` (Polars Calamine goes through `fastexcel`).

## 0.1.0 — 2026-08-17

First public extract: CLI, redaction by default, synthetic fixture tests, ruff/pyright/pytest on 3.12–3.14.
