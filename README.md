# xlsx-autopsy

Decompose a huge Excel workbook without opening it.

If the file is 300MB and the interesting logic lives in pivot caches, stop
double-clicking. Treat the workbook as a zip of report definitions: connections,
pivot field maps, formulas, and the rendered sheets. Then rebuild the model in
SQL or whatever BI tool you actually want to live in.

## Why this exists

Some production "reports" are Excel workbooks that nobody should open.

- The grid is a rendering.
- The model is in pivot caches, named ranges, and OLEDB connections.
- Opening the file in Excel is how you lose an afternoon.

This is a forensics CLI, not a spreadsheet editor.

## Install

```bash
pip install -e .
# or
uv sync
```

Python 3.12, 3.13, or 3.14. Local default is 3.14.

## Usage

```bash
xlsx-autopsy --excel workbook.xlsx
xlsx-autopsy --excel workbook.xlsx -o out --skip-formulas
python -m xlsx_autopsy --excel workbook.xlsx
```

Each run wipes `reconstruction.duckdb`, `report_blueprint.json`, and `parquet/`
in the output directory so two workbooks never commingle. Pass `--keep-outputs`
only when you are deliberately appending.

Default output directory is `out/`:

- `report_blueprint.json` — workbook meta, connections, resolved pivots
- `reconstruction.duckdb` — queryable extract
- `parquet/` — one file per sheet

Connection strings are **redacted by default**, including brace-wrapped OLEDB
secrets that contain semicolons. If you actually need the raw string, pass
`--include-connection-secrets`. Do not commit that output.

A corrupt workbook or a failed formula worker exits **1**. Success is not silent.

Optional TOML:

```toml
[xlsx_autopsy]
workbook_path = "workbook.xlsx"
output_dir = "out"
sst_truncate = 5000
```

```bash
xlsx-autopsy --config xlsx-autopsy.toml
```

## What it extracts

| Surface | Why it matters |
| --- | --- |
| Pivot tables + cache fields | Field indexes become names. This is usually the report. |
| Pivot cache sources | SQL / table lineage hiding in the cache definition. |
| Connections | Where the workbook actually drinks from. Secrets stripped. |
| Formulas | Shared formulas are deduped so a 300MB file does not explode. |
| Sheet values | Calamine reads the rendered grid when you need it. |

## Engineering bar

This is not a gist with a README taped on.

- [PYTHON_STYLEGUIDE.md](PYTHON_STYLEGUIDE.md) — normative Python standard
- [FORMAT_LINT_TYPECHECK_README.md](FORMAT_LINT_TYPECHECK_README.md) — ruff / pyright / pytest
- [CONTRIBUTING.md](CONTRIBUTING.md) — how to send a change

```bash
uv sync --group dev
uv run ruff check src tests
uv run ruff format --check src tests
uv run pyright src tests
uv run pytest
```

## Safety

This repo ships **no workbooks, no connection strings, no company config**.

The fixture workbook in `tests/` is synthetic. If you fork this and point it
at a real file, keep the outputs out of git.

## License

MIT. See [LICENSE](LICENSE).
