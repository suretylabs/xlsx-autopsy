# xlsx-autopsy

[![ci](https://github.com/suretylabs/xlsx-autopsy/actions/workflows/ci.yml/badge.svg)](https://github.com/suretylabs/xlsx-autopsy/actions/workflows/ci.yml)
[![pypi](https://img.shields.io/pypi/v/xlsx-autopsy.svg)](https://pypi.org/project/xlsx-autopsy/)

Decompose a huge Excel workbook without opening it.

The grid is a rendering. The report lives in pivot caches and connections.
Those caches store field **indexes**. This CLI resolves them to names,
redacts OLEDB secrets, and lifts the model into DuckDB and Parquet.

Inside the xlsx:

```json
{"row_fields": [2], "data_fields": [{"fld": 11}]}
```

In `report_blueprint.json`:

```json
{"row_field_names": ["Region"], "data_field_names": [{"field_name": "Premium"}]}
```

That mapping is usually the report. The sheet is just how Excel drew it.

```bash
uvx xlsx-autopsy report.xlsx
```

This is a report autopsy. Not a spreadsheet editor. Not a malware scanner.
Not another xlsx-to-DataFrame dump.

## Why this exists

Some production "reports" are 300MB workbooks nobody should open. The people
who built them left. The logic is in the zip, not the grid.

| What people try | What they actually get |
| --- | --- |
| Double-click in Excel | The rendering. Then you wait. |
| pandas / DuckDB `read_xlsx` / `xlsx2csv` | The grid as a table. Pivot indexes stay indexes. Connections are gone. |
| `openpyxl` / `xlrd` | A cell library. You still write the reconstruction. |
| `xlwings` | Automates Excel. You still opened Excel. |
| `oletools` / ExcelRaven | Macros, malware, leaked passwords. Different job. |
| This | Field `2` becomes `Region`. Secrets stripped. Formulas streamed. Queryable output. |

Opening the file is how you lose an afternoon.

## Install

**uv-native.** There is no pip path.

```bash
uvx xlsx-autopsy workbook.xlsx
uv tool install xlsx-autopsy
```

From a clone:

```bash
uv sync
uv run xlsx-autopsy workbook.xlsx
```

Pin a SHA with `uvx --from git+https://github.com/suretylabs/xlsx-autopsy xlsx-autopsy`.

Contributors:

```bash
uv sync --group dev
```

Python 3.12, 3.13, or 3.14. Local default is 3.14 (`.python-version`).

## Usage

```bash
uv run xlsx-autopsy workbook.xlsx
uv run xlsx-autopsy workbook.xlsx -o out --skip-formulas
uv run python -m xlsx_autopsy workbook.xlsx
uv run xlsx-autopsy -V
```

Each run wipes `reconstruction.duckdb`, `report_blueprint.json`, and `parquet/`
in the output directory so two workbooks never commingle. Pass `--keep-outputs`
only when you are deliberately appending.

Default output directory is `out/`:

| Artifact | What you get |
| --- | --- |
| `report_blueprint.json` | Workbook meta, redacted connections, resolved pivots |
| `reconstruction.duckdb` | Queryable extract: sheets, formulas, SST, metadata tables |
| `parquet/` | One file per sheet |

Connection strings are **redacted by default**, including brace-wrapped OLEDB
secrets that contain semicolons. If you actually need the raw string, pass
`--include-connection-secrets`. Do not commit that output.

A missing file, a corrupt zip, or a failed formula worker exits **1**.
Success is not silent.

Optional TOML: copy [`xlsx-autopsy.toml.example`](xlsx-autopsy.toml.example)
to `xlsx-autopsy.toml` (gitignored).

```bash
uv run xlsx-autopsy --config xlsx-autopsy.toml
```

## What it extracts

| Surface | Why it matters |
| --- | --- |
| Pivot tables + cache fields | Field indexes become names. This is usually the report. |
| Pivot cache sources | SQL / table lineage hiding in the cache definition. |
| Connections | Where the workbook actually drinks from. Secrets stripped. |
| Formulas | Shared formulas are deduped so a 300MB file does not explode. |
| Shared strings | The lookup text sitting next to the numbers. |
| Sheet values | Calamine reads the rendered grid when you need it. |

## Engineering bar

This is not a gist with a README taped on.

- [PYTHON_STYLEGUIDE.md](PYTHON_STYLEGUIDE.md) — normative Python standard
- [FORMAT_LINT_TYPECHECK_README.md](FORMAT_LINT_TYPECHECK_README.md) — ruff / pyright / pytest
- [CONTRIBUTING.md](CONTRIBUTING.md) — how to send a change
- [SECURITY.md](SECURITY.md) — how to report a hole

```bash
uv sync --group dev
uv run ruff check src tests
uv run ruff format --check src tests
uv run pyright src tests
uv run pytest
```

CI runs that pack on 3.12, 3.13, and 3.14.

## Safety

This repo ships **no workbooks, no connection strings, no company config**.

The fixture workbook in `tests/` is synthetic. If you point this at a real
file, keep the outputs out of git.

## License

MIT. See [LICENSE](LICENSE).
