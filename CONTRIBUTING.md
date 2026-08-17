# Contributing

This is a small public tool. Keep it that way.

The package manager is **uv**. Do not document or add a pip, poetry, or pdm path.

## Bar

1. Read `PYTHON_STYLEGUIDE.md`. It is normative, not decorative.
2. Run the commands in `FORMAT_LINT_TYPECHECK_README.md` before you push.
3. Do not add dependencies without a justification in the PR.
4. Do not commit a real workbook, a connection string, or anything from `out/`.

## Dev loop

```pwsh
uv sync --group dev
uv run xlsx-autopsy path/to/workbook.xlsx -o out --skip-formulas
uv run ruff check src tests
uv run ruff format --check src tests
uv run pyright src tests
uv run pytest
```

## Secrets

Default extract redacts OLEDB/ODBC credentials. If you change `xlsx_autopsy.secrets`
or `extract_connections`, add or update a test that would fail if a password leaked
into `report_blueprint.json`.
