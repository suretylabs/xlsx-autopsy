# Linting, Formatting, and Type Checking

**Last Updated**: 2026-08-17
**Python Version**: 3.12+ (project targets 3.14)
**Tooling**: Ruff (lint + format), Pyright (types), uv (deps), pytest (tests)

## Overview

Ruff-first workflow. Pyright for types. uv for a reproducible environment.
`PYTHON_STYLEGUIDE.md` is the normative prose standard. This file is the operator
manual for the tools that enforce the mechanical subset.

This repository is small on purpose. Run the checks over `src` and `tests`. Do not
invent a scoped wrapper just because a larger monorepo needed one.

## Tooling summary

| Tool | Job | Config |
| --- | --- | --- |
| Ruff | lint + format | `ruff.toml` |
| Pyright | type check | `pyrightconfig.json` |
| uv | lockfile + runner | `pyproject.toml`, `uv.lock` |
| pytest | tests | `[tool.pytest.ini_options]` |

Never use Pylance/Black/isort/flake8 as the style source of truth. Ruff owns that.

## Install

```pwsh
uv sync --group dev
```

Python 3.12, 3.13, or 3.14. Local default is 3.14 (`.python-version`). The console script is `xlsx-autopsy`.

## Commands that must pass before a push

```pwsh
uv run ruff check src tests
uv run ruff format --check src tests
uv run pyright src tests
uv run pytest
```

Apply format only to what you touched when you can:

```pwsh
uv run ruff format src/xlsx_autopsy/secrets.py
uv run ruff check --fix src/xlsx_autopsy/secrets.py
```

`ruff check` passing is **not** proof of formatting. CI runs `ruff format --check .`.

## Ruff

Enabled families (see `ruff.toml`, do not fork this list in prose):

- `E` / `W` pycodestyle
- `F` Pyflakes
- `I` import sorting
- `N` naming
- `UP` pyupgrade
- `TID` tidy imports
- `B` bugbear
- `D` Google docstrings
- `T20` no `print()`

Line length is 120. That is calibrated to stay above normal authoring and still
catch runaways. Do not bikeshed it down to 88 in a drive-by PR.

Docstring convention is Google (`D203` and `D213` ignored as the conventional pair).
Tests omit `D` because test names are the documentation unless a fixture has a real
contract.

## Type checking

```pwsh
uv run pyright src tests
```

- Mode is `strict` with selected noise from incomplete third-party stubs turned down.
- Prefer `cast()` plus a comment over a bare `# type: ignore`.
- When an ignore is required, name the code and the reason:

  ```python
  tree = et.parse(handle)  # type: ignore[assignment] # lxml/etree stub mismatch
  ```

- Do not add pandas/numpy stub packs this package does not use.
- `reportMissingTypeStubs` is off because Calamine/lxml stubs are incomplete. That
  is not permission to skip annotations on *our* functions.

## Typing conventions

- Python 3.12+ standard library typing. Use `typing_extensions` only for a feature
  that is not in `typing` yet, and justify it.
- Prefer built-in generics: `list[str]`, `dict[str, Any]`, `X | None`.
- `from __future__ import annotations` goes directly under the module docstring.
- Annotate every non-test public and private function. Nested helpers can inherit
  clarity from the enclosing signature when they stay tiny.

## Tests

- Prove the default connection extract cannot emit the fixture password.
- Prove `--include-connection-secrets` is opt-in, not ambient.
- Fixture workbooks are built in memory under `tests/`. Do not check in a real `.xlsx`.

```pwsh
uv run pytest
uv run pytest tests/test_secrets.py
```

## VS Code

Workspace settings belong in `.vscode/settings.json` if you want them. Recommended:

```jsonc
{
  "python.languageServer": "Pylance",
  "python.analysis.diagnosticMode": "openFilesOnly",
  "[python]": {
    "editor.defaultFormatter": "charliermarsh.ruff",
    "editor.formatOnSave": true,
    "editor.codeActionsOnSave": {
      "source.fixAll.ruff": "explicit",
      "source.organizeImports.ruff": "explicit"
    }
  }
}
```

Type-checking mode lives in `pyrightconfig.json`, not in editor settings.
Ruff behavior lives in `ruff.toml`, not in editor settings.

## CI

`.github/workflows/ci.yml` runs the same four commands on Python 3.12, 3.13, and 3.14.
If you change a local check, change CI in the same PR.

## Updating tools

```pwsh
uv add --dev ruff
uv add --dev pyright
```

Pin changes go through `uv.lock` in the same commit. Do not "just run latest" on
a contributor machine and forget the lockfile.
