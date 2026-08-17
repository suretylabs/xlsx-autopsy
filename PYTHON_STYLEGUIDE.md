# Python Styleguide

**Last Updated**: 2026-08-17
**Python Version**: 3.12+ (project targets 3.14)
**Applies To**: All non-test `.py` files in this repository

## Rule Scope and Enforcement

- This document is the normative engineering standard for non-test Python code. `ruff.toml` defines the
  automatically enforceable subset; a rule that is not represented by a Ruff check is still required.
- Ruff is the source of truth for mechanical formatting, import organization, and lint categories. The concrete
  formatter values, enabled rules, ignores, and per-file exceptions belong in `ruff.toml`; do not duplicate them
  here. Ruff is an initial pass, not a substitute for the required manual review below.
- Test-only Ruff exceptions (for example missing docstrings) are tooling boundaries. They do not weaken the
  prose rules for runtime code. Operator-facing CLI output may use Rich; everything else uses `logging`.
- When a rule is not mechanically enforceable, reviewers and agents should report the exception explicitly rather
  than silently weakening the rule or changing the shared Ruff configuration.

## General Principles

- Use Google-style docstrings for all non-test modules, classes, and concrete functions or methods.
- Include type hints for all non-test function parameters and return types.
- Support Python 3.12+. The project targets Python 3.14. Syntax or APIs newer than 3.12 require an explicit
  change to the supported-version contract before use.
- Follow PEP 8 with Ruff as the primary linter/formatter.
- This repository is uv-native. Document and run `uv sync`, `uv run`, `uv add`,
  and `uv lock`. Do not add a pip, poetry, or pdm install path.

## Configuration

- Put script-level runtime or operator-tunable configuration in a top-level `config` dict unless a project config
  file is used. This package reads optional `xlsx-autopsy.toml` plus CLI/env; do not invent a second config channel.
- Immutable module constants, parser defaults, and internal structural sets are not required to use the `config`
  dict. Keep those as named constants when that makes their scope and immutability clearer.
- Never default to a machine-specific workbook path. The CLI requires `--excel` or a local `workbook.xlsx`.

## Docstrings

- Every non-test module must have a top-level docstring.
- Every standalone script or CLI entry module must have a top-level docstring with a copyable usage example.
- Use `uv run` for execution examples. Prefer the installed console script when one exists.
- Include a `uv run python -m package.module` example when the module supports module execution. Include both
  forms only when both entry points are supported; do not document an invocation that the file does not provide.
- Every non-test class and concrete function or method, including private helpers, should have a Google-style
  docstring. A concise one-line docstring is sufficient when the behavior is obvious from the name, signature,
  and implementation.
- Use expanded `Args:`, `Returns:`, and `Raises:` sections when behavior is not obvious or when the function
  has meaningful failure modes, side effects, platform-specific behavior, filesystem or network boundaries, or
  safety invariants. Document the contract, not information already clear from the type signature.
- Prefer an expanded docstring whenever it adds useful context for a human engineer or an AI coding agent.
  High-value context includes responsibility, lifecycle or ownership, invariants, source-of-truth relationships,
  failure and recovery posture, operational safety boundaries, and meaningful neighboring symbols or entry points.
- Use this repository's vocabulary in docstrings: workbook, pivot cache, connection, blueprint, redaction.
  Name exact paths, symbols, configuration keys, and CLI flags when they materially improve search or handoff;
  avoid brittle line-number references.
- Treat docstrings as semantic navigation surfaces: explain why a behavior exists and how it relates to adjacent
  concepts, not just what each line of the implementation does.
- Do not add expansion that merely restates an obvious type signature or duplicates implementation details likely
  to drift. When a concise docstring communicates the complete contract, keep it concise.
- A local nested helper does not need a separate docstring when its purpose is clear from the enclosing function.
  Document non-obvious nested-helper behavior in the enclosing function's prose instead.
- Test modules and test-only helpers may omit docstrings unless they expose a reusable fixture/helper or document
  a non-obvious contract needed by other tests.

**Example**:

```python
"""xlsx-autopsy: decompose a huge Excel workbook without opening it.

Usage:
    uv run xlsx-autopsy --excel workbook.xlsx
    uv run python -m xlsx_autopsy --excel workbook.xlsx -o out --skip-formulas
"""
```

## Libraries

- Do not add a dependency unless it is required for a real capability. Propose the addition in the PR and justify
  it. `uv add` only after that decision. Do not pin a new stack (HTTP client, ORM, cloud SDK) "just in case."
- Prefer the libraries already in `pyproject.toml`: Polars, DuckDB, Calamine/fastexcel, lxml, Rich.

## Logging

- Use the `logging` module instead of `print()` for diagnostic output.
- The CLI may use Rich (`console.print`, `console.rule`, `console.log`) as the operator surface. File history still
  goes through `logging` so a run is reconstructable after the terminal scroll is gone.
- Configure logging in `main()` or another entrypoint path; do not configure logging as an import-time side effect.
- Use `logger.info()` for general messages, `logger.warning()` for warnings, `logger.error()` for errors.
- Never log raw connection strings, passwords, tokens, or workbook cell values that could contain credentials.
  Connection extract is redacted unless `--include-connection-secrets` is explicitly passed.

## Imports

- Ensure no unused imports; remove any that are not referenced in the code.
- Avoid inline imports; all imports must be at the top of the file.
- Use `from` imports only when necessary; prefer absolute imports for clarity.
- First-party imports use the `xlsx_autopsy` package name.
- For modules requiring future annotations, place `from __future__ import annotations` immediately after the
  module docstring and before any other imports.

## Module export surfaces (`__all__`)

- Prefer an explicit `__all__` on modules that intentionally export a shared surface for other modules: package
  facades, CLI re-export shims, and leaf shared-constant/schema modules.
- Leaf shared-constant modules are a special case: sibling code often does `from ._constants import NAME`, and
  unused-global checkers are typically file-local. Sibling imports do not count as uses inside the defining file.
  Listing those names in `__all__` is the supported way to mark intentional public surface.
- Keep `__all__` in **definition order** with the body when the module is a shared-constant surface. Prefer a
  drift test when the export list is large enough that manual review will miss additions.
- Do **not** add `__all__` to every module "for completeness." Ordinary implementation modules whose module-level
  names are used in the same file do not need an export list.
- Do **not** disable unused-global checks to silence shared-constant false positives. Fix the export surface.

## Polars

- Polars is strict. Do not mix types in a column and hope the writer sorts it out.
- Check `df.dtypes` before casts. Handle nulls explicitly (`fill_null` or conditionals) before integer/float casts.
- Do not `cast(pl.Boolean)` from Utf8. Use a conditional:
  `pl.when(pl.col("flag") == "1").then(True).otherwise(False)`.
- Parse date strings with `str.strptime` (or an existing date dtype). Do not smash formats together.

## Public-safe outputs

- Default extract redacts OLEDB/ODBC secrets. Tests must prove the default path cannot emit a fixture password.
- Do not commit workbooks, `out/`, DuckDB files, or a real `xlsx-autopsy.toml`.
- Fixture workbooks in `tests/` are synthetic. If a test needs a secret-shaped string, keep it obviously fake.

## Validation

### Automated initial pass

```pwsh
uv run ruff check src tests
uv run ruff format --check src tests
uv run pyright src tests
uv run pytest
```

- Run focused tests for the affected behavior first. Escalate to the full suite before a push.
- Ruff format check is not implied by `ruff check`. Both must pass.

### Required manual review

- After the automated pass, an engineer must manually validate this styleguide for every changed non-test `.py`
  file. Review the complete affected module, not only lines reported by Ruff.
- Confirm module usage documentation, type hints, configuration posture, logging behavior, import structure, and
  docstring depth.
- Specifically ask whether expanded docstrings add semantic value. Expand them when they clarify contracts,
  domain terminology, ownership, invariants, failure posture, safety constraints, or navigation to related code.
  Do not expand them solely to satisfy a uniform visual format.
- Record intentional deviations in the review or handoff, including why the deviation is safe and whether it
  should become a documented repository exception.
