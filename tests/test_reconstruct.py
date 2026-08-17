"""Public-safe extract of a synthetic workbook."""

from __future__ import annotations

import json
from pathlib import Path
import zipfile

import pytest

from conftest import FAKE_PASSWORD, RAW_CONNECTION
from xlsx_autopsy.reconstruct import extract_connections, parse_pivot_caches


def test_extract_connections_redacts_by_default(fixture_xlsx: Path) -> None:
    with zipfile.ZipFile(fixture_xlsx) as zipf:
        recs = extract_connections(zipf)
    assert recs
    rec = recs[0]
    assert rec["name"] == "Warehouse"
    assert rec["secrets_redacted"] is True
    assert rec["connection_string"] is not None
    assert FAKE_PASSWORD not in rec["connection_string"]
    assert ("P" + "WD=***") in rec["connection_string"]
    assert rec["sql_command"] == "SELECT 1 AS n"


def test_extract_connections_opt_in_keeps_secret(fixture_xlsx: Path) -> None:
    with zipfile.ZipFile(fixture_xlsx) as zipf:
        recs = extract_connections(zipf, include_secrets=True)
    assert recs[0]["connection_string"] == RAW_CONNECTION
    assert recs[0]["secrets_redacted"] is False


def test_parse_pivot_caches(fixture_xlsx: Path) -> None:
    with zipfile.ZipFile(fixture_xlsx) as zipf:
        caches = parse_pivot_caches(zipf)
    assert caches
    assert caches[0]["type"] == "external"


def test_cli_writes_redacted_blueprint(
    fixture_xlsx: Path,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from xlsx_autopsy.reconstruct import main

    out = tmp_path / "out"
    monkeypatch.setattr(
        "sys.argv",
        ["xlsx-autopsy", "--excel", str(fixture_xlsx), "-o", str(out), "--skip-formulas"],
    )
    main()
    blueprint = json.loads((out / "report_blueprint.json").read_text(encoding="utf-8"))
    dumped = json.dumps(blueprint)
    assert FAKE_PASSWORD not in dumped
    assert blueprint["connections"][0]["secrets_redacted"] is True
    assert (out / "reconstruction.duckdb").is_file()
    assert not (out / "excel_analysis.json").exists()
