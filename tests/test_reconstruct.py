"""Public-safe extract of a synthetic workbook."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import zipfile

import duckdb
import pytest

from conftest import FAKE_PASSWORD, RAW_CONNECTION, write_fixture_xlsx
from xlsx_autopsy.reconstruct import (
    _worker_formula_task,
    extract_connections,
    extract_formulas,
    parse_pivot_caches,
    sanitize_table_name,
    scout_formula_sheets,
)


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


def test_sanitize_punctuation_only_sheet_name() -> None:
    digest = hashlib.sha256(b"!!!").hexdigest()[:8]
    assert sanitize_table_name("!!!") == f"sheet_{digest}"
    assert sanitize_table_name("Summary") == "Summary"


def test_extract_formulas_and_shared_strings(fixture_xlsx: Path) -> None:
    with zipfile.ZipFile(fixture_xlsx) as zipf:
        formulas = extract_formulas(zipf)
    texts = {row[2] for row in formulas}
    cells = {row[1] for row in formulas}
    assert "C1" in cells
    assert "B1+B2" in texts
    assert any(text is not None and "B2+1" in text for text in texts)


def test_scout_omits_sheets_past_cap(tmp_path: Path) -> None:
    path = tmp_path / "many.xlsx"
    with zipfile.ZipFile(path, "w") as zipf:
        zipf.writestr("xl/worksheets/sheet1.xml", "<worksheet><sheetData/></worksheet>")
        zipf.writestr("xl/worksheets/sheet2.xml", "<worksheet><f>A1</f></worksheet>")
        zipf.writestr("xl/worksheets/sheet10.xml", "<worksheet><f>B1</f></worksheet>")
    with zipfile.ZipFile(path) as zipf:
        scout = scout_formula_sheets(zipf, max_sheets=1)
    assert scout.hits == []
    assert scout.omitted_by_cap == ["xl/worksheets/sheet2.xml", "xl/worksheets/sheet10.xml"]


def test_scout_forces_parse_on_error(fixture_xlsx: Path) -> None:
    with zipfile.ZipFile(fixture_xlsx) as zipf:

        def boom(_name: str) -> zipfile.ZipInfo:
            raise RuntimeError("scout window unreadable")

        zipf.getinfo = boom  # type: ignore[method-assign]
        scout = scout_formula_sheets(zipf)
    assert "xl/worksheets/sheet1.xml" in scout.hits
    assert "xl/worksheets/sheet1.xml" in scout.forced_parse


def test_worker_csv_round_trips_into_duckdb(fixture_xlsx: Path, tmp_path: Path) -> None:
    csv_path = tmp_path / "formulas.csv"
    result = _worker_formula_task(
        str(fixture_xlsx),
        ["xl/worksheets/sheet1.xml"],
        str(csv_path),
        10,
        10,
    )
    assert result.error is None
    assert result.count == 2
    con = duckdb.connect(str(tmp_path / "check.duckdb"))
    try:
        con.execute("CREATE TABLE formulas (sheet_file TEXT, cell TEXT, formula TEXT)")
        csv_sql = str(csv_path).replace("\\", "/")
        con.execute(f"COPY formulas FROM '{csv_sql}' (FORMAT CSV, HEADER true, DELIMITER ',', QUOTE '\"', ESCAPE '\"')")
        rows = con.execute("SELECT cell FROM formulas ORDER BY cell").fetchall()
        assert [row[0] for row in rows] == ["C1", "C2"]
    finally:
        con.close()


def test_worker_missing_workbook_is_error(tmp_path: Path) -> None:
    csv_path = tmp_path / "formulas.csv"
    result = _worker_formula_task(
        str(tmp_path / "missing.xlsx"),
        ["xl/worksheets/sheet1.xml"],
        str(csv_path),
        10,
        10,
    )
    assert result.error is not None
    assert result.count == 0
    assert not csv_path.exists()


def _run_cli(monkeypatch: pytest.MonkeyPatch, argv: list[str]) -> None:
    from xlsx_autopsy.reconstruct import main

    monkeypatch.setattr("sys.argv", argv)
    main()


def test_cli_extracts_formulas_sst_and_punct_sheet(
    fixture_xlsx: Path,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    out = tmp_path / "out"
    _run_cli(
        monkeypatch,
        ["xlsx-autopsy", "--excel", str(fixture_xlsx), "-o", str(out)],
    )
    blueprint = json.loads((out / "report_blueprint.json").read_text(encoding="utf-8"))
    names = {sheet["name"] for sheet in blueprint["processed_sheets"]}
    tables = {sheet.get("table") for sheet in blueprint["processed_sheets"]}
    assert "Summary" in names
    assert "!!!" in names
    digest = hashlib.sha256(b"!!!").hexdigest()[:8]
    assert f"sheet_{digest}" in tables

    con = duckdb.connect(str(out / "reconstruction.duckdb"), read_only=True)
    try:
        formula_rows = con.execute("SELECT cell, formula FROM formulas ORDER BY cell").fetchall()
        cells = {row[0] for row in formula_rows}
        texts = {row[1] for row in formula_rows}
        assert len(formula_rows) == 2
        assert "C1" in cells
        assert "B1+B2" in texts
        assert any("B2+1" in str(text) for text in texts)
        sst = {row[0] for row in con.execute("SELECT s_full FROM sst").fetchall()}
        assert {"premium", "exposure"} <= sst
    finally:
        con.close()


def test_cli_wipes_prior_outputs(
    fixture_xlsx: Path,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    out = tmp_path / "out"
    _run_cli(
        monkeypatch,
        ["xlsx-autopsy", "--excel", str(fixture_xlsx), "-o", str(out), "--skip-formulas"],
    )
    poison = out / "parquet" / "poison.parquet"
    poison.write_bytes(b"stale")
    second = write_fixture_xlsx(tmp_path / "second.xlsx")
    _run_cli(
        monkeypatch,
        ["xlsx-autopsy", "--excel", str(second), "-o", str(out), "--skip-formulas"],
    )
    assert not poison.exists()
    assert (out / "reconstruction.duckdb").is_file()


def test_cli_missing_workbook_exits_1(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        "sys.argv",
        ["xlsx-autopsy", "--excel", str(tmp_path / "nope.xlsx"), "-o", str(tmp_path / "out")],
    )
    with pytest.raises(SystemExit) as excinfo:
        from xlsx_autopsy.reconstruct import main

        main()
    assert excinfo.value.code == 1


def test_cli_corrupt_zip_exits_1(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    bad = tmp_path / "bad.xlsx"
    bad.write_text("not a zip", encoding="utf-8")
    monkeypatch.setattr(
        "sys.argv",
        ["xlsx-autopsy", "--excel", str(bad), "-o", str(tmp_path / "out")],
    )
    with pytest.raises(SystemExit) as excinfo:
        from xlsx_autopsy.reconstruct import main

        main()
    assert excinfo.value.code == 1
