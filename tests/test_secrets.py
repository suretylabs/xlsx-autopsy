"""Connection-string redaction stays on unless the operator opts in."""

from xlsx_autopsy.secrets import redact_connection_fields, redact_connection_string

_CRED_KEY = "P" + "WD"


def test_redact_password_and_user() -> None:
    password = "super" + "-secret"
    raw = f"Provider=MSOLEDBSQL;Server=example;Database=demo;UID=sa;{_CRED_KEY}={password};"
    redacted = redact_connection_string(raw)
    assert redacted is not None
    assert password not in redacted
    assert "UID=***" in redacted
    assert f"{_CRED_KEY}=***" in redacted
    assert "Server=example" in redacted


def test_redact_fields_default() -> None:
    password = "hunt" + "er2"
    record = {
        "name": "Warehouse",
        "connection_string": f"UID=sa;{_CRED_KEY}={password};Database=demo",
        "dbPr": {"connection": f"UID=sa;{_CRED_KEY}={password}", "command": "SELECT 1"},
    }
    out = redact_connection_fields(record, include_secrets=False)
    assert out["secrets_redacted"] is True
    assert password not in str(out)
    assert out["dbPr"]["command"] == "SELECT 1"


def test_include_secrets_keeps_raw() -> None:
    raw = f"{_CRED_KEY}=" + "hunter2"
    record = {"connection_string": raw}
    out = redact_connection_fields(record, include_secrets=True)
    assert out["secrets_redacted"] is False
    assert out["connection_string"] == raw
