"""Redact credentials that hide in Excel connection XML.

Default extract must not write passwords, user ids, or cloud keys into
``report_blueprint.json``. Opt in with ``--include-connection-secrets``.
"""

from __future__ import annotations

import re
from typing import Any

_SECRET_KEY_RE = re.compile(
    r"(?i)\b("
    r"Password|Pwd|PWD|"
    r"User ID|UID|UserName|Username|User|"
    r"AccountKey|SharedAccessKey|SharedAccessSignature|"
    r"Secret|Token|AccessToken|ApiKey|API Key|Key"
    r")\s*=\s*([^;]*)"
)

_REDACTED = "***"


def redact_connection_string(value: str | None) -> str | None:
    """Mask password/user/key pairs inside an OLEDB/ODBC connection string."""
    if value is None:
        return None
    return _SECRET_KEY_RE.sub(lambda match: f"{match.group(1)}={_REDACTED}", value)


def redact_connection_fields(record: dict[str, Any], *, include_secrets: bool) -> dict[str, Any]:
    """Return a copy of a connection record with secrets stripped unless opted in."""
    out = dict(record)
    if include_secrets:
        out["secrets_redacted"] = False
        return out

    out["connection_string"] = redact_connection_string(out.get("connection_string"))
    db_pr = out.get("dbPr")
    if isinstance(db_pr, dict):
        nested = dict(db_pr)
        if "connection" in nested:
            nested["connection"] = redact_connection_string(nested.get("connection"))
        out["dbPr"] = nested
    out["secrets_redacted"] = True
    return out
