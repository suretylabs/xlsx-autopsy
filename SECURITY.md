# Security

Default extract **redacts** OLEDB/ODBC credentials, including brace-wrapped and
quoted values. `--include-connection-secrets` is opt-in and writes raw strings
to disk. Do not commit that output.

## Report a vulnerability

Use [GitHub Security Advisories](https://github.com/suretylabs/xlsx-autopsy/security/advisories/new)
on this repository.

Do **not** open a public issue with a real workbook, a connection string, a
password, or a screenshot of either.

## What this tool is not

xlsx-autopsy is a local forensics CLI. It does not phone home. It does not
upload your workbook. It does not claim to be a malware sandbox.
