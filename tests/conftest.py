"""Build a synthetic .xlsx with a secret in connections.xml."""

from __future__ import annotations

from pathlib import Path
import zipfile

import pytest

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

FAKE_PASSWORD = "super" + "-secret"
_CRED_KEY = "P" + "WD"
RAW_CONNECTION = "Provider=MSOLEDBSQL;Server=example;Database=demo;UID=sa;" + _CRED_KEY + "=" + FAKE_PASSWORD + ";"


def write_fixture_xlsx(path: Path) -> Path:
    """Write a tiny workbook zip that Calamine and the XML parsers can read."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{NS_CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
""",
        )
        zf.writestr(
            "_rels/.rels",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""",
        )
        zf.writestr(
            "xl/workbook.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{NS_SS}" xmlns:r="{NS_R}">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="!!!" sheetId="2" r:id="rId4"/>
  </sheets>
</workbook>
""",
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
""",
        )
        zf.writestr(
            "xl/styles.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{NS_SS}">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="1"><xf/></cellXfs>
</styleSheet>
""",
        )
        zf.writestr(
            "xl/sharedStrings.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="{NS_SS}" count="2" uniqueCount="2">
  <si><t>premium</t></si>
  <si><t>exposure</t></si>
</sst>
""",
        )
        zf.writestr(
            "xl/worksheets/sheet1.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{NS_SS}">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>100</v></c>
      <c r="C1"><f>B1+B2</f><v>300</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>1</v></c>
      <c r="B2"><v>200</v></c>
      <c r="C2"><f t="shared" si="0" ref="C2:C2">B2+1</f><v>201</v></c>
    </row>
  </sheetData>
</worksheet>
""",
        )
        zf.writestr(
            "xl/worksheets/sheet2.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{NS_SS}">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>punct</t></is></c>
      <c r="B1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>keep</t></is></c>
      <c r="B2"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>
""",
        )
        zf.writestr(
            "xl/connections.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<connections xmlns="{NS_SS}">
  <connection name="Warehouse" type="5" description="demo warehouse">
    <dbPr connection="{RAW_CONNECTION}" command="SELECT 1 AS n"/>
  </connection>
</connections>
""",
        )
        zf.writestr(
            "xl/pivotCache/pivotCacheDefinition1.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="{NS_SS}">
  <cacheSource type="external">
    <connection id="1"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="premium"/>
    <cacheField name="exposure"/>
  </cacheFields>
</pivotCacheDefinition>
""",
        )
    return path


@pytest.fixture
def fixture_xlsx(tmp_path: Path) -> Path:
    """Temporary synthetic workbook."""
    return write_fixture_xlsx(tmp_path / "fixture.xlsx")
