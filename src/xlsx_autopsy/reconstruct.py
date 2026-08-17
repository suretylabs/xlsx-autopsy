r"""xlsx-autopsy: decompose a huge Excel workbook without opening it.

Excel is treated as a report-definition zip, not a data source. The useful
model usually lives in pivot caches, connections, and formulas — not the
rendered grid.

Extracts:
- Sheet values (via Calamine)
- Formulas (shared formulas are deduped)
- Pivot tables + cache field names + cache sources
- Named ranges
- Database connections (connection strings redacted by default)

Outputs:
- report_blueprint.json
- reconstruction.duckdb
- parquet/ per sheet

Examples:
    uv run xlsx-autopsy --excel workbook.xlsx
    uv run xlsx-autopsy --excel workbook.xlsx -o out --skip-formulas
    uv run python -m xlsx_autopsy --excel workbook.xlsx
"""

from __future__ import annotations

import argparse
from concurrent.futures import ProcessPoolExecutor
import csv
from dataclasses import dataclass
import hashlib
import json
import logging
import math
import multiprocessing
import os
from pathlib import Path
import re
import shutil
import time
import tomllib
import traceback
from typing import Any, TypedDict, cast
import xml.etree.ElementTree as ElementTree
import zipfile

import duckdb
import polars as pl
from rich.console import Console
from rich.progress import BarColumn, Progress, SpinnerColumn, TextColumn, TimeElapsedColumn

from xlsx_autopsy.secrets import redact_connection_fields

try:
    from lxml import etree as et  # type: ignore[import-untyped] # lxml stubs incomplete
except ModuleNotFoundError:
    et = ElementTree

type XmlElement = Any

DEFAULT_EXCEL_PATH = Path("workbook.xlsx")

console = Console()
logger = logging.getLogger("xlsx_autopsy")


# -----------------------------------------------------------------------------
# Type Definitions
# -----------------------------------------------------------------------------
class SheetInfo(TypedDict):
    """Sheet metadata from workbook.xml."""

    name: str | None
    sheetId: str | None
    rId: str | None


class DefinedName(TypedDict):
    """Named range definition."""

    name: str | None
    text: str | None


class PivotCache(TypedDict):
    """Pivot cache mapping."""

    cacheId: str | None
    rId: str | None


class WorkbookMeta(TypedDict):
    """Workbook metadata structure."""

    sheets: list[SheetInfo]
    defined_names: list[DefinedName]
    calc_props: dict[str, Any]
    pivot_caches: list[PivotCache]


class ProcessedSheet(TypedDict, total=False):
    """Information about a processed sheet."""

    name: str
    table: str
    rows: int
    cols: int
    status: str


class FailedSheet(TypedDict):
    """Information about a failed sheet."""

    name: str
    error: str


class AnalysisReport(TypedDict):
    """Complete analysis report structure."""

    workbook_path: str
    metadata: WorkbookMeta | dict[str, Any]
    connections: list[dict[str, Any]]
    pivot_caches: list[dict[str, Any]]
    pivot_tables: list[dict[str, Any]]
    pivot_cache_defs: dict[str, Any]
    pivot_tables_resolved: list[dict[str, Any]]
    package_inventory: dict[str, Any]
    processed_sheets: list[ProcessedSheet]
    failed_sheets: list[FailedSheet]
    formula_warnings: list[str]


@dataclass(frozen=True)
class FormulaScoutResult:
    """Sheets the formula extractor should parse, plus what the scout skipped."""

    hits: list[str]
    omitted_by_cap: list[str]
    forced_parse: list[str]


@dataclass(frozen=True)
class FormulaWorkerResult:
    """Outcome of one formula-extract worker process."""

    count: int
    csv_path: str
    error: str | None = None


# -----------------------------------------------------------------------------
# Utils & Config
# -----------------------------------------------------------------------------
def load_toml_config(config_path: Path) -> dict[str, Any]:
    """Load configuration from a TOML file.

    Args:
        config_path: Path to the TOML configuration file.

    Returns:
        Dictionary containing the parsed TOML data, or empty dict if file doesn't exist.
    """
    if not config_path.exists():
        return {}
    with config_path.open("rb") as f:
        return tomllib.load(f)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments for the Excel forensics script.

    Returns:
        Parsed command-line arguments.
    """
    parser = argparse.ArgumentParser(
        prog="xlsx-autopsy",
        description="Decompose a huge Excel workbook without opening it.",
    )
    parser.add_argument("--excel", type=Path, help="Input .xlsx path")
    parser.add_argument("--input", "-i", type=Path, help="Alias for --excel")
    parser.add_argument("--output", "-o", type=Path, help="Output directory")
    parser.add_argument(
        "--keep-outputs",
        action="store_true",
        help="Do not wipe prior DuckDB/parquet/blueprint in the output directory.",
    )
    parser.add_argument(
        "--skip-formulas",
        action="store_true",
        help="Skip cell-formula extract. Faster when you only need pivots and values.",
    )
    parser.add_argument(
        "--include-connection-secrets",
        action="store_true",
        help="Keep raw OLEDB/ODBC connection strings. Default redacts passwords and user ids.",
    )
    parser.add_argument("--config", default=None, help="Optional TOML config (xlsx-autopsy.toml)")
    return parser.parse_args()


def get_config(args: argparse.Namespace) -> dict[str, Any]:
    """Merge CLI, Config, and Env Vars.

    Args:
        args: Parsed command-line arguments.

    Returns:
        Configuration dictionary with resolved paths and settings.
    """
    raw = load_toml_config(Path(args.config)) if args.config else {}
    section = raw.get("xlsx_autopsy", {})

    def _resolve(arg_val: Any, env_key: str, toml_key: str, default: Any) -> Any:
        """Resolve a config value using CLI, env, TOML, then default precedence."""
        if arg_val is not None and arg_val != "":
            return arg_val
        env_val = os.getenv(env_key)
        if env_val is not None and env_val != "":
            return env_val
        if toml_key in section:
            return section.get(toml_key)
        return default

    input_path = args.excel if args.excel else args.input

    sst_trunc_raw = _resolve(None, "XLSX_AUTOPSY_SST_TRUNCATE", "sst_truncate", 5000)
    try:
        sst_truncate = int(sst_trunc_raw)
    except (TypeError, ValueError):
        sst_truncate = 5000
    if sst_truncate < 0:
        sst_truncate = 0

    return {
        "workbook_path": Path(_resolve(input_path, "XLSX_AUTOPSY_WORKBOOK", "workbook_path", str(DEFAULT_EXCEL_PATH))),
        "output_dir": Path(_resolve(args.output, "XLSX_AUTOPSY_OUTPUT", "output_dir", "out")),
        "db_name": _resolve(None, "XLSX_AUTOPSY_DB", "db_name", "reconstruction.duckdb"),
        "parquet_dir": _resolve(None, "XLSX_AUTOPSY_PARQUET", "parquet_dir", "parquet"),
        "sst_truncate": sst_truncate,
        "skip_formulas": args.skip_formulas,
        "include_connection_secrets": bool(args.include_connection_secrets),
        "keep_outputs": bool(args.keep_outputs),
        "formula_extract": section.get("formula_extract") or {},
    }


def cleanup_outputs(output_dir: Path, db_path: Path, parquet_dir: Path) -> None:
    """Clean up output files and directories before a fresh run.

    Args:
        output_dir: Base output directory.
        db_path: Path to the DuckDB database file.
        parquet_dir: Directory containing Parquet files.

    Returns:
        None
    """
    if db_path.exists():
        db_path.unlink()
    for p in ["report_blueprint.json"]:
        fp = output_dir / p
        if fp.exists():
            fp.unlink()
    if parquet_dir.exists():
        shutil.rmtree(parquet_dir)
    parquet_dir.mkdir(parents=True, exist_ok=True)


def sanitize_table_name(name: str) -> str:
    """Sanitize a sheet name to a non-empty DuckDB table name.

    Excel allows names that are punctuation-only (``!!!``). Those must not
    collapse to an empty identifier.

    Args:
        name: Original sheet name.

    Returns:
        Alphanumeric/underscore table name. Punctuation-only names become
        ``sheet_<sha256-8>``.
    """
    cleaned = "".join(c if c.isalnum() else "_" for c in name).strip("_")
    if cleaned:
        return cleaned
    digest = hashlib.sha256(name.encode("utf-8")).hexdigest()[:8]
    return f"sheet_{digest}"


def log_status(msg: str, start_time: float | None = None) -> None:
    """Log a status message with optional elapsed time.

    Args:
        msg: Message to log.
        start_time: Optional start time for elapsed time calculation.

    Returns:
        None
    """
    time_str = ""
    if start_time:
        time_str = f"[{time.perf_counter() - start_time:7.1f}s] "
    console.log(f"[dim]{time_str}[/dim]{msg}")
    logger.info(f"{time_str.strip()} {msg}")


# -----------------------------------------------------------------------------
# JSON Sanitizer
# -----------------------------------------------------------------------------
def _json_sanitize(obj: Any) -> Any:
    """Recursively sanitize objects for JSON export.

    Args:
        obj: Object to sanitize.

    Returns:
        Sanitized object suitable for JSON serialization.
    """
    if obj is None or isinstance(obj, (bool, int, str)):
        return obj
    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return str(obj)
        return obj
    if isinstance(obj, dict):
        return {str(k): _json_sanitize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [_json_sanitize(x) for x in obj]
    if isinstance(obj, Path):
        return str(obj)
    # Handle lxml attributes specifically
    if hasattr(obj, "items"):
        return {str(k): str(v) for k, v in dict(obj).items()}
    return str(obj)


def write_outputs(
    analysis: AnalysisReport | dict[str, Any], output_dir: Path, db_path: Path, parquet_dir: Path
) -> None:
    """Write reconstruction blueprint results to JSON and display output paths.

    Args:
        analysis: Analysis dictionary to write to JSON.
        output_dir: Output directory.
        db_path: Path to DuckDB database.
        parquet_dir: Directory containing Parquet files.

    Returns:
        None
    """
    # Primary output
    json_path = output_dir / "report_blueprint.json"
    with json_path.open("w", encoding="utf-8") as f:
        safe = _json_sanitize(analysis)
        json.dump(safe, f, indent=2, ensure_ascii=False)

    console.print(f"   [green]JSON:[/green]    {json_path}")
    console.print(f"   [green]DuckDB:[/green]  {db_path}")
    console.print(f"   [green]Parquet:[/green] {parquet_dir}")


# -----------------------------------------------------------------------------
# XML Reconstruction (The Heavy Lifting)
# -----------------------------------------------------------------------------
def safe_open(zipf: zipfile.ZipFile, member: str) -> zipfile.ZipExtFile | None:
    """Safely open a member from a ZIP file, returning None if not found.

    Args:
        zipf: ZIP file object.
        member: Member name to open.

    Returns:
        File-like object or None if member not found.
    """
    try:
        # zipfile.ZipFile.open() returns IO[bytes] at type-check time but ZipExtFile at runtime
        return cast(zipfile.ZipExtFile, zipf.open(member))
    except KeyError:
        return None


def strip_ns_inplace(elem: XmlElement) -> None:
    """Strip XML namespaces in-place to simplify .find()/.xpath() queries."""
    for e in elem.iter():
        if isinstance(e.tag, str) and "}" in e.tag:
            e.tag = e.tag.split("}", 1)[1]


def parse_workbook_rels(zipf: zipfile.ZipFile) -> dict[str, str]:
    """Parse xl/_rels/workbook.xml.rels to map rId -> xl/* target."""
    rels: dict[str, str] = {}
    fh = safe_open(zipf, "xl/_rels/workbook.xml.rels")
    if not fh:
        return rels
    try:
        tree = et.parse(fh)
        root = tree.getroot()
        strip_ns_inplace(root)
        for r in root.findall("Relationship"):
            rid = r.get("Id")
            target = r.get("Target")
            if not rid or not target:
                continue
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = f"xl/{target.lstrip('./')}"
            rels[rid] = target
    except Exception:  # noqa: BLE001
        logger.debug("Failed to parse workbook relationships")
    return rels


def extract_connections(
    zipf: zipfile.ZipFile,
    *,
    include_secrets: bool = False,
) -> list[dict[str, Any]]:
    """Extract database connections from Excel connections.xml.

    Connection strings are redacted unless ``include_secrets`` is True.

    Args:
        zipf: ZIP file object of the Excel workbook.
        include_secrets: Keep raw OLEDB/ODBC credentials. Default False.

    Returns:
        List of connection dictionaries with name, type, description, SQL, and connection string.
    """
    connections = []
    fh = safe_open(zipf, "xl/connections.xml")
    if not fh:
        return []

    try:
        tree = et.parse(fh)
        root = tree.getroot()
        strip_ns_inplace(root)

        for conn in root.findall("connection"):
            c_info = {
                "name": conn.get("name"),
                "type": conn.get("type"),
                "desc": conn.get("description"),
                "sql_command": None,
                "connection_string": None,
            }
            for tag in ["dbPr", "oledbPr"]:
                pr = conn.find(tag)
                if pr is not None:
                    c_info["sql_command"] = pr.get("command")
                    c_info["connection_string"] = pr.get("connection")
            connections.append(redact_connection_fields(c_info, include_secrets=include_secrets))
    except Exception:  # noqa: BLE001
        logger.debug("Failed to parse connections.xml")
    return connections


def parse_workbook_meta(zipf: zipfile.ZipFile) -> WorkbookMeta:
    """Extract sheets and global definitions from workbook.xml.

    Args:
        zipf: ZIP file object of the Excel workbook.

    Returns:
        Dictionary containing sheets, defined names, and calculation properties.
    """
    data: WorkbookMeta = {"sheets": [], "defined_names": [], "calc_props": {}, "pivot_caches": []}
    fh = safe_open(zipf, "xl/workbook.xml")
    if not fh:
        return data

    tree = et.parse(fh)
    root = tree.getroot()
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    # Sheets
    sheets_el = root.find(f"{ns}sheets")
    if sheets_el is not None:
        for s in sheets_el.findall(f"{ns}sheet"):
            sheet_info: SheetInfo = {
                "name": s.attrib.get("name"),
                "sheetId": s.attrib.get("sheetId"),
                "rId": s.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"),
            }
            data["sheets"].append(sheet_info)

    # Defined Names (Named Ranges)
    dn_el = root.find(f"{ns}definedNames")
    if dn_el is not None:
        for dn in dn_el.findall(f"{ns}definedName"):
            defined_name: DefinedName = {"name": dn.attrib.get("name"), "text": dn.text}
            data["defined_names"].append(defined_name)

    # Pivot cache mapping: cacheId -> rId (to resolve pivotCacheDefinition parts via workbook rels)
    pc_el = root.find(f"{ns}pivotCaches")
    if pc_el is not None:
        for pc in pc_el.findall(f"{ns}pivotCache"):
            pivot_cache: PivotCache = {
                "cacheId": pc.attrib.get("cacheId"),
                "rId": pc.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"),
            }
            data["pivot_caches"].append(pivot_cache)

    return data


def parse_pivot_caches(zipf: zipfile.ZipFile) -> list[dict[str, Any]]:
    """Extract pivot cache definitions from Excel pivot cache files.

    Args:
        zipf: ZIP file object of the Excel workbook.

    Returns:
        List of pivot cache definition dictionaries with path, type, and SQL snippets.
    """
    defs = []
    members = [n for n in zipf.namelist() if n.startswith("xl/pivotCache/pivotCacheDefinition")]
    for member in members:
        try:
            tree = et.parse(zipf.open(member))
            root = tree.getroot()
            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            cache_source = root.find(f"{ns}cacheSource")
            info = {"path": member, "type": None, "sql": None}
            if cache_source is not None:
                info["type"] = cache_source.attrib.get("type")
                # Look for embedded SQL/MDX in the cache source
                info["snippet"] = et.tostring(cache_source, encoding="unicode")[:500]
            defs.append(info)
        except Exception:  # noqa: BLE001
            logger.debug("Failed to parse pivot cache definition: %s", member)
    return defs


def parse_pivot_cache_definitions(
    zipf: zipfile.ZipFile,
    workbook_meta: WorkbookMeta | dict[str, Any],
) -> dict[str, dict[str, Any]]:
    """Resolve pivot cache definitions keyed by cacheId.

    Why this matters:
    - Pivot tables store field references by index.
    - Pivot cache definitions contain the ordered cacheFields list (field names),
      and the cacheSource tells you where the pivot data comes from.

    Args:
        zipf: ZIP file object.
        workbook_meta: Workbook metadata containing pivot cache mappings.

    Returns:
        Dictionary mapping cache IDs to their definitions.
    """
    rels = parse_workbook_rels(zipf)
    out: dict[str, dict[str, Any]] = {}

    for pc in workbook_meta.get("pivot_caches") or []:
        cache_id = pc.get("cacheId")
        rid = pc.get("rId")
        if not cache_id or not rid:
            continue
        target = rels.get(rid)
        if not target:
            continue

        fh = safe_open(zipf, target)
        if not fh:
            continue

        try:
            tree = et.parse(fh)
            root = tree.getroot()
            strip_ns_inplace(root)

            info: dict[str, Any] = {
                "cacheId": str(cache_id),
                "definition_path": target,
                "source_type": None,
                "worksheet_source": None,
                "external_source": None,
                "cache_fields": [],
            }

            cache_source = root.find("cacheSource")
            if cache_source is not None:
                info["source_type"] = cache_source.get("type")
                ws_src = cache_source.find("worksheetSource")
                if ws_src is not None:
                    info["worksheet_source"] = {
                        "sheet": ws_src.get("sheet"),
                        "ref": ws_src.get("ref"),
                        "name": ws_src.get("name"),
                    }
                ext_src = cache_source.find("externalSource")
                if ext_src is not None:
                    info["external_source"] = {"connectionId": ext_src.get("connectionId")}

            cache_fields_el = root.find("cacheFields")
            if cache_fields_el is not None:
                info["cache_fields"] = [cf.get("name") for cf in cache_fields_el.findall("cacheField")]

            out[str(cache_id)] = info
        except Exception:
            continue

    return out


def parse_detailed_pivots(zipf: zipfile.ZipFile) -> list[dict[str, Any]]:
    """Extract detailed Pivot Table definitions (Rows, Cols, Data fields).

    Args:
        zipf: ZIP file object of the Excel workbook.

    Returns:
        List of pivot table definition dictionaries containing path, name,
        cache ID, and field lists for rows, columns, pages, and data.
    """
    pivots = []
    members = [n for n in zipf.namelist() if n.startswith("xl/pivotTables/pivotTable")]
    for member in members:
        try:
            # Many pivotTable*.xml files are large; namespaces also vary. Strip namespaces to simplify.
            root = et.fromstring(zipf.read(member))
            strip_ns_inplace(root)

            # Basic identifiers
            pt_name = root.get("name")
            cache_id = root.get("cacheId")

            # Field references by index
            row_fields: list[str] = []
            col_fields: list[str] = []
            page_fields: list[str] = []
            data_fields: list[dict[str, Any]] = []

            rf = root.find("rowFields")
            if rf is not None:
                row_fields = _collect_attr_values(rf, "field", "x")

            cf = root.find("colFields")
            if cf is not None:
                col_fields = _collect_attr_values(cf, "field", "x")

            # pageFields uses <pageField fld="...">, not <field x="...">
            pf = root.find("pageFields")
            if pf is not None:
                page_fields = _collect_attr_values(pf, "pageField", "fld")

            # dataFields uses <dataField fld="..." subtotal="sum|count|...">
            df = root.find("dataFields")
            if df is not None:
                for d in df.findall("dataField"):
                    data_fields.append(
                        {
                            "fld": d.get("fld"),
                            "subtotal": d.get("subtotal"),
                            "name": d.get("name"),
                            "showDataAs": d.get("showDataAs"),
                            "baseField": d.get("baseField"),
                            "baseItem": d.get("baseItem"),
                        }
                    )

            # Pivot field properties (axis placement is often encoded here)
            pivot_fields: list[dict[str, Any]] = []
            pfs = root.find("pivotFields")
            if pfs is not None:
                for pf_i, pfi in enumerate(pfs.findall("pivotField")):
                    pivot_fields.append(
                        {
                            "index": pf_i,
                            "axis": pfi.get("axis"),
                            "dataField": pfi.get("dataField"),
                            "showAll": pfi.get("showAll"),
                            "defaultSubtotal": pfi.get("defaultSubtotal"),
                            "sortType": pfi.get("sortType"),
                            "outline": pfi.get("outline"),
                            "compact": pfi.get("compact"),
                            "subtotalTop": pfi.get("subtotalTop"),
                        }
                    )

            pivots.append(
                {
                    "path": member,
                    "name": pt_name,
                    "cacheId": cache_id,
                    "row_fields": row_fields,
                    "col_fields": col_fields,
                    "page_fields": page_fields,
                    "data_fields": data_fields,
                    "pivot_fields": pivot_fields,
                }
            )
        except Exception:  # noqa: BLE001
            logger.debug("Failed to parse pivot table: %s", member)
    return pivots


def resolve_pivot_tables(
    pivot_tables: list[dict[str, Any]],
    cache_defs_by_id: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    """Resolve pivot field indexes to cache field names."""
    resolved: list[dict[str, Any]] = []

    def _idx_to_name(cache_fields: list[Any], idx: Any) -> str | None:
        """Resolve a pivot field index into a cache-field name."""
        if idx is None:
            return None
        try:
            i = int(str(idx))
        except Exception:
            return None
        if i < 0 or i >= len(cache_fields):
            return None
        return cache_fields[i]

    for pt in pivot_tables or []:
        cache_id = pt.get("cacheId")
        cache_def = cache_defs_by_id.get(str(cache_id)) if cache_id is not None else None
        cache_fields = (cache_def or {}).get("cache_fields") or []

        out = dict(pt)
        out["cache_definition"] = cache_def

        out["row_field_names"] = [_idx_to_name(cache_fields, x) for x in (pt.get("row_fields") or [])]
        out["col_field_names"] = [_idx_to_name(cache_fields, x) for x in (pt.get("col_fields") or [])]
        out["page_field_names"] = [_idx_to_name(cache_fields, x) for x in (pt.get("page_fields") or [])]

        # data_fields is list[dict], each has fld which is the field index
        df_res: list[dict[str, Any]] = []
        for d in pt.get("data_fields") or []:
            d2 = dict(d)
            d2["field_name"] = _idx_to_name(cache_fields, d.get("fld"))
            df_res.append(d2)
        out["data_field_names"] = df_res

        resolved.append(out)
    return resolved


def list_excel_package_inventory(zipf: zipfile.ZipFile) -> dict[str, list[str]]:
    """Inventory pivot/chart/slicer-related parts for debugging/report lineage."""
    inv: dict[str, list[str]] = {
        "pivotTables": [],
        "pivotCache": [],
        "slicerCaches": [],
        "timelineCaches": [],
        "charts": [],
        "drawings": [],
    }
    for n in zipf.namelist():
        if n.startswith("xl/pivotTables/"):
            inv["pivotTables"].append(n)
        elif n.startswith("xl/pivotCache/"):
            inv["pivotCache"].append(n)
        elif n.startswith("xl/slicerCaches/"):
            inv["slicerCaches"].append(n)
        elif n.startswith("xl/timelineCaches/"):
            inv["timelineCaches"].append(n)
        elif n.startswith("xl/charts/"):
            inv["charts"].append(n)
        elif n.startswith("xl/drawings/"):
            inv["drawings"].append(n)
    for k in inv:
        inv[k] = sorted(inv[k])
    return inv


def _collect_attr_values(parent: XmlElement, child_tag: str, attr_name: str) -> list[str]:
    """Collect non-null attribute values from XML children.

    Args:
        parent: Parent XML element containing matching child nodes.
        child_tag: Child tag name to search for.
        attr_name: Attribute name to extract from each matching child.

    Returns:
        Ordered list of non-null attribute values.
    """
    values: list[str] = []
    for child in parent.findall(child_tag):
        value = child.get(attr_name)
        if value is not None:
            values.append(value)
    return values


def _tag_endswith(elem: XmlElement, local: str) -> bool:
    """Check if an XML element's tag matches a local name, ignoring namespaces.

    Args:
        elem: The XML element to check.
        local: The local tag name to match (e.g., 'c' for cell).

    Returns:
        True if the tag matches the local name, False otherwise.
    """
    t = elem.tag
    return isinstance(t, str) and (t == local or t.endswith("}" + local))


def _format_formula_with_meta(elem: XmlElement, txt: str) -> str:
    """Prefix formula text with metadata extracted from the formula element."""
    meta = []
    if elem.get("t"):
        meta.append(f"t={elem.get('t')}")
    if elem.get("si"):
        meta.append(f"si={elem.get('si')}")
    if elem.get("ref"):
        meta.append(f"ref={elem.get('ref')}")
    return f"[{','.join(meta)}] {txt}" if meta else txt


# Fast pre-scan to avoid iterparse on huge sheets that contain no formulas.
# This is a cheap bytes regex search for real <f> tags.
_FORMULA_TAG_RE_BYTES = re.compile(rb"<(?:[A-Za-z0-9_]+:)?f(?:\s|>)")


def _sheet_part_sort_key(name: str) -> tuple[int, str]:
    """Sort ``sheet10.xml`` after ``sheet2.xml`` instead of lexicographically."""
    match = re.search(r"sheet(\d+)\.xml$", name)
    return (int(match.group(1)), name) if match else (10**9, name)


def _worksheet_members(zipf: zipfile.ZipFile, max_sheets: int) -> tuple[list[str], list[str]]:
    """Return worksheet parts up to ``max_sheets``, plus the omitted tail."""
    members = [n for n in zipf.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
    members.sort(key=_sheet_part_sort_key)
    return members[:max_sheets], members[max_sheets:]


def scout_formula_sheets(
    zipf: zipfile.ZipFile,
    *,
    max_sheets: int = 300,
    max_bytes_per_sheet: int = 25_000_000,
) -> FormulaScoutResult:
    """Return worksheet parts that should be parsed for formulas.

    Scout is an optimization, not a filter of last resort. A sheet is included
    when a ``<f>`` tag appears in the scout window, when the part is larger than
    the window (the tag may live past it), or when the scout itself errors.
    Sheets past ``max_sheets`` are recorded, not silently dropped.
    """
    selected, omitted = _worksheet_members(zipf, max_sheets)
    hits: list[str] = []
    forced_parse: list[str] = []
    for name in selected:
        try:
            info = zipf.getinfo(name)
            with zipf.open(name) as handle:
                head = handle.read(max_bytes_per_sheet)
            if _FORMULA_TAG_RE_BYTES.search(head):
                hits.append(name)
            elif info.file_size > max_bytes_per_sheet:
                hits.append(name)
                forced_parse.append(name)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Formula scout failed for %s (%s); forcing parse", name, exc)
            hits.append(name)
            forced_parse.append(name)
    return FormulaScoutResult(hits=hits, omitted_by_cap=omitted, forced_parse=forced_parse)


def extract_formulas(
    zipf: zipfile.ZipFile,
    config: dict[str, Any] | None = None,
) -> list[tuple[str, str | None, str | None]]:
    """Stream formulas from worksheet parts using iterparse, with safety caps.

    We do a cheap bytes pre-scan to avoid parsing huge sheets that contain no formulas.
    Shared formulas are deduped so we do not explode output size.

    Returns:
        List of tuples: (sheet_file, cell_ref, formula_text)
    """
    cfg = (config or {}).get("formula_extract", {}) if isinstance(config, dict) else {}

    max_sheets = int(cfg.get("max_sheets", 300))
    scout_enabled = bool(cfg.get("scout_enabled", True))
    scout_bytes = int(cfg.get("scout_bytes", 25_000_000))

    max_cells_per_sheet = int(cfg.get("max_cells_per_sheet", 1_500_000))
    max_formulas_per_sheet = int(cfg.get("max_formulas_per_sheet", 200_000))

    formulas: list[tuple[str, str | None, str | None]] = []

    if scout_enabled:
        scout = scout_formula_sheets(
            zipf,
            max_sheets=max_sheets,
            max_bytes_per_sheet=scout_bytes,
        )
        sheet_members = scout.hits
        if not sheet_members:
            return formulas
    else:
        sheet_members, _omitted = _worksheet_members(zipf, max_sheets)

    for sheet in sheet_members:
        fh = safe_open(zipf, sheet)
        if not fh:
            continue

        current_cell_ref: str | None = None

        # Shared formulas: Excel repeats <f t="shared" si="..."/> across huge ranges,
        # often with empty text. Capture only the unique shared formulas that include text.
        shared_text_by_si: dict[str, str] = {}

        cell_count = 0
        formula_count = 0

        try:
            # We use start/end to track <c r="A1"> boundaries
            for event, elem in et.iterparse(fh, events=("start", "end")):
                if event == "start" and _tag_endswith(elem, "c"):
                    current_cell_ref = elem.get("r")

                elif event == "end" and _tag_endswith(elem, "f"):
                    txt = (elem.text or "").strip()
                    f_type = (elem.get("t") or "").strip().lower()
                    si = (elem.get("si") or "").strip()

                    # Handle shared formulas efficiently
                    if f_type == "shared" and si:
                        # Most shared formula tags have no text. Skip them.
                        if not txt:
                            elem.clear()
                            continue

                        # Keep only the first text-bearing instance per (sheet, si)
                        key = f"{sheet}::si={si}"
                        if key not in shared_text_by_si:
                            shared_text_by_si[key] = txt
                            formulas.append((sheet, current_cell_ref, _format_formula_with_meta(elem, txt)))
                            formula_count += 1

                        elem.clear()
                        if formula_count >= max_formulas_per_sheet:
                            break
                        continue

                    # Non-shared formulas: keep only if there is actual text
                    if txt:
                        formulas.append((sheet, current_cell_ref, _format_formula_with_meta(elem, txt)))
                        formula_count += 1

                    elem.clear()
                    if formula_count >= max_formulas_per_sheet:
                        break

                elif event == "end" and _tag_endswith(elem, "c"):
                    current_cell_ref = None
                    cell_count += 1
                    elem.clear()
                    if cell_count >= max_cells_per_sheet:
                        break

        except Exception:
            continue

    return formulas


def stream_shared_strings(
    zipf: zipfile.ZipFile,
    conn: duckdb.DuckDBPyConnection,
    output_dir: Path,
    sst_truncate: int,
) -> None:
    """Stream and process shared strings from Excel workbook for efficient storage.

    Parses the sharedStrings.xml file, truncates long strings if necessary,
    and bulk loads the data into a DuckDB table for querying.

    Args:
        zipf: ZIP file object of the Excel workbook.
        conn: DuckDB connection to the database.
        output_dir: Directory to store temporary CSV file.
        sst_truncate: Maximum length for string truncation; 0 for no truncation.

    Returns:
        None
    """
    conn.execute("DROP TABLE IF EXISTS sst")
    conn.execute("""
        CREATE TABLE sst(
            idx BIGINT,
            s_full TEXT,
            s_trunc TEXT,
            s_hash TEXT,
            s_len BIGINT
        );
    """)

    fh = safe_open(zipf, "xl/sharedStrings.xml")
    if not fh:
        return

    sst_csv = output_dir / "temp_sst.csv"

    # 2. Stream XML to CSV (Fast I/O)
    with sst_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        idx = 0
        # iterparse is memory efficient for massive XML files
        for _event, elem in et.iterparse(fh, events=("end",)):
            if elem.tag.endswith("si"):
                texts = [
                    t.text
                    for t in elem.iterfind(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                    if t.text
                ]
                text = "".join(texts)
                length = len(text)

                truncate_enabled = sst_truncate > 0 and length > sst_truncate
                s_trunc = text[:sst_truncate] if truncate_enabled else text
                s_hash = hashlib.sha256(text.encode("utf-8")).hexdigest() if truncate_enabled else None

                writer.writerow([idx, text, s_trunc, s_hash, length])
                idx += 1
                elem.clear()  # Free memory

    # 3. Bulk Load to DuckDB (Zero WAL Thrashing)
    conn.execute(
        f"COPY sst FROM '{str(sst_csv).replace(os.sep, '/')}' (FORMAT CSV, DELIMITER ',', QUOTE '\"', ESCAPE '\"')"
    )
    sst_csv.unlink()


def _worker_formula_task(
    excel_path: str, sheet_members: list[str], output_csv: str, max_formulas: int, max_cells: int
) -> FormulaWorkerResult:
    """Parse a subset of worksheets and write formula rows to CSV.

    On any exception the CSV is discarded so the parent never COPY-loads a
    partial or missing file as if it were success.
    """
    count = 0
    try:
        with zipfile.ZipFile(excel_path, "r") as zipf, open(output_csv, "w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow(["sheet_file", "cell", "formula"])

            for sheet in sheet_members:
                fh = safe_open(zipf, sheet)
                if not fh:
                    continue

                current_cell_ref = None
                sheet_formula_count = 0
                sheet_cell_count = 0
                shared_text_by_si: dict[str, str] = {}

                for event, elem in et.iterparse(fh, events=("start", "end")):
                    if event == "start" and _tag_endswith(elem, "c"):
                        current_cell_ref = elem.get("r")

                    elif event == "end" and _tag_endswith(elem, "f"):
                        txt = (elem.text or "").strip()
                        f_type = (elem.get("t") or "").strip().lower()
                        si = (elem.get("si") or "").strip()

                        val_to_write = None

                        if f_type == "shared" and si:
                            if txt:
                                key = f"{sheet}::si={si}"
                                if key not in shared_text_by_si:
                                    shared_text_by_si[key] = txt
                                    val_to_write = _format_formula_with_meta(elem, txt)
                        elif txt:
                            val_to_write = _format_formula_with_meta(elem, txt)

                        if val_to_write:
                            writer.writerow([sheet, current_cell_ref, val_to_write])
                            count += 1
                            sheet_formula_count += 1

                        elem.clear()
                        if sheet_formula_count >= max_formulas:
                            break

                    elif event == "end" and _tag_endswith(elem, "c"):
                        current_cell_ref = None
                        sheet_cell_count += 1
                        elem.clear()
                        if sheet_cell_count >= max_cells:
                            break
    except Exception as exc:  # noqa: BLE001
        Path(output_csv).unlink(missing_ok=True)
        return FormulaWorkerResult(count=0, csv_path=output_csv, error=f"{type(exc).__name__}: {exc}")
    return FormulaWorkerResult(count=count, csv_path=output_csv)


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def main() -> None:
    """Main entry point for the Excel forensics script.

    Parses arguments, loads configuration, processes the Excel workbook,
    extracts metadata and data, and saves results to various formats.

    Args:
        None

    Returns:
        None
    """
    args = parse_args()
    start_global = time.perf_counter()

    # 1. Config
    config = get_config(args)
    excel_path = config["workbook_path"]
    output_dir = config["output_dir"]
    db_path = output_dir / config["db_name"]
    parquet_dir = output_dir / config["parquet_dir"]
    sst_truncate = config["sst_truncate"]
    skip_formulas = config["skip_formulas"]

    if not Path(excel_path).is_file():
        console.print(f"[bold red]Workbook not found:[/bold red] {excel_path}")
        console.print("Pass --excel path/to/workbook.xlsx")
        raise SystemExit(1)

    console.rule("[bold blue]xlsx-autopsy[/bold blue]")
    log_status(f"Input: {excel_path}")

    if config["keep_outputs"]:
        log_status("Keeping prior outputs [--keep-outputs]")
    else:
        log_status("Resetting outputs...")
        cleanup_outputs(output_dir, db_path, parquet_dir)

    output_dir.mkdir(parents=True, exist_ok=True)
    parquet_dir.mkdir(parents=True, exist_ok=True)

    # --- Setup Logging ---
    logger.setLevel(logging.INFO)
    if logger.hasHandlers():
        logger.handlers.clear()
    log_file = output_dir / "reconstruction_history.log"
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))
    logger.addHandler(fh)

    logger.info("-" * 80)
    logger.info(f"RUN STARTED | Input: {excel_path}")

    analysis: AnalysisReport = {
        # Blueprint used to recreate this Excel report in SQL + a BI tool
        "workbook_path": str(excel_path),
        "metadata": {},
        "connections": [],
        "pivot_caches": [],
        "pivot_tables": [],
        "pivot_cache_defs": {},
        "pivot_tables_resolved": [],
        "package_inventory": {},
        "processed_sheets": [],
        "failed_sheets": [],
        "formula_warnings": [],
    }

    try:
        with zipfile.ZipFile(excel_path, "r") as zipf:
            con = duckdb.connect(str(db_path))
            # --- PHASE 1: Report Structure (XML) ---
            log_status("Extracting report structure (sheets, connections, pivots)")

            wb_meta = parse_workbook_meta(zipf)
            analysis["metadata"] = wb_meta

            analysis["connections"] = extract_connections(
                zipf,
                include_secrets=config["include_connection_secrets"],
            )
            analysis["pivot_caches"] = parse_pivot_caches(zipf)
            analysis["pivot_tables"] = parse_detailed_pivots(zipf)
            analysis["package_inventory"] = list_excel_package_inventory(zipf)

            # Resolve pivot fields -> names via pivot cache definitions
            cache_defs_by_id = parse_pivot_cache_definitions(zipf, wb_meta)
            analysis["pivot_cache_defs"] = cache_defs_by_id
            analysis["pivot_tables_resolved"] = resolve_pivot_tables(analysis["pivot_tables"], cache_defs_by_id)

            # --- PHASE 2: Formulas ---
            # This is slow but necessary for "Everything"
            # Always create formulas table (even if empty) for consistency
            con.execute("DROP TABLE IF EXISTS formulas")
            con.execute("CREATE TABLE formulas (sheet_file TEXT, cell TEXT, formula TEXT)")

            if skip_formulas:
                log_status("Skipping calculation logic (cell formulas) [--skip-formulas enabled]")
                log_status("  ⚠ Formulas table created but left empty for schema consistency")
            else:
                log_status("Extracting calculation logic (Multiprocessed)...")

                cfg = (config or {}).get("formula_extract", {})
                scout = scout_formula_sheets(
                    zipf,
                    max_sheets=int(cfg.get("max_sheets", 300)),
                    max_bytes_per_sheet=int(cfg.get("scout_bytes", 25_000_000)),
                )
                candidates = scout.hits
                if scout.omitted_by_cap:
                    warning = (
                        f"Formula scout omitted {len(scout.omitted_by_cap)} sheet(s) "
                        f"past max_sheets={cfg.get('max_sheets', 300)}"
                    )
                    log_status(f"  ⚠ {warning}")
                    analysis["formula_warnings"].append(warning)
                if scout.forced_parse:
                    warning = f"Formula scout forced parse on {len(scout.forced_parse)} sheet(s)"
                    log_status(f"  ⚠ {warning}")
                    analysis["formula_warnings"].append(warning)

                if not candidates:
                    log_status("  ⚠ No formulas found (scout returned 0 sheets)")
                else:
                    user_limit = int(cfg.get("max_workers", 4))
                    num_workers = min(user_limit, multiprocessing.cpu_count())
                    num_workers = min(num_workers, len(candidates))

                    chunk_size = math.ceil(len(candidates) / num_workers)
                    chunks = [candidates[i : i + chunk_size] for i in range(0, len(candidates), chunk_size)]

                    max_f = int(cfg.get("max_formulas_per_sheet", 200_000))
                    max_c = int(cfg.get("max_cells_per_sheet", 1_500_000))

                    log_status(f"  • Spawning {num_workers} workers for {len(candidates)} sheets")

                    worker_results: list[FormulaWorkerResult] = []
                    with ProcessPoolExecutor(max_workers=num_workers) as executor:
                        futures = []
                        for i, chunk in enumerate(chunks):
                            csv_path = output_dir / f"temp_formulas_worker_{i}.csv"
                            futures.append(
                                executor.submit(
                                    _worker_formula_task,
                                    str(excel_path),
                                    chunk,
                                    str(csv_path),
                                    max_f,
                                    max_c,
                                )
                            )
                        worker_results = [future.result() for future in futures]

                    total_captured = 0
                    worker_errors: list[str] = []
                    for result in worker_results:
                        csv_file = Path(result.csv_path)
                        if result.error:
                            worker_errors.append(result.error)
                            analysis["formula_warnings"].append(f"Formula worker failed: {result.error}")
                            csv_file.unlink(missing_ok=True)
                            continue
                        if not csv_file.is_file():
                            worker_errors.append(f"missing CSV {result.csv_path}")
                            analysis["formula_warnings"].append(f"Formula worker missing CSV: {result.csv_path}")
                            continue
                        if result.count > 0:
                            csv_sql = str(csv_file).replace("\\", "/")
                            con.execute(
                                f"COPY formulas FROM '{csv_sql}' "
                                "(FORMAT CSV, HEADER true, DELIMITER ',', QUOTE '\"', ESCAPE '\"')"
                            )
                        total_captured += result.count
                        csv_file.unlink(missing_ok=True)

                    log_status(f"  ✓ Captured {total_captured:,} formulas")
                    if worker_errors:
                        raise RuntimeError(f"{len(worker_errors)} formula worker(s) failed: {worker_errors[0]}")

            # --- PHASE 3: Shared Strings (The "Ghost Data" Hunt) ---
            log_status("Extracting lookup text and dimensions (shared strings)")
            stream_shared_strings(zipf, con, output_dir, sst_truncate)

            # --- PHASE 4: Data Extraction (Polars + Calamine) ---
            # Use the sheet list we already parsed from XML! No fastexcel needed.
            sheets_to_process = [s["name"] for s in wb_meta["sheets"] if s["name"] is not None]
            log_status(f"Found {len(sheets_to_process)} sheets to extract")

            # Track seen table names to prevent collision (e.g. 'Sheet' vs 'Sheet!')
            seen_tables = set()

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("{task.completed}/{task.total}"),
                TimeElapsedColumn(),
                transient=False,
            ) as progress:
                task = progress.add_task("Extracting...", total=len(sheets_to_process))

                for sheet_name in sheets_to_process:
                    base_name = sanitize_table_name(sheet_name)
                    sanitized_name = base_name
                    counter = 1
                    while sanitized_name in seen_tables:
                        sanitized_name = f"{base_name}_{counter}"
                        counter += 1
                    seen_tables.add(sanitized_name)

                    progress.update(task, description=f"Sheet: [bold]{sheet_name}[/bold]")

                    try:
                        # infer_schema_length=0 ensures EVERYTHING is a string (Forensic Fidelity)
                        df = pl.read_excel(
                            str(excel_path),
                            sheet_name=sheet_name,
                            engine="calamine",
                            has_header=False,
                            infer_schema_length=0,
                            raise_if_empty=False,
                        )

                        if df.height == 0:
                            empty_sheet = cast(ProcessedSheet, {"name": sheet_name, "status": "empty"})
                            analysis["processed_sheets"].append(empty_sheet)
                            progress.advance(task)
                            continue

                        # 1. Parquet
                        pq_path = parquet_dir / f"{sanitized_name}.parquet"
                        df.write_parquet(pq_path)
                        # 2. DuckDB
                        con.execute(
                            f"CREATE OR REPLACE TABLE '{sanitized_name}' AS "
                            f"SELECT * FROM read_parquet('{str(pq_path)}')"
                        )

                        processed_sheet = cast(
                            ProcessedSheet,
                            {"name": sheet_name, "table": sanitized_name, "rows": df.height, "cols": df.width},
                        )
                        analysis["processed_sheets"].append(processed_sheet)

                    except Exception as e:
                        progress.console.print(f"  [red]Error on {sheet_name}: {e}[/red]")
                        failed_sheet = cast(FailedSheet, {"name": sheet_name, "error": str(e)})
                        analysis["failed_sheets"].append(failed_sheet)

                    progress.advance(task)

        # --- PHASE 5: Finalize ---
        log_status("Writing metadata tables to DuckDB...")

        # Normalize optional collections for safe len() and iteration
        analysis["connections"] = analysis.get("connections") or []
        analysis["pivot_tables"] = analysis.get("pivot_tables") or []
        analysis["pivot_caches"] = analysis.get("pivot_caches") or []

        if analysis["connections"]:
            try:
                tmp_conn = output_dir / "temp_connections.json"
                with tmp_conn.open("w", encoding="utf-8") as f:
                    json.dump(_json_sanitize(analysis["connections"]), f)
                con.execute(
                    "CREATE OR REPLACE TABLE _meta_connections AS SELECT * FROM read_json_auto(?)",
                    [str(tmp_conn).replace(os.sep, "/")],
                )
                tmp_conn.unlink()
                log_status(f"  ✓ Wrote {len(analysis['connections'])} connection(s) to _meta_connections")
            except Exception as e:
                log_status(f"  ✗ Failed to write connections metadata: {e}")

        if analysis["pivot_tables"]:
            try:
                tmp_piv = output_dir / "temp_pivots.json"
                with tmp_piv.open("w", encoding="utf-8") as f:
                    json.dump(_json_sanitize(analysis["pivot_tables"]), f)
                con.execute(
                    "CREATE OR REPLACE TABLE _meta_pivots AS SELECT * FROM read_json_auto(?)",
                    [str(tmp_piv).replace(os.sep, "/")],
                )
                tmp_piv.unlink()
                log_status(f"  ✓ Wrote {len(analysis['pivot_tables'])} pivot table(s) to _meta_pivots")
            except Exception as e:
                log_status(f"  ✗ Failed to write pivot metadata: {e}")

        if analysis["pivot_caches"]:
            try:
                tmp_cache = output_dir / "temp_pivot_caches.json"
                with tmp_cache.open("w", encoding="utf-8") as f:
                    json.dump(_json_sanitize(analysis["pivot_caches"]), f)
                con.execute(
                    "CREATE OR REPLACE TABLE _meta_pivot_caches AS SELECT * FROM read_json_auto(?)",
                    [str(tmp_cache).replace(os.sep, "/")],
                )
                tmp_cache.unlink()
                log_status(f"  ✓ Wrote {len(analysis['pivot_caches'])} pivot cache(s) to _meta_pivot_caches")
            except Exception as e:
                log_status(f"  ✗ Failed to write pivot cache metadata: {e}")

        if analysis.get("pivot_cache_defs"):
            try:
                tmp_cache_defs = output_dir / "temp_pivot_cache_defs.json"
                with tmp_cache_defs.open("w", encoding="utf-8") as f:
                    json.dump(_json_sanitize(analysis["pivot_cache_defs"]), f)
                con.execute(
                    "CREATE OR REPLACE TABLE _meta_pivot_cache_defs AS SELECT * FROM read_json_auto(?)",
                    [str(tmp_cache_defs).replace(os.sep, "/")],
                )
                tmp_cache_defs.unlink()
                log_status("  ✓ Wrote pivot cache definitions to _meta_pivot_cache_defs")
            except Exception as e:
                log_status(f"  ✗ Failed to write pivot cache defs metadata: {e}")

        if analysis.get("pivot_tables_resolved"):
            try:
                tmp_piv_res = output_dir / "temp_pivots_resolved.json"
                with tmp_piv_res.open("w", encoding="utf-8") as f:
                    json.dump(_json_sanitize(analysis["pivot_tables_resolved"]), f)
                con.execute(
                    "CREATE OR REPLACE TABLE _meta_pivots_resolved AS SELECT * FROM read_json_auto(?)",
                    [str(tmp_piv_res).replace(os.sep, "/")],
                )
                tmp_piv_res.unlink()
                log_status("  ✓ Wrote resolved pivots to _meta_pivots_resolved")
            except Exception as e:
                log_status(f"  ✗ Failed to write resolved pivots metadata: {e}")

        # Summary statistics
        log_status("\n📊 Extraction Summary:")
        log_status(f"  • Sheets processed: {len(analysis['processed_sheets'])}")
        log_status(f"  • Sheets failed: {len(analysis['failed_sheets'])}")
        formulas_count = con.execute("SELECT COUNT(*) FROM formulas").fetchone()
        log_status(f"  • Formulas captured: {(formulas_count[0] if formulas_count else 0):,}")
        log_status(f"  • Pivot tables: {len(analysis.get('pivot_tables') or [])}")
        log_status(f"  • Pivot caches: {len(analysis.get('pivot_caches') or [])}")
        log_status(f"  • Pivot tables resolved: {len(analysis.get('pivot_tables_resolved') or [])}")
        log_status(f"  • Connections: {len(analysis.get('connections') or [])}")
        sst_exists = con.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_name = 'sst'").fetchone()
        if sst_exists and sst_exists[0]:
            shared_strings_count = con.execute("SELECT COUNT(*) FROM sst").fetchone()
            sst_n = shared_strings_count[0] if shared_strings_count else 0
        else:
            sst_n = 0
        log_status(f"  • Shared strings: {sst_n:,}")

        write_outputs(analysis, output_dir, db_path, parquet_dir)

        duration = time.perf_counter() - start_global
        console.rule("[bold green]Complete[/bold green]")
        console.print(f"Total Time: [bold white]{duration:.2f}s[/bold white]")
        console.print("DB Schema:")
        con.sql("SHOW TABLES").show()
        con.close()

    except Exception as exc:
        console.print(f"[bold red]CRITICAL FAILURE:[/bold red] {exc}")
        traceback.print_exc()
        try:
            con.close()  # type: ignore[possibly-undefined]
        except Exception:
            pass
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
