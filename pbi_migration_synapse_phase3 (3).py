
"""PBIP Metadata Reader and Migration Utility.

Safety: migrates physical sources only. Native SQL is parsed and rewritten with SQLGlot T-SQL ASTs. Supports Synapse-to-Synapse migrations including right-query expansion, unified merged-query output bindings for downstream M and TMDL sourceColumn, native SQL output-contract protection, immutable SQLGlot payloads, implicit qualifier aliases, per-expression navigation, and native SQL inventory. Semantic table/column/measure names and DAX are protected.

Edit only ROOT_FOLDER and DRY_RUN in USER SETTINGS.

Stages:
  DRY_RUN=True, no mappings: initial inventory.
  DRY_RUN=True, mappings present: mapping validation, dependency analysis,
  recommendations written to Mapping Template.xlsx, no PBIP edits.
  DRY_RUN=False, mappings present: apply valid mappings directly to PBIP TMDL.
"""
from __future__ import annotations

# =============================================================================
# USER SETTINGS - EDIT ONLY THESE VALUES
# =============================================================================
ROOT_FOLDER = r"C:\Your\Report Root Folder"
DRY_RUN = True
ALIGN_TARGET_COLUMN_NAMES = True  # False = legacy-alias output contracts
RENAME_SEMANTIC_COLUMNS = True    # Phase 3: False = source-only (Phase 2), True = full semantic + report rename
# =============================================================================

import hashlib
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

try:
    import sqlglot
    from sqlglot import exp
    from sqlglot.errors import ParseError, UnsupportedError
    from sqlglot.optimizer.scope import traverse_scope
except ImportError as exc:
    raise SystemExit(
        "SQLGlot is required for safe native SQL migration. "
        "Install it with: python -m pip install sqlglot"
    ) from exc

LOG_NAME = "Metadata Log.xlsx"
MAPPING_NAME = "Mapping Template.xlsx"
TABLE_MAP = "Table and Column Mapping"
SOURCE_MAP = "Data Source Mapping"

TABLE_HEADERS = [
    "Legacy Table", "Target Table", "Legacy Column Name", "Legacy Data Type",
    "Target Column Name", "Target Data Type", "Retain Legacy Data Type",
    "Dependency Notes",
]
SOURCE_HEADERS = [
    "Legacy Source", "Legacy Schema", "Legacy Item",
    "Target Source", "Target Schema", "Target Item",
    "Source SQL Dialect", "Target SQL Dialect",
    "Legacy Catalog", "Target Catalog",
]

LOG_SHEETS = {
    "Run Summary": ["Item", "Value"],
    "Parameters": ["Parameter Name", "Value", "Is Current", "Data Type", "Referenced By"],
    "Query & Source Inventory": [
        "Object Type", "Table Mode", "Query Name", "Query Type", "Source Type",
        "Source", "Server", "Database", "Schema", "Item",
        "Native Query", "Query Status", "Origin", "Source File", "Contains Join",
        "Contains Merge", "Contains Append", "Contains Unpivot", "Migration Complexity", "Notes",
    ],
    "Files Processed": [
        "Run Timestamp", "Input File", "Input Path", "File Type", "Classification",
        "Relevant Mapping Found", "Changed", "Action", "Write Status", "Message",
        "Input SHA-256", "Generated SHA-256",
    ],
    "Table Changes": [
        "Input File", "File Type", "Query or Model Table", "Change Location",
        "Old Table", "New Table", "SQL Alias", "Alias Preserved",
        "Replacement Count", "Result",
    ],
    "Column Changes": [
        "Input File", "Model Table", "Model Column", "Old Source Column",
        "New Source Column", "Changed In", "Model Name Changed", "Result",
    ],
    "Data Type Decisions": [
        "Input File", "Table", "Model Column", "Legacy Type", "Target Type",
        "Retain Legacy Data Type", "Final Type", "Decision Source",
        "Rule Description", "SQL Cast Added", "M Step Added", "TMDL Updated", "Result",
    ],
    "Type Dependency Analysis": [
        "Legacy Table", "Legacy Column", "Legacy Data Type", "Target Table",
        "Target Column", "Target Data Type", "Type Mismatch",
        "Relationship Dependency", "DAX Dependency", "M Query Dependency",
        "Native SQL Dependency", "Unpivot Dependency", "Other Dependency",
        "System Recommendation", "Mapping Template Selection", "Selection Source",
        "Dependency Notes", "Source Files", "Validation Status",
    ],
    "Native SQL Changes": [
        "Input File", "Query Name", "Change Type", "Old Value", "New Value",
        "Occurrences", "Safe Replacement", "Message",
    ],
    "M Query Changes": [
        "Input File", "Query Name", "M Step", "Change Type", "Old Value",
        "New Value", "Result",
    ],
    "Unpivot Inventory": [
        "Query Name", "Source File", "Unpivot Function", "Input Step",
        "Unpivoted Columns", "Retained Columns", "Attribute Column", "Value Column",
        "Legacy Column Name", "Target Column Name", "Legacy Name Becomes Row Value",
        "Legacy Name Retained", "Retention Method", "Validation Status", "Details",
    ],
    "Semantic & Report Changes": [
        "Input File", "Layer", "Object", "Change Type",
        "Old Value", "New Value", "Occurrences", "Result",
    ],
    "Protected Content": [
        "Input File", "Protected Type", "Object Name", "Count", "Action", "Reason",
    ],
    "Warnings & Manual Review": [
        "Input File", "Warning Code", "Category", "Object", "Warning", "Impact",
        "Recommended Review",
    ],
    "Errors": [
        "Input File", "Stage", "Error Code", "Error Message",
        "Existing Content Preserved", "Processing Continued",
    ],
}

DARK_BLUE = "1F4E78"
WHITE = "FFFFFF"
LIGHT_BLUE = "D9EAF7"
ORANGE = "FCE4D6"
LIGHT_RED = "F4CCCC"
THIN_GRAY = Side(style="thin", color="D9E1F2")

@dataclass
class Paths:
    root: Path
    migration: Path
    pbip: Path
    semantic: Path
    definition: Path
    tables: Path
    expressions: Path
    relationships: list[Path]
    report: Path | None = None
    roles: list[Path] = field(default_factory=list)
    cultures: list[Path] = field(default_factory=list)
    perspectives: list[Path] = field(default_factory=list)

@dataclass
class MappingRow:
    excel_row: int
    legacy_table: str
    target_table: str
    legacy_column: str
    legacy_type: str
    target_column: str
    target_type: str
    selection: str
    notes: str
    system_recommendation: str = ""
    dependency_notes: str = ""
    dependencies: dict[str, list[str]] = field(default_factory=dict)

@dataclass
class SourceMapping:
    legacy_source: str
    legacy_schema: str
    legacy_item: str
    target_source: str
    target_schema: str
    target_item: str
    source_sql_dialect: str = ""
    target_sql_dialect: str = ""
    legacy_catalog: str = ""
    target_catalog: str = ""


def clean(v) -> str:
    return "" if v is None else str(v).strip()


def norm(v: str) -> str:
    return re.sub(r"\s+", " ", clean(v)).casefold()


def sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def read_text(path: Path) -> str:
    for enc in ("utf-8-sig", "utf-16", "utf-8"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeError:
            pass
    return path.read_text(errors="replace")


def atomic_write(path: Path, text: str) -> None:
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(text, encoding="utf-8")
    temp.replace(path)


def discover_paths(root_text: str) -> Paths:
    root = Path(root_text).expanduser().resolve()
    if not root.is_dir():
        raise FileNotFoundError(f"ROOT_FOLDER does not exist: {root}")
    migration = root / "Migration"
    migration.mkdir(exist_ok=True)
    candidates = sorted(p for p in root.iterdir() if p.is_dir() and p.name.casefold().endswith(" pbip"))
    if len(candidates) != 1:
        raise ValueError(f"Expected exactly one folder ending in ' Pbip' under {root}; found {len(candidates)}")
    pbip = candidates[0]
    semantic_models = sorted(pbip.glob("*.SemanticModel"))
    if len(semantic_models) != 1:
        raise ValueError(f"Expected exactly one .SemanticModel folder under {pbip}; found {len(semantic_models)}")
    semantic = semantic_models[0]
    definition = semantic / "definition"
    tables = definition / "tables"
    if not tables.is_dir():
        raise FileNotFoundError(f"TMDL tables folder not found: {tables}")
    expressions = definition / "expressions.tmdl"
    relationship_files = [p for p in definition.glob("*.tmdl") if "relationship" in p.name.casefold()]
    model = definition / "model.tmdl"
    if model.exists(): relationship_files.append(model)
    reports = sorted(pbip.glob("*.Report"))
    report_dir = reports[0] if len(reports) == 1 else None
    roles_files = sorted((definition / "roles").glob("*.tmdl")) if (definition / "roles").is_dir() else []
    cultures_files = sorted((definition / "cultures").glob("*.tmdl")) if (definition / "cultures").is_dir() else []
    perspectives_files = sorted((definition / "perspectives").glob("*.tmdl")) if (definition / "perspectives").is_dir() else []
    return Paths(
        root, migration, pbip, semantic, definition, tables, expressions,
        sorted(set(relationship_files)), report_dir, roles_files, cultures_files, perspectives_files,
    )


def table_name(text: str, fallback: str) -> str:
    m = re.search(r"(?mi)^\s*table\s+(.+?)\s*$", text)
    return m.group(1).strip(" '\"") if m else fallback


def split_expression_blocks(path: Path, text: str):
    if path.name.casefold() != "expressions.tmdl":
        return [(table_name(text, path.stem), text)]
    matches = list(re.finditer(r"(?mi)^\s*expression\s+(.+?)\s*=", text))
    if not matches: return [("expressions", text)]
    out = []
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        out.append((m.group(1).strip(" '\""), text[m.start():end]))
    return out


def create_or_upgrade_mapping_template(path: Path) -> None:
    if path.exists():
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.active.title = TABLE_MAP
        wb.create_sheet(SOURCE_MAP)
    for title, headers in ((TABLE_MAP, TABLE_HEADERS), (SOURCE_MAP, SOURCE_HEADERS)):
        ws = wb[title] if title in wb.sheetnames else wb.create_sheet(title)
        existing = [clean(c.value) for c in ws[1]]
        if not any(existing):
            for i, h in enumerate(headers, 1): ws.cell(1, i, h)
        else:
            for h in headers:
                if h not in existing:
                    existing.append(h)
                    ws.cell(1, len(existing), h)
        style_tab(ws, add_table=False)
    ws = wb[TABLE_MAP]
    header_map = {clean(c.value): c.column for c in ws[1]}
    retain_col = get_column_letter(header_map["Retain Legacy Data Type"])
    # Remove old validations to avoid duplicates on repeated runs.
    ws.data_validations.dataValidation = []
    dv = DataValidation(type="list", formula1='"Yes,No,Review"', allow_blank=True)
    dv.error = "Select Yes, No, or Review from the dropdown."
    dv.errorTitle = "Invalid selection"
    dv.prompt = "Choose Yes, No, or Review."
    dv.promptTitle = "Retain legacy data type"
    dv.showErrorMessage = True
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(f"{retain_col}2:{retain_col}10000")
    wb.save(path)


def style_tab(ws, add_table: bool = True) -> None:
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 34
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=DARK_BLUE)
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN_GRAY)
    if ws.max_row > 1:
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        vals = [len(clean(ws.cell(r, col).value)) for r in range(1, min(ws.max_row, 250) + 1)]
        width = min(max(max(vals or [10]) + 2, 13), 55)
        if clean(ws.cell(1, col).value) in {"Notes", "Dependency Notes", "Details", "Message", "Warning", "Recommended Review"}:
            width = 55
        ws.column_dimensions[get_column_letter(col)].width = width
    if add_table and ws.max_row >= 2 and ws.max_column >= 1:
        # Clear existing tables before recreating exactly once.
        ws._tables.clear()
        name = "Tbl" + re.sub(r"[^A-Za-z0-9]", "", ws.title)
        name = name[:245]
        tab = Table(displayName=name, ref=f"A1:{get_column_letter(ws.max_column)}{ws.max_row}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False, showLastColumn=False)
        ws.add_table(tab)
    elif ws.max_row >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row,1)}"


def parse_mapping(path: Path):
    wb = load_workbook(path, data_only=False)
    mappings, source_maps = [], []
    ws = wb[TABLE_MAP]
    h = {clean(c.value): c.column for c in ws[1]}
    for r in range(2, ws.max_row + 1):
        get = lambda name: clean(ws.cell(r, h[name]).value) if name in h else ""
        values = [get(x) for x in TABLE_HEADERS]
        if not any(values): continue
        if not ((get("Legacy Table") and get("Target Table")) or (get("Legacy Column Name") and get("Target Column Name"))):
            continue
        mappings.append(MappingRow(r, get("Legacy Table"), get("Target Table"), get("Legacy Column Name"), get("Legacy Data Type"), get("Target Column Name"), get("Target Data Type"), get("Retain Legacy Data Type"), get("Dependency Notes")))
    ws2 = wb[SOURCE_MAP]
    h2 = {clean(c.value): c.column for c in ws2[1]}
    for r in range(2, ws2.max_row + 1):
        get = lambda name: clean(ws2.cell(r, h2[name]).value) if name in h2 else ""
        if not any(get(x) for x in SOURCE_HEADERS): continue
        if get("Legacy Source") and get("Target Source"):
            source_maps.append(SourceMapping(*(get(x) for x in SOURCE_HEADERS)))
    return mappings, source_maps


def normalize_type(value: str) -> str:
    v = norm(value).replace(" ", "")
    groups = {
        "text": {"text", "string", "varchar", "nvarchar", "char", "nchar"},
        "integer": {"int", "integer", "int16", "int32", "int64", "bigint", "smallint", "tinyint"},
        "decimal": {"decimal", "numeric", "fixeddecimal", "currency"},
        "number": {"number", "double", "float", "single", "real"},
        "date": {"date"},
        "datetime": {"datetime", "datetime2", "datetimezone", "timestamp"},
        "boolean": {"bool", "boolean", "bit", "logical"},
        "binary": {"binary", "varbinary"},
    }
    for group, items in groups.items():
        if v in items: return group
    return v


def all_model_files(paths: Paths) -> list[Path]:
    files = sorted(paths.definition.rglob("*.tmdl"))
    return files


def find_column_references(text: str, table: str, column: str) -> bool:
    if not column: return False
    patterns = [
        rf"'?'?{re.escape(table)}'?'?\s*\[\s*{re.escape(column)}\s*\]" if table else r"(?!x)x",
        rf"\[\s*{re.escape(column)}\s*\]",
        rf'"{re.escape(column)}"',
        rf"sourceColumn\s*:\s*{re.escape(column)}\b",
    ]
    return any(re.search(p, text, re.I) for p in patterns)


def dependency_analysis(paths: Paths, mappings: list[MappingRow]):
    """Analyze structural source dependencies without inspecting or rewriting DAX."""
    files = all_model_files(paths)
    contents = {path: read_text(path) for path in files}
    results = []

    for mapping in mappings:
        deps = {key: [] for key in ("relationship", "m", "sql", "unpivot", "other")}
        source_files = set()
        mismatch = (
            bool(mapping.legacy_type and mapping.target_type)
            and normalize_type(mapping.legacy_type) != normalize_type(mapping.target_type)
        )

        for path, full_text in contents.items():
            m_region, _, _ = extract_m_region(full_text)
            sql_regions = native_sql_regions(full_text)
            sql_text = "\n".join(region[2] for region in sql_regions)

            legacy_ref = find_column_references(
                m_region + "\n" + sql_text,
                mapping.legacy_table,
                mapping.legacy_column,
            )
            target_ref = find_column_references(
                m_region + "\n" + sql_text,
                mapping.target_table,
                mapping.target_column,
            )
            relationship_ref = find_column_references(
                full_text if path in paths.relationships else "",
                mapping.legacy_table,
                mapping.legacy_column,
            )
            if not (legacy_ref or target_ref or relationship_ref):
                continue

            source_files.add(str(path.relative_to(paths.semantic.parent)))
            if relationship_ref:
                deps["relationship"].append(path.name)
            if m_region and (legacy_ref or target_ref):
                deps["m"].append(path.name)
                if re.search(r"(?i)Table\.Unpivot(?:OtherColumns)?", m_region):
                    deps["unpivot"].append(path.name)
                if re.search(r"(?i)(sortByColumn|hierarchy\s+|level\s+)", full_text):
                    deps["other"].append(path.name)
            if sql_regions and (legacy_ref or target_ref):
                deps["sql"].append(path.name)

        for key in deps:
            deps[key] = sorted(set(deps[key]))

        if not mapping.legacy_type:
            recommendation = "No"
            note = "Legacy data type is blank; legacy retention is not required."
        elif mapping.target_type and normalize_type(mapping.legacy_type) == normalize_type(mapping.target_type):
            recommendation = "No"
            note = "Legacy and target data types normalize to the same semantic type."
        elif mapping.target_type:
            recommendation = "Review"
            note = "Legacy and target data types differ; user review is required."
        else:
            recommendation = "Review"
            note = "Target data type is blank; user review is required."

        mapping.system_recommendation = recommendation
        mapping.dependency_notes = note
        mapping.dependencies = deps
        selection = mapping.selection if mapping.selection in {"Yes", "No", "Review"} else recommendation
        selection_source = "User Selection" if mapping.selection in {"Yes", "No", "Review"} else "Code Recommendation"

        results.append({
            "Legacy Table": mapping.legacy_table,
            "Legacy Column": mapping.legacy_column,
            "Legacy Data Type": mapping.legacy_type,
            "Target Table": mapping.target_table,
            "Target Column": mapping.target_column,
            "Target Data Type": mapping.target_type,
            "Type Mismatch": "Yes" if mismatch else "No",
            "Relationship Dependency": ", ".join(deps["relationship"]),
            "DAX Dependency": "Not inspected",
            "M Query Dependency": ", ".join(deps["m"]),
            "Native SQL Dependency": ", ".join(deps["sql"]),
            "Unpivot Dependency": ", ".join(deps["unpivot"]),
            "Other Dependency": ", ".join(deps["other"]),
            "System Recommendation": recommendation,
            "Mapping Template Selection": selection,
            "Selection Source": selection_source,
            "Dependency Notes": note,
            "Source Files": "; ".join(sorted(source_files)),
            "Validation Status": "Completed",
        })
    return results

def update_mapping_recommendations(path: Path, mappings: list[MappingRow]) -> None:
    wb = load_workbook(path)
    ws = wb[TABLE_MAP]
    h = {clean(c.value): c.column for c in ws[1]}
    for m in mappings:
        retain_cell = ws.cell(m.excel_row, h["Retain Legacy Data Type"])
        notes_cell = ws.cell(m.excel_row, h["Dependency Notes"])
        if clean(retain_cell.value) not in {"Yes", "No", "Review"}:
            retain_cell.value = m.system_recommendation
        if not clean(notes_cell.value):
            notes_cell.value = m.dependency_notes
    style_tab(ws, add_table=False)
    style_tab(wb[SOURCE_MAP], add_table=False)
    wb.save(path)


def detect_source(block: str):
    m = re.search(r'Sql\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', block, re.I)
    if m:
        server, database = m.group(1), m.group(2)
        nav = re.search(r'\[\s*Schema\s*=\s*"([^"]+)"\s*,\s*Item\s*=\s*"([^"]+)"', block, re.I)
        schema, item = (nav.group(1), nav.group(2)) if nav else ("", "")
        return "SQL", f"{server};{database}", server, database, schema, item
    patterns = [
        ("Excel", r'Excel\.Workbook\s*\(\s*File\.Contents\s*\(\s*"([^"]+)"'),
        ("SharePoint", r'SharePoint\.(?:Files|Contents|Tables)\s*\(\s*"([^"]+)"'),
        ("Folder", r'Folder\.(?:Files|Contents)\s*\(\s*"([^"]+)"'),
        ("CSV", r'Csv\.Document\s*\(\s*File\.Contents\s*\(\s*"([^"]+)"'),
        ("Web", r'Web\.Contents\s*\(\s*"([^"]+)"'),
        ("OData", r'OData\.Feed\s*\(\s*"([^"]+)"'),
    ]
    for typ, pat in patterns:
        m = re.search(pat, block, re.I)
        if m: return typ, m.group(1), "", "", "", ""
    return "Other", "", "", "", "", ""


def _sql_physical_sources(sql_text: str):
    """Return every physical SQL source found through SQLGlot scope analysis.

    CTE and derived-table references are excluded because traverse_scope exposes
    them as Scope objects rather than exp.Table sources. Repeated FROM/JOIN uses
    are aggregated by physical object and alias while retaining occurrence count.
    """
    normalized_sql, function_audit = normalize_synapse_sql_functions(sql_text)
    records = {}
    parse_errors = []
    try:
        statements = sqlglot.parse(normalized_sql, read="tsql")
    except ParseError as exc:
        return [], function_audit, [str(exc)]

    for statement_index, statement in enumerate(statements, 1):
        if statement is None:
            continue
        try:
            scopes = list(traverse_scope(statement))
        except Exception as exc:
            parse_errors.append(f"Scope analysis failed for statement {statement_index}: {exc}")
            continue
        for scope_index, scope in enumerate(scopes, 1):
            for alias, source in scope.sources.items():
                if not isinstance(source, exp.Table):
                    continue
                catalog = clean(source.catalog)
                schema = clean(source.db)
                item = clean(source.name)
                if not item:
                    continue
                alias_name = clean(alias) or item
                physical = ".".join(part for part in (catalog, schema, item) if part)
                key = (
                    norm(catalog), norm(schema), norm(item), norm(alias_name),
                )
                record = records.setdefault(key, {
                    "Catalog": catalog,
                    "Schema": schema,
                    "Item": item,
                    "Alias": alias_name,
                    "Physical Object": physical,
                    "Occurrences": 0,
                    "Scopes": set(),
                })
                record["Occurrences"] += 1
                record["Scopes"].add(f"statement {statement_index}/scope {scope_index}")

    output = []
    for record in records.values():
        record["Scopes"] = "; ".join(sorted(record["Scopes"]))
        output.append(record)
    output.sort(key=lambda row: (
        norm(row["Catalog"]), norm(row["Schema"]),
        norm(row["Item"]), norm(row["Alias"]),
    ))
    return output, function_audit, parse_errors


def _native_source_inventory_rows(
    query: str,
    block: str,
    source_file: str,
    origin: str,
    query_status: str,
    server: str,
    database: str,
    contains_join: str,
    contains_merge: str,
    contains_append: str,
    contains_unpivot: str,
):
    """Create one Query & Source Inventory row per embedded physical SQL source."""
    rows = []
    sql_regions = native_sql_regions(block)
    for payload_index, (_, _, encoded_sql) in enumerate(sql_regions, 1):
        decoded_sql = decode_m_sql_string(encoded_sql)
        physical_sources, function_audit, parse_errors = _sql_physical_sources(decoded_sql)
        for physical in physical_sources:
            notes = (
                f"Discovered from native SQL payload {payload_index} using SQLGlot AST; "
                f"alias={physical['Alias']}; occurrences={physical['Occurrences']}; "
                f"scopes={physical['Scopes']}"
            )
            rows.append({
                "Object Type": "Native SQL Source",
                "Query Name": query,
                "Query Type": "Native SQL Physical Source",
                "Source Type": "SQL",
                "Source": physical["Physical Object"],
                "Server": server,
                "Database": database,
                "Schema": physical["Schema"],
                "Item": physical["Item"],
                "Native Query": "Yes",
                "Query Status": query_status,
                "Origin": origin,
                "Source File": source_file,
                "Contains Join": contains_join,
                "Contains Merge": contains_merge,
                "Contains Append": contains_append,
                "Contains Unpivot": contains_unpivot,
                "Migration Complexity": "High",
                "Notes": notes,
            })
        if parse_errors:
            rows.append({
                "Object Type": "Native SQL Source Parse Error",
                "Query Name": query,
                "Query Type": "Native SQL Physical Source",
                "Source Type": "SQL",
                "Source": "",
                "Server": server,
                "Database": database,
                "Schema": "",
                "Item": "",
                "Native Query": "Yes",
                "Query Status": query_status,
                "Origin": origin,
                "Source File": source_file,
                "Contains Join": contains_join,
                "Contains Merge": contains_merge,
                "Contains Append": contains_append,
                "Contains Unpivot": contains_unpivot,
                "Migration Complexity": "High",
                "Notes": "Native SQL source extraction failed: " + " | ".join(parse_errors),
            })
    return rows


def _deduplicate_inventory_rows(rows):
    """Deduplicate identical inventory rows without collapsing separate aliases."""
    seen = set()
    output = []
    for row in rows:
        key = tuple(norm(row.get(field)) for field in (
            "Object Type", "Query Name", "Source Type", "Server", "Database",
            "Schema", "Item", "Source", "Source File", "Notes",
        ))
        if key in seen:
            continue
        seen.add(key)
        output.append(row)
    return output


def inventory(paths: Paths):
    rows, unpivots = [], []
    _parameter_defs, current_parameters = parameter_definitions(paths)
    files = sorted(paths.tables.glob("*.tmdl")) + ([paths.expressions] if paths.expressions.exists() else [])
    for path in files:
        text_value = read_text(path)
        for query, block in split_expression_blocks(path, text_value):
            resolved_block = resolve_parameters_for_inventory(block, current_parameters)
            source_type, source, server, database, schema, item = detect_source(resolved_block)
            native = bool(native_sql_regions(block))
            merge = bool(re.search(r"Table\.(NestedJoin|FuzzyNestedJoin)", block))
            join = bool(re.search(r"Table\.Join", block))
            append = bool(re.search(r"Table\.Combine", block))
            unpivot = bool(re.search(r"Table\.Unpivot", block))
            status = "Disabled" if re.search(
                r"(?i)(loadEnabled|enableLoad)\s*[:=]\s*false|PBI_IsEnabled\s*=\s*false",
                block,
            ) else "Active"
            if native:
                query_type = "Native Query"
            elif merge:
                query_type = "Merge Query"
            elif append:
                query_type = "Append Query"
            elif source_type == "Excel":
                query_type = "Excel Import"
            elif source_type == "SharePoint":
                query_type = "SharePoint Import"
            else:
                query_type = "Table Navigation"

            origin = "expressions.tmdl" if path == paths.expressions else "tables/*.tmdl"
            relative_file = str(path.relative_to(paths.semantic.parent))
            join_flag = "Yes" if join else "No"
            merge_flag = "Yes" if merge else "No"
            append_flag = "Yes" if append else "No"
            unpivot_flag = "Yes" if unpivot else "No"

            # Query-level connector row.
            rows.append({
                "Object Type": "Expression" if path == paths.expressions else "Table",
                "Query Name": query,
                "Query Type": query_type,
                "Source Type": source_type,
                "Source": source,
                "Server": server,
                "Database": database,
                "Schema": schema,
                "Item": item,
                "Native Query": "Yes" if native else "No",
                "Query Status": status,
                "Origin": origin,
                "Source File": relative_file,
                "Contains Join": join_flag,
                "Contains Merge": merge_flag,
                "Contains Append": append_flag,
                "Contains Unpivot": unpivot_flag,
                "Migration Complexity": "High" if native else "Medium" if merge or append or unpivot else "Low",
                "Notes": "Connector/query-level inventory row.",
            })

            # Physical sources embedded inside [Query="..."] or Value.NativeQuery.
            if native:
                rows.extend(_native_source_inventory_rows(
                    query=query,
                    block=block,
                    source_file=relative_file,
                    origin=origin,
                    query_status=status,
                    server=server,
                    database=database,
                    contains_join=join_flag,
                    contains_merge=merge_flag,
                    contains_append=append_flag,
                    contains_unpivot=unpivot_flag,
                ))

            for unpivot_match in re.finditer(
                r"(?is)Table\.(UnpivotOtherColumns|Unpivot)\s*\((.*?)\)", block
            ):
                strings = re.findall(r'"((?:[^"\\]|\\.)*)"', unpivot_match.group(2))
                unpivots.append({
                    "Query Name": query,
                    "Source File": relative_file,
                    "Unpivot Function": "Table." + unpivot_match.group(1),
                    "Input Step": unpivot_match.group(2).split(",", 1)[0].strip(),
                    "Unpivoted Columns": ", ".join(strings[:-2]) if unpivot_match.group(1) == "Unpivot" else "",
                    "Retained Columns": ", ".join(strings[:-2]) if unpivot_match.group(1) == "UnpivotOtherColumns" else "",
                    "Attribute Column": strings[-2] if len(strings) >= 2 else "",
                    "Value Column": strings[-1] if strings else "",
                    "Legacy Column Name": "",
                    "Target Column Name": "",
                    "Legacy Name Becomes Row Value": "Yes",
                    "Legacy Name Retained": "Not evaluated",
                    "Retention Method": "",
                    "Validation Status": "Detected",
                    "Details": "Former column names become row values in the attribute column.",
                })
    return _deduplicate_inventory_rows(rows), unpivots


# =============================================================================
# PARAMETERS, COMPLETE INVENTORY, AND STORAGE MODES
# =============================================================================
@dataclass
class ParameterValue:
    name: str
    value: str
    is_current: bool
    data_type: str
    references: list[str] = field(default_factory=list)
    source_related: bool = False


def _unquote_m(value: str) -> str:
    value = clean(value).rstrip(',')
    if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        return value[1:-1].replace('""', '"')
    return value


def _parameter_scalar(token: str):
    token = clean(token).rstrip(',')
    quoted = re.fullmatch(r'"((?:[^"]|"")*)"', token, re.DOTALL)
    if quoted:
        return quoted.group(1).replace('""', '"'), 'Text'
    if re.fullmatch(r'(?i)true|false', token):
        return token.casefold(), 'Logical'
    if re.fullmatch(r'-?\d+', token):
        return token, 'Whole Number'
    if re.fullmatch(r'-?\d+\.\d+', token):
        return token, 'Decimal Number'
    if re.fullmatch(r'(?i)null', token):
        return '', 'Any'
    return None, ''

def _parameter_definition(name: str, block: str):
    marked = bool(re.search(r'(?i)IsParameterQuery\s*=\s*true', block))
    metadata = bool(re.search(r'(?i)(?:CurrentValue|DefaultValue|SuggestedValues|AllowedValues)\s*=', block))
    if not marked and not metadata:
        return None
    type_match = re.search(r'(?i)\bType\s*=\s*"([^"]+)"', block)
    data_type = type_match.group(1).strip() if type_match else ''
    header = re.search(r'(?is)^\s*expression\s+(?:#["](?:[^"]|"")+["]|.+?)\s*=\s*', block)
    body = block[header.end():] if header else block
    expression = re.split(r'(?i)\s+meta\s*\[', body, maxsplit=1)[0].strip()
    current, inferred = _parameter_scalar(expression)
    current = '' if current is None else current
    data_type = data_type or inferred
    current_match = re.search(
        r'(?i)(?:CurrentValue|DefaultValue)\s*=\s*("(?:[^"]|"")*"|-?\d+(?:\.\d+)?|true|false|null)',
        block,
    )
    if current_match:
        parsed, inferred = _parameter_scalar(current_match.group(1))
        if parsed is not None:
            current = parsed
            data_type = data_type or inferred
    values = []
    for list_match in re.finditer(r'\{([^{}]*)\}', block, re.DOTALL):
        for token in re.findall(r'"(?:[^"]|"")*"|-?\d+(?:\.\d+)?|true|false|null', list_match.group(1), re.I):
            parsed, inferred = _parameter_scalar(token)
            if parsed is not None and parsed not in values:
                values.append(parsed)
            data_type = data_type or inferred
    if current not in values:
        values.insert(0, current)
    return {'name': name, 'values': values or [''], 'current': current, 'data_type': data_type}

def _parameter_token(name: str) -> str:
    return rf'(?<![A-Za-z0-9_])(?:#"{re.escape(name)}"|{re.escape(name)})(?![A-Za-z0-9_])'

def parameter_definitions(paths: Paths):
    if not paths.expressions.exists():
        return {}, {}
    definitions = {}
    for name, block in split_expression_blocks(paths.expressions, read_text(paths.expressions)):
        definition = _parameter_definition(name, block)
        if definition:
            definitions[name] = definition
    current = {name: definition['current'] for name, definition in definitions.items()}
    return definitions, current

def resolve_parameters_for_inventory(block: str, current: dict[str, str]) -> str:
    resolved = block
    for name in sorted(current, key=len, reverse=True):
        value = current[name]
        if value == '':
            continue
        literal = '"' + value.replace('"', '""') + '"'
        resolved = re.sub(_parameter_token(name), lambda _match, replacement=literal: replacement, resolved, flags=re.I)
    return resolved

def _source_role_for_reference(block: str, parameter_name: str) -> str:
    token = _parameter_token(parameter_name)
    rules = (
        ('Server', rf'(?is)Sql\.Database\s*\(\s*{token}\s*,'),
        ('Database', rf'(?is)Sql\.Database\s*\([^,]+,\s*{token}(?:\s*[,\)])'),
        ('Server', rf'(?is)Databricks\.Catalogs\s*\(\s*{token}\s*,'),
        ('HTTP Path', rf'(?is)Databricks\.Catalogs\s*\([^,]+,\s*{token}(?:\s*[,\)])'),
        ('Catalog', rf'(?is)(?:Catalog|Database)\s*=\s*{token}'),
        ('Schema', rf'(?is)(?:Schema|SchemaName)\s*=\s*{token}'),
        ('Item', rf'(?is)(?:Item|Table|Name)\s*=\s*{token}'),
    )
    for role, pattern in rules:
        if re.search(pattern, block):
            return role
    return ''

def discover_parameters(paths: Paths):
    raw_definitions, current = parameter_definitions(paths)
    references = {name: set() for name in raw_definitions}
    source_related = {name: False for name in raw_definitions}
    roles = {}
    scan_files = sorted(paths.tables.glob('*.tmdl')) + ([paths.expressions] if paths.expressions.exists() else [])
    for path in scan_files:
        for query_name, block in split_expression_blocks(path, read_text(path)):
            for name in raw_definitions:
                if path == paths.expressions and query_name == name:
                    continue
                if not re.search(_parameter_token(name), block, re.I):
                    continue
                references[name].add(query_name)
                role = _source_role_for_reference(block, name)
                if role:
                    source_related[name] = True
                    roles[(name, query_name)] = role
    rows, values_by_name = [], {}
    for name in sorted(raw_definitions, key=str.casefold):
        definition = raw_definitions[name]
        values_by_name[name] = []
        for value in definition['values']:
            parameter = ParameterValue(
                name=name, value=value, is_current=(value == definition['current']),
                data_type=definition['data_type'], references=sorted(references[name], key=str.casefold),
                source_related=source_related[name],
            )
            values_by_name[name].append(parameter)
            for reference in parameter.references or ['']:
                rows.append({
                    'Parameter Name': name, 'Value': value,
                    'Is Current': 'Yes' if parameter.is_current else 'No',
                    'Data Type': parameter.data_type, 'Referenced By': reference,
                    '_Source Related': 'Yes' if parameter.source_related else 'No',
                })
    return rows, values_by_name, roles

def detect_table_modes(paths: Paths) -> tuple[dict[str,str], str]:
    modes={}
    for path in sorted(paths.tables.glob('*.tmdl')):
        text=read_text(path)
        name=tmdl_table_name(text, path.stem)
        found=[]
        for pattern in (
            r'(?mi)^\s*(?:storageMode|mode)\s*:\s*([A-Za-z][A-Za-z0-9_-]*)\s*$',
            r'(?mi)^\s*(?:storageMode|mode)\s*=\s*([A-Za-z][A-Za-z0-9_-]*)\s*$',
        ):
            found.extend(m.group(1) for m in re.finditer(pattern, text))
        normalized=[]
        for mode in found:
            aliases={'directquery':'DirectQuery','import':'Import','dual':'Dual','directlake':'DirectLake'}
            normalized.append(aliases.get(mode.casefold(), mode))
        unique=list(dict.fromkeys(normalized))
        modes[name]=unique[0] if len(unique)==1 else 'Mixed' if unique else ''
    model_modes=sorted({m for m in modes.values() if m})
    model_mode=model_modes[0] if len(model_modes)==1 else 'Mixed' if len(model_modes)>1 else 'Not specified'
    return modes, model_mode


def enrich_inventory(rows: list[dict], table_modes: dict[str,str], definitions: dict[str,list[ParameterValue]]) -> list[dict]:
    current={name: next((v.value for v in vals if v.is_current), '') for name,vals in definitions.items()}
    for row in rows:
        q=clean(row.get('Query Name'))
        row['Table Mode']=table_modes.get(q, '')
        # Resolve parameter tokens that may have survived literal-only connector detection.
        for field in ('Server','Database','Schema','Item','Source'):
            value=clean(row.get(field))
            key=value[2:-1] if value.startswith('#"') and value.endswith('"') else value
            if key in current and current[key]:
                row[field]=current[key]
    return rows


def _parameter_target(value: str, sources: list[SourceMapping]) -> str:
    """Return a unique authorized target for an exact legacy source component."""
    candidates=set()
    for s in sources:
        for old,new in (
            (s.legacy_source,s.target_source), (s.legacy_catalog,s.target_catalog),
            (s.legacy_schema,s.target_schema), (s.legacy_item,s.target_item),
        ):
            if old and new and norm(old)==norm(value) and norm(old)!=norm(new):
                candidates.add(new)
    return next(iter(candidates)) if len(candidates)==1 else ''


def migrate_current_source_parameters(value: str, path: Path, definitions: dict[str,list[ParameterValue]], sources: list[SourceMapping], log, status: str) -> str:
    """Change only current, source-related parameter literals with unique mapping matches."""
    if path.name.casefold() != 'expressions.tmdl':
        return value
    updated=value
    for name, values in definitions.items():
        current=next((v for v in values if v.is_current and v.source_related), None)
        if not current or not current.value:
            continue
        target=_parameter_target(current.value, sources)
        if not target:
            continue
        blocks=split_expression_blocks(path, updated)
        rebuilt=[]
        changed=False
        for block_name, block in blocks:
            if block_name != name:
                rebuilt.append(block)
                continue
            pattern=re.compile(r'(?m)(^\s*)("'+re.escape(current.value.replace('"','""'))+r'")(?=\s*(?:meta\s*\[|$))')
            new_literal='"'+target.replace('"','""')+'"'
            new_block,count=pattern.subn(lambda m:m.group(1)+new_literal, block, count=1)
            if count:
                changed=True
                rebuilt.append(new_block)
                log['M Query Changes'].append({
                    'Input File': path.name, 'Query Name': name, 'M Step': 'Parameter',
                    'Change Type': 'Current source parameter mapping', 'Old Value': current.value,
                    'New Value': target, 'Result': status,
                })
            else:
                rebuilt.append(block)
        if changed:
            updated=''.join(rebuilt)
    return updated


def scoped_replace(text: str, old: str, new: str):
    """Replace structured physical-source references only; never bare words."""
    if not old or not new or old == new:
        return text, 0
    output = text
    total = 0

    # Exact quoted source value, for example "Location" in M column lists.
    pattern = re.compile(r'"' + re.escape(old) + r'"', re.IGNORECASE)
    output, count = pattern.subn(lambda _: '"' + new + '"', output)
    total += count

    # Exact bracketed source column, for example [Location].
    pattern = re.compile(r'\[' + re.escape(old) + r'\]', re.IGNORECASE)
    output, count = pattern.subn(lambda _: '[' + new + ']', output)
    total += count

    # Qualified SQL source column, for example d.Location.
    pattern = re.compile(
        r'(?P<prefix>\b[A-Za-z_][A-Za-z0-9_]*\s*\.\s*)'
        r'(?P<name>' + re.escape(old) + r')\b',
        re.IGNORECASE,
    )
    output, count = pattern.subn(lambda m: m.group('prefix') + new, output)
    total += count
    return output, total

PROTECTED_DECLARATIONS = {
    "table": re.compile(r"(?mi)^\s*table\s+(.+?)\s*$"),
    "column": re.compile(r"(?mi)^\s*column\s+(.+?)\s*$"),
    "measure": re.compile(r"(?mi)^\s*measure\s+(.+?)\s*="),
    "hierarchy": re.compile(r"(?mi)^\s*hierarchy\s+(.+?)\s*$"),
    "expression": re.compile(r"(?mi)^\s*expression\s+(.+?)\s*=\s*$"),
}


def semantic_declarations(value: str):
    return {
        kind: [match.group(1).strip() for match in pattern.finditer(value)]
        for kind, pattern in PROTECTED_DECLARATIONS.items()
    }


def measure_blocks(value: str):
    starts = list(re.finditer(r"(?mi)^\s*measure\s+.+?\s*=", value))
    blocks = []
    for match in starts:
        following = re.search(
            r"(?mi)^\s*(?:measure|column|hierarchy|partition|annotation)\s+",
            value[match.end():],
        )
        end = match.end() + following.start() if following else len(value)
        blocks.append(value[match.start():end])
    return blocks


def validate_protected_content(path: Path, original: str, proposed: str, allowed_column_renames: dict[str, tuple[str, str]] | None = None):
    errors = []
    before = semantic_declarations(original)
    after = semantic_declarations(proposed)
    if allowed_column_renames is None:
        for kind in PROTECTED_DECLARATIONS:
            if before[kind] != after[kind]:
                errors.append(f"Protected {kind} names changed in {path.name}")
        if measure_blocks(original) != measure_blocks(proposed):
            errors.append(f"DAX measure content changed in {path.name}")
        return errors
    for kind in PROTECTED_DECLARATIONS:
        if kind == "column":
            expected_cols = []
            for raw_col in before["column"]:
                name_part = raw_col.split("=", 1)[0].strip()
                unq = _unquote_identifier(name_part)
                if norm(unq) in allowed_column_renames:
                    expected_cols.append(norm(allowed_column_renames[norm(unq)][1]))
                else:
                    expected_cols.append(norm(unq))
            actual_cols = [norm(_unquote_identifier(c.split("=", 1)[0].strip())) for c in after["column"]]
            if expected_cols != actual_cols:
                errors.append(f"Unauthorized column declaration change in {path.name}")
        elif before[kind] != after[kind]:
            errors.append(f"Protected {kind} names changed in {path.name}")
    return errors


def append_safety_errors(log, path: Path, errors):
    for message in errors:
        log["Errors"].append({
            "Input File": path.name,
            "Stage": "Protected-content validation",
            "Error Code": "SEMANTIC_CONTENT_CHANGED",
            "Error Message": message,
            "Existing Content Preserved": "Yes",
            "Processing Continued": "No",
        })

def extract_m_region(text: str):
    """Return the M partition/expression region and absolute boundaries."""
    start_match = re.search(r"(?mi)^\s*(?:source|expression\s+.+?)\s*=\s*$", text)
    if not start_match:
        return "", -1, -1
    start = start_match.end()
    end_match = re.search(
        r"(?mi)^\s*(?:partition|measure|column|hierarchy|annotation|relationship|table|expression)\s+",
        text[start:],
    )
    end = start + end_match.start() if end_match else len(text)
    return text[start:end], start, end


def _scan_m_string(text: str, quote_start: int):
    """Scan an M string, honoring doubled quotes and backslash escapes."""
    index = quote_start + 1
    while index < len(text):
        if text[index] == '"':
            if index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            return index
        if text[index] == "\\" and index + 1 < len(text):
            index += 2
            continue
        index += 1
    return -1


def native_sql_regions(text: str):
    """Find SQL payloads in Value.NativeQuery and [Query=] M expressions."""
    regions = []
    for marker in re.finditer(r"(?i)(Value\.NativeQuery\s*\(|\[\s*Query\s*=)", text):
        quote = text.find('"', marker.end())
        if quote < 0:
            continue
        end = _scan_m_string(text, quote)
        if end > quote:
            regions.append((quote + 1, end, text[quote + 1:end]))
    return regions


def _unquote_identifier(value: str) -> str:
    value = clean(value)
    if len(value) < 2:
        return value
    if value.startswith("'") and value.endswith("'"):
        return value[1:-1].replace("''", "'")
    if value.startswith("[") and value.endswith("]"):
        return value[1:-1].replace("]]", "]")
    if value.startswith('"') and value.endswith('"'):
        return value[1:-1].replace('""', '"')
    if value.startswith("`") and value.endswith("`"):
        return value[1:-1].replace("``", "`")
    return value


def _normalize_object_name(value: str) -> str:
    parts = [_unquote_identifier(part).casefold() for part in re.split(r"\s*\.\s*", clean(value)) if clean(part)]
    return ".".join(parts)


def _object_matches(found: str, expected: str) -> bool:
    found_norm = _normalize_object_name(found)
    expected_norm = _normalize_object_name(expected)
    if not found_norm or not expected_norm:
        return False
    return found_norm == expected_norm or found_norm.endswith("." + expected_norm) or expected_norm.endswith("." + found_norm)


def normalize_sql_dialect(value: str, fallback: str = "tsql") -> str:
    """Normalize user-friendly dialect names to SQLGlot dialect names."""
    value = norm(value).replace(" ", "").replace("_", "").replace("-", "")
    aliases = {
        "": fallback,
        "synapse": "tsql",
        "azuresynapse": "tsql",
        "sqlserver": "tsql",
        "mssql": "tsql",
        "tsql": "tsql",
        "databricks": "databricks",
        "databrickssql": "databricks",
        "spark": "spark",
        "sparksql": "spark",
    }
    return aliases.get(value, value or fallback)


def infer_sql_dialects(source_maps: list[SourceMapping], sql_text: str = "") -> tuple[str, str]:
    """Use workbook dialect settings first, then connector/source hints."""
    explicit_source = next((m.source_sql_dialect for m in source_maps if clean(m.source_sql_dialect)), "")
    explicit_target = next((m.target_sql_dialect for m in source_maps if clean(m.target_sql_dialect)), "")
    source_hint = " ".join(m.legacy_source for m in source_maps).casefold()
    target_hint = " ".join(m.target_source for m in source_maps).casefold()

    if not explicit_source:
        explicit_source = "databricks" if any(x in source_hint for x in ("databricks", "spark")) else "tsql"
    if not explicit_target:
        explicit_target = "databricks" if any(x in target_hint for x in ("databricks", "spark")) else "tsql"

    return normalize_sql_dialect(explicit_source), normalize_sql_dialect(explicit_target)


def decode_m_sql_string(value: str) -> str:
    """Decode common M string escapes before giving SQL to SQLGlot."""
    return (
        value.replace('""', '"')
        .replace('#(lf)', '\n')
        .replace('#(cr)', '\r')
        .replace('#(tab)', '\t')
    )


def encode_m_sql_string(value: str, original_payload: str) -> str:
    """Re-encode SQL using the same multiline convention as the original M string."""
    encoded = value.replace('"', '""')
    if '#(lf)' in original_payload:
        encoded = encoded.replace('\r\n', '\n').replace('\n', '#(lf)')
    if '#(cr)' in original_payload:
        encoded = encoded.replace('\r', '#(cr)')
    if '#(tab)' in original_payload:
        encoded = encoded.replace('\t', '#(tab)')
    return encoded


def table_parts(table: exp.Table) -> tuple[str, str, str]:
    return (
        clean(table.catalog),
        clean(table.db),
        clean(table.name),
    )


def mapping_object_parts(value: str) -> tuple[str, str, str]:
    parts = [_unquote_identifier(part) for part in re.split(r"\s*\.\s*", clean(value)) if clean(part)]
    if len(parts) == 1:
        return "", "", parts[0]
    if len(parts) == 2:
        return "", parts[0], parts[1]
    return parts[-3], parts[-2], parts[-1]


def table_matches_mapping(table: exp.Table, object_name: str) -> bool:
    expected_catalog, expected_schema, expected_name = mapping_object_parts(object_name)
    actual_catalog, actual_schema, actual_name = table_parts(table)
    if not expected_name or norm(actual_name) != norm(expected_name):
        return False
    if expected_schema and norm(actual_schema) != norm(expected_schema):
        return False
    if expected_catalog and norm(actual_catalog) != norm(expected_catalog):
        return False
    return True


def target_table_expression(mapping: MappingRow, source_maps: list[SourceMapping]) -> exp.Table:
    catalog, schema, name = mapping_object_parts(mapping.target_table)

    # A matching Data Source Mapping can supply catalog/schema/item independently.
    for source_map in source_maps:
        legacy_object = ".".join(filter(None, [source_map.legacy_catalog, source_map.legacy_schema, source_map.legacy_item]))
        if table_name_matches(mapping.legacy_table, legacy_object):
            catalog = source_map.target_catalog or catalog
            schema = source_map.target_schema or schema
            name = source_map.target_item or name
            break

    return exp.Table(
        this=exp.to_identifier(name),
        db=exp.to_identifier(schema) if schema else None,
        catalog=exp.to_identifier(catalog) if catalog else None,
    )


def table_name_matches(left: str, right: str) -> bool:
    left_parts = [norm(_unquote_identifier(part)) for part in clean(left).split('.') if clean(part)]
    right_parts = [norm(_unquote_identifier(part)) for part in clean(right).split('.') if clean(part)]
    if not left_parts or not right_parts:
        return False
    shorter, longer = (left_parts, right_parts) if len(left_parts) <= len(right_parts) else (right_parts, left_parts)
    return longer[-len(shorter):] == shorter


def projection_source_column(expression: exp.Expression) -> exp.Column | None:
    """Return a direct projected column, including a CAST(Column AS type)."""
    node = expression.this if isinstance(expression, exp.Alias) else expression
    if isinstance(node, exp.Column):
        return node
    if isinstance(node, (exp.Cast, exp.TryCast)) and isinstance(node.this, exp.Column):
        return node.this
    return None


def sqlglot_cast_type(legacy_type: str, target_dialect: str) -> str:
    normalized = normalize_type(legacy_type)
    if normalized == "integer":
        return "BIGINT"
    if normalized in {"decimal", "number"}:
        return "DOUBLE" if target_dialect in {"databricks", "spark"} else "FLOAT"
    if normalized == "date":
        return "DATE"
    if normalized == "datetime":
        return "TIMESTAMP" if target_dialect in {"databricks", "spark"} else "DATETIME2"
    if normalized == "boolean":
        return "BOOLEAN" if target_dialect in {"databricks", "spark"} else "BIT"
    return "STRING" if target_dialect in {"databricks", "spark"} else "NVARCHAR(MAX)"


def normalize_synapse_sql_functions(sql_text: str):
    rules = (
        (re.compile(r"(?is)\{\s*fn\s+CURDATE\s*\(\s*\)\s*\}"), "CAST(GETDATE() AS date)", "{fn CURDATE()}"),
        (re.compile(r"(?i)\bCURDATE\s*\(\s*\)"), "CAST(GETDATE() AS date)", "CURDATE()"),
        (re.compile(r"(?is)\{\s*fn\s+NOW\s*\(\s*\)\s*\}"), "GETDATE()", "{fn NOW()}"),
        (re.compile(r"(?i)\bCURRENT_DATE\b"), "CAST(GETDATE() AS date)", "CURRENT_DATE"),
    )
    output, audit = sql_text, []
    for pattern, target_value, source_label in rules:
        output, count = pattern.subn(target_value, output)
        if count:
            audit.append({"kind":"function", "old_value":source_label, "new_value":target_value, "occurrences":count, "source_dialect":"tsql", "target_dialect":"tsql"})
    return output, audit

def _physical_sources(scope):
    return {norm(alias): source for alias, source in scope.sources.items() if isinstance(source, exp.Table)}

def _one_table_mapping(table, mappings):
    """Resolve one unique legacy-to-target table pair from repeated column rows."""
    pairs = {}
    for mapping in mappings:
        if not table_matches_mapping(table, mapping.legacy_table):
            continue
        if not clean(mapping.legacy_table) or not clean(mapping.target_table):
            continue
        key = (norm(mapping.legacy_table), norm(mapping.target_table))
        pairs.setdefault(key, mapping)
    return next(iter(pairs.values())) if len(pairs) == 1 else None


def _table_mapping_conflicts(table, mappings):
    pairs = {
        (norm(mapping.legacy_table), norm(mapping.target_table))
        for mapping in mappings
        if clean(mapping.target_table)
        and table_matches_mapping(table, mapping.legacy_table)
    }
    return sorted(pairs) if len(pairs) > 1 else []

def _one_column_mapping(table, column_name, mappings):
    """Resolve a column mapping only inside its owning physical table."""
    if table is None or not clean(column_name):
        return None
    candidates = []
    for mapping in mappings:
        selection = mapping.selection if mapping.selection in {"Yes", "No", "Review"} else mapping.system_recommendation
        if selection == "Review" or not clean(mapping.target_column):
            continue
        if not table_matches_mapping(table, mapping.legacy_table):
            continue
        if norm(mapping.legacy_column) != norm(column_name):
            continue
        candidates.append(mapping)
    if not candidates:
        return None

    unique = {}
    for mapping in candidates:
        key = (
            norm(mapping.legacy_table), norm(mapping.target_table),
            norm(mapping.legacy_column), norm(mapping.target_column),
        )
        unique.setdefault(key, mapping)
    candidates = list(unique.values())

    # Identity mappings are authoritative for the physical source table.
    if any(norm(m.legacy_column) == norm(m.target_column) for m in candidates):
        return None
    target_names = {norm(m.target_column) for m in candidates}
    return candidates[0] if len(target_names) == 1 else None

def _projection_column(projection):
    node = projection.this if isinstance(projection, exp.Alias) else projection
    if isinstance(node, exp.Column): return node
    if isinstance(node, (exp.Cast, exp.TryCast)) and isinstance(node.this, exp.Column): return node.this
    return None

def _preserve_projection_outputs(select, changed_columns):
    aliases_added = casts_added = 0
    output = []
    output_renames = {}
    for projection in select.expressions:
        column = _projection_column(projection)
        change = changed_columns.get(id(column)) if column is not None else None
        if change:
            m = change["mapping"]
            old_name = change["old_name"]
            new_name = change.get("new_name") or m.target_column
            if not isinstance(projection, exp.Alias):
                value = projection.copy()
                selection = m.selection if m.selection in {"Yes", "No"} else m.system_recommendation
                cast_applied = False
                if not change.get("from_child_scope") and selection == "Yes" and normalize_type(m.legacy_type) != normalize_type(m.target_type):
                    value = exp.cast(value, sqlglot_cast_type(m.legacy_type, "tsql"))
                    casts_added += 1
                    cast_applied = True
                if ALIGN_TARGET_COLUMN_NAMES:
                    projection = exp.alias_(value, new_name, quoted=True) if cast_applied else value
                    if norm(old_name) != norm(new_name):
                        output_renames[norm(old_name)] = (new_name, m)
                else:
                    projection = exp.alias_(value, old_name, quoted=True)
                    aliases_added += 1
            elif ALIGN_TARGET_COLUMN_NAMES and norm(projection.alias) == norm(old_name) and norm(old_name) != norm(new_name):
                # Self-alias (e.g. d.Location AS Location): align alias to target name
                # and propagate to any enclosing subquery/CTE scope. Custom author aliases
                # (e.g. AVG(d.Crude_Benchmark_Reading) AS [Crude Benchmark]) are preserved.
                projection = exp.alias_(projection.this.copy(), new_name, quoted=True)
                output_renames[norm(old_name)] = (new_name, m)
        output.append(projection)
    select.set("expressions", output)
    return aliases_added, casts_added, output_renames

def _remaining_legacy_physical_refs(tree, mappings):
    problems = []
    for scope in traverse_scope(tree):
        tables = _physical_sources(scope)
        for column in scope.columns:
            qualifier = norm(column.table)
            table = tables.get(qualifier)
            if table is None and not qualifier:
                candidates = [t for t in tables.values() if _one_column_mapping(t, column.name, mappings)]
                table = candidates[0] if len(candidates) == 1 else None
            mapping = _one_column_mapping(table, column.name, mappings)
            if mapping:
                identity = ".".join(x for x in (clean(table.catalog), clean(table.db), clean(table.name)) if x)
                problems.append(f"Legacy physical reference remains: {identity}.{column.name}")
    return sorted(set(problems))

def rewrite_sql_ast(sql_text, mappings, source_maps):
    normalized_sql, audit = normalize_synapse_sql_functions(sql_text)
    errors = []
    try:
        tree = sqlglot.parse_one(normalized_sql, read="tsql")
    except ParseError as exc:
        return sql_text, audit, [f"SQLGlot could not parse normalized Synapse SQL: {exc}"]
    stars = list(tree.find_all(exp.Star))
    if stars:
        audit.append({"kind":"star", "occurrences":len(stars), "source_dialect":"tsql", "target_dialect":"tsql", "message":"SELECT * / alias.* preserved; mapped target table supplies columns at runtime"})
    scope_output_renames = {}
    for scope in traverse_scope(tree):
        tables = _physical_sources(scope)
        child_scopes = {
            norm(alias): src for alias, src in scope.sources.items()
            if not isinstance(src, exp.Table)
        }
        changed, grouped = {}, {}
        for column in list(scope.columns):
            qualifier = norm(column.table)
            table = tables.get(qualifier)
            mapping = _one_column_mapping(table, column.name, mappings)
            target_col_name = mapping.target_column if mapping is not None else ""
            from_child_scope = False

            # Check derived subquery / CTE child scope when qualified
            if mapping is None and qualifier and qualifier in child_scopes:
                child_renames = scope_output_renames.get(id(child_scopes[qualifier]), {})
                if norm(column.name) in child_renames:
                    target_col_name, mapping = child_renames[norm(column.name)]
                    from_child_scope = True

            # Unqualified column: check physical tables first, then child scopes
            if mapping is None and not qualifier:
                candidates = []
                for alias, candidate_table in tables.items():
                    candidate_mapping = _one_column_mapping(candidate_table, column.name, mappings)
                    if candidate_mapping:
                        candidates.append((alias, candidate_table, candidate_mapping))
                if len(candidates) == 1:
                    alias, table, mapping = candidates[0]
                    target_col_name = mapping.target_column
                    column.set("table", exp.to_identifier(alias))
                elif not candidates and child_scopes:
                    child_hits = []
                    for _calias, cscope in child_scopes.items():
                        cren = scope_output_renames.get(id(cscope), {})
                        if norm(column.name) in cren:
                            child_hits.append(cren[norm(column.name)])
                    unique_targets = {norm(tcol) for tcol, _m in child_hits}
                    if len(unique_targets) == 1 and child_hits:
                        target_col_name, mapping = child_hits[0]
                        from_child_scope = True
                    elif len(unique_targets) > 1:
                        errors.append(
                            f"AMBIGUOUS_SUBQUERY_COLUMN_RENAME: unqualified column '{column.name}' "
                            "maps to conflicting target names across child subqueries/CTEs."
                        )
                        continue

            if mapping is None or not target_col_name:
                continue
            old_name = column.name
            column.set("this", exp.to_identifier(target_col_name))
            changed[id(column)] = {"mapping": mapping, "old_name": old_name, "new_name": target_col_name, "from_child_scope": from_child_scope}
            if not from_child_scope:
                key = (mapping.legacy_table, mapping.legacy_column, mapping.target_table, mapping.target_column)
                item = grouped.setdefault(key, {"mapping": mapping, "aliases": set(), "occurrences": 0})
                item["aliases"].add(norm(column.table)); item["occurrences"] += 1

        select = scope.expression if isinstance(scope.expression, exp.Select) else None
        alias_count = cast_count = 0
        scope_renames = {}
        if select and changed:
            alias_count, cast_count, scope_renames = _preserve_projection_outputs(select, changed)
        # Propagate renames through SELECT * / alias.* or UNION branches to enclosing scopes
        if ALIGN_TARGET_COLUMN_NAMES and select:
            for proj in select.expressions:
                if isinstance(proj, exp.Star):
                    for _calias, cscope in child_scopes.items():
                        for old_k, val in scope_output_renames.get(id(cscope), {}).items():
                            scope_renames.setdefault(old_k, val)
                    for _talias, tbl in tables.items():
                        for m in mappings:
                            if table_matches_mapping(tbl, m.legacy_table) and clean(m.legacy_column) and clean(m.target_column) and norm(m.legacy_column) != norm(m.target_column):
                                scope_renames.setdefault(norm(m.legacy_column), (m.target_column, m))
        elif ALIGN_TARGET_COLUMN_NAMES and isinstance(scope.expression, exp.Union):
            for cscope in getattr(scope, "union_scopes", []) or []:
                for old_k, val in scope_output_renames.get(id(cscope), {}).items():
                    scope_renames.setdefault(old_k, val)
        if scope_renames:
            scope_output_renames[id(scope)] = scope_renames

        for item in grouped.values():
            m = item["mapping"]
            audit.append({"kind":"column", "legacy_table":m.legacy_table, "target_table":m.target_table, "legacy_column":m.legacy_column, "target_column":m.target_column, "aliases":", ".join(sorted(a for a in item["aliases"] if a)), "occurrences":item["occurrences"], "projection_aliases_added":alias_count, "casts_added":cast_count, "source_dialect":"tsql", "target_dialect":"tsql"})
    for table in list(tree.find_all(exp.Table)):
        conflicts = _table_mapping_conflicts(table, mappings)
        if conflicts:
            errors.append(
                f"Conflicting target table mappings for {table.sql(dialect='tsql')}: {conflicts}"
            )
            continue
        mapping = _one_table_mapping(table, mappings)
        if mapping is None:
            continue
        alias = table.args.get("alias")
        implicit_alias_required = _implicit_legacy_table_alias_required(tree, table)
        legacy_object_name = clean(table.name)
        target_table = target_table_expression(mapping, source_maps)
        if alias:
            target_table.set("alias", alias.copy())
        elif implicit_alias_required and legacy_object_name:
            target_table.set("alias", exp.TableAlias(this=exp.to_identifier(legacy_object_name)))
        table.replace(target_table)
        audit.append({
            "kind":"table", "legacy_table":mapping.legacy_table,
            "target_table":mapping.target_table, "alias":target_table.alias_or_name,
            "occurrences":1, "source_dialect":"tsql", "target_dialect":"tsql",
            "alias_preservation": "Existing explicit alias preserved" if alias else "Implicit object-name qualifier preserved explicitly" if implicit_alias_required else "No alias required",
        })
    try:
        rewritten = tree.sql(dialect="tsql", pretty=True)
        reparsed = sqlglot.parse_one(rewritten, read="tsql")
    except (ParseError, UnsupportedError) as exc:
        return sql_text, audit, [f"Generated Synapse SQL is invalid: {exc}"]
    if re.search(r"(?i)\bCURDATE\s*\(", rewritten) or re.search(r"(?is)\{\s*fn\s+CURDATE", rewritten): errors.append("CURDATE remained after Synapse function normalization")
    errors.extend(_remaining_legacy_physical_refs(reparsed, mappings))
    return rewritten, audit, errors

def replace_native_sql_payloads(
    value: str,
    mappings: list[MappingRow],
    source_maps: list[SourceMapping],
) -> tuple[str, list[dict], list[str]]:
    regions = native_sql_regions(value)
    if not regions:
        return value, [], []

    output = value
    all_audit = []
    all_errors = []
    for start, end, original_payload in reversed(regions):
        decoded_sql = decode_m_sql_string(original_payload)
        rewritten_sql, audit, errors = rewrite_sql_ast(decoded_sql, mappings, source_maps)
        all_audit.extend(audit)
        all_errors.extend(errors)
        if rewritten_sql != decoded_sql:
            encoded_sql = encode_m_sql_string(rewritten_sql, original_payload)
            output = output[:start] + encoded_sql + output[end:]

    return output, all_audit, all_errors

def _m_query_reference_tokens(value: str):
    """Mask external Power Query references so physical column rewrites cannot rename them."""
    region, start, end = extract_m_region(value)
    if start < 0:
        return value, {}
    local_steps = {
        norm(step.get("name", "").strip('#"'))
        for step in parse_m_steps(region)
        if clean(step.get("name"))
    }
    protected = {}
    masked = region
    candidates = []
    candidates.extend(re.finditer(r'#"((?:[^"]|"")+)"', region))
    candidates.extend(re.finditer(
        r'(?mi)^[ \t]*(?:#"(?:[^"]|"")+"|[A-Za-z_][A-Za-z0-9_ ]*)\s*=\s*([A-Za-z_][A-Za-z0-9_ ]*)\s*,?\s*$',
        region,
    ))
    spans = []
    for match in candidates:
        if match.re.pattern.startswith('#"'):
            name = match.group(1).replace('""', '"')
            token_text = match.group(0)
            span = match.span(0)
        else:
            name = match.group(1).strip()
            token_text = match.group(1)
            span = match.span(1)
        if norm(name) in local_steps:
            continue
        spans.append((span[0], span[1], token_text))
    for token_index, (token_start, token_end, token_text) in enumerate(sorted(spans, reverse=True)):
        placeholder = f"__PBI_QUERY_REFERENCE_{token_index}_{sha(token_text)[:12]}__"
        protected[placeholder] = token_text
        masked = masked[:token_start] + placeholder + masked[token_end:]
    return value[:start] + masked + value[end:], protected


def _restore_m_query_reference_tokens(value: str, protected: dict[str, str]) -> str:
    restored = value
    for placeholder, original in protected.items():
        restored = restored.replace(placeholder, original)
    return restored


def replace_non_native_m_column(value: str, legacy: str, target: str):
    """Update source-facing M references while preserving RenameColumns outputs and query references."""
    region, start, end = extract_m_region(value)
    if start < 0 or not legacy or not target or legacy == target:
        return value, 0
    masked_value, query_references = _m_query_reference_tokens(value)
    region, start, end = extract_m_region(masked_value)
    protected = {}
    masked_region = region
    pair_pattern = re.compile(
        r'(\{\s*"(?:[^"\r\n]|"")*"\s*,\s*)'
        r'("(?:[^"\r\n]|"")*")'
        r'(\s*\})',
        re.DOTALL,
    )
    for call in _balanced_m_calls(
        masked_region,
        ("Table.Combine", "Table.NestedJoin", "Table.Join", "Table.FuzzyNestedJoin", "Table.ExpandTableColumn"),
    ):
        token = f"__PBI_MULTI_TABLE_CALL_{len(protected)}_{sha(call['body'])[:12]}__"
        protected[token] = call["function"] + "(" + call["body"] + ")"
        masked_region = masked_region[:call["start"]] + token + masked_region[call["end"]:]

    for call in _balanced_m_calls(masked_region, ("Table.RenameColumns",)):
        args = split_m_top_level(call["body"])
        if len(args) < 2:
            continue
        rename_pairs = args[1]

        def protect_destination(match):
            token = f"__PBI_RENAME_OUTPUT_{len(protected)}_{sha(match.group(2))[:12]}__"
            protected[token] = match.group(2)
            return match.group(1) + token + match.group(3)

        masked_pairs = pair_pattern.sub(protect_destination, rename_pairs)
        if masked_pairs == rename_pairs:
            continue
        args[1] = masked_pairs
        rewritten = call["function"] + "(" + ", ".join(args) + ")"
        masked_region = masked_region[:call["start"]] + rewritten + masked_region[call["end"]:]

    updated = masked_region
    total = 0
    patterns = (
        (re.compile(r"\[" + re.escape(legacy) + r"\]", re.IGNORECASE), "[" + target + "]"),
        (re.compile(r'"' + re.escape(legacy) + r'"', re.IGNORECASE), '"' + target + '"'),
    )
    for pattern, replacement in patterns:
        updated, count = pattern.subn(lambda _: replacement, updated)
        total += count
    for token, original_output in protected.items():
        updated = updated.replace(token, original_output)
    result = masked_value[:start] + updated + masked_value[end:]
    result = _restore_m_query_reference_tokens(result, query_references)
    return result, total


def replace_tmdl_source_bindings(value: str, legacy_source: str, target_source: str):
    """Update matching sourceColumn properties while preserving semantic column names."""
    if not legacy_source or not target_source or legacy_source == target_source:
        return value, 0, []
    column_starts = list(re.finditer(r"(?mi)^\s*column\s+(.+?)\s*$", value))
    changes = []
    output = value
    for index in range(len(column_starts) - 1, -1, -1):
        match = column_starts[index]
        start = match.end()
        end = column_starts[index + 1].start() if index + 1 < len(column_starts) else len(value)
        block = value[start:end]
        source_match = re.search(r"(?mi)^(?P<prefix>[ \t]*sourceColumn[ \t]*:[ \t]*)(?P<name>[^\r\n]+?)[ \t]*$", block)
        if not source_match:
            continue
        current = _unquote_identifier(source_match.group("name"))
        if current.casefold() != _unquote_identifier(legacy_source).casefold():
            continue
        new_block = block[:source_match.start()] + source_match.group("prefix") + target_source + block[source_match.end():]
        output = output[:start] + new_block + output[end:]
        changes.append(match.group(1).strip())
    return output, len(changes), sorted(changes)


def validate_relationships(paths, proposed_table_files: dict, proposed_relationship_files: dict | None = None):
    declared = {}
    for path, value in proposed_table_files.items():
        table_match = re.search(r"(?mi)^\s*table\s+(.+?)\s*$", value)
        if not table_match:
            continue
        table_name = _unquote_identifier(table_match.group(1)).casefold()
        declared[table_name] = {
            _unquote_identifier(match.group(1)).casefold()
            for match in re.finditer(r"(?mi)^\s*column\s+(.+?)\s*$", value)
        }
    errors = []
    endpoint_pattern = re.compile(
        r"(?mi)^\s*(?:fromColumn|toColumn)\s*:\s*"
        r"(?:'([^']+)'|([^.\s]+))\."
        r"(?:'([^']+)'|([^\s]+))\s*$"
    )
    rel_overrides = proposed_relationship_files or {}
    for relationship_file in paths.relationships:
        relationship_text = rel_overrides.get(relationship_file, read_text(relationship_file))
        for match in endpoint_pattern.finditer(relationship_text):
            table_name = _unquote_identifier(match.group(1) or match.group(2))
            column_name = _unquote_identifier(match.group(3) or match.group(4))
            if table_name.casefold() not in declared:
                errors.append(f"Missing relationship table: {table_name}")
            elif column_name.casefold() not in declared[table_name.casefold()]:
                errors.append(f"Missing relationship column: {table_name}.{column_name}")
    return errors


def clean_preview_log(log):
    valid_files = {
        clean(row.get("Input File"))
        for tab in ("Table Changes", "Column Changes", "Native SQL Changes", "M Query Changes", "Semantic & Report Changes")
        for row in log.get(tab, [])
        if clean(row.get("Input File"))
    }
    for row in log.get("Files Processed", []):
        if clean(row.get("Input File")) not in valid_files and row.get("Changed") == "Proposed":
            row.update({
                "Relevant Mapping Found": "No",
                "Changed": "No",
                "Action": "Unchanged",
                "Write Status": "Not Written",
                "Message": "No applicable source-layer changes.",
                "Generated SHA-256": "",
            })


def unique_dependency_rows(rows):
    seen = set()
    output = []
    for row in rows:
        key = tuple(norm(row.get(name)) for name in (
            "Legacy Table", "Target Table", "Legacy Column", "Target Column", "Legacy Data Type", "Target Data Type"
        ))
        if key not in seen:
            seen.add(key)
            output.append(row)
    return output


def tmdl_table_name(value: str, fallback: str = "") -> str:
    match = re.search(r"(?mi)^\s*table\s+(.+?)\s*$", value)
    return _unquote_identifier(match.group(1)) if match else fallback


def is_native_query(value: str) -> bool:
    return bool(native_sql_regions(value))


def _mapping_applies_to_non_native(value: str, mapping: MappingRow, source_maps: list[SourceMapping]) -> bool:
    m_region, _, _ = extract_m_region(value)
    if not m_region:
        return False
    candidates = [mapping.legacy_table, mapping.target_table]
    for source_map in source_maps:
        if _object_matches(mapping.legacy_table, ".".join(filter(None, [source_map.legacy_schema, source_map.legacy_item]))):
            candidates.extend([source_map.legacy_item, source_map.target_item])
    return any(candidate and re.search(re.escape(candidate.split(".")[-1]), m_region, re.IGNORECASE) for candidate in candidates)



M_QUERY_FUNCTIONS = {
    "Table.NestedJoin": "merge",
    "Table.Join": "join",
    "Table.FuzzyNestedJoin": "fuzzy merge",
    "Table.Combine": "append",
    "Table.ExpandTableColumn": "expand",
    "Table.ExpandRecordColumn": "expand record",
    "Table.SelectColumns": "select columns",
    "Table.RemoveColumns": "remove columns",
    "Table.RenameColumns": "rename columns",
    "Table.TransformColumnTypes": "type conversion",
    "Table.Unpivot": "unpivot",
    "Table.UnpivotOtherColumns": "unpivot other columns",
}


def split_m_top_level(text: str, delimiter: str = ",") -> list[str]:
    """Split M arguments while respecting strings, braces, brackets, and parentheses."""
    parts, start = [], 0
    paren = brace = bracket = 0
    in_string = False
    index = 0
    while index < len(text):
        char = text[index]
        if in_string:
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    index += 2
                    continue
                in_string = False
            index += 1
            continue
        if char == '"':
            in_string = True
        elif char == '(':
            paren += 1
        elif char == ')':
            paren = max(0, paren - 1)
        elif char == '{':
            brace += 1
        elif char == '}':
            brace = max(0, brace - 1)
        elif char == '[':
            bracket += 1
        elif char == ']':
            bracket = max(0, bracket - 1)
        elif char == delimiter and paren == 0 and brace == 0 and bracket == 0:
            parts.append(text[start:index].strip())
            start = index + 1
        index += 1
    parts.append(text[start:].strip())
    return parts


def parse_m_steps(m_text: str) -> list[dict]:
    """Parse top-level let bindings without renaming or evaluating M."""
    body_match = re.search(r"(?is)\blet\b(?P<body>.*)\bin\b\s*(?P<output>#[^\r\n]+|[^\r\n]+)\s*$", m_text)
    if not body_match:
        return []
    steps = []
    for assignment in split_m_top_level(body_match.group("body")):
        match = re.match(r'(?is)^\s*(?P<name>#["](?:[^"]|"")+[\"]|[A-Za-z_][A-Za-z0-9_]*)\s*=\s*(?P<expr>.*)$', assignment)
        if not match:
            continue
        name = match.group("name").strip()
        expression = match.group("expr").strip()
        operations = [label for function, label in M_QUERY_FUNCTIONS.items() if function in expression]
        refs = set(re.findall(r'#["]((?:[^"]|"")+)[\"]', expression))
        refs.update(re.findall(r'(?<![.\w])([A-Za-z_][A-Za-z0-9_]*)\s*(?=[,})\]])', expression))
        refs.discard(name.strip('#"'))
        steps.append({"name": name, "expression": expression, "operations": operations, "references": sorted(refs)})
    return steps


def query_reference_names(m_text: str) -> set[str]:
    """Return quoted query references used by joins, merges, appends, or direct references."""
    names = set(re.findall(r'#["]((?:[^"]|"")+)[\"]', m_text))
    for function in ("Table.NestedJoin", "Table.Join", "Table.FuzzyNestedJoin", "Table.Combine"):
        for match in re.finditer(re.escape(function) + r"\s*\((.*?)\)", m_text, re.IGNORECASE | re.DOTALL):
            names.update(re.findall(r'#["]((?:[^"]|"")+)[\"]', match.group(1)))
    return names


def build_m_query_graph(file_texts: dict[Path, str]) -> dict[str, dict]:
    graph = {}
    for path, value in file_texts.items():
        region, _, _ = extract_m_region(value)
        if not region:
            continue
        query_name = tmdl_table_name(value, path.stem)
        steps = parse_m_steps(region)
        operations = sorted({operation for step in steps for operation in step["operations"]})
        graph[query_name.casefold()] = {
            "name": query_name,
            "path": path,
            "steps": steps,
            "references": {name.casefold() for name in query_reference_names(region)},
            "operations": operations,
        }
    return graph


def transitive_query_dependencies(graph: dict[str, dict], query_name: str) -> set[str]:
    visited, pending = set(), [query_name.casefold()]
    while pending:
        current = pending.pop()
        if current in visited:
            continue
        visited.add(current)
        pending.extend(graph.get(current, {}).get("references", set()) - visited)
    visited.discard(query_name.casefold())
    return visited


def mapped_query_names(graph: dict[str, dict], mappings: list[MappingRow], sources: list[SourceMapping]) -> set[str]:
    result = set()
    object_names = {
        clean(value).casefold()
        for mapping in mappings
        for value in (mapping.legacy_table, mapping.target_table)
        if clean(value)
    }
    object_names.update(
        clean(value).casefold()
        for source in sources
        for value in (source.legacy_item, source.target_item)
        if clean(value)
    )
    for key, node in graph.items():
        original = read_text(node["path"]).casefold()
        if any(name.split(".")[-1] in original for name in object_names):
            result.add(key)
    return result


def m_mapping_reaches_source(
    graph: dict[str, dict],
    mapped_queries: set[str],
    current_query: str,
    mapping: MappingRow,
    value: str,
    current_parameters=None,
) -> bool:
    direct = _direct_navigation_table_mapping(value, [mapping], current_parameters)
    if direct is None:
        return False
    return table_name_matches(direct.legacy_table, mapping.legacy_table)


def replace_m_expand_inputs(value: str, legacy: str, target: str):
    """Update only input-column lists of Table.ExpandTableColumn; preserve output names."""
    region, start, end = extract_m_region(value)
    if start < 0:
        return value, 0
    pattern = re.compile(r"(?is)Table\.ExpandTableColumn\s*\((.*?)\)")
    count = 0
    def rewrite(match):
        nonlocal count
        args = split_m_top_level(match.group(1))
        if len(args) < 3:
            return match.group(0)
        new_input = re.sub(r'"' + re.escape(legacy) + r'"', '"' + target + '"', args[2], flags=re.IGNORECASE)
        if new_input == args[2]:
            return match.group(0)
        args[2] = new_input
        count += 1
        return "Table.ExpandTableColumn(" + ", ".join(args) + ")"
    updated = pattern.sub(rewrite, region)
    return value[:start] + updated + value[end:], count


def replace_m_join_keys(value: str, legacy: str, target: str):
    """Update key lists in NestedJoin, Join, and FuzzyNestedJoin."""
    region, start, end = extract_m_region(value)
    if start < 0:
        return value, 0
    functions = r"Table\.(?:NestedJoin|Join|FuzzyNestedJoin)"
    pattern = re.compile(r"(?is)(" + functions + r")\s*\((.*?)\)")
    count = 0
    def rewrite(match):
        nonlocal count
        args = split_m_top_level(match.group(2))
        changed = False
        for index, argument in enumerate(args):
            if argument.strip().startswith("{"):
                replacement = re.sub(r'"' + re.escape(legacy) + r'"', '"' + target + '"', argument, flags=re.IGNORECASE)
                if replacement != argument:
                    args[index] = replacement
                    changed = True
        if not changed:
            return match.group(0)
        count += 1
        return match.group(1) + "(" + ", ".join(args) + ")"
    updated = pattern.sub(rewrite, region)
    return value[:start] + updated + value[end:], count


def replace_m_append_dependencies(value: str, legacy: str, target: str):
    """Update explicit column references around Table.Combine while preserving query names."""
    region, start, end = extract_m_region(value)
    if start < 0 or "Table.Combine" not in region:
        return value, 0
    updated = region
    patterns = (
        re.compile(r"\[" + re.escape(legacy) + r"\]", re.IGNORECASE),
        re.compile(r'"' + re.escape(legacy) + r'"', re.IGNORECASE),
    )
    total = 0
    for pattern in patterns:
        updated, count = pattern.subn(lambda match: match.group(0).replace(legacy, target), updated)
        total += count
    return value[:start] + updated + value[end:], total


def classify_file_with_lineage(value: str, graph_node: dict | None) -> str:
    if is_native_query(value):
        region, _, _ = extract_m_region(value)
        trailing_steps = parse_m_steps(region)
        extra_ops = sorted({op for step in trailing_steps for op in step["operations"]})
        return "Native SQL + M steps" if extra_ops else "Native SQL"
    operations = graph_node.get("operations", []) if graph_node else []
    return "Non-native M" + (" - " + ", ".join(operations) if operations else "")

def _m_escape(value: str) -> str:
    return clean(value).replace('"', '""')


def _m_field(value: str) -> str:
    return '#"' + _m_escape(value) + '"'


def _balanced_m_calls(value: str, function_names: tuple[str, ...]):
    """Locate M calls with nested-safe parenthesis/string scanning."""
    pattern = re.compile(
        r"(?i)(" + "|".join(re.escape(name) for name in function_names) + r")\s*\("
    )
    calls = []
    for match in pattern.finditer(value):
        opening = value.find("(", match.start(), match.end())
        depth, index, in_string = 1, opening + 1, False
        while index < len(value) and depth:
            char = value[index]
            if in_string:
                if char == '"':
                    if index + 1 < len(value) and value[index + 1] == '"':
                        index += 2
                        continue
                    in_string = False
                index += 1
                continue
            if char == '"':
                in_string = True
            elif char == '(':
                depth += 1
            elif char == ')':
                depth -= 1
            index += 1
        if depth == 0:
            calls.append({
                "function": match.group(1),
                "start": match.start(),
                "end": index,
                "body": value[opening + 1:index - 1],
            })
    return sorted(calls, key=lambda item: item["start"], reverse=True)


def _parse_static_text_list(expression: str):
    value = expression.strip()
    if not (value.startswith("{") and value.endswith("}")):
        return None
    result = []
    for item in split_m_top_level(value[1:-1]):
        match = re.fullmatch(r'"((?:[^"]|"")*)"', item.strip(), re.DOTALL)
        if not match:
            return None
        result.append(match.group(1).replace('""', '"'))
    return result


def _render_text_list(values: list[str]) -> str:
    return "{" + ", ".join('"' + _m_escape(value) + '"' for value in values) + "}"


def _applicable_unpivot_mappings(value, query_name, mappings, graph, mapped_queries, current_parameters=None):
    """Reuse existing query-lineage logic. No extra user mapping is required."""
    result = []
    for mapping in mappings:
        selection = mapping.selection if mapping.selection in {"Yes", "No", "Review"} else mapping.system_recommendation
        if selection == "Review" or not clean(mapping.legacy_column) or not clean(mapping.target_column):
            continue
        try:
            applicable = m_mapping_reaches_source(
                graph, mapped_queries, query_name, mapping, value, current_parameters
            )
        except Exception:
            m_region, _, _ = extract_m_region(value)
            applicable = any(
                table and table.split(".")[-1].casefold() in m_region.casefold()
                for table in (mapping.legacy_table, mapping.target_table)
            )
        if applicable:
            result.append(mapping)
    return result


def _wrap_attribute_preservation(unpivot_expression, attribute_column, replacements):
    """Restore legacy row values inline without creating or renaming an M step."""
    if not replacements:
        return unpivot_expression
    record = ", ".join(
        f'{_m_field(target)} = "{_m_escape(legacy)}"'
        for target, legacy in sorted(replacements.items(), key=lambda pair: norm(pair[0]))
    )
    return (
        "Table.TransformColumns(\n"
        f"    {unpivot_expression},\n"
        f'    {{{{"{_m_escape(attribute_column)}", '
        f"each Record.FieldOrDefault([{record}], _, _), type text}}}}\n"
        ")"
    )


def migrate_unpivot_preservation(
    value, query_name, input_file, mappings, graph, mapped_queries, native, status,
    extra_mappings=None, log=None, current_parameters=None,
):
    """Use target input names and preserve legacy Attribute values automatically."""
    region, start, end = extract_m_region(value)
    if start < 0:
        return value, [], []

    scoped = list(_applicable_unpivot_mappings(
        value, query_name, mappings, graph, mapped_queries, current_parameters
    ))
    if extra_mappings:
        seen_scoped = {
            (norm(m.legacy_table), norm(m.legacy_column), norm(m.target_column))
            for m in scoped
        }
        for m in extra_mappings:
            key = (norm(m.legacy_table), norm(m.legacy_column), norm(m.target_column))
            if key not in seen_scoped:
                scoped.append(m)
                seen_scoped.add(key)
    legacy_to_target = {
        norm(mapping.legacy_column): clean(mapping.target_column)
        for mapping in scoped
        if norm(mapping.legacy_column) != norm(mapping.target_column)
    }
    target_to_legacy = {
        norm(mapping.target_column): clean(mapping.legacy_column)
        for mapping in scoped
        if norm(mapping.legacy_column) != norm(mapping.target_column)
    }
    target_spelling = {
        norm(mapping.target_column): clean(mapping.target_column)
        for mapping in scoped
    }

    changed_region = region
    inventory, errors = [], []
    calls = _balanced_m_calls(
        changed_region, ("Table.Unpivot", "Table.UnpivotOtherColumns")
    )
    for call in calls:
        args = split_m_top_level(call["body"])
        if len(args) < 4:
            errors.append(f"{input_file}: malformed {call['function']} call.")
            continue
        columns = _parse_static_text_list(args[1])
        attribute_match = re.fullmatch(r'"((?:[^"]|"")*)"', args[2].strip(), re.DOTALL)
        value_match = re.fullmatch(r'"((?:[^"]|"")*)"', args[3].strip(), re.DOTALL)
        if columns is None or not attribute_match or not value_match:
            errors.append(
                f"{input_file}: dynamic {call['function']} arguments cannot be safely rewritten."
            )
            continue

        attribute_column = attribute_match.group(1).replace('""', '"')
        value_column = value_match.group(1).replace('""', '"')
        original_columns = list(columns)

        sql_uses_legacy_aliases = native and not ALIGN_TARGET_COLUMN_NAMES
        if not sql_uses_legacy_aliases:
            columns = [legacy_to_target.get(norm(column), column) for column in columns]
        input_updated = columns != original_columns

        if call["function"].casefold() == "table.unpivot":
            replacements = {
                target_spelling[norm(column)]: target_to_legacy[norm(column)]
                for column in columns
                if norm(column) in target_to_legacy
            }
        else:
            # UnpivotOtherColumns lists columns to preserve. Every applicable
            # mapped target column outside that list becomes an Attribute value.
            preserved = {norm(column) for column in columns}
            replacements = {
                clean(mapping.target_column): clean(mapping.legacy_column)
                for mapping in scoped
                if norm(mapping.target_column) != norm(mapping.legacy_column)
                and norm(mapping.target_column) not in preserved
                and not sql_uses_legacy_aliases
            }
            if log is not None and scoped:
                log["Warnings & Manual Review"].append({
                    "Input File": input_file,
                    "Warning Code": "UNPIVOT_OTHER_COLUMNS_EXTRA_TARGET_COLS",
                    "Category": "Unpivot",
                    "Object": query_name,
                    "Warning": (
                        f"{call['function']} is used on a mapped table. Any brand-new extra columns "
                        "present in the target table will be unpivoted into Attribute/Value rows."
                    ),
                    "Impact": "Unexpected new Attribute values may appear if the target table has extra columns.",
                    "Recommended Review": "Confirm whether a preceding Table.SelectColumns step is needed to restrict unpivoted columns.",
                })

        args[1] = _render_text_list(columns)
        unpivot = call["function"] + "(" + ", ".join(args) + ")"
        rewritten = _wrap_attribute_preservation(
            unpivot, attribute_column, replacements
        )
        changed_region = (
            changed_region[:call["start"]]
            + rewritten
            + changed_region[call["end"]:]
        )

        method = (
            "Native SQL output alias preservation"
            if sql_uses_legacy_aliases
            else "Inline target-to-legacy Attribute mapping"
        )
        if replacements:
            for target_column, legacy_value in sorted(replacements.items(), key=lambda pair: norm(pair[0])):
                inventory.append({
                    "Input File": input_file,
                    "Query Name": query_name,
                    "Unpivot Function": call["function"],
                    "Attribute Column": attribute_column,
                    "Value Column": value_column,
                    "Legacy Attribute Value": legacy_value,
                    "Target Source Column": target_column,
                    "Input Reference Updated": "Yes" if input_updated else "No",
                    "Attribute Preservation Required": "Yes",
                    "Attribute Preservation Method": method,
                    "Attribute Value Preserved": "Yes",
                    "Mappings Applied": f"{target_column} -> {legacy_value}",
                    "Result": status,
                })
        else:
            inventory.append({
                "Input File": input_file,
                "Query Name": query_name,
                "Unpivot Function": call["function"],
                "Attribute Column": attribute_column,
                "Value Column": value_column,
                "Legacy Attribute Value": "",
                "Target Source Column": "",
                "Input Reference Updated": "Yes" if input_updated else "No",
                "Attribute Preservation Required": "No",
                "Attribute Preservation Method": method,
                "Attribute Value Preserved": "Not Applicable",
                "Mappings Applied": "0",
                "Result": status,
            })

    return value[:start] + changed_region + value[end:], inventory, errors


def validate_unpivot_inventory(value, rows):
    """Block writes if a required inline preservation mapping is missing."""
    region, _, _ = extract_m_region(value)
    errors = []
    for row in rows:
        if row.get("Attribute Preservation Required") != "Yes":
            continue
        if row.get("Attribute Preservation Method") == "Native SQL output alias preservation":
            continue
        target = clean(row.get("Target Source Column"))
        legacy = clean(row.get("Legacy Attribute Value"))
        attribute = clean(row.get("Attribute Column"))
        required = (
            "Record.FieldOrDefault",
            _m_field(target),
            '"' + _m_escape(legacy) + '"',
            '"' + _m_escape(attribute) + '"',
        )
        if not all(fragment in region for fragment in required):
            errors.append(
                f"Unpivot preservation missing for {row.get('Query Name')}."
                f"{attribute}: {target} -> {legacy}"
            )
    return errors

def _native_sql_preserved_aliases(sql_audit):
    """Legacy output names intentionally preserved by native SQL projections."""
    if ALIGN_TARGET_COLUMN_NAMES:
        return set()
    return {
        norm(row.get("legacy_column"))
        for row in sql_audit
        if row.get("kind") == "column" and clean(row.get("legacy_column"))
    }


def _native_sql_target_bindings(sql_audit, output_contract: set[str], mappings: list[MappingRow]):
    """Return legacy->target column bindings exposed by rewritten native SQL."""
    if not ALIGN_TARGET_COLUMN_NAMES:
        return {}, []
    bindings, matched_rows = {}, []
    seen_rows = set()
    for row in sql_audit:
        if row.get("kind") != "column":
            continue
        legacy = clean(row.get("legacy_column"))
        target = clean(row.get("target_column"))
        legacy_table = clean(row.get("legacy_table"))
        if not legacy or not target or norm(legacy) == norm(target):
            continue
        if output_contract and norm(target) not in output_contract:
            continue
        bindings[norm(legacy)] = target
        for m in mappings:
            if (table_name_matches(m.legacy_table, legacy_table)
                    and norm(m.legacy_column) == norm(legacy)
                    and norm(m.target_column) == norm(target)):
                key = (norm(m.legacy_table), norm(m.legacy_column), norm(m.target_column))
                if key not in seen_rows:
                    seen_rows.add(key)
                    matched_rows.append(m)
    return bindings, matched_rows


def _native_mapping_crosses_alias_boundary(mapping, preserved_aliases):
    """False means downstream M must keep the legacy SQL output name."""
    return norm(mapping.legacy_column) not in preserved_aliases


def _inventory_native_source_keys(inventory_rows):
    return {
        (
            norm(row.get("Query Name")), norm(row.get("Schema")),
            norm(row.get("Item")), norm(row.get("Source File")),
        )
        for row in inventory_rows
        if row.get("Object Type") == "Native SQL Source"
    }


def validate_native_source_inventory(paths: Paths, inventory_rows):
    """Verify every SQLGlot physical source is represented in inventory."""
    errors = []
    keys = _inventory_native_source_keys(inventory_rows)
    files = sorted(paths.tables.glob("*.tmdl")) + ([paths.expressions] if paths.expressions.exists() else [])
    for path in files:
        value = read_text(path)
        relative_file = str(path.relative_to(paths.semantic.parent))
        for query, block in split_expression_blocks(path, value):
            for _, _, encoded_sql in native_sql_regions(block):
                decoded_sql = decode_m_sql_string(encoded_sql)
                sources, _, parse_errors = _sql_physical_sources(decoded_sql)
                if parse_errors:
                    errors.append(
                        f"Native SQL source inventory parse failed for {query}: {' | '.join(parse_errors)}"
                    )
                    continue
                for source in sources:
                    key = (
                        norm(query), norm(source.get("Schema")),
                        norm(source.get("Item")), norm(relative_file),
                    )
                    if key not in keys:
                        errors.append(
                            f"Native SQL source missing from inventory: {query} -> {source.get('Physical Object')}"
                        )
    return sorted(set(errors))



def _unique_table_pairs(mappings: list[MappingRow]):
    """Collapse repeated column rows into unique legacy/target physical table pairs."""
    pairs = {}
    conflicts = {}
    for mapping in mappings:
        if not clean(mapping.legacy_table) or not clean(mapping.target_table):
            continue
        legacy_key = _normalize_object_name(mapping.legacy_table)
        pair_key = (legacy_key, _normalize_object_name(mapping.target_table))
        pairs.setdefault(pair_key, mapping)
        conflicts.setdefault(legacy_key, set()).add(pair_key[1])
    valid = []
    for (legacy_key, _), mapping in pairs.items():
        if len(conflicts.get(legacy_key, set())) == 1:
            valid.append(mapping)
    return valid


def _navigation_record_pattern():
    """Match static Power Query navigation records regardless of field order."""
    return re.compile(
        r'(?is)Source\s*\{\s*\[(?P<fields>.*?)\]\s*\}\s*\[Data\]'
    )


def _m_record_field(fields: str, field_name: str):
    match = re.search(
        rf'(?i)\b{re.escape(field_name)}\s*=\s*"((?:[^\"]|\"\")*)"',
        fields,
    )
    return match.group(1).replace('""', '"') if match else ""


def _replace_m_record_field(fields: str, field_name: str, new_value: str):
    pattern = re.compile(
        rf'(?i)(\b{re.escape(field_name)}\s*=\s*")((?:[^\"]|\"\")*)(")'
    )
    return pattern.subn(
        lambda m: m.group(1) + _m_escape(new_value) + m.group(3),
        fields,
        count=1,
    )


def _table_mapping_for_navigation(schema: str, item: str, mappings: list[MappingRow]):
    """Resolve a static M navigation source to one unique table mapping."""
    found = []
    navigation_name = ".".join(part for part in (schema, item) if part)
    for mapping in _unique_table_pairs(mappings):
        if table_name_matches(mapping.legacy_table, navigation_name):
            found.append(mapping)
    targets = {_normalize_object_name(m.target_table) for m in found}
    return found[0] if len(targets) == 1 else None


def _column_mappings_for_table(mapping: MappingRow, mappings: list[MappingRow]):
    """Return only column mappings belonging to one physical legacy table lineage."""
    output = []
    seen = set()
    for candidate in mappings:
        selection = candidate.selection if candidate.selection in {"Yes", "No", "Review"} else candidate.system_recommendation
        if selection == "Review":
            continue
        if not table_name_matches(candidate.legacy_table, mapping.legacy_table):
            continue
        if not clean(candidate.legacy_column) or not clean(candidate.target_column):
            continue
        key = (norm(candidate.legacy_column), norm(candidate.target_column))
        if key in seen:
            continue
        seen.add(key)
        output.append(candidate)
    return output


def _migrate_navigation_block(
    block: str,
    query_name: str,
    input_file: str,
    mappings: list[MappingRow],
    source_maps: list[SourceMapping],
    log,
    status: str,
):
    """Migrate static Schema/Item navigation and table-scoped M columns in one query block."""
    changes = []
    matched_lineages = []

    def rewrite_navigation(match):
        fields = match.group("fields")
        legacy_schema = _m_record_field(fields, "Schema")
        legacy_item = _m_record_field(fields, "Item")
        table_mapping = _table_mapping_for_navigation(legacy_schema, legacy_item, mappings)
        if table_mapping is None:
            return match.group(0)

        target_catalog, target_schema, target_item = mapping_object_parts(table_mapping.target_table)
        for source_map in source_maps:
            legacy_object = ".".join(filter(None, [
                source_map.legacy_catalog,
                source_map.legacy_schema,
                source_map.legacy_item,
            ]))
            if table_name_matches(table_mapping.legacy_table, legacy_object):
                target_schema = source_map.target_schema or target_schema
                target_item = source_map.target_item or target_item
                break
        if not target_item:
            return match.group(0)

        new_fields = fields
        schema_count = item_count = 0
        if target_schema and norm(target_schema) != norm(legacy_schema):
            new_fields, schema_count = _replace_m_record_field(new_fields, "Schema", target_schema)
        if norm(target_item) != norm(legacy_item):
            new_fields, item_count = _replace_m_record_field(new_fields, "Item", target_item)
        if not (schema_count or item_count):
            return match.group(0)

        matched_lineages.append(table_mapping)
        changes.append({
            "mapping": table_mapping,
            "legacy": ".".join(filter(None, [legacy_schema, legacy_item])),
            "target": ".".join(filter(None, [target_schema, target_item])),
            "count": 1,
        })
        return match.group(0)[:match.start("fields") - match.start()] + new_fields + match.group(0)[match.end("fields") - match.start():]

    updated = _navigation_record_pattern().sub(rewrite_navigation, block)

    # Apply columns only from the physical table actually matched by this block.
    for lineage in matched_lineages:
        for column_mapping in _column_mappings_for_table(lineage, mappings):
            if norm(column_mapping.legacy_column) == norm(column_mapping.target_column):
                continue
            updated, count = replace_non_native_m_column(
                updated,
                column_mapping.legacy_column,
                column_mapping.target_column,
            )
            if count:
                log["M Query Changes"].append({
                    "Input File": input_file,
                    "Query Name": query_name,
                    "M Step": "Existing source-facing transformations",
                    "Change Type": "Table-scoped mapped column references",
                    "Old Value": column_mapping.legacy_column,
                    "New Value": column_mapping.target_column,
                    "Result": f"{status}; {count} occurrence(s)",
                })

    for change in changes:
        log["Table Changes"].append({
            "Input File": input_file,
            "File Type": "Power Query M",
            "Query or Model Table": query_name,
            "Change Location": "Power Query navigation record",
            "Old Table": change["legacy"],
            "New Table": change["target"],
            "SQL Alias": "",
            "Alias Preserved": "Not Applicable",
            "Replacement Count": change["count"],
            "Result": status,
        })
        log["M Query Changes"].append({
            "Input File": input_file,
            "Query Name": query_name,
            "M Step": "Source/Navigation",
            "Change Type": "Table mapping derived Schema/Item replacement",
            "Old Value": change["legacy"],
            "New Value": change["target"],
            "Result": f"{status}; 1 navigation record",
        })
    return updated, changes


def migrate_m_navigation_per_query(
    path: Path,
    value: str,
    mappings: list[MappingRow],
    source_maps: list[SourceMapping],
    log,
    status: str,
):
    """Process every expression in expressions.tmdl independently."""
    if path.name.casefold() != "expressions.tmdl":
        query_name = tmdl_table_name(value, path.stem)
        return _migrate_navigation_block(
            value, query_name, path.name, mappings, source_maps, log, status
        )[0]

    blocks = split_expression_blocks(path, value)
    if not blocks:
        return value
    rebuilt = []
    for query_name, block in blocks:
        migrated, _ = _migrate_navigation_block(
            block, query_name, path.name, mappings, source_maps, log, status
        )
        rebuilt.append(migrated)
    return "".join(rebuilt)


def validate_mapped_legacy_navigation(path: Path, value: str, mappings: list[MappingRow]):
    """Block writes when a mapped legacy Schema/Item navigation record remains."""
    errors = []
    for query_name, block in split_expression_blocks(path, value):
        for match in _navigation_record_pattern().finditer(block):
            fields = match.group("fields")
            schema = _m_record_field(fields, "Schema")
            item = _m_record_field(fields, "Item")
            mapping = _table_mapping_for_navigation(schema, item, mappings)
            if mapping is not None:
                errors.append(
                    f"MAPPED_LEGACY_SOURCE_REMAINS: {query_name} -> "
                    f"{'.'.join(filter(None, [schema, item]))}; expected {mapping.target_table}"
                )
    return errors

def mask_native_sql_payloads(value: str):
    """Mask already encoded native SQL so M rewriters cannot alter it."""
    regions = native_sql_regions(value)
    if not regions:
        return value, []
    masked = value
    payloads = []
    for ordinal, (start, end, payload) in enumerate(reversed(regions)):
        token = f"__PBI_NATIVE_SQL_{len(regions)-1-ordinal}_{sha(payload)[:16]}__"
        if token in value:
            raise ValueError("Native SQL mask token collision")
        payloads.append((token, payload))
        masked = masked[:start] + token + masked[end:]
    payloads.reverse()
    return masked, payloads


def restore_native_sql_payloads(value: str, payloads):
    restored, errors = value, []
    for token, payload in payloads:
        count = restored.count(token)
        if count != 1:
            errors.append(f"Native SQL placeholder {token} occurred {count} times; expected once")
            continue
        restored = restored.replace(token, payload, 1)
    return restored, errors


def native_sql_payload_fingerprints(value: str):
    return [sha(payload) for _, _, payload in native_sql_regions(value)]


def validate_native_sql_immutable(expected_sql_value: str, final_value: str):
    if native_sql_payload_fingerprints(expected_sql_value) != native_sql_payload_fingerprints(final_value):
        return ["NATIVE_SQL_MUTATED_BY_M_REWRITER: embedded SQL changed after SQLGlot rewrite"]
    return []


def _implicit_legacy_table_alias_required(tree, table: exp.Table) -> bool:
    if table.args.get("alias"):
        return False
    legacy_name = norm(table.name)
    return bool(legacy_name) and any(norm(column.table) == legacy_name for column in tree.find_all(exp.Column))

def _query_output_names(expression: exp.Expression) -> set[str]:
    """Return names exposed by the outermost native SQL query."""
    node = expression
    while isinstance(node, (exp.Subquery, exp.Paren)):
        node = node.this
    if isinstance(node, exp.Union):
        return _query_output_names(node.this)
    if not isinstance(node, exp.Select):
        node = node.find(exp.Select) if isinstance(node, exp.Expression) else None
    if node is None:
        return set()
    names = set()
    for projection in node.expressions:
        if isinstance(projection, exp.Star):
            continue
        name = clean(projection.alias_or_name)
        if name:
            names.add(norm(name))
    return names


def native_sql_output_contract(value: str):
    """Return normalized outer SELECT outputs from every native SQL payload."""
    names, errors = set(), []
    for index, (_, _, payload) in enumerate(native_sql_regions(value), 1):
        try:
            tree = sqlglot.parse_one(decode_m_sql_string(payload), read="tsql")
        except ParseError as exc:
            errors.append(f"Native SQL output contract parse failed for payload {index}: {exc}")
            continue
        names.update(_query_output_names(tree))
    return names, errors


def _quoted_m_values(expression: str) -> list[str]:
    return [value.replace('""', '"') for value in re.findall(r'"((?:[^\"]|\"\")*)"', expression)]


def validate_post_native_m_contract(value: str, output_contract: set[str], mappings: list[MappingRow]):
    """Reject post-native M references to hidden physical target columns."""
    if not output_contract:
        return []
    region, _, _ = extract_m_region(value)
    if not region:
        return []
    functions = (
        "Table.RenameColumns", "Table.TransformColumnTypes", "Table.SelectColumns",
        "Table.RemoveColumns", "Table.ReorderColumns", "Table.DuplicateColumn",
        "Table.NestedJoin", "Table.Join", "Table.FuzzyNestedJoin", "Table.Group",
        "Table.Unpivot", "Table.UnpivotOtherColumns",
    )
    calls = _balanced_m_calls(region, functions)
    errors = []
    for mapping in mappings:
        legacy, target = norm(mapping.legacy_column), norm(mapping.target_column)
        if not legacy or not target or legacy == target:
            continue
        if legacy not in output_contract or target in output_contract:
            continue
        for call in calls:
            if target in {norm(item) for item in _quoted_m_values(call["body"])}:
                errors.append(
                    "POST_NATIVE_M_COLUMN_NOT_EXPOSED: "
                    f"{call['function']} references {mapping.target_column}, but native SQL "
                    f"exposes preserved alias {mapping.legacy_column}."
                )
    return sorted(set(errors))


def _syn_record_field_expression(record_text: str, field_name: str) -> str:
    marker = re.search(rf'(?i)\b{re.escape(field_name)}\s*=\s*', record_text)
    if not marker:
        return ''
    start = marker.end()
    paren = brace = bracket = 0
    in_string = False
    index = start
    while index < len(record_text):
        char = record_text[index]
        if in_string:
            if char == '"':
                if index + 1 < len(record_text) and record_text[index + 1] == '"':
                    index += 2; continue
                in_string = False
            index += 1; continue
        if char == '"': in_string = True
        elif char == '(': paren += 1
        elif char == ')': paren = max(0, paren - 1)
        elif char == '{': brace += 1
        elif char == '}': brace = max(0, brace - 1)
        elif char == '[': bracket += 1
        elif char == ']':
            if paren == 0 and brace == 0 and bracket == 0:
                break
            bracket = max(0, bracket - 1)
        elif char == ',' and paren == 0 and brace == 0 and bracket == 0:
            break
        index += 1
    return record_text[start:index].strip()


def _direct_navigation_table_mapping(value: str, mappings: list[MappingRow], current_parameters=None):
    """Resolve the physical table directly navigated by one non-native query (literal or parameterized)."""
    region, _, _ = extract_m_region(value)
    if not region:
        return None
    params = current_parameters or {}

    def _resolve_field(fields: str, field_name: str) -> str:
        val = _m_record_field(fields, field_name)
        if val:
            return val
        if params:
            expr = _syn_record_field_expression(fields, field_name)
            ident = _parse_query_identifier(expr)
            if ident:
                actual = next((k for k in params if norm(k) == norm(ident)), "")
                if actual and params[actual]:
                    return params[actual]
        return ""

    for match in _navigation_record_pattern().finditer(region):
        fields = match.group("fields")
        schema = _resolve_field(fields, "Schema")
        item = _resolve_field(fields, "Item")
        if not item:
            continue
        mapping = _table_mapping_for_navigation(schema, item, mappings)
        if mapping is not None:
            return mapping
        # Recovery/idempotency: navigation may already use the target object.
        object_name = ".".join(part for part in (schema, item) if part)
        candidates = [
            pair for pair in _unique_table_pairs(mappings)
            if table_name_matches(pair.target_table, object_name)
        ]
        targets = {_normalize_object_name(pair.target_table) for pair in candidates}
        if len(targets) == 1:
            return candidates[0]
    return None


def _tmdl_declared_column_names(value: str) -> set[str]:
    """Return normalized TMDL column and sourceColumn names declared in a query block."""
    names = {
        norm(_unquote_identifier(m.group(1)))
        for m in re.finditer(r'(?mi)^\s*column\s+(.+?)\s*$', value)
    }
    names.update(
        norm(_unquote_identifier(m.group(1)))
        for m in re.finditer(r'(?mi)^\s*sourceColumn\s*:\s*(.+?)\s*$', value)
    )
    return {n for n in names if n}


def _query_direct_column_map(value: str, mappings: list[MappingRow], current_parameters=None):
    """Return legacy-to-target columns for a query's own physical navigation or native SQL."""
    if is_native_query(value):
        if not ALIGN_TARGET_COLUMN_NAMES:
            return {}
        result, conflicts = {}, set()
        for _, _, encoded_sql in native_sql_regions(value):
            sources, _audit, _errors = _sql_physical_sources(decode_m_sql_string(encoded_sql))
            for row in sources:
                physical = row.get("Physical Object") or ".".join(
                    p for p in (row.get("Schema"), row.get("Item")) if p
                )
                for mapping in mappings:
                    selection = mapping.selection if mapping.selection in {"Yes", "No", "Review"} else mapping.system_recommendation
                    if selection == "Review" or not clean(mapping.legacy_column) or not clean(mapping.target_column):
                        continue
                    if not table_name_matches(mapping.legacy_table, physical):
                        continue
                    legacy = norm(mapping.legacy_column)
                    target = clean(mapping.target_column)
                    if legacy == norm(target):
                        continue
                    if legacy in result and norm(result[legacy]) != norm(target):
                        conflicts.add(legacy)
                    else:
                        result[legacy] = target
        if conflicts:
            declared_sources = _tmdl_declared_column_names(value)
            for legacy in list(conflicts):
                # Prefer the target column actually exposed by this query's TMDL or outer SELECT
                matched_targets = set()
                for _, _, encoded_sql in native_sql_regions(value):
                    sources, _audit, _errors = _sql_physical_sources(decode_m_sql_string(encoded_sql))
                    for row in sources:
                        physical = row.get("Physical Object") or ".".join(
                            p for p in (row.get("Schema"), row.get("Item")) if p
                        )
                        for m in mappings:
                            if table_name_matches(m.legacy_table, physical) and norm(m.legacy_column) == legacy and clean(m.target_column):
                                if norm(m.target_column) in declared_sources:
                                    matched_targets.add(clean(m.target_column))
                if len(matched_targets) == 1:
                    result[legacy] = next(iter(matched_targets))
                else:
                    result.pop(legacy, None)
        return result
    table_mapping = _direct_navigation_table_mapping(value, mappings, current_parameters)
    if table_mapping is None:
        return {}
    result = {}
    conflicts = set()
    for mapping in _column_mappings_for_table(table_mapping, mappings):
        legacy = norm(mapping.legacy_column)
        target = clean(mapping.target_column)
        if not legacy or not target or legacy == norm(target):
            continue
        if legacy in result and norm(result[legacy]) != norm(target):
            conflicts.add(legacy)
        else:
            result[legacy] = target
    for legacy in conflicts:
        result.pop(legacy, None)
    region, _, _ = extract_m_region(value)
    if region and result:
        for call in reversed(_balanced_m_calls(region, ("Table.RenameColumns",))):
            args = split_m_top_level(call["body"])
            if len(args) < 2:
                continue
            for m_pair in re.finditer(
                r'\{\s*"((?:[^"\r\n]|"")*)"\s*,\s*"((?:[^"\r\n]|"")*)"\s*\}',
                args[1], re.DOTALL,
            ):
                src_col = m_pair.group(1).replace('""', '"')
                if norm(src_col) in result:
                    result.pop(norm(src_col), None)
    return result


def _combine_binding_maps(binding_maps: list[dict[str, str]]):
    combined, conflicts = {}, set()
    for binding_map in binding_maps:
        for legacy, current in binding_map.items():
            key = norm(legacy)
            if not key or not clean(current):
                continue
            if key in combined and norm(combined[key]) != norm(current):
                conflicts.add(key)
            else:
                combined[key] = current
    for key in conflicts:
        combined.pop(key, None)
    return combined, conflicts


def _direct_reference_dependencies(value: str):
    region, _, _ = extract_m_region(value)
    if not region:
        return []
    local = {norm(step.get("name", "").strip('#"')) for step in parse_m_steps(region)}
    result = []
    for step in parse_m_steps(region):
        name = _parse_query_identifier(step.get("expression", ""))
        if name and norm(name) not in local:
            result.append(name)
    return result


def _append_dependencies(value: str):
    region, _, _ = extract_m_region(value)
    if not region:
        return []
    result = []
    for call in _balanced_m_calls(region, ("Table.Combine",)):
        args = split_m_top_level(call["body"])
        if not args:
            continue
        members = args[0].strip()
        if members.startswith("{") and members.endswith("}"):
            for member in split_m_top_level(members[1:-1]):
                name = _parse_query_identifier(member)
                if name:
                    result.append(name)
    return result


def _all_merge_base_queries(value: str, originals_by_query: dict[str, str]) -> list[str]:
    """Return external left-side base queries across single or chained merge calls."""
    region, _, _ = extract_m_region(value)
    if not region:
        return []
    local_steps = {norm(step.get("name", "").strip('#"')) for step in parse_m_steps(region)}
    calls = sorted(
        _balanced_m_calls(region, ("Table.NestedJoin", "Table.Join", "Table.FuzzyNestedJoin")),
        key=lambda item: item["start"],
    )
    bases, seen = [], set()
    for call in calls:
        args = split_m_top_level(call["body"])
        if len(args) < 4:
            continue
        left_name = _parse_query_identifier(args[0])
        if left_name and norm(left_name) not in local_steps and left_name.casefold() in originals_by_query:
            if left_name.casefold() not in seen:
                seen.add(left_name.casefold())
                bases.append(left_name)
    return bases


def _all_expanded_output_bindings(value: str, originals_by_query: dict[str, str], mappings: list[MappingRow], binding_cache=None, visiting=None, current_parameters=None):
    """Resolve expanded output bindings across ALL Table.ExpandTableColumn calls in a query."""
    region, _, _ = extract_m_region(value)
    if not region:
        return {}, []
    nested = _merge_nested_lineage(value)
    result, errors = {}, []
    for call in sorted(_balanced_m_calls(region, ("Table.ExpandTableColumn",)), key=lambda item: item["start"]):
        args = split_m_top_level(call["body"])
        if len(args) < 3:
            continue
        nested_match = re.fullmatch(r'"((?:[^"]|"")*)"', args[1].strip(), re.DOTALL)
        inputs = _parse_static_text_list(args[2])
        outputs = _parse_static_text_list(args[3]) if len(args) >= 4 else inputs
        if not nested_match or inputs is None or outputs is None or len(inputs) != len(outputs):
            continue
        nested_name = nested_match.group(1).replace('""', '"')
        right_query = nested.get(norm(nested_name), "")
        if not right_query:
            continue
        right_map, local_errors = resolved_query_output_bindings(
            right_query, originals_by_query, mappings, binding_cache, visiting, current_parameters
        )
        errors.extend(local_errors)
        if not right_map:
            continue
        reverse_right = {norm(t): l for l, t in right_map.items()}
        for input_name, output_name in zip(inputs, outputs):
            if ALIGN_TARGET_COLUMN_NAMES and norm(output_name) == norm(input_name):
                if norm(input_name) in right_map:
                    target_input = right_map[norm(input_name)]
                    if norm(target_input) != norm(output_name):
                        result[norm(output_name)] = target_input
                elif norm(input_name) in reverse_right:
                    legacy_key = reverse_right[norm(input_name)]
                    if legacy_key != norm(input_name):
                        result[legacy_key] = input_name
    return result, sorted(set(errors))


def resolved_query_output_bindings(query_name, originals_by_query, mappings, cache=None, visiting=None, current_parameters=None):
    """Resolve names exposed by direct sources, references, chained merges, and (when aligned) appends."""
    cache = cache if cache is not None else {}
    visiting = visiting if visiting is not None else set()
    key = query_name.casefold()
    if key in cache:
        return dict(cache[key]), []
    if key in visiting:
        return {}, [f"QUERY_DEPENDENCY_CYCLE: {query_name}"]
    value = originals_by_query.get(key, "")
    if not value:
        return {}, []
    visiting.add(key)
    errors, maps = [], []
    direct = _query_direct_column_map(value, mappings, current_parameters)
    if direct:
        maps.append(direct)
    for dependency in _direct_reference_dependencies(value):
        dep_map, dep_errors = resolved_query_output_bindings(
            dependency, originals_by_query, mappings, cache, visiting, current_parameters)
        errors.extend(dep_errors)
        if dep_map:
            maps.append(dep_map)
    for base_query in _all_merge_base_queries(value, originals_by_query):
        dep_map, dep_errors = resolved_query_output_bindings(
            base_query, originals_by_query, mappings, cache, visiting, current_parameters)
        errors.extend(dep_errors)
        if dep_map:
            maps.append(dep_map)
    expanded_map, expand_errors = _all_expanded_output_bindings(
        value, originals_by_query, mappings, cache, visiting, current_parameters
    )
    errors.extend(expand_errors)
    if expanded_map:
        maps.append(expanded_map)
    if ALIGN_TARGET_COLUMN_NAMES:
        for append_branch in _append_dependencies(value):
            dep_map, dep_errors = resolved_query_output_bindings(
                append_branch, originals_by_query, mappings, cache, visiting, current_parameters)
            errors.extend(dep_errors)
            if dep_map:
                maps.append(dep_map)
    combined, conflicts = _combine_binding_maps(maps)
    if conflicts:
        errors.append("QUERY_OUTPUT_MAPPING_CONFLICT: " + ", ".join(sorted(conflicts)))
    visiting.remove(key)
    cache[key] = dict(combined)
    return combined, sorted(set(errors))


def migrate_join_keys_with_lineage(value, query_name, originals_by_query, mappings, binding_cache=None, current_parameters=None):
    """Rewrite left and right join-key lists independently from branch lineage."""
    region, start, end = extract_m_region(value)
    if start < 0:
        return value, [], []
    current_bindings, errors = resolved_query_output_bindings(
        query_name, originals_by_query, mappings, binding_cache, None, current_parameters)
    changed_region, changes = region, []
    for call in _balanced_m_calls(changed_region, ("Table.NestedJoin", "Table.Join", "Table.FuzzyNestedJoin")):
        args = split_m_top_level(call["body"])
        if len(args) < 4:
            continue
        left_query = _parse_query_identifier(args[0])
        right_query = _parse_query_identifier(args[2])
        left_map = current_bindings
        if left_query and left_query.casefold() in originals_by_query:
            left_map, local_errors = resolved_query_output_bindings(
                left_query, originals_by_query, mappings, binding_cache, None, current_parameters)
            errors.extend(local_errors)
        right_map = {}
        if right_query:
            right_map, local_errors = resolved_query_output_bindings(
                right_query, originals_by_query, mappings, binding_cache, None, current_parameters)
            errors.extend(local_errors)
        left_keys = _parse_static_text_list(args[1])
        right_keys = _parse_static_text_list(args[3])
        if left_keys is None or right_keys is None or len(left_keys) != len(right_keys):
            continue
        new_left = [left_map.get(norm(name), name) for name in left_keys]
        new_right = [right_map.get(norm(name), name) for name in right_keys]
        if new_left == left_keys and new_right == right_keys:
            continue
        args[1], args[3] = _render_text_list(new_left), _render_text_list(new_right)
        rewritten = call["function"] + "(" + ", ".join(args) + ")"
        changed_region = changed_region[:call["start"]] + rewritten + changed_region[call["end"]:]
        changes.append({"function":call["function"], "left_query":left_query or query_name,
            "right_query":right_query, "old_left":left_keys, "new_left":new_left,
            "old_right":right_keys, "new_right":new_right})
    return value[:start] + changed_region + value[end:], changes, sorted(set(errors))


def migrate_append_contracts(value, query_name, originals_by_query, mappings, binding_cache=None, current_parameters=None):
    """Reconcile Table.Combine members with forward adaptation when aligned."""
    region, start, end = extract_m_region(value)
    if start < 0:
        return value, [], []
    changed_region, changes, errors = region, [], []
    for call in _balanced_m_calls(changed_region, ("Table.Combine",)):
        args = split_m_top_level(call["body"])
        if not args:
            continue
        members_text = args[0].strip()
        if not (members_text.startswith("{") and members_text.endswith("}")):
            continue
        members = split_m_top_level(members_text[1:-1])
        if not ALIGN_TARGET_COLUMN_NAMES:
            continue
        branch_info = []
        unified_targets = {}
        target_conflicts = {}
        for member in members:
            branch = _parse_query_identifier(member)
            binding_map = {}
            if branch:
                binding_map, local_errors = resolved_query_output_bindings(
                    branch, originals_by_query, mappings, binding_cache, None, current_parameters
                )
                errors.extend(local_errors)
            branch_info.append((member, branch, binding_map))
            for legacy_key, target_col in binding_map.items():
                key = norm(legacy_key)
                if key in unified_targets and norm(unified_targets[key]) != norm(target_col):
                    target_conflicts.setdefault(key, {unified_targets[key]}).add(target_col)
                else:
                    unified_targets[key] = target_col
        if target_conflicts:
            for key, targets in sorted(target_conflicts.items()):
                errors.append(
                    f"APPEND_BRANCH_TARGET_CONFLICT: {query_name} Table.Combine branches map "
                    f"'{key}' to conflicting target columns: {', '.join(sorted(targets))}"
                )
            continue
        if not unified_targets:
            continue
        rewritten_members, local_changes = [], []
        for member, branch, binding_map in branch_info:
            if not branch:
                rewritten_members.append(member)
                continue
            branch_text = originals_by_query.get(branch.casefold(), "")
            declared_cols = _tmdl_declared_column_names(branch_text) if branch_text else set()
            branch_target_outputs = {norm(v) for v in binding_map.values()}
            pairs = []
            for legacy_key, target_col in sorted(unified_targets.items()):
                if norm(target_col) in branch_target_outputs:
                    continue
                if legacy_key in binding_map and norm(binding_map[legacy_key]) == norm(target_col):
                    continue
                legacy_spelling = next(
                    (m.legacy_column for m in mappings
                     if norm(m.legacy_column) == legacy_key
                     and norm(m.target_column) == norm(target_col)),
                    legacy_key,
                )
                if norm(legacy_spelling) == norm(target_col):
                    continue
                if declared_cols and norm(target_col) in declared_cols and legacy_key in declared_cols:
                    errors.append(
                        f"APPEND_FORWARD_RENAME_COLLISION: branch '{branch}' in {query_name} "
                        f"already declares both '{legacy_spelling}' and '{target_col}'."
                    )
                    continue
                pairs.append((legacy_spelling, target_col))
            if not pairs:
                rewritten_members.append(member)
                continue
            rendered = ", ".join(
                '{"' + _m_escape(old_col) + '", "' + _m_escape(new_col) + '"}'
                for old_col, new_col in pairs
            )
            adapted = "Table.RenameColumns(" + member.strip() + ", {" + rendered + "}, MissingField.Ignore)"
            rewritten_members.append(adapted)
            local_changes.append({"branch": branch, "pairs": pairs, "direction": "forward"})
        if not local_changes:
            continue
        args[0] = "{" + ", ".join(rewritten_members) + "}"
        rewritten = "Table.Combine(" + ", ".join(args) + ")"
        changed_region = changed_region[:call["start"]] + rewritten + changed_region[call["end"]:]
        changes.extend(local_changes)
    return value[:start] + changed_region + value[end:], changes, sorted(set(errors))


def _parse_query_identifier(expression: str):
    value = expression.strip()
    quoted = re.fullmatch(r'#"((?:[^"]|"")+)"', value)
    if quoted:
        return quoted.group(1).replace('""', '"')
    plain = re.fullmatch(r'[A-Za-z_][A-Za-z0-9_ ]*', value)
    return plain.group(0).strip() if plain else ""


def _merge_lineage_for_query(value: str):
    """Return the left/right query names from the first static merge call."""
    region, _, _ = extract_m_region(value)
    if not region:
        return "", ""
    calls = _balanced_m_calls(
        region, ("Table.NestedJoin", "Table.Join", "Table.FuzzyNestedJoin")
    )
    if not calls:
        return "", ""
    # _balanced_m_calls is reverse sorted, so use the earliest call in the query.
    call = sorted(calls, key=lambda item: item["start"])[0]
    args = split_m_top_level(call["body"])
    if len(args) < 4:
        return "", ""
    return _parse_query_identifier(args[0]), _parse_query_identifier(args[2])


def _expanded_output_map(value: str, right_column_map: dict[str, str]):
    """Map only columns actually introduced by right-side ExpandTableColumn calls."""
    region, _, _ = extract_m_region(value)
    result = {}
    if not region or not right_column_map:
        return result
    for call in _balanced_m_calls(region, ("Table.ExpandTableColumn",)):
        args = split_m_top_level(call["body"])
        if len(args) < 3:
            continue
        inputs = _parse_static_text_list(args[2])
        outputs = _parse_static_text_list(args[3]) if len(args) >= 4 else inputs
        if inputs is None or outputs is None or len(inputs) != len(outputs):
            continue
        for input_name, output_name in zip(inputs, outputs):
            target_input = right_column_map.get(norm(input_name), input_name)
            # Preserve explicit output aliases in ExpandTableColumn. If the output
            # name matches the input name, follow the target column; otherwise
            # retain the explicit output name.
            target_output = (
                target_input if norm(output_name) == norm(input_name)
                else output_name
            )
            result[norm(output_name)] = target_output
    return result


def _replace_row_field_references(expression: str, replacements: dict[str, str]):
    """Rewrite [Field] and [#"Field"] row references inside one M expression."""
    updated = expression
    total = 0
    # Longest legacy names first avoids partial-like surprises in diagnostics.
    for legacy_norm, target in sorted(replacements.items(), key=lambda item: -len(item[0])):
        def quoted_repl(match):
            nonlocal total
            name = match.group(1).replace('""', '"')
            if norm(name) != legacy_norm:
                return match.group(0)
            total += 1
            return '[#"' + _m_escape(target) + '"]'
        updated = re.sub(r'\[#"((?:[^"]|"")+)"\]', quoted_repl, updated)

        def plain_repl(match):
            nonlocal total
            name = match.group(1)
            if norm(name) != legacy_norm:
                return match.group(0)
            total += 1
            if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_]*', target):
                return '[' + target + ']'
            return '[#"' + _m_escape(target) + '"]'
        updated = re.sub(r'\[([A-Za-z_][A-Za-z0-9_ ]*)\]', plain_repl, updated)
    return updated, total


def migrate_merged_addcolumn_references(
    value: str,
    originals_by_query: dict[str, str],
    mappings: list[MappingRow],
    binding_cache=None,
    current_parameters=None,
):
    """Propagate merged upstream target names into Table.AddColumn row formulas across chained merges."""
    left_query, right_query = _merge_lineage_for_query(value)
    base_queries = _all_merge_base_queries(value, originals_by_query)
    if not left_query and not right_query and not base_queries:
        return value, [], []
    left_maps = []
    for base_q in (base_queries or ([left_query] if left_query else [])):
        bm, _ = resolved_query_output_bindings(
            base_q, originals_by_query, mappings, binding_cache, None, current_parameters
        )
        if bm:
            left_maps.append(bm)
    left_map, _ = _combine_binding_maps(left_maps)
    expanded_map, _ = _all_expanded_output_bindings(
        value, originals_by_query, mappings, binding_cache, None, current_parameters
    )
    replacements = dict(left_map)
    replacements.update(expanded_map)
    if not replacements:
        return value, [], []

    region, start, end = extract_m_region(value)
    if start < 0:
        return value, [], []
    changes = []
    changed_region = region
    for call in _balanced_m_calls(changed_region, ("Table.AddColumn",)):
        args = split_m_top_level(call["body"])
        if len(args) < 3:
            continue
        old_formula = args[2]
        new_formula, count = _replace_row_field_references(old_formula, replacements)
        if not count:
            continue
        args[2] = new_formula
        rewritten = "Table.AddColumn(" + ", ".join(args) + ")"
        changed_region = changed_region[:call["start"]] + rewritten + changed_region[call["end"]:]
        changes.append({
            "left_query": left_query,
            "right_query": right_query,
            "old_formula": old_formula,
            "new_formula": new_formula,
            "count": count,
        })
    return value[:start] + changed_region + value[end:], changes, []


def _query_text_index(file_texts: dict[Path, str]):
    """Index every table and expressions.tmdl block by query name."""
    result = {}
    for path, value in file_texts.items():
        for query_name, block in split_expression_blocks(path, value):
            result[query_name.casefold()] = block
    return result

def _merge_nested_lineage(value: str):
    region, _, _ = extract_m_region(value)
    result = {}
    if not region:
        return result
    for call in _balanced_m_calls(region, ("Table.NestedJoin", "Table.FuzzyNestedJoin")):
        args = split_m_top_level(call["body"])
        if len(args) < 5:
            continue
        right_query = _parse_query_identifier(args[2])
        nested_match = re.fullmatch(r'"((?:[^"]|"")*)"', args[4].strip(), re.DOTALL)
        if right_query and nested_match:
            result[norm(nested_match.group(1).replace('""', '"'))] = right_query
    return result


def migrate_join_expansions(value: str, originals_by_query: dict[str, str], mappings: list[MappingRow], binding_cache=None, current_parameters=None):
    region, start, end = extract_m_region(value)
    if start < 0:
        return value, [], {}, []
    nested = _merge_nested_lineage(value)
    changed_region, changes, bindings, errors = region, [], {}, []
    for call in _balanced_m_calls(changed_region, ("Table.ExpandTableColumn",)):
        args = split_m_top_level(call["body"])
        if len(args) < 3:
            continue
        nested_match = re.fullmatch(r'"((?:[^"]|"")*)"', args[1].strip(), re.DOTALL)
        inputs = _parse_static_text_list(args[2])
        outputs = _parse_static_text_list(args[3]) if len(args) >= 4 else inputs
        if not nested_match or inputs is None or outputs is None or len(inputs) != len(outputs):
            continue
        nested_name = nested_match.group(1).replace('""', '"')
        right_query = nested.get(norm(nested_name), "")
        right_map, _ = resolved_query_output_bindings(right_query, originals_by_query, mappings, binding_cache, None, current_parameters) if right_query else ({}, [])
        if not right_map:
            continue
        new_inputs, new_outputs, local = [], [], {}
        for legacy_input, legacy_output in zip(inputs, outputs):
            target_input = right_map.get(norm(legacy_input), legacy_input)
            target_output = target_input if norm(legacy_output) == norm(legacy_input) else legacy_output
            new_inputs.append(target_input); new_outputs.append(target_output)
            if norm(target_output) != norm(legacy_output):
                local[norm(legacy_output)] = target_output
        if new_inputs == inputs and new_outputs == outputs:
            continue
        args[2] = _render_text_list(new_inputs)
        if len(args) >= 4:
            args[3] = _render_text_list(new_outputs)
        rewritten = "Table.ExpandTableColumn(" + ", ".join(args) + ")"
        changed_region = changed_region[:call["start"]] + rewritten + changed_region[call["end"]:]
        bindings.update(local)
        changes.append({"right_query":right_query,"old_inputs":inputs,"new_inputs":new_inputs,"old_outputs":outputs,"new_outputs":new_outputs})
    return value[:start]+changed_region+value[end:], changes, bindings, errors


def _merged_query_output_bindings(value: str, originals_by_query: dict[str, str], mappings: list[MappingRow], binding_cache=None, current_parameters=None):
    binding_maps = []
    for base_q in _all_merge_base_queries(value, originals_by_query):
        left_map, _ = resolved_query_output_bindings(
            base_q, originals_by_query, mappings, binding_cache, None, current_parameters
        )
        if left_map:
            binding_maps.append(left_map)
    expanded_map, _ = _all_expanded_output_bindings(
        value, originals_by_query, mappings, binding_cache, None, current_parameters
    )
    if expanded_map:
        binding_maps.append(expanded_map)
    combined, _conflicts = _combine_binding_maps(binding_maps)
    return combined


def propagate_bindings_to_downstream_m(value: str, bindings: dict[str, str]):
    updated, changes = value, []
    for legacy_norm, target in bindings.items():
        region, _, _ = extract_m_region(updated)
        candidates = re.findall(r'"((?:[^"]|"")*)"', region)
        candidates += re.findall(r'\[#"((?:[^"]|"")*)"\]', region)
        candidates += re.findall(r'\[([A-Za-z_][A-Za-z0-9_ ]*)\]', region)
        legacy = next((c.replace('""','"') for c in candidates if norm(c.replace('""','"')) == legacy_norm), "")
        if not legacy:
            continue
        updated, count = replace_non_native_m_column(updated, legacy, target)
        if count:
            changes.append((legacy, target, count))
    return updated, changes


def apply_query_output_bindings(value: str, bindings: dict[str, str]):
    updated, changes = value, []
    for legacy_norm, target in bindings.items():
        source_values = [_unquote_identifier(m.group(1)) for m in re.finditer(r'(?mi)^\s*sourceColumn\s*:\s*(.+?)\s*$', updated)]
        current = next((v for v in source_values if norm(v) == legacy_norm), "")
        if not current or norm(current) == norm(target):
            continue
        updated, count, semantic_columns = replace_tmdl_source_bindings(updated, current, target)
        if count:
            changes.append({"legacy":current,"target":target,"semantic_columns":semantic_columns})
    return updated, changes


# =============================================================================
# PHASE 3: SEMANTIC MODEL & REPORT LAYER COLUMN ALIGNMENT
# =============================================================================
import json

TMDL_RESERVED_KEYWORDS = {
    "table", "column", "measure", "hierarchy", "level", "partition",
    "annotation", "relationship", "model", "role", "culture", "perspective",
    "expression", "perspectivetable", "perspectivecolumn", "tablepermission",
    "calculationgroup", "calculationitem", "true", "false", "null", "ref",
    "lineageTag", "sourceColumn", "sortByColumn", "dataType", "formatString",
}

DAX_ITERATOR_FUNCTIONS = {
    fn.casefold() for fn in {
        "sumx", "averagex", "minx", "maxx", "countx", "countax",
        "filter", "addcolumns", "selectcolumns", "summarize", "summarizecolumns",
        "generate", "generateall", "rankx", "concatenatex", "productx",
        "medianx", "percentilex.inc", "percentilex.exc", "stdevx.p", "stdevx.s",
        "varx.p", "varx.s", "topn", "sample", "groupby",
    }
}


def format_tmdl_identifier(name: str) -> str:
    """Format an identifier according to TMDL quoting rules."""
    raw = _unquote_identifier(name)
    if not raw:
        return raw
    needs_quotes = (
        not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", raw)
        or raw.casefold() in {k.casefold() for k in TMDL_RESERVED_KEYWORDS}
    )
    if needs_quotes:
        return "'" + raw.replace("'", "''") + "'"
    return raw


def format_tmdl_relationship_endpoint(table_name: str, column_name: str) -> str:
    return f"{format_tmdl_identifier(table_name)}.{format_tmdl_identifier(column_name)}"


def _tmdl_table_columns_with_source(value: str) -> list[dict]:
    """Return declared columns in a table TMDL with their semantic name and sourceColumn."""
    starts = list(re.finditer(r"(?mi)^(?P<indent>[ \t]*)column[ \t]+(?P<name>[^\r\n=]+?)[ \t]*$", value))
    columns = []
    for idx, match in enumerate(starts):
        start = match.end()
        end = starts[idx + 1].start() if idx + 1 < len(starts) else len(value)
        next_decl = re.search(r"(?mi)^[ \t]*(?:measure|hierarchy|partition|annotation|table)\s+", value[start:end])
        if next_decl:
            end = start + next_decl.start()
        block = value[start:end]
        src_match = re.search(r"(?mi)^[ \t]*sourceColumn[ \t]*:[ \t]*(?P<src>[^\r\n]+?)[ \t]*$", block)
        columns.append({
            "semantic": _unquote_identifier(match.group("name")),
            "raw_name": match.group("name").strip(),
            "source": _unquote_identifier(src_match.group("src")) if src_match else "",
            "is_calculated": False,
        })
    return columns


def _tmdl_table_member_names(value: str) -> dict[str, set[str]]:
    """Return normalized names of columns, measures, and hierarchies in a table TMDL."""
    cols = {
        norm(_unquote_identifier(m.group(1)))
        for m in re.finditer(r"(?mi)^[ \t]*column[ \t]+([^=\r\n]+?)(?:[ \t]*=.*)?[ \t]*$", value)
    }
    measures = {
        norm(_unquote_identifier(m.group(1)))
        for m in re.finditer(r"(?mi)^[ \t]*measure[ \t]+(.+?)[ \t]*=", value)
    }
    hierarchies = {
        norm(_unquote_identifier(m.group(1)))
        for m in re.finditer(r"(?mi)^[ \t]*hierarchy[ \t]+(.+?)[ \t]*$", value)
    }
    return {"columns": cols, "measures": measures, "hierarchies": hierarchies}


def build_semantic_rename_plan(
    originals: dict[Path, str],
    proposed: dict[Path, str],
    tables_dir: Path,
):
    """Derive the authoritative per-table semantic column rename plan from Phase 2 sourceColumn updates."""
    rename_plan: dict[str, dict[str, tuple[str, str]]] = {}
    table_display_names: dict[str, str] = {}
    table_measures: dict[str, set[str]] = {}
    errors: list[str] = []

    for path, orig_text in originals.items():
        if path.parent != tables_dir:
            continue
        prop_text = proposed.get(path, orig_text)
        tbl_name = tmdl_table_name(orig_text, path.stem)
        tbl_key = norm(tbl_name)
        table_display_names[tbl_key] = tbl_name
        members = _tmdl_table_member_names(orig_text)
        table_measures[tbl_key] = members["measures"]

        orig_cols = {norm(c["semantic"]): c for c in _tmdl_table_columns_with_source(orig_text)}
        prop_cols = {norm(c["semantic"]): c for c in _tmdl_table_columns_with_source(prop_text)}

        planned_for_table: dict[str, tuple[str, str]] = {}
        target_owners: dict[str, list[str]] = {}

        for col_key, orig_info in orig_cols.items():
            prop_info = prop_cols.get(col_key)
            if not prop_info:
                continue
            old_src = orig_info["source"]
            new_src = prop_info["source"]
            sem_name = orig_info["semantic"]
            if not old_src or not new_src or norm(old_src) == norm(new_src):
                continue
            # Preserve custom business-friendly semantic names!
            if norm(sem_name) != norm(old_src):
                continue
            planned_for_table[col_key] = (sem_name, new_src)
            target_owners.setdefault(norm(new_src), []).append(sem_name)

        if not planned_for_table:
            continue

        valid_for_table: dict[str, tuple[str, str]] = {}
        renamed_away = set(planned_for_table.keys())
        for col_key, (old_sem, new_sem) in planned_for_table.items():
            new_key = norm(new_sem)
            if len(target_owners.get(new_key, [])) > 1:
                errors.append(
                    f"SEMANTIC_COLUMN_RENAME_COLLISION: {tbl_name} has multiple columns "
                    f"({', '.join(sorted(target_owners[new_key]))}) mapping to '{new_sem}'."
                )
                continue
            existing_other_cols = members["columns"] - renamed_away
            if new_key in existing_other_cols or new_key in members["measures"] or new_key in members["hierarchies"]:
                errors.append(
                    f"SEMANTIC_COLUMN_RENAME_COLLISION: {tbl_name}.{old_sem} -> {new_sem} "
                    f"collides with an existing column, measure, or hierarchy in {tbl_name}."
                )
                continue
            valid_for_table[col_key] = (old_sem, new_sem)

        if valid_for_table:
            rename_plan[tbl_key] = valid_for_table

    return rename_plan, table_display_names, table_measures, sorted(set(errors))


def _dax_parse_first_arg_table(text: str, open_paren_idx: int) -> str:
    """Extract the bare/quoted table name if it is the first argument of a DAX iterator."""
    i = open_paren_idx + 1
    n = len(text)
    while i < n and text[i].isspace():
        i += 1
    if i >= n:
        return ""
    if text[i] == "'":
        j = i + 1
        while j < n:
            if text[j] == "'":
                if j + 1 < n and text[j + 1] == "'":
                    j += 2
                    continue
                name = text[i + 1:j].replace("''", "'")
                k = j + 1
                while k < n and text[k].isspace():
                    k += 1
                if k < n and text[k] == ",":
                    return name
                return ""
            j += 1
        return ""
    m = re.match(r"([A-Za-z_][A-Za-z0-9_]*)\s*,", text[i:])
    if m:
        return m.group(1)
    m_wrap = re.match(
        r"(?i)(?:ALL|ALLSELECTED|ALLEXCEPT|VALUES|DISTINCT|FILTER)\s*\(\s*(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_]*))",
        text[i:],
    )
    if m_wrap:
        return m_wrap.group(1).replace("''", "'") if m_wrap.group(1) else m_wrap.group(2)
    return ""


def rewrite_dax_expression(
    dax: str,
    rename_plan: dict[str, dict[str, tuple[str, str]]],
    current_table: str = "",
    table_measures: dict[str, set[str]] | None = None,
):
    """Lexer-aware DAX column rewriter (strings, comments, qualified, iterator-scoped unqualified)."""
    if not dax or not rename_plan:
        return dax, [], []

    measures_by_table = table_measures or {}
    out = []
    changes = []
    warnings = []
    i = 0
    n = len(dax)
    paren_depth = 0
    iterator_stack: list[tuple[int, str]] = []

    def _active_row_table() -> str:
        return iterator_stack[-1][1] if iterator_stack else current_table

    def _encode_bracket(col: str) -> str:
        return "[" + col.replace("]", "]]") + "]"

    while i < n:
        ch = dax[i]
        nxt = dax[i + 1] if i + 1 < n else ""

        if (ch == "-" and nxt == "-") or (ch == "/" and nxt == "/"):
            end = dax.find("\n", i + 2)
            end = n if end < 0 else end
            out.append(dax[i:end])
            i = end
            continue

        if ch == "/" and nxt == "*":
            end = dax.find("*/", i + 2)
            end = n if end < 0 else end + 2
            out.append(dax[i:end])
            i = end
            continue

        if ch == '"':
            j = i + 1
            while j < n:
                if dax[j] == '"':
                    if j + 1 < n and dax[j + 1] == '"':
                        j += 2
                        continue
                    j += 1
                    break
                j += 1
            out.append(dax[i:j])
            i = j
            continue

        if ch == "(":
            paren_depth += 1
            out.append(ch)
            i += 1
            continue
        if ch == ")":
            if iterator_stack and iterator_stack[-1][0] == paren_depth:
                iterator_stack.pop()
            paren_depth = max(0, paren_depth - 1)
            out.append(ch)
            i += 1
            continue

        if ch == "'":
            j = i + 1
            while j < n:
                if dax[j] == "'":
                    if j + 1 < n and dax[j + 1] == "'":
                        j += 2
                        continue
                    break
                j += 1
            if j < n and dax[j] == "'":
                tbl_token = dax[i:j + 1]
                tbl_name = dax[i + 1:j].replace("''", "'")
                k = j + 1
                while k < n and dax[k] in " \t":
                    k += 1
                if k < n and dax[k] == "[":
                    m_end = k + 1
                    while m_end < n:
                        if dax[m_end] == "]":
                            if m_end + 1 < n and dax[m_end + 1] == "]":
                                m_end += 2
                                continue
                            break
                        m_end += 1
                    if m_end < n and dax[m_end] == "]":
                        col_name = dax[k + 1:m_end].replace("]]", "]")
                        tbl_map = rename_plan.get(norm(tbl_name), {})
                        if norm(col_name) in tbl_map:
                            old_sem, new_sem = tbl_map[norm(col_name)]
                            out.append(tbl_token + dax[j + 1:k] + _encode_bracket(new_sem))
                            changes.append((f"{tbl_name}[{old_sem}]", f"{tbl_name}[{new_sem}]"))
                        else:
                            out.append(dax[i:m_end + 1])
                        i = m_end + 1
                        continue
                out.append(tbl_token)
                i = j + 1
                continue

        if ch.isalpha() or ch == "_":
            j = i + 1
            while j < n and (dax[j].isalnum() or dax[j] in "_."):
                j += 1
            ident = dax[i:j]
            k = j
            while k < n and dax[k] in " \t":
                k += 1
            if k < n and dax[k] == "[":
                m_end = k + 1
                while m_end < n:
                    if dax[m_end] == "]":
                        if m_end + 1 < n and dax[m_end + 1] == "]":
                            m_end += 2
                            continue
                        break
                    m_end += 1
                if m_end < n and dax[m_end] == "]":
                    col_name = dax[k + 1:m_end].replace("]]", "]")
                    tbl_map = rename_plan.get(norm(ident), {})
                    if norm(col_name) in tbl_map:
                        old_sem, new_sem = tbl_map[norm(col_name)]
                        out.append(ident + dax[j:k] + _encode_bracket(new_sem))
                        changes.append((f"{ident}[{old_sem}]", f"{ident}[{new_sem}]"))
                    else:
                        out.append(dax[i:m_end + 1])
                    i = m_end + 1
                    continue
            if k < n and dax[k] == "(" and ident.casefold() in DAX_ITERATOR_FUNCTIONS:
                iter_table = _dax_parse_first_arg_table(dax, k)
                if iter_table:
                    paren_depth += 1
                    iterator_stack.append((paren_depth, iter_table))
                    out.append(dax[i:k + 1])
                    i = k + 1
                    continue
            out.append(ident)
            i = j
            continue

        if ch == "[":
            m_end = i + 1
            while m_end < n:
                if dax[m_end] == "]":
                    if m_end + 1 < n and dax[m_end + 1] == "]":
                        m_end += 2
                        continue
                    break
                m_end += 1
            if m_end < n and dax[m_end] == "]":
                col_name = dax[i + 1:m_end].replace("]]", "]")
                active_tbl = _active_row_table()
                tbl_key = norm(active_tbl)
                tbl_map = rename_plan.get(tbl_key, {})
                if tbl_key and norm(col_name) in tbl_map:
                    if norm(col_name) in measures_by_table.get(tbl_key, set()):
                        warnings.append(
                            f"DAX_UNQUALIFIED_MEASURE_SHADOW: [{col_name}] in {active_tbl} "
                            "matches a measure name; left unchanged."
                        )
                    else:
                        old_sem, new_sem = tbl_map[norm(col_name)]
                        out.append(_encode_bracket(new_sem))
                        changes.append((f"[{old_sem}]", f"[{new_sem}]"))
                        i = m_end + 1
                        continue
                out.append(dax[i:m_end + 1])
                i = m_end + 1
                continue

        out.append(ch)
        i += 1

    return "".join(out), changes, sorted(set(warnings))


def rewrite_table_tmdl_semantic_columns(
    value: str,
    table_name_str: str,
    rename_plan: dict[str, dict[str, tuple[str, str]]],
    table_measures: dict[str, set[str]],
):
    """Apply Phase 3 semantic column renames to one table TMDL file."""
    tbl_key = norm(table_name_str)
    own_map = rename_plan.get(tbl_key, {})
    changes = []
    warnings = []

    m_masks = {}
    masked = value
    m_part_pattern = re.compile(
        r"(?mi)^([ \t]*partition[ \t]+[^\r\n]+=[ \t]*m[ \t]*$)(.*?)(?=^[ \t]*(?:partition|measure|column|hierarchy|annotation|table)\s+|\Z)",
        re.DOTALL,
    )
    for idx, match in enumerate(reversed(list(m_part_pattern.finditer(masked)))):
        token = f"__PBI_P3_MPART_{idx}_{sha(match.group(0))[:12]}__"
        m_masks[token] = match.group(0)
        masked = masked[:match.start()] + token + masked[match.end():]

    if own_map:
        def _repl_col_decl(match):
            indent, raw_name, trailing = match.group(1), match.group(2), match.group(3)
            unq = _unquote_identifier(raw_name)
            if norm(unq) not in own_map:
                return match.group(0)
            old_sem, new_sem = own_map[norm(unq)]
            formatted = format_tmdl_identifier(new_sem)
            changes.append(("TMDL column declaration", old_sem, new_sem))
            return f"{indent}column {formatted}{trailing}"

        masked = re.sub(r"(?mi)^([ \t]*)column[ \t]+([^=\r\n]+?)([ \t]*)$", _repl_col_decl, masked)

        def _repl_sort_by(match):
            prefix, raw_name, trailing = match.group(1), match.group(2), match.group(3)
            unq = _unquote_identifier(raw_name)
            if norm(unq) not in own_map:
                return match.group(0)
            old_sem, new_sem = own_map[norm(unq)]
            formatted = format_tmdl_identifier(new_sem)
            changes.append(("TMDL sortByColumn", old_sem, new_sem))
            return f"{prefix}{formatted}{trailing}"

        masked = re.sub(r"(?mi)^([ \t]*sortByColumn[ \t]*:[ \t]*)([^\r\n]+?)([ \t]*)$", _repl_sort_by, masked)

        def _repl_level_col(match):
            prefix, raw_name, trailing = match.group(1), match.group(2), match.group(3)
            unq = _unquote_identifier(raw_name)
            if norm(unq) not in own_map:
                return match.group(0)
            old_sem, new_sem = own_map[norm(unq)]
            formatted = format_tmdl_identifier(new_sem)
            changes.append(("TMDL hierarchy level column", old_sem, new_sem))
            return f"{prefix}{formatted}{trailing}"

        masked = re.sub(r"(?mi)^([ \t]*column[ \t]*:[ \t]*)([^\r\n]+?)([ \t]*)$", _repl_level_col, masked)

    dax_block_pattern = re.compile(
        r"(?mi)^(?P<header>[ \t]*(?:measure|column|calculationItem)[ \t]+[^\r\n=]+=[ \t]*)"
        r"(?P<body>(?:```[\s\S]*?```|[^\r\n]*(?:\r?\n(?![ \t]*(?:formatString|displayFolder|lineageTag|dataType|isHidden|summarizeBy|sortByColumn|description|annotation|changedProperty|measure|column|hierarchy|partition|calculationItem|table|__PBI_P3_MPART_)\b)[^\r\n]*)*))",
    )

    def _repl_dax_block(match):
        header, body = match.group("header"), match.group("body")
        rewritten, dax_changes, dax_warnings = rewrite_dax_expression(
            body, rename_plan, current_table=table_name_str, table_measures=table_measures
        )
        warnings.extend(dax_warnings)
        for old_ref, new_ref in dax_changes:
            changes.append(("DAX expression", old_ref, new_ref))
        return header + rewritten

    masked = dax_block_pattern.sub(_repl_dax_block, masked)

    calc_part_pattern = re.compile(
        r"(?mi)^(?P<header>[ \t]*source[ \t]*=[ \t]*)"
        r"(?P<body>(?:```[\s\S]*?```|[^\r\n]*(?:\r?\n(?![ \t]*(?:annotation|partition|measure|column|hierarchy|table|__PBI_P3_MPART_)\b)[^\r\n]*)*))",
    )
    masked = calc_part_pattern.sub(_repl_dax_block, masked)

    for token, original_part in m_masks.items():
        masked = masked.replace(token, original_part)

    return masked, changes, sorted(set(warnings))


def rewrite_relationships_tmdl(
    value: str,
    rename_plan: dict[str, dict[str, tuple[str, str]]],
):
    """Rewrite fromColumn/toColumn endpoints in relationships.tmdl or model.tmdl."""
    if not rename_plan:
        return value, []
    changes = []
    endpoint_pattern = re.compile(
        r"(?mi)^(?P<prefix>[ \t]*(?:fromColumn|toColumn)[ \t]*:[ \t]*)(?P<endpoint>[^\r\n]+?)(?P<trailing>[ \t]*)$"
    )

    def _parse_ep(ep: str):
        parser = globals().get("_parse_tmdl_relationship_endpoint")
        if callable(parser):
            return parser(ep)
        m = re.fullmatch(
            r"\s*(?:'(?P<qt>(?:[^']|'')+)'|(?P<pt>[^.\s][^.]*?))\s*\.\s*"
            r"(?:'(?P<qc>(?:[^']|'')+)'|(?P<pc>\S(?:.*?\S)?))\s*",
            clean(ep),
        )
        if not m:
            return "", ""
        t = m.group("qt").replace("''", "'") if m.group("qt") is not None else m.group("pt")
        c = m.group("qc").replace("''", "'") if m.group("qc") is not None else m.group("pc")
        return clean(t), clean(c)

    def _repl(match):
        prefix, ep, trailing = match.group("prefix"), match.group("endpoint"), match.group("trailing")
        tbl, col = _parse_ep(ep)
        if not tbl or not col:
            return match.group(0)
        tbl_map = rename_plan.get(norm(tbl), {})
        if norm(col) not in tbl_map:
            return match.group(0)
        old_sem, new_sem = tbl_map[norm(col)]
        new_ep = format_tmdl_relationship_endpoint(tbl, new_sem)
        changes.append((f"{tbl}.{old_sem}", f"{tbl}.{new_sem}"))
        return f"{prefix}{new_ep}{trailing}"

    return endpoint_pattern.sub(_repl, value), changes


def rewrite_role_tmdl(
    value: str,
    rename_plan: dict[str, dict[str, tuple[str, str]]],
    table_measures: dict[str, set[str]],
):
    """Rewrite DAX filter expressions inside roles/*.tmdl tablePermission declarations."""
    if not rename_plan:
        return value, [], []
    changes, warnings = [], []
    perm_pattern = re.compile(
        r"(?mi)^(?P<header>[ \t]*tablePermission[ \t]+(?P<tbl>'(?:[^']|'')+'|[^=\r\n]+?)[ \t]*=[ \t]*)"
        r"(?P<body>(?:```[\s\S]*?```|[^\r\n]*(?:\r?\n(?![ \t]*(?:tablePermission|annotation|role|modelPermission)\b)[^\r\n]*)*))",
    )

    def _repl(match):
        header, raw_tbl, body = match.group("header"), match.group("tbl"), match.group("body")
        tbl_name = _unquote_identifier(raw_tbl)
        rewritten, dax_changes, dax_warnings = rewrite_dax_expression(
            body, rename_plan, current_table=tbl_name, table_measures=table_measures
        )
        warnings.extend(dax_warnings)
        for old_ref, new_ref in dax_changes:
            changes.append((tbl_name, old_ref, new_ref))
        return header + rewritten

    return perm_pattern.sub(_repl, value), changes, sorted(set(warnings))


def rewrite_culture_or_perspective_tmdl(
    value: str,
    rename_plan: dict[str, dict[str, tuple[str, str]]],
    kind: str,
):
    """Rewrite table/column references in cultures/*.tmdl or perspectives/*.tmdl."""
    if not rename_plan:
        return value, []
    lines = value.splitlines(keepends=True)
    out = []
    changes = []
    current_table = ""
    tbl_kw = "perspectiveTable" if kind == "perspective" else "table"
    col_kw = "perspectiveColumn" if kind == "perspective" else "column"
    tbl_re = re.compile(rf"(?i)^([ \t]*{tbl_kw}[ \t]+)(.+?)([ \t]*\r?\n?)$")
    col_re = re.compile(rf"(?i)^([ \t]*{col_kw}[ \t]+)(.+?)([ \t]*\r?\n?)$")

    for line in lines:
        m_tbl = tbl_re.match(line)
        if m_tbl:
            current_table = _unquote_identifier(m_tbl.group(2))
            out.append(line)
            continue
        m_col = col_re.match(line)
        if m_col and current_table:
            col_name = _unquote_identifier(m_col.group(2))
            tbl_map = rename_plan.get(norm(current_table), {})
            if norm(col_name) in tbl_map:
                old_sem, new_sem = tbl_map[norm(col_name)]
                formatted = format_tmdl_identifier(new_sem)
                changes.append((current_table, old_sem, new_sem))
                out.append(f"{m_col.group(1)}{formatted}{m_col.group(3)}")
                continue
        out.append(line)
    return "".join(out), changes


def _rewrite_query_ref_string(value: str, rename_plan: dict[str, dict[str, tuple[str, str]]]):
    """Rewrite Power BI queryRef strings like 'FactOrders.LocCode' or 'Sum(FactOrders.OrderAmt)'."""
    m = re.match(r"^(?P<agg>[A-Za-z_][A-Za-z0-9_]*\()?(?P<table>[^().]+)\.(?P<col>[^().]+)(?P<close>\))?$", value)
    if not m:
        return value, None
    tbl, col = m.group("table"), m.group("col")
    tbl_map = rename_plan.get(norm(tbl), {})
    if norm(col) not in tbl_map:
        return value, None
    old_sem, new_sem = tbl_map[norm(col)]
    rebuilt = f"{m.group('agg') or ''}{tbl}.{new_sem}{m.group('close') or ''}"
    return rebuilt, (f"{tbl}.{old_sem}", f"{tbl}.{new_sem}")


def rewrite_pbi_report_json(
    raw_text: str,
    rename_plan: dict[str, dict[str, tuple[str, str]]],
):
    """Rewrite semantic Column bindings in PBIR JSON or legacy report.json (including From-aliases & embedded JSON)."""
    if not rename_plan or not raw_text.strip():
        return raw_text, [], []
    try:
        data = json.loads(raw_text)
    except Exception as exc:
        return raw_text, [], [f"REPORT_JSON_PARSE_FAILED: {exc}"]

    changes = []

    def _walk(node, from_aliases: dict[str, str]):
        if isinstance(node, list):
            for item in node:
                _walk(item, from_aliases)
            return
        if not isinstance(node, dict):
            return

        local_aliases = from_aliases
        if isinstance(node.get("From"), list):
            local_aliases = dict(from_aliases)
            for entry in node["From"]:
                if isinstance(entry, dict) and isinstance(entry.get("Name"), str) and isinstance(entry.get("Entity"), str):
                    local_aliases[entry["Name"]] = entry["Entity"]

        if "Column" in node and isinstance(node["Column"], dict):
            col_obj = node["Column"]
            expr = col_obj.get("Expression")
            prop = col_obj.get("Property")
            if isinstance(expr, dict) and isinstance(prop, str):
                src_ref = expr.get("SourceRef")
                if isinstance(src_ref, dict):
                    entity = src_ref.get("Entity") or local_aliases.get(src_ref.get("Source", ""), "")
                    if isinstance(entity, str) and norm(entity) in rename_plan:
                        tbl_map = rename_plan[norm(entity)]
                        if norm(prop) in tbl_map:
                            old_sem, new_sem = tbl_map[norm(prop)]
                            col_obj["Property"] = new_sem
                            changes.append(("Report Column binding", f"{entity}.{old_sem}", f"{entity}.{new_sem}"))

        qref_renamed = None
        if isinstance(node.get("queryRef"), str):
            new_qref, qchange = _rewrite_query_ref_string(node["queryRef"], rename_plan)
            if qchange:
                node["queryRef"] = new_qref
                qref_renamed = qchange
                changes.append(("Report queryRef", qchange[0], qchange[1]))
        if qref_renamed and isinstance(node.get("nativeQueryRef"), str):
            old_col = qref_renamed[0].split(".", 1)[-1]
            new_col = qref_renamed[1].split(".", 1)[-1]
            if norm(node["nativeQueryRef"]) == norm(old_col):
                node["nativeQueryRef"] = new_col
                changes.append(("Report nativeQueryRef", old_col, new_col))

        for key in ("config", "filters", "dataTransforms", "query"):
            val = node.get(key)
            if isinstance(val, str) and val.lstrip().startswith(("{", "[")):
                try:
                    nested = json.loads(val)
                except Exception:
                    continue
                before_count = len(changes)
                _walk(nested, local_aliases)
                if len(changes) > before_count:
                    node[key] = json.dumps(nested, separators=(",", ":"))

        for k, v in node.items():
            if k in {"From", "Column"}:
                continue
            if isinstance(v, (dict, list)):
                _walk(v, local_aliases)

    _walk(data, {})
    if not changes:
        return raw_text, [], []

    is_pretty = "\n" in raw_text.strip()
    rendered = json.dumps(data, indent=2, ensure_ascii=False) + "\n" if is_pretty else json.dumps(data, separators=(",", ":"), ensure_ascii=False)
    try:
        json.loads(rendered)
    except Exception as exc:
        return raw_text, [], [f"REPORT_JSON_VALIDATION_FAILED: {exc}"]
    return rendered, changes, []


def discover_report_json_files(paths: Paths) -> list[Path]:
    """Return all PBIR and legacy JSON files under the sibling *.Report folder."""
    report_dir = getattr(paths, "report", None)
    if not report_dir or not report_dir.is_dir():
        return []
    candidates = []
    definition_dir = report_dir / "definition"
    if definition_dir.is_dir():
        for p in sorted(definition_dir.rglob("*.json")):
            if p.name.casefold() in {"visual.json", "page.json", "report.json"} or p.name.casefold().endswith(".bookmark.json"):
                candidates.append(p)
    legacy_report = report_dir / "report.json"
    if legacy_report.exists() and legacy_report not in candidates:
        candidates.append(legacy_report)
    return candidates


def apply_phase3_semantic_and_report_migrations(
    paths: Paths,
    originals: dict[Path, str],
    proposed: dict[Path, str],
    log: dict,
    status: str,
):
    """Execute Phase 3 across table TMDL, relationships, roles, cultures, perspectives, and .Report JSON."""
    if not (ALIGN_TARGET_COLUMN_NAMES and RENAME_SEMANTIC_COLUMNS):
        return {}, {}, {}, []

    rename_plan, table_display_names, table_measures, plan_errors = build_semantic_rename_plan(
        originals, proposed, paths.tables
    )
    if plan_errors or not rename_plan:
        return {}, {}, rename_plan, plan_errors

    extra_originals: dict[Path, str] = {}
    extra_proposed: dict[Path, str] = {}
    errors: list[str] = []

    def _log_p3(input_file: str, layer: str, obj: str, change_type: str, old_val: str, new_val: str, count: int = 1):
        log.setdefault("Semantic & Report Changes", []).append({
            "Input File": input_file,
            "Layer": layer,
            "Object": obj,
            "Change Type": change_type,
            "Old Value": old_val,
            "New Value": new_val,
            "Occurrences": count,
            "Result": status,
        })

    for path in sorted(paths.tables.glob("*.tmdl")):
        if path not in proposed:
            continue
        tbl_name = tmdl_table_name(originals[path], path.stem)
        updated, changes, warnings = rewrite_table_tmdl_semantic_columns(
            proposed[path], tbl_name, rename_plan, table_measures
        )
        proposed[path] = updated
        grouped = {}
        for ctype, old_v, new_v in changes:
            grouped[(ctype, old_v, new_v)] = grouped.get((ctype, old_v, new_v), 0) + 1
        for (ctype, old_v, new_v), cnt in sorted(grouped.items()):
            _log_p3(path.name, "Semantic Table", tbl_name, ctype, old_v, new_v, cnt)
            if ctype == "TMDL column declaration":
                log["Column Changes"].append({
                    "Input File": path.name, "Model Table": tbl_name,
                    "Model Column": new_v, "Old Source Column": old_v,
                    "New Source Column": new_v,
                    "Changed In": "Phase 3 TMDL semantic column declaration",
                    "Model Name Changed": "Yes", "Result": status,
                })
        for w in warnings:
            log["Warnings & Manual Review"].append({
                "Input File": path.name, "Warning Code": "DAX_UNQUALIFIED_MEASURE_SHADOW",
                "Category": "Phase 3 DAX", "Object": tbl_name, "Warning": w,
                "Impact": "Unqualified reference left unchanged to avoid altering a measure call.",
                "Recommended Review": "Verify whether the reference targets the measure or the renamed column.",
            })

    for rel_path in paths.relationships:
        if not rel_path.exists():
            continue
        orig = read_text(rel_path)
        updated, rel_changes = rewrite_relationships_tmdl(orig, rename_plan)
        if updated != orig:
            extra_originals[rel_path] = orig
            extra_proposed[rel_path] = updated
            for old_ep, new_ep in rel_changes:
                _log_p3(rel_path.name, "Relationships", "Relationship", "fromColumn/toColumn", old_ep, new_ep)

    for role_path in getattr(paths, "roles", []):
        if not role_path.exists():
            continue
        orig = read_text(role_path)
        updated, role_changes, role_warnings = rewrite_role_tmdl(orig, rename_plan, table_measures)
        if updated != orig:
            extra_originals[role_path] = orig
            extra_proposed[role_path] = updated
            for tbl_name, old_ref, new_ref in role_changes:
                _log_p3(role_path.name, "RLS Role", tbl_name, "tablePermission DAX", old_ref, new_ref)
        for w in role_warnings:
            log["Warnings & Manual Review"].append({
                "Input File": role_path.name, "Warning Code": "DAX_UNQUALIFIED_MEASURE_SHADOW",
                "Category": "Phase 3 RLS", "Object": role_path.stem, "Warning": w,
                "Impact": "Unqualified RLS reference left unchanged.",
                "Recommended Review": "Confirm RLS filter column reference.",
            })

    for kind, file_list in (("culture", getattr(paths, "cultures", [])), ("perspective", getattr(paths, "perspectives", []))):
        for cp_path in file_list:
            if not cp_path.exists():
                continue
            orig = read_text(cp_path)
            updated, cp_changes = rewrite_culture_or_perspective_tmdl(orig, rename_plan, kind)
            if updated != orig:
                extra_originals[cp_path] = orig
                extra_proposed[cp_path] = updated
                for tbl_name, old_col, new_col in cp_changes:
                    _log_p3(cp_path.name, kind.title(), tbl_name, f"{kind} column", old_col, new_col)

    for json_path in discover_report_json_files(paths):
        orig = read_text(json_path)
        updated, json_changes, json_errors = rewrite_pbi_report_json(orig, rename_plan)
        if json_errors:
            errors.extend(f"{json_path.name}: {e}" for e in json_errors)
            continue
        if updated != orig:
            extra_originals[json_path] = orig
            extra_proposed[json_path] = updated
            rel_name = str(json_path.relative_to(paths.pbip)) if paths.pbip in json_path.parents else json_path.name
            grouped = {}
            for ctype, old_v, new_v in json_changes:
                grouped[(ctype, old_v, new_v)] = grouped.get((ctype, old_v, new_v), 0) + 1
            for (ctype, old_v, new_v), cnt in sorted(grouped.items()):
                _log_p3(rel_name, "Report Visual/JSON", json_path.stem, ctype, old_v, new_v, cnt)

    return extra_originals, extra_proposed, rename_plan, sorted(set(errors))


def apply_migrations(paths: Paths, mappings: list[MappingRow], sources: list[SourceMapping], log, write_files: bool = True):
    files = sorted(paths.tables.glob("*.tmdl")) + ([paths.expressions] if paths.expressions.exists() else [])
    status = "Applied" if write_files else "Proposed"
    originals = {path: read_text(path) for path in files}
    proposed = dict(originals)
    graph = build_m_query_graph(originals)
    original_queries = _query_text_index(originals)
    query_binding_cache = {}
    mapped_queries = mapped_query_names(graph, mappings, sources)
    blocking_errors = []

    parameter_rows, parameter_definitions, parameter_roles = discover_parameters(paths)
    current_parameters = {
        name: next((entry.value for entry in entries if getattr(entry, "is_current", False)), "")
        for name, entries in parameter_definitions.items()
    }
    for path in files:
        updated = originals[path]
        updated = migrate_current_source_parameters(updated, path, parameter_definitions, sources, log, status)
        # Migrate each expressions.tmdl query independently and derive
        # Schema/Item navigation replacements from table mappings.
        updated = migrate_m_navigation_per_query(
            path, updated, mappings, sources, log, status
        )
        query_name = tmdl_table_name(updated, path.stem)
        native = is_native_query(updated)
        affected = set()

        # Source/navigation changes stay inside M. Mask native SQL first so
        # schema/item replacements cannot alter SQL identifiers or comments.
        early_masks = []
        if is_native_query(updated):
            updated, early_masks = mask_native_sql_payloads(updated)
        region, region_start, region_end = extract_m_region(updated)
        if region_start >= 0:
            migrated_region = region
            for source in sources:
                for old, new, change_type in (
                    (source.legacy_source, source.target_source, "Connection source"),
                    (source.legacy_schema, source.target_schema, "Source schema"),
                    (source.legacy_item, source.target_item, "Source item"),
                    (source.legacy_catalog, source.target_catalog, "Catalog"),
                ):
                    if old and new and old != new:
                        migrated_region, count = scoped_replace(migrated_region, old, new)
                        if count:
                            log["M Query Changes"].append({
                                "Input File": path.name, "Query Name": query_name,
                                "M Step": "Source/Navigation", "Change Type": change_type,
                                "Old Value": old, "New Value": new,
                                "Result": f"{status}; {count} occurrence(s)",
                            })
            updated = updated[:region_start] + migrated_region + updated[region_end:]
        if early_masks:
            updated, early_restore_errors = restore_native_sql_payloads(updated, early_masks)
            for early_error in early_restore_errors:
                blocking_errors.append(f"{path.name}: {early_error}")
                log["Errors"].append({
                    "Input File": path.name, "Stage": "Pre-SQL native payload protection",
                    "Error Code": "NATIVE_SQL_MASK_RESTORE_FAILED",
                    "Error Message": early_error, "Existing Content Preserved": "Yes",
                    "Processing Continued": "No",
                })

        native_preserved_aliases = set()
        native_output_contract = set()
        native_target_bindings = {}
        native_scoped_mappings = []
        sqlglot_value = None
        native_sql_masks = []
        if native:
            candidate, sql_audit, sql_errors = replace_native_sql_payloads(updated, mappings, sources)
            native_output_contract, contract_errors = native_sql_output_contract(candidate)
            native_target_bindings, native_scoped_mappings = _native_sql_target_bindings(
                sql_audit, native_output_contract, mappings
            )
            native_preserved_aliases = (
                set() if ALIGN_TARGET_COLUMN_NAMES
                else (_native_sql_preserved_aliases(sql_audit) | native_output_contract)
            )
            sql_errors.extend(contract_errors)
            sqlglot_value = candidate
            updated, native_sql_masks = mask_native_sql_payloads(candidate)
            if sql_errors:
                for error in sql_errors:
                    blocking_errors.append(f"{path.name}: {error}")
                    log["Errors"].append({
                        "Input File": path.name,
                        "Stage": "SQLGlot Synapse AST validation",
                        "Error Code": "NATIVE_SQL_AST_REWRITE_FAILED",
                        "Error Message": error,
                        "Existing Content Preserved": "Yes",
                        "Processing Continued": "No",
                    })
            for record in sql_audit:
                if record["kind"] == "function":
                    log["Native SQL Changes"].append({
                        "Input File": path.name, "Query Name": query_name,
                        "Change Type": "Target SQL function normalization",
                        "Old Value": record["old_value"], "New Value": record["new_value"],
                        "Occurrences": record["occurrences"], "Safe Replacement": "Yes",
                        "Message": status,
                    })
                elif record["kind"] == "table":
                    log["Table Changes"].append({
                        "Input File": path.name, "File Type": "Native SQL",
                        "Query or Model Table": query_name,
                        "Change Location": "SQLGlot AST table node",
                        "Old Table": record["legacy_table"], "New Table": record["target_table"],
                        "SQL Alias": record["alias"], "Alias Preserved": "Yes",
                        "Replacement Count": record["occurrences"],
                        "Result": f"{status}; {record.get('alias_preservation', 'Alias preserved')}",
                    })
                elif record["kind"] == "column":
                    affected.add((record["legacy_table"], record["legacy_column"]))
                    log["Native SQL Changes"].append({
                        "Input File": path.name, "Query Name": query_name,
                        "Change Type": "SQLGlot AST table-scoped column rewrite",
                        "Old Value": f"{record['legacy_table']}.{record['legacy_column']}",
                        "New Value": f"{record['target_table']}.{record['target_column']}",
                        "Occurrences": record["occurrences"], "Safe Replacement": "Yes",
                        "Message": f"{status}; aliases={record['aliases']}; projection aliases={record['projection_aliases_added']}; casts={record['casts_added']}",
                    })
                elif record["kind"] == "star":
                    log["Native SQL Changes"].append({
                        "Input File": path.name, "Query Name": query_name,
                        "Change Type": "SELECT star preserved",
                        "Old Value": "SELECT * / alias.*", "New Value": "Unchanged",
                        "Occurrences": record["occurrences"], "Safe Replacement": "Yes",
                        "Message": record["message"],
                    })

            if native_output_contract:
                log["Native SQL Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "Change Type": "Native SQL output contract",
                    "Old Value": "Downstream M source schema",
                    "New Value": ", ".join(sorted(native_output_contract)),
                    "Occurrences": len(native_output_contract), "Safe Replacement": "Yes",
                    "Message": f"{status}; downstream M references to these aliases are protected",
                })

        query_output_bindings = {}
        if not native:
            updated, join_key_changes, join_key_errors = migrate_join_keys_with_lineage(
                updated, query_name, original_queries, mappings, query_binding_cache, current_parameters
            )
            for change in join_key_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": change["function"], "Change Type": "Branch-scoped join key mapping",
                    "Old Value": f"left={change['old_left']}; right={change['old_right']}",
                    "New Value": f"left={change['new_left']}; right={change['new_right']}",
                    "Result": f"{status}; left={change['left_query']}; right={change['right_query']}",
                })
            for error in join_key_errors:
                blocking_errors.append(f"{path.name}: {error}")
            updated, append_changes, append_errors = migrate_append_contracts(
                updated, query_name, original_queries, mappings, query_binding_cache, current_parameters
            )
            for change in append_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Table.Combine", "Change Type": "Append-local forward target reconciliation",
                    "Old Value": ", ".join(old_col for old_col, _new_col in change["pairs"]),
                    "New Value": ", ".join(new_col for _old_col, new_col in change["pairs"]),
                    "Result": f"{status}; branch={change['branch']}",
                })
            for error in append_errors:
                blocking_errors.append(f"{path.name}: {error}")
            inherited_bindings, inherited_errors = resolved_query_output_bindings(
                query_name, original_queries, mappings, query_binding_cache, None, current_parameters
            )
            query_output_bindings.update(inherited_bindings)
            for inherited_error in inherited_errors:
                blocking_errors.append(f"{path.name}: {inherited_error}")
            updated, expand_changes, expand_bindings, expand_errors = migrate_join_expansions(updated, original_queries, mappings, query_binding_cache, current_parameters)
            query_output_bindings.update(expand_bindings)
            for change in expand_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Table.ExpandTableColumn",
                    "Change Type": "Right-query target inputs and propagated output names",
                    "Old Value": f"inputs={change['old_inputs']}; outputs={change['old_outputs']}",
                    "New Value": f"inputs={change['new_inputs']}; outputs={change['new_outputs']}",
                    "Result": f"{status}; right query={change['right_query']}",
                })
            for expand_error in expand_errors:
                blocking_errors.append(f"{path.name}: {expand_error}")
            query_output_bindings.update(_merged_query_output_bindings(updated, original_queries, mappings, query_binding_cache, current_parameters))
            updated, downstream_changes = propagate_bindings_to_downstream_m(updated, query_output_bindings)
            for legacy, target, count in downstream_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Downstream merged-query transformations",
                    "Change Type": "Query output binding propagation",
                    "Old Value": legacy, "New Value": target,
                    "Result": f"{status}; {count} occurrence(s)",
                })
        elif ALIGN_TARGET_COLUMN_NAMES and native_target_bindings:
            query_binding_cache[query_name.casefold()] = dict(native_target_bindings)
            updated, post_join_changes, post_join_errors = migrate_join_keys_with_lineage(
                updated, query_name, original_queries, mappings, query_binding_cache, current_parameters
            )
            for change in post_join_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": change["function"], "Change Type": "Post-native branch-scoped join key mapping",
                    "Old Value": f"left={change['old_left']}; right={change['old_right']}",
                    "New Value": f"left={change['new_left']}; right={change['new_right']}",
                    "Result": f"{status}; left={change['left_query']}; right={change['right_query']}",
                })
            for error in post_join_errors:
                blocking_errors.append(f"{path.name}: {error}")
            updated, post_expand_changes, post_expand_bindings, post_expand_errors = migrate_join_expansions(
                updated, original_queries, mappings, query_binding_cache, current_parameters
            )
            for change in post_expand_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Table.ExpandTableColumn",
                    "Change Type": "Post-native right-query target inputs and output names",
                    "Old Value": f"inputs={change['old_inputs']}; outputs={change['old_outputs']}",
                    "New Value": f"inputs={change['new_inputs']}; outputs={change['new_outputs']}",
                    "Result": f"{status}; right query={change['right_query']}",
                })
            for error in post_expand_errors:
                blocking_errors.append(f"{path.name}: {error}")
            effective_native_bindings = dict(native_target_bindings)
            effective_native_bindings.update(post_expand_bindings)
            query_binding_cache[query_name.casefold()] = dict(effective_native_bindings)
            updated, post_native_changes = propagate_bindings_to_downstream_m(updated, effective_native_bindings)
            for legacy, target, count in post_native_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Post-native M transformations",
                    "Change Type": "Native SQL target column propagation",
                    "Old Value": legacy, "New Value": target,
                    "Result": f"{status}; {count} occurrence(s)",
                })
            updated, binding_changes = apply_query_output_bindings(updated, effective_native_bindings)
            for change in binding_changes:
                for semantic_column in change["semantic_columns"]:
                    log["Column Changes"].append({
                        "Input File": path.name, "Model Table": query_name,
                        "Model Column": semantic_column,
                        "Old Source Column": change["legacy"],
                        "New Source Column": change["target"],
                        "Changed In": "TMDL sourceColumn from native SQL target alignment",
                        "Model Name Changed": "No", "Result": status,
                    })
        # Native SQL queries can still contain M steps after Value.NativeQuery.
        # Non-native and post-native transformations use the same lineage-aware M pass.
        for mapping in mappings:
            # expressions.tmdl is a multi-query container. Its non-native M
            # transformations were handled per expression above; never apply
            # container-wide column substitutions here.
            if path.name.casefold() == "expressions.tmdl":
                continue
            selection = mapping.selection if mapping.selection in {"Yes", "No", "Review"} else mapping.system_recommendation
            if selection == "Review" or not m_mapping_reaches_source(graph, mapped_queries, query_name, mapping, updated, current_parameters):
                continue

            # Native SQL is the lineage boundary when SQLGlot preserved the
            # legacy projection alias. Downstream M must retain that legacy name.
            if native and not _native_mapping_crosses_alias_boundary(
                mapping, native_preserved_aliases
            ):
                continue

            total_m = 0
            candidate, count = replace_m_expand_inputs(updated, mapping.legacy_column, mapping.target_column)
            if count:
                updated, total_m = candidate, total_m + count
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Table.ExpandTableColumn", "Change Type": "Input columns only; output names preserved",
                    "Old Value": mapping.legacy_column, "New Value": mapping.target_column,
                    "Result": f"{status}; {count} expansion(s)",
                })

            if native and not ALIGN_TARGET_COLUMN_NAMES:
                candidate, count = replace_m_join_keys(updated, mapping.legacy_column, mapping.target_column)
                if count:
                    updated, total_m = candidate, total_m + count
                    log["M Query Changes"].append({
                        "Input File": path.name, "Query Name": query_name,
                        "M Step": "Join/Merge", "Change Type": "Post-native join key mapping",
                        "Old Value": mapping.legacy_column, "New Value": mapping.target_column,
                        "Result": f"{status}; {count} join/merge step(s)",
                    })

            candidate, count = replace_m_append_dependencies(updated, mapping.legacy_column, mapping.target_column)
            if count:
                updated, total_m = candidate, total_m + count
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Table.Combine", "Change Type": "Append-related column dependency",
                    "Old Value": mapping.legacy_column, "New Value": mapping.target_column,
                    "Result": f"{status}; {count} occurrence(s)",
                })

            candidate, count = replace_non_native_m_column(updated, mapping.legacy_column, mapping.target_column)
            if count:
                updated, total_m = candidate, total_m + count
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Existing source-facing transformations", "Change Type": "Mapped column references",
                    "Old Value": mapping.legacy_column, "New Value": mapping.target_column,
                    "Result": f"{status}; {count} occurrence(s)",
                })

            source_count = 0
            semantic_columns = []
            if not native:
                candidate, source_count, semantic_columns = replace_tmdl_source_bindings(updated, mapping.legacy_column, mapping.target_column)
                if source_count:
                    updated = candidate
                    for semantic_column in semantic_columns:
                        log["Column Changes"].append({
                            "Input File": path.name, "Model Table": query_name,
                            "Model Column": semantic_column,
                            "Old Source Column": mapping.legacy_column,
                            "New Source Column": mapping.target_column,
                            "Changed In": "TMDL sourceColumn", "Model Name Changed": "No", "Result": status,
                        })

            if total_m or source_count:
                affected.add((mapping.legacy_table, mapping.legacy_column))

        # After upstream non-native queries adopt target column names,
        # rewrite row-context references in merged-query Table.AddColumn formulas
        # from the actual left/right lineage. This handles [Location] ->
        # [Location_Code] without allowing the unexpanded right-side Location
        # mapping to bleed into the formula.
        if not native:
            updated, addcolumn_changes, addcolumn_errors = migrate_merged_addcolumn_references(
                updated, original_queries, mappings, query_binding_cache, current_parameters
            )
            for change in addcolumn_changes:
                log["M Query Changes"].append({
                    "Input File": path.name,
                    "Query Name": query_name,
                    "M Step": "Table.AddColumn",
                    "Change Type": "Merged-query row-field lineage propagation",
                    "Old Value": change["old_formula"],
                    "New Value": change["new_formula"],
                    "Result": (
                        f"{status}; {change['count']} field reference(s); "
                        f"left={change['left_query']}; right={change['right_query']}"
                    ),
                })
            for addcolumn_error in addcolumn_errors:
                blocking_errors.append(f"{path.name}: {addcolumn_error}")
            updated, binding_changes = apply_query_output_bindings(updated, query_output_bindings)
            for change in binding_changes:
                for semantic_column in change["semantic_columns"]:
                    log["Column Changes"].append({
                        "Input File": path.name, "Model Table": query_name,
                        "Model Column": semantic_column,
                        "Old Source Column": change["legacy"],
                        "New Source Column": change["target"],
                        "Changed In": "TMDL sourceColumn from merged-query output binding",
                        "Model Name Changed": "No", "Result": status,
                    })

        for mapping in mappings:
            if (mapping.legacy_table, mapping.legacy_column) not in affected:
                continue
            selection = mapping.selection if mapping.selection in {"Yes", "No", "Review"} else mapping.system_recommendation
            log["Data Type Decisions"].append({
                "Input File": path.name, "Table": mapping.legacy_table,
                "Model Column": mapping.legacy_column,
                "Legacy Type": mapping.legacy_type, "Target Type": mapping.target_type,
                "Retain Legacy Data Type": selection,
                "Final Type": mapping.legacy_type if selection == "Yes" else mapping.target_type,
                "Decision Source": "Mapping Template",
                "Rule Description": mapping.dependency_notes or mapping.notes,
                "SQL Cast Added": "Evaluated by SQLGlot" if native and selection == "Yes" else "No",
                "M Step Added": "No", "TMDL Updated": "No" if native else "Yes", "Result": status,
            })

        updated, unpivot_rows, unpivot_errors = migrate_unpivot_preservation(
            updated, query_name, path.name, mappings, graph,
            mapped_queries, native, status,
            extra_mappings=native_scoped_mappings if native else None,
            log=log, current_parameters=current_parameters,
        )
        log["Unpivot Inventory"].extend(unpivot_rows)
        for unpivot_error in unpivot_errors:
            blocking_errors.append(unpivot_error)
            log["Errors"].append({
                "Input File": path.name,
                "Stage": "Unpivot lineage migration",
                "Error Code": "UNPIVOT_DYNAMIC_OR_UNSAFE",
                "Error Message": unpivot_error,
                "Existing Content Preserved": "Yes",
                "Processing Continued": "No",
            })
        if native:
            updated, restore_errors = restore_native_sql_payloads(updated, native_sql_masks)
            protection_errors = restore_errors + validate_native_sql_immutable(sqlglot_value, updated)
            contract_validation_errors = validate_post_native_m_contract(updated, native_output_contract, mappings)
            for protection_error in protection_errors:
                blocking_errors.append(f"{path.name}: {protection_error}")
                log["Errors"].append({
                    "Input File": path.name, "Stage": "Native SQL immutability validation",
                    "Error Code": "NATIVE_SQL_MUTATED_BY_M_REWRITER",
                    "Error Message": protection_error,
                    "Existing Content Preserved": "Yes", "Processing Continued": "No",
                })
            for contract_error in contract_validation_errors:
                blocking_errors.append(f"{path.name}: {contract_error}")
                log["Errors"].append({
                    "Input File": path.name, "Stage": "Post-native M output-contract validation",
                    "Error Code": "POST_NATIVE_M_COLUMN_NOT_EXPOSED",
                    "Error Message": contract_error,
                    "Existing Content Preserved": "Yes", "Processing Continued": "No",
                })
        proposed[path] = updated

    p3_originals, p3_proposed, p3_rename_plan, p3_errors = apply_phase3_semantic_and_report_migrations(
        paths, originals, proposed, log, status
    )
    blocking_errors.extend(p3_errors)
    all_files = list(files) + list(p3_proposed.keys())
    for p3_path in p3_proposed:
        originals[p3_path] = p3_originals[p3_path]
        proposed[p3_path] = p3_proposed[p3_path]

    validation_errors = list(blocking_errors)
    unpivot_rows_by_file = {}
    for row in log["Unpivot Inventory"]:
        unpivot_rows_by_file.setdefault(clean(row.get("Input File")), []).append(row)
    for path in files:
        validation_errors.extend(
            validate_unpivot_inventory(
                proposed[path], unpivot_rows_by_file.get(path.name, [])
            )
        )
    for path in files:
        tbl_key = norm(tmdl_table_name(originals[path], path.stem))
        allowed_p3 = (p3_rename_plan.get(tbl_key, {}) if (ALIGN_TARGET_COLUMN_NAMES and RENAME_SEMANTIC_COLUMNS and p3_rename_plan) else None)
        validation_errors.extend(validate_protected_content(path, originals[path], proposed[path], allowed_p3))
        validation_errors.extend(validate_mapped_legacy_navigation(path, proposed[path], mappings))
    validation_errors.extend(validate_relationships(
        paths,
        {p: proposed[p] for p in files if p.parent == paths.tables},
        {p: proposed[p] for p in p3_proposed if p in paths.relationships},
    ))

    if validation_errors:
        for error in sorted(set(validation_errors)):
            if not any(row.get("Error Message") == error for row in log["Errors"]):
                log["Errors"].append({
                    "Input File": "Semantic Model", "Stage": "Pre-write validation",
                    "Error Code": "MODEL_INTEGRITY_VALIDATION_FAILED",
                    "Error Message": error, "Existing Content Preserved": "Yes", "Processing Continued": "No",
                })
        proposed = dict(originals)
        write_files = False

    for path in all_files:
        changed = proposed[path] != originals[path]
        query_name = tmdl_table_name(originals[path], path.stem)
        graph_node = graph.get(query_name.casefold())
        if changed and write_files:
            atomic_write(path, proposed[path])
        log["Files Processed"].append({
            "Run Timestamp": datetime.now().isoformat(timespec="seconds"),
            "Input File": path.name, "Input Path": str(path), "File Type": "TMDL",
            "Classification": classify_file_with_lineage(originals[path], graph_node),
            "Relevant Mapping Found": "Yes" if changed else "No",
            "Changed": "Yes" if changed and write_files else "Proposed" if changed else "No",
            "Action": "Modified" if changed and write_files else "Would Modify" if changed else "Unchanged",
            "Write Status": "Success" if changed and write_files else "Preview Only" if changed else "Not Written",
            "Message": "Synapse source migration applied." if changed and write_files else "Synapse source migration previewed." if changed else "No applicable changes.",
            "Input SHA-256": sha(originals[path]), "Generated SHA-256": sha(proposed[path]) if changed else "",
        })

def build_log_workbook(path: Path, log: dict[str,list[dict]], summary: list[tuple[str,object]]) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    for title, headers in LOG_SHEETS.items():
        ws = wb.create_sheet(title)
        ws.append(headers)
        if title == "Run Summary":
            for item, value in summary: ws.append([item, value])
        else:
            for row in log.get(title, []):
                ws.append([row.get(h, "") for h in headers])
                if title == "Parameters" and row.get("Is Current") == "No" and row.get("_Source Related") == "Yes":
                    for cell in ws[ws.max_row]:
                        cell.fill = PatternFill("solid", fgColor="FFF2CC")
        style_tab(ws, add_table=True)
    wb.save(path)


def main() -> int:
    paths = discover_paths(ROOT_FOLDER)
    mapping_path = paths.migration / MAPPING_NAME
    log_path = paths.migration / LOG_NAME
    create_or_upgrade_mapping_template(mapping_path)
    mappings, sources = parse_mapping(mapping_path)
    inventory_rows, unpivot_rows = inventory(paths)
    parameter_rows, parameter_definitions, parameter_roles = discover_parameters(paths)
    table_modes, semantic_model_mode = detect_table_modes(paths)
    inventory_rows = enrich_inventory(inventory_rows, table_modes, parameter_definitions)
    dependency_rows = dependency_analysis(paths, mappings) if mappings else []
    if DRY_RUN and mappings:
        update_mapping_recommendations(mapping_path, mappings)
        # reload selections after recommendations were populated
        mappings, sources = parse_mapping(mapping_path)
    valid_count = len(mappings) + len(sources)
    mode = "Initial Inventory" if DRY_RUN and not valid_count else "Mapping Validation and Migration Preview" if DRY_RUN else "Migration"
    effective_dry = DRY_RUN or valid_count == 0
    log = {sheet: [] for sheet in LOG_SHEETS}
    log["Query & Source Inventory"] = inventory_rows
    log["Parameters"] = parameter_rows
    log["Unpivot Inventory"] = unpivot_rows
    log["Type Dependency Analysis"] = dependency_rows
    for inventory_error in validate_native_source_inventory(paths, inventory_rows):
        log["Errors"].append({
            "Input File": "Query & Source Inventory",
            "Stage": "Native SQL source inventory validation",
            "Error Code": "NATIVE_SQL_SOURCE_INVENTORY_INCOMPLETE",
            "Error Message": inventory_error,
            "Existing Content Preserved": "Yes",
            "Processing Continued": "No",
        })
    table_files = sorted(paths.tables.glob("*.tmdl"))
    all_scan = table_files + ([paths.expressions] if paths.expressions.exists() else [])
    if effective_dry and valid_count == 0:
        for p in all_scan:
            original = read_text(p)
            log["Files Processed"].append({
                "Run Timestamp": datetime.now().isoformat(timespec="seconds"),
                "Input File": p.name,
                "Input Path": str(p),
                "File Type": "TMDL",
                "Classification": "Expressions" if p == paths.expressions else "Table",
                "Relevant Mapping Found": "No",
                "Changed": "No",
                "Action": "Scanned",
                "Write Status": "Dry run",
                "Message": "Inventory completed; no mappings defined.",
                "Input SHA-256": sha(original),
                "Generated SHA-256": "",
            })
    else:
        apply_migrations(paths, mappings, sources, log, write_files=(not effective_dry))
        if not effective_dry:
            inventory_rows, unpivot_rows = inventory(paths)
            parameter_rows, parameter_definitions, parameter_roles = discover_parameters(paths)
            table_modes, semantic_model_mode = detect_table_modes(paths)
            inventory_rows = enrich_inventory(inventory_rows, table_modes, parameter_definitions)
            log["Query & Source Inventory"] = inventory_rows
            log["Parameters"] = parameter_rows
            log["Unpivot Inventory"] = unpivot_rows
            for inventory_error in validate_native_source_inventory(paths, inventory_rows):
                log["Errors"].append({
                    "Input File": "Query & Source Inventory",
                    "Stage": "Post-migration native SQL source inventory validation",
                    "Error Code": "NATIVE_SQL_SOURCE_INVENTORY_INCOMPLETE",
                    "Error Message": inventory_error,
                    "Existing Content Preserved": "Yes",
                    "Processing Continued": "No",
                })
    active = sum(1 for r in inventory_rows if r["Query Status"] == "Active")
    disabled = sum(1 for r in inventory_rows if r["Query Status"] == "Disabled")
    native = sum(1 for r in inventory_rows if r["Native Query"] == "Yes")
    clean_preview_log(log)
    dependency_rows = unique_dependency_rows(dependency_rows)
    mismatches = sum(1 for r in dependency_rows if r["Type Mismatch"] == "Yes")
    recs = {x: sum(1 for r in dependency_rows if r["System Recommendation"] == x) for x in ("Yes", "No", "Review")}
    selections = {x: sum(1 for r in dependency_rows if r["Mapping Template Selection"] == x) for x in ("Yes", "No", "Review")}
    type_decision_count = len(log["Data Type Decisions"])
    native_sql_change_count = len(log["Native SQL Changes"])
    m_query_change_count = len(log["M Query Changes"])
    modified = sum(1 for r in log["Files Processed"] if r["Changed"] == "Yes")
    unchanged = sum(1 for r in log["Files Processed"] if r["Changed"] == "No")
    warning_count = len(log["Warnings & Manual Review"])
    error_count = len(log["Errors"])
    if valid_count == 0:
        message = "Mapping not defined. Dry run completed. No PBIP files were modified."
    elif effective_dry:
        message = "Migration preview completed. Recommendations and potential edits were written to the Mapping Template and Metadata Log. No PBIP files were modified."
    else:
        message = "Migration completed. Valid mappings were applied and the Metadata Log was refreshed."
    summary = [
        ("Report Name", paths.root.name), ("Run Timestamp", datetime.now().isoformat(timespec="seconds")),
        ("Execution Mode", "Migration Synapse -> Synapse"), ("Root Folder", str(paths.root)), ("PBIP Folder", str(paths.pbip)),
        ("Semantic Model", paths.semantic.name), ("Semantic Model Mode", semantic_model_mode), ("Migration Folder", str(paths.migration)),
        ("Mapping Template Status", "Populated" if valid_count else "Headers only"),
        ("Table and Column Mapping Rows", len(mappings)), ("Data Source Mapping Rows", len(sources)),
        ("TMDL Table Files Scanned", len(table_files)), ("Expressions File Scanned", "Yes" if paths.expressions.exists() else "No"),
        ("Active Queries Found", active), ("Disabled Queries Found", disabled), ("Native Queries Found", native),
        ("Unpivots Found", len(unpivot_rows)), ("Type Mismatches Found", mismatches),
        ("System Type Recommendations", f"Yes: {recs['Yes']}; No: {recs['No']}; Review: {recs['Review']}"),
        ("Mapping Template Selections", f"Yes: {selections['Yes']}; No: {selections['No']}; Review: {selections['Review']}"),
        ("Data Type Decisions Generated", type_decision_count),
        ("Native SQL Changes Generated", native_sql_change_count),
        ("M Query Changes Generated", m_query_change_count),
        ("Unpivot Inventory Rows", len(log["Unpivot Inventory"])),
        ("Files Proposed for Modification", sum(1 for r in log["Files Processed"] if r["Changed"] == "Proposed")),
        ("Files Modified", modified), ("Files Unchanged", unchanged), ("Warnings", warning_count),
        ("Manual Review Items", selections["Review"]), ("Errors", error_count),
        ("PBIP Files Modified", "No" if effective_dry else "Yes" if modified else "No"),
        ("Overall Status", "Completed" if not error_count else "Completed with errors"), ("Run Message", message),
    ]
    build_log_workbook(log_path, log, summary)
    print("\nPower BI PBIP migration completed")
    print(f"Execution mode: {mode}")
    print(f"Metadata log: {log_path}")
    print(f"Mapping template: {mapping_path}")
    print(message)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
