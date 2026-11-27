
"""PBIP Metadata Reader and Migration Utility.

Safety: migrates physical sources only. Converts Azure Synapse PBIP semantic models to Databricks using Data Source Mapping and Table and Column Mapping. Native SQL is parsed as T-SQL, rewritten via SQLGlot AST lineage, and emitted as Databricks SQL with single-query CTE conversion for temp-table batches, Databricks string/DATEADD/DATEDIFF normalization, and fail-closed target-dialect validation. Preserves pre-existing parameter boundaries, Source-to-SQL-to-Result named-binding contracts, indentation-safe navigation rendering, missing-field resilience, and verified atomic writes. Semantic table/column/measure names and DAX are protected.

Edit only ROOT_FOLDER and DRY_RUN in USER SETTINGS. Each immediate report folder under ROOT_FOLDER is processed independently; per-report Mapping Template.xlsx and Metadata Log.xlsx live in that report folder/Migration.

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
ROOT_FOLDER = r"C:\Your\Migration Batch Folder"
DRY_RUN = True
ALIGN_TARGET_COLUMN_NAMES = True  # False = legacy-alias output contracts
RENAME_SEMANTIC_COLUMNS = True    # Phase 3: False = source-only (Phase 2), True = full semantic + report rename
# =============================================================================

import hashlib
import os
import re
from types import MappingProxyType
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

def _env_bool(name: str, default: bool) -> bool:
    """Read a strict optional boolean environment override."""
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return default
    value = raw.strip().casefold()
    if value in {"1", "true", "yes", "y", "on"}:
        return True
    if value in {"0", "false", "no", "n", "off"}:
        return False
    raise ValueError(f"{name} must be one of true/false, yes/no, 1/0; received: {raw!r}")

def apply_runtime_settings() -> None:
    """Apply optional deployment overrides without editing migration logic.

    Supported values:
      PBI_MIGRATOR_ROOT_FOLDER
      PBI_MIGRATOR_DRY_RUN
      PBI_MIGRATOR_ALIGN_TARGET_COLUMN_NAMES
    The in-file USER SETTINGS remain the defaults.
    """
    global ROOT_FOLDER, DRY_RUN, ALIGN_TARGET_COLUMN_NAMES, RENAME_SEMANTIC_COLUMNS
    root_override = os.getenv("PBI_MIGRATOR_ROOT_FOLDER")
    if root_override and root_override.strip():
        ROOT_FOLDER = root_override.strip()
    DRY_RUN = _env_bool("PBI_MIGRATOR_DRY_RUN", DRY_RUN)
    ALIGN_TARGET_COLUMN_NAMES = _env_bool(
        "PBI_MIGRATOR_ALIGN_TARGET_COLUMN_NAMES", ALIGN_TARGET_COLUMN_NAMES
    )
    RENAME_SEMANTIC_COLUMNS = _env_bool(
        "PBI_MIGRATOR_RENAME_SEMANTIC_COLUMNS", RENAME_SEMANTIC_COLUMNS
    )

def validate_runtime_settings() -> None:
    """Fail early on placeholder or empty execution settings."""
    if not clean(ROOT_FOLDER):
        raise ValueError("ROOT_FOLDER is blank. Set USER SETTINGS or PBI_MIGRATOR_ROOT_FOLDER.")
    normalized = ROOT_FOLDER.replace("/", "\\").casefold()
    if normalized == r"c:\your\migration batch folder".casefold():
        raise ValueError("ROOT_FOLDER still contains the template placeholder.")

LOG_NAME = "Metadata Log.xlsx"
MAPPING_NAME = "Mapping Template.xlsx"
TABLE_MAP = "Table and Column Mapping"
SOURCE_MAP = "Data Source Mapping"

TABLE_HEADERS = [
    "Legacy Table", "Target Table", "Legacy Column Name", "Legacy Data Type",
    "Target Column Name", "Target Data Type", "Retain Legacy Data Type",
    "Notes", "Dependency Notes",
]
SOURCE_HEADERS = [
    "Legacy Server", "Legacy Database", "Legacy Schema", "Legacy Item",
    "Target Server", "Target HTTP Path", "Target Catalog", "Target Schema", "Target Item",
]

LOG_SHEETS = {
    "Run Summary": ["Item", "Value"],
    "Parameters": [
        "Parameter Name", "Value", "Is Current", "Data Type",
        "Query Tables Using Parameter",
    ],
    "Query & Source Inventory": [
        "Object Type", "Query Name", "Table Mode", "Inventory Level", "Connection Type",
        "Referenced Queries", "Source", "Server", "Database", "Schema", "Item", "Native Query",
        "Origin", "Source File", "Connection Expression", "Uses Parameters",
        "Parameter Names", "Parameter Roles", "Resolved Parameter Values",
        "Contains Join", "Contains Merge", "Contains Append", "Contains Unpivot",
        "References Other Queries", "Migration Complexity",
        "Discovery Method", "Data Source Mapping Found", "Target Server", "Target HTTP Path",
        "Target Catalog", "Target Schema", "Target Item", "Notes",
    ],
    "Connection Changes": [
        "Input File", "Query Name", "Classification", "Legacy Connector", "Target Connector",
        "Legacy Server", "Legacy Database", "Target Server", "Target HTTP Path",
        "Target Catalog", "Target Schema", "Target Item", "Result",
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
        "Dependency Risk Summary", "Dependency Notes", "Source Files", "Validation Status",
    ],
    "Native SQL Changes": [
        "Input File", "Query Name", "Change Type", "Old Value", "New Value",
        "New Value (With Parameters)", "Occurrences", "Safe Replacement", "Message",
        "Source Dialect", "Target Dialect", "Validation Stage",
        "Original SHA-256", "Generated SHA-256",
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
    "Target Metadata Validation": [
        "Platform", "Target Server", "Target Database/Catalog", "Target Schema",
        "Target Object", "Target Column", "Legacy Provenance", "Validation Status",
        "Error Code", "Details",
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
    risk_summary: str = ""
    system_recommendation: str = ""
    dependency_notes: str = ""
    dependencies: dict[str, list[str]] = field(default_factory=dict)

# Optional, fail-closed physical-column compatibility overrides.
# Keys and values are normalized only for matching; emitted identifiers retain
# the spelling supplied in ``target_column``. Leave empty unless a specific,
# reviewed legacy/table/column/target-table exception is required.
PHYSICAL_COLUMN_COMPATIBILITY_OVERRIDES: dict[tuple[str, str, str], dict[str, str]] = {}

@dataclass
class SourceMapping:
    legacy_server: str
    legacy_database: str
    legacy_schema: str
    legacy_item: str
    target_server: str
    target_http_path: str
    target_catalog: str
    target_schema: str
    target_item: str

@dataclass
class NativeConnectionTarget:
    """Connection-level target from Data Source Mapping; no table identity."""
    target_server: str
    target_http_path: str
    target_catalog: str

# Databricks connector behavior is migration-code policy, not user input.
SOURCE_SQL_DIALECT = None
TARGET_SQL_DIALECT = "databricks"
TARGET_CONNECTOR = "Databricks.Catalogs"
TARGET_ITEM_KIND = "Table"
DATABRICKS_IMPLEMENTATION = "2.0"
ENABLE_FOLDING = True


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
    """Write UTF-8 text atomically without creating backup files."""
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(text, encoding="utf-8")
    temp.replace(path)

def verified_atomic_write(path: Path, text: str) -> tuple[bool, str]:
    """Atomically write and verify the exact generated content."""
    atomic_write(path, text)
    actual = read_text(path)
    if actual != text:
        return False, f"Post-write content verification failed for {path}"
    if sha(actual) != sha(text):
        return False, f"Post-write SHA-256 verification failed for {path}"
    return True, ""

def discover_paths(root_text: str) -> Paths:
    """Resolve one report directory into its PBIP semantic-model paths.

    Migration artifacts are written to the report-specific Migration subfolder.
    """
    root = Path(root_text).expanduser().resolve()
    if not root.is_dir():
        raise FileNotFoundError(f"Report folder does not exist: {root}")
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

def discover_report_roots(batch_root_text: str) -> list[Path]:
    """Return immediate report directories under ROOT_FOLDER.

    A report directory is recognized by exactly one immediate child folder whose
    name ends in `` Pbip``. For backward compatibility, ROOT_FOLDER itself is
    treated as a single report when it directly contains exactly one such folder.
    """
    batch_root = Path(batch_root_text).expanduser().resolve()
    if not batch_root.is_dir():
        raise FileNotFoundError(f"ROOT_FOLDER does not exist: {batch_root}")
    direct_pbip = [p for p in batch_root.iterdir() if p.is_dir() and p.name.casefold().endswith(" pbip")]
    if len(direct_pbip) == 1:
        return [batch_root]
    reports = []
    for child in sorted((p for p in batch_root.iterdir() if p.is_dir()), key=lambda p: p.name.casefold()):
        pbip_children = [p for p in child.iterdir() if p.is_dir() and p.name.casefold().endswith(" pbip")]
        if len(pbip_children) == 1:
            reports.append(child)
    if not reports:
        raise ValueError(
            f"No report folders found under {batch_root}. Expected immediate subfolders containing one '* Pbip' folder each."
        )
    return reports

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
    ws.sheet_view.showGridLines = True
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
                cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
    for col in range(1, ws.max_column + 1):
        vals = [len(clean(ws.cell(r, col).value)) for r in range(1, min(ws.max_row, 250) + 1)]
        width = min(max(max(vals or [10]) + 2, 13), 55)
        if clean(ws.cell(1, col).value) in {"Notes", "Dependency Risk Summary", "Dependency Notes", "Details", "Message", "Warning", "Recommended Review"}:
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
        mappings.append(MappingRow(
            excel_row=r,
            legacy_table=get("Legacy Table"),
            target_table=get("Target Table"),
            legacy_column=get("Legacy Column Name"),
            legacy_type=get("Legacy Data Type"),
            target_column=get("Target Column Name"),
            target_type=get("Target Data Type"),
            selection=get("Retain Legacy Data Type"),
            notes=get("Notes"),
        ))
    ws2 = wb[SOURCE_MAP]
    h2 = {clean(c.value): c.column for c in ws2[1]}
    for r in range(2, ws2.max_row + 1):
        get = lambda name: clean(ws2.cell(r, h2[name]).value) if name in h2 else ""
        values = [get(x) for x in SOURCE_HEADERS]
        if not any(values):
            continue
        required = [
            get("Legacy Server"), get("Legacy Database"),
            get("Target Server"), get("Target HTTP Path"), get("Target Catalog"),
        ]
        if not all(required):
            continue
        source_maps.append(SourceMapping(*values))
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

def build_dependency_risk_summary(mapping: MappingRow, deps: dict[str, list[str]], m_operations: set[str]) -> str:
    """Return a concise, evidence-based dependency and potential-impact summary."""
    legacy_family = normalize_type(mapping.legacy_type)
    target_family = normalize_type(mapping.target_type)
    type_text = f"{mapping.legacy_type or 'blank'} -> {mapping.target_type or 'blank'}"
    family_changed = bool(legacy_family and target_family and legacy_family != target_family)
    type_missing = not mapping.legacy_type or not mapping.target_type

    labels = []
    if deps.get("relationship"):
        labels.append("Relationship")
    if {"merge", "join", "fuzzy merge"}.intersection(m_operations):
        labels.append("Query contains Merge/Join")
    if deps.get("unpivot"):
        labels.append("Unpivot")
    if deps.get("sql"):
        labels.append("Native SQL")
    if {"append"}.intersection(m_operations):
        labels.append("Append")
    if {"expand", "expand record"}.intersection(m_operations):
        labels.append("Expand")
    if deps.get("m") and not any(label in labels for label in ("Query contains Merge/Join", "Unpivot", "Append", "Expand")):
        labels.append("Power Query")
    if deps.get("other"):
        labels.append("Semantic model")
    if not labels:
        labels.append("No structural dependency")

    if type_missing:
        severity = "MEDIUM" if any(deps.values()) else "LOW"
        impact = "Type information is incomplete; review is needed before assessing compatibility."
    elif deps.get("relationship") and family_changed:
        severity = "HIGH"
        impact = "Type-family change may break relationship compatibility and model filtering."
    elif {"merge", "join", "fuzzy merge"}.intersection(m_operations) and family_changed:
        severity = "MEDIUM"
        impact = "The containing query has a merge/join and the type family changes; this scan did not confirm that the mapped column is a join key, so review compatibility."
    elif deps.get("unpivot") and family_changed:
        severity = "HIGH"
        impact = "Type-family change may alter unpivoted values and downstream transformations."
    elif deps.get("sql") and family_changed:
        severity = "MEDIUM"
        impact = "SQL output type changes; downstream consumers may expect the legacy type."
    elif deps.get("m") and family_changed:
        severity = "MEDIUM"
        impact = "Power Query steps may depend on the legacy type and could refresh differently."
    elif any(deps.values()):
        severity = "LOW"
        impact = "Dependency exists, but no semantic type-family change was detected."
    elif family_changed:
        severity = "MEDIUM"
        impact = "Type family changes, but no structural dependency was detected."
    else:
        severity = "LOW"
        impact = "No significant structural type risk was detected."

    return f"{severity} | {' + '.join(labels)} | {type_text} | {impact}"

def dependency_analysis(paths: Paths, mappings: list[MappingRow]):
    """Analyze structural source dependencies without inspecting or rewriting DAX.

    This function preserves the established recommendation rules and enriches only
    the human-readable Dependency Notes written to the mapping template and log.
    """
    files = all_model_files(paths)
    contents = {path: read_text(path) for path in files}
    results = []

    operation_patterns = (
        ("merge", r"(?i)Table\.(?:NestedJoin|FuzzyNestedJoin)\s*\("),
        ("join", r"(?i)Table\.Join\s*\("),
        ("expand", r"(?i)Table\.ExpandTableColumn\s*\("),
        ("expand record", r"(?i)Table\.ExpandRecordColumn\s*\("),
        ("append", r"(?i)Table\.Combine\s*\("),
        ("unpivot", r"(?i)Table\.Unpivot\s*\("),
        ("unpivot other columns", r"(?i)Table\.UnpivotOtherColumns\s*\("),
        ("select columns", r"(?i)Table\.SelectColumns\s*\("),
        ("remove columns", r"(?i)Table\.RemoveColumns\s*\("),
        ("rename columns", r"(?i)Table\.RenameColumns\s*\("),
        ("type conversion", r"(?i)Table\.TransformColumnTypes\s*\("),
    )

    for mapping in mappings:
        deps = {key: [] for key in ("relationship", "m", "sql", "unpivot", "other")}
        source_files = set()
        dependency_queries = {key: set() for key in deps}
        m_operations = set()
        active_queries = set()
        disabled_queries = set()

        mismatch = (
            bool(mapping.legacy_type and mapping.target_type)
            and normalize_type(mapping.legacy_type) != normalize_type(mapping.target_type)
        )

        for path, full_text in contents.items():
            m_region, _, _ = extract_m_region(full_text)
            sql_regions = native_sql_regions(full_text)
            sql_text = "\n".join(region[2] for region in sql_regions)
            query_name = table_name(full_text, path.stem)

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

            relative_file = str(path.relative_to(paths.semantic.parent))
            source_files.add(relative_file)

            query_disabled = bool(re.search(
                r"(?i)(loadEnabled|enableLoad)\s*[:=]\s*false|PBI_IsEnabled\s*=\s*false",
                full_text,
            ))
            if query_disabled:
                disabled_queries.add(query_name)
            else:
                active_queries.add(query_name)

            if relationship_ref:
                deps["relationship"].append(path.name)
                dependency_queries["relationship"].add(query_name)

            if m_region and (legacy_ref or target_ref):
                deps["m"].append(path.name)
                dependency_queries["m"].add(query_name)

                for operation, pattern in operation_patterns:
                    if re.search(pattern, m_region):
                        m_operations.add(operation)

                if re.search(r"(?i)Table\.Unpivot(?:OtherColumns)?", m_region):
                    deps["unpivot"].append(path.name)
                    dependency_queries["unpivot"].add(query_name)

                if re.search(r"(?i)(sortByColumn|hierarchy\s+|level\s+)", full_text):
                    deps["other"].append(path.name)
                    dependency_queries["other"].add(query_name)

            if sql_regions and (legacy_ref or target_ref):
                deps["sql"].append(path.name)
                dependency_queries["sql"].add(query_name)

        for key in deps:
            deps[key] = sorted(set(deps[key]), key=str.casefold)

        # Preserve the established recommendation logic exactly.
        if not mapping.legacy_type:
            recommendation = "No"
            type_note = "Legacy data type is blank; legacy retention is not required."
        elif (
            mapping.target_type
            and normalize_type(mapping.legacy_type) == normalize_type(mapping.target_type)
        ):
            recommendation = "No"
            type_note = (
                "Legacy and target data types normalize to the same semantic type "
                f"({mapping.legacy_type} -> {mapping.target_type})."
            )
        elif mapping.target_type:
            recommendation = "Review"
            type_note = (
                f"Data type changes from {mapping.legacy_type} to {mapping.target_type}; "
                "user review is required."
            )
        else:
            recommendation = "Review"
            type_note = "Target data type is blank; user review is required."

        note_parts = [type_note]

        if deps["relationship"]:
            queries = sorted(dependency_queries["relationship"], key=str.casefold)
            note_parts.append(
                "Relationship dependency detected"
                + (f" for {', '.join(queries)}" if queries else "")
                + f" in {', '.join(deps['relationship'])}. "
                "Review relationship-key type compatibility before changing retention behavior."
            )

        if deps["sql"]:
            queries = sorted(dependency_queries["sql"], key=str.casefold)
            note_parts.append(
                "Native SQL dependency detected"
                + (f" for {', '.join(queries)}" if queries else "")
                + f" in {', '.join(deps['sql'])}. "
                "The migrated SQL may need to preserve the legacy output contract for downstream consumers."
            )

        if deps["m"]:
            queries = sorted(dependency_queries["m"], key=str.casefold)
            operation_text = ", ".join(sorted(m_operations, key=str.casefold))
            m_note = (
                "Power Query dependency detected"
                + (f" for {', '.join(queries)}" if queries else "")
                + f" in {', '.join(deps['m'])}."
            )
            if operation_text:
                m_note += f" Query-level operations present: {operation_text}. These operations do not by themselves prove that this mapped column is a key or operation input."
            m_note += " Review downstream transformation compatibility before changing retention behavior."
            note_parts.append(m_note)

        if deps["unpivot"]:
            queries = sorted(dependency_queries["unpivot"], key=str.casefold)
            note_parts.append(
                "Unpivot dependency detected"
                + (f" for {', '.join(queries)}" if queries else "")
                + f" in {', '.join(deps['unpivot'])}. "
                "Source column names or values may have downstream significance after unpivoting."
            )

        if deps["other"]:
            queries = sorted(dependency_queries["other"], key=str.casefold)
            note_parts.append(
                "Additional semantic dependency detected"
                + (f" for {', '.join(queries)}" if queries else "")
                + f" in {', '.join(deps['other'])}."
            )

        if active_queries:
            note_parts.append(
                "Active query context: "
                + ", ".join(sorted(active_queries, key=str.casefold)) + "."
            )
        if disabled_queries:
            note_parts.append(
                "Disabled query context: "
                + ", ".join(sorted(disabled_queries, key=str.casefold)) + "."
            )

        if not any(deps.values()):
            note_parts.append(
                "No relationship, native SQL, Power Query, unpivot, or other structural dependency "
                "was detected for this mapping."
            )

        note_parts.append("DAX is protected and was not inspected for dependency analysis.")
        note = " ".join(part for part in note_parts if clean(part))

        mapping.risk_summary = build_dependency_risk_summary(mapping, deps, m_operations)
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
            "Dependency Risk Summary": mapping.risk_summary,
            "Dependency Notes": note,
            "Source Files": "; ".join(sorted(source_files, key=str.casefold)),
            "Validation Status": "Completed",
        })
    return results

def update_mapping_recommendations(path: Path, mappings: list[MappingRow]) -> None:
    wb = load_workbook(path)
    ws = wb[TABLE_MAP]
    h = {clean(c.value): c.column for c in ws[1]}
    for m in mappings:
        retain_cell = ws.cell(m.excel_row, h["Retain Legacy Data Type"])
        summary_cell = ws.cell(m.excel_row, h["Notes"])
        notes_cell = ws.cell(m.excel_row, h["Dependency Notes"])
        if clean(retain_cell.value) not in {"Yes", "No", "Review"}:
            retain_cell.value = m.system_recommendation
        # Notes is a concise system-generated risk summary and is refreshed each analysis run.
        summary_cell.value = m.risk_summary
        # Keep the established protection for detailed user-entered dependency notes.
        if not clean(notes_cell.value):
            notes_cell.value = m.dependency_notes
    style_tab(ws, add_table=False)
    style_tab(wb[SOURCE_MAP], add_table=False)
    wb.save(path)

def detect_source(block: str):
    """Detect Synapse or Databricks connectors and their navigation context."""
    synapse = re.search(
        r'Sql\.Database\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', block, re.I
    )
    if synapse:
        server, database = synapse.group(1), synapse.group(2)
        nav = re.search(
            r'\[\s*Schema\s*=\s*"([^"]+)"\s*,\s*Item\s*=\s*"([^"]+)"', block, re.I
        )
        schema, item = (nav.group(1), nav.group(2)) if nav else ("", "")
        return "SQL", f"{server};{database}", server, database, schema, item

    databricks = re.search(
        r'Databricks\.Catalogs\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"', block, re.I
    )
    if databricks:
        server, http_path = databricks.group(1), databricks.group(2)
        catalog_option = re.search(r'\bCatalog\s*=\s*"([^"]+)"', block, re.I)
        database_nav = re.search(
            r'\[\s*Name\s*=\s*"([^"]+)"\s*,\s*Kind\s*=\s*"Database"\s*\]', block, re.I
        )
        catalog = (
            catalog_option.group(1) if catalog_option else
            database_nav.group(1) if database_nav else ""
        )
        schema_nav = re.search(
            r'\[\s*Name\s*=\s*"([^"]+)"\s*,\s*Kind\s*=\s*"Schema"\s*\]', block, re.I
        )
        item_nav = re.search(
            r'\[\s*Name\s*=\s*"([^"]+)"\s*,\s*Kind\s*=\s*"(?:Table|View)"\s*\]', block, re.I
        )
        schema = schema_nav.group(1) if schema_nav else ""
        item = item_nav.group(1) if item_nav else ""
        return "Databricks", f"{server};{http_path};{catalog}", server, catalog, schema, item

    patterns = [
        ("Excel", r'Excel\.Workbook\s*\(\s*File\.Contents\s*\(\s*"([^"]+)"'),
        ("SharePoint", r'SharePoint\.(?:Files|Contents|Tables)\s*\(\s*"([^"]+)"'),
        ("Folder", r'Folder\.(?:Files|Contents)\s*\(\s*"([^"]+)"'),
        ("CSV", r'Csv\.Document\s*\(\s*File\.Contents\s*\(\s*"([^"]+)"'),
        ("Web", r'Web\.Contents\s*\(\s*"([^"]+)"'),
        ("OData", r'OData\.Feed\s*\(\s*"([^"]+)"'),
        ("JSON", r'Json\.Document\s*\(\s*(?:File\.Contents\s*\(\s*)?"([^"]+)"'),
    ]
    for typ, pat in patterns:
        match = re.search(pat, block, re.I)
        if match:
            return typ, match.group(1), "", "", "", ""
    return "Other", "", "", "", "", ""

@dataclass
class DiscoveredNativeSql:
    sql: str
    construct: str
    argument_expression: str
    # Complete M expression that originally produced SQL. For named consumers such
    # as Value.NativeQuery(Source, SQL), this is the SQL binding expression rather
    # than the token ``SQL``. It is authoritative for parameter provenance.
    parameter_expression: str = ""
    binding_name: str = ""
    resolution: str = "Static M expression"
    # Exact source coordinates captured during discovery. These coordinates are
    # authoritative for migration and remove the need to rediscover a payload by
    # comparing SQL text after SQLGlot has reformatted it.
    write_start: int = -1
    write_end: int = -1
    write_mode: str = ""
    original_encoded_payload: str = ""

@dataclass(frozen=True)
class MBindingSpan:
    """One top-level M let binding with immutable absolute source coordinates."""
    name: str
    expression: str
    expression_start: int
    expression_end: int
    ordinal: int = 0

@dataclass(frozen=True)
class NativeQueryContract:
    """One native consumer resolved in its smallest enclosing M let scope."""
    source: MBindingSpan | None
    sql: MBindingSpan | None
    result: MBindingSpan | None
    source_expression: str
    sql_expression: str
    source_reference: str
    sql_reference: str
    result_reference: str
    consumer_start: int
    consumer_end: int
    sql_argument_start: int
    sql_argument_end: int
    scope_start: int
    scope_end: int
    scope_bindings: tuple[MBindingSpan, ...]

def _stable_m_binding_spans(value: str) -> dict[str, MBindingSpan]:
    """Return the single authoritative binding directory for discovery/write-back.

    The existing structural let parser owns the coordinates. Consumers must carry
    these coordinates forward and must never rediscover a binding after SQL
    evaluation or SQLGlot formatting.
    """
    _region, region_start, _region_end, rows = _m_let_bindings_with_spans(value)
    if region_start < 0:
        return {}
    directory = {}
    for row in rows:
        name = clean(row.get("name_value"))
        start = region_start + int(row.get("expr_start", -1))
        end = region_start + int(row.get("expr_end", -1))
        expression = row.get("expression", "")
        if not name or start < 0 or end <= start or end > len(value):
            continue
        # The source slice, not a normalized copy, is authoritative.
        source_expression = value[start:end]
        directory[norm(name)] = MBindingSpan(
            name=name,
            expression=source_expression,
            expression_start=start,
            expression_end=end,
            ordinal=int(row.get("ordinal", 0)),
        )
    return directory

def _inventory_m_bindings(value: str) -> dict[str, str]:
    """Return binding expressions from the same parser used for write coordinates."""
    return {
        key: span.expression.strip()
        for key, span in _stable_m_binding_spans(value).items()
    }

def _strip_balanced_parentheses(expression: str) -> str:
    value = clean(expression)
    while value.startswith('(') and value.endswith(')'):
        depth = 0
        in_string = False
        closes_at_end = False
        for index, char in enumerate(value):
            if in_string:
                if char == '"':
                    if index + 1 < len(value) and value[index + 1] == '"':
                        continue
                    in_string = False
                continue
            if char == '"': in_string = True
            elif char == '(': depth += 1
            elif char == ')':
                depth -= 1
                if depth == 0:
                    closes_at_end = index == len(value) - 1
                    break
        if not closes_at_end:
            break
        value = value[1:-1].strip()
    return value

def _evaluate_static_m_text(expression: str, bindings: dict[str, str], parameters: dict[str, str], visiting=None):
    """Evaluate only static M text used to build native SQL; never execute M."""
    visiting = set() if visiting is None else set(visiting)
    expression = _strip_balanced_parentheses(expression)
    scalar, _data_type = _parameter_scalar(expression)
    if scalar is not None:
        return scalar
    identifier = _m_identifier_name(expression)
    if identifier:
        parameter_name = next((name for name in parameters if norm(name) == norm(identifier)), '')
        if parameter_name:
            return parameters.get(parameter_name, '')
        key = norm(identifier)
        if key in visiting or key not in bindings:
            return None
        visiting.add(key)
        return _evaluate_static_m_text(bindings[key], bindings, parameters, visiting)
    parts = split_m_top_level(expression, delimiter='&')
    if len(parts) > 1:
        values = [_evaluate_static_m_text(part, bindings, parameters, visiting) for part in parts]
        return None if any(value is None for value in values) else ''.join(values)
    combine = re.fullmatch(r'(?is)Text\.Combine\s*\((.*)\)', expression)
    if combine:
        args = split_m_top_level(combine.group(1))
        if not args:
            return None
        members = clean(args[0])
        if not (members.startswith('{') and members.endswith('}')):
            return None
        values = [_evaluate_static_m_text(member, bindings, parameters, visiting)
                  for member in split_m_top_level(members[1:-1])]
        if any(value is None for value in values):
            return None
        separator = ''
        if len(args) > 1:
            separator = _evaluate_static_m_text(args[1], bindings, parameters, visiting)
            if separator is None:
                return None
        return separator.join(values)
    return None

def _m_record_field_expression(record_text: str, field_name: str) -> str:
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

def _discovered_sql_write_location(
    block: str,
    construct: str,
    argument_expression: str,
    binding_name: str,
    sql_text: str,
    argument_start: int = -1,
    argument_end: int = -1,
):
    """Return the authoritative native-SQL write span.

    Named SQL arguments must resolve to their let-binding expression. If that
    binding is unavailable, discovery fails closed and never falls back to the
    Value.NativeQuery consumer argument.
    """
    if binding_name:
        binding = _binding_expression_span(block, binding_name)
        if binding:
            expression = binding["expression"]
            literal = re.fullmatch(r'\s*"((?:[^"]|"")*)"\s*', expression, re.DOTALL)
            if literal:
                encoded = literal.group(1)
            else:
                try:
                    encoded = encode_m_sql_string(sql_text, "")
                except TypeError:
                    encoded = encode_m_sql_string(sql_text)
            return {
                "start": binding["start"],
                "end": binding["end"],
                "mode": "binding_expression",
                "encoded": encoded,
            }
        # Named-binding discovery is authoritative. Returning an invalid span
        # makes materialization fail closed with WRITE_POSITION_NOT_FOUND.
        return {"start": -1, "end": -1, "mode": "", "encoded": ""}

    # Inline SQL, including concatenation and Text.Combine, must be replaced by
    # the complete second Value.NativeQuery argument captured by the call parser.
    if 0 <= argument_start < argument_end <= len(block):
        expression = block[argument_start:argument_end]
        literal = re.fullmatch(r'\s*"((?:[^"]|"")*)"\s*', expression, re.DOTALL)
        encoded = literal.group(1) if literal else ""
        if not encoded:
            try:
                encoded = encode_m_sql_string(sql_text, "")
            except TypeError:
                encoded = encode_m_sql_string(sql_text)
        return {
            "start": argument_start,
            "end": argument_end,
            "mode": "argument_expression",
            "encoded": encoded,
        }

    return {"start": -1, "end": -1, "mode": "", "encoded": ""}

def _candidate_m_let_scopes(value: str):
    """Parse every lexical M let scope with absolute binding coordinates.

    TMDL embeds M below ``source =`` or ``expression ... =``. Nested M expressions
    and expression blocks can therefore be invisible to a single outer-region
    parser. Each ``let`` token is parsed independently through the established
    balanced binding parser and then rebased to the original source text.
    """
    scopes = []
    prefix = 'source =\n'
    for match in re.finditer(r'(?i)\blet\b', value):
        start = match.start()
        wrapped = prefix + value[start:]
        _region, region_start, region_end, rows = _m_let_bindings_with_spans(wrapped)
        if region_start < 0 or not rows:
            continue
        adjustment = start - len(prefix)
        bindings = []
        for row in rows:
            expr_start = adjustment + region_start + int(row.get('expr_start', -1))
            expr_end = adjustment + region_start + int(row.get('expr_end', -1))
            name = clean(row.get('name_value'))
            if not name or expr_start < 0 or expr_end <= expr_start or expr_end > len(value):
                continue
            bindings.append(MBindingSpan(
                name=name,
                expression=value[expr_start:expr_end],
                expression_start=expr_start,
                expression_end=expr_end,
                ordinal=int(row.get('ordinal', 0)),
            ))
        if not bindings:
            continue
        absolute_end = min(len(value), adjustment + region_end)
        key = (start, absolute_end, tuple((b.name, b.expression_start, b.expression_end) for b in bindings))
        if any(existing[0] == key for existing in scopes):
            continue
        scopes.append((key, start, absolute_end, tuple(bindings)))
    return [(start, end, bindings) for _key, start, end, bindings in scopes]

def _lexical_scope_for_position(value: str, position: int):
    """Return the smallest parsed let scope containing an absolute position."""
    candidates = [row for row in _candidate_m_let_scopes(value)
                  if row[0] <= position <= row[1]]
    if not candidates:
        return None
    return min(candidates, key=lambda row: (row[1] - row[0], -row[0]))

def _contract_binding_directory(contract: NativeQueryContract) -> dict[str, str]:
    return {norm(binding.name): binding.expression.strip()
            for binding in contract.scope_bindings}

def _targeted_named_binding_span(value: str, binding_name: str, before_position: int):
    """Find the nearest line-level M binding before a consumer, fail-closed.

    This is a resilience fallback for very large generated SQL expressions where a
    broad lexical-scope parse may omit one binding. It recognizes only a declaration
    at the beginning of a line, then scans the expression with M string/comment and
    nesting rules until the next top-level comma or ``in`` token.
    """
    if not binding_name:
        return None
    quoted = '#"' + binding_name.replace('"', '""') + '"'
    names = [re.escape(binding_name)]
    if not binding_name.startswith('#"'):
        names.append(re.escape(quoted))
    pattern = re.compile(
        r'(?mi)^[ \t]*(?P<name>' + '|'.join(names) + r')[ \t]*=[ \t]*'
    )
    declarations = [match for match in pattern.finditer(value, 0, before_position)]
    if not declarations:
        return None
    match = declarations[-1]
    start = match.end()
    index = start
    paren = brace = bracket = 0
    in_string = False
    line_comment = False
    block_comment = 0
    while index < before_position:
        char = value[index]
        nxt = value[index + 1] if index + 1 < len(value) else ''
        if line_comment:
            if char in '\r\n':
                line_comment = False
            index += 1
            continue
        if block_comment:
            if char == '/' and nxt == '*':
                block_comment += 1; index += 2; continue
            if char == '*' and nxt == '/':
                block_comment -= 1; index += 2; continue
            index += 1
            continue
        if in_string:
            if char == '"':
                if nxt == '"':
                    index += 2; continue
                in_string = False
            index += 1
            continue
        if char == '/' and nxt == '/':
            line_comment = True; index += 2; continue
        if char == '/' and nxt == '*':
            block_comment = 1; index += 2; continue
        if char == '"': in_string = True
        elif char == '(': paren += 1
        elif char == ')': paren = max(0, paren - 1)
        elif char == '{': brace += 1
        elif char == '}': brace = max(0, brace - 1)
        elif char == '[': bracket += 1
        elif char == ']': bracket = max(0, bracket - 1)
        elif paren == 0 and brace == 0 and bracket == 0:
            if char == ',':
                break
            if (char == 'i' or char == 'I') and re.match(r'(?i)in\b', value[index:]):
                previous = value[index - 1] if index else ' '
                if previous.isspace():
                    break
        index += 1
    end = index
    while end > start and value[end - 1].isspace():
        end -= 1
    if end <= start:
        return None
    parsed_name = _parse_query_identifier(match.group('name')) or match.group('name')
    return MBindingSpan(
        name=parsed_name,
        expression=value[start:end],
        expression_start=start,
        expression_end=end,
        ordinal=0,
    )

def _nearest_named_binding(value: str, binding_name: str, before_position: int,
                           scope_bindings=()):
    candidates = [binding for binding in scope_bindings
                  if norm(binding.name) == norm(binding_name)
                  and binding.expression_end <= before_position]
    for _scope_start, _scope_end, bindings in _candidate_m_let_scopes(value):
        candidates.extend(
            binding for binding in bindings
            if norm(binding.name) == norm(binding_name)
            and binding.expression_end <= before_position
        )
    stable = _stable_m_binding_spans(value).get(norm(binding_name))
    if stable is not None and stable.expression_end <= before_position:
        candidates.append(stable)
    if candidates:
        unique = {(row.expression_start, row.expression_end, norm(row.name)): row
                  for row in candidates}
        return max(unique.values(), key=lambda row: row.expression_start)
    return _targeted_named_binding_span(value, binding_name, before_position)

def _native_query_contracts(block: str):
    """Return every Value.NativeQuery with resilient lexical binding resolution."""
    contracts = []
    seen = set()
    for call in sorted(_balanced_m_calls(block, ('Value.NativeQuery',)), key=lambda row: row['start']):
        args = split_m_top_level_with_spans(call['body'])
        if len(args) < 2:
            continue
        opening = block.find('(', call['start'], call['end'])
        if opening < 0:
            continue
        scope = _lexical_scope_for_position(block, call['start'])
        scope_start, scope_end, scope_bindings = scope if scope else (0, len(block), tuple(_stable_m_binding_spans(block).values()))
        directory = {norm(row.name): row for row in scope_bindings}
        source_expression = args[0]['text'].strip()
        sql_expression = args[1]['text'].strip()
        source_name = _m_identifier_name(source_expression)
        sql_name = _m_identifier_name(sql_expression)
        source = directory.get(norm(source_name)) if source_name else None
        sql = directory.get(norm(sql_name)) if sql_name else None
        if source is None and source_name:
            source = _nearest_named_binding(block, source_name, call['start'], scope_bindings)
        if sql is None and sql_name:
            sql = _nearest_named_binding(block, sql_name, call['start'], scope_bindings)
        merged = {(row.expression_start, row.expression_end, norm(row.name)): row
                  for row in scope_bindings}
        for row in (source, sql):
            if row is not None:
                merged[(row.expression_start, row.expression_end, norm(row.name))] = row
        effective_bindings = tuple(sorted(merged.values(), key=lambda row: row.expression_start))
        result = next((row for row in effective_bindings
                       if row.expression_start <= call['start'] and call['end'] <= row.expression_end), None)
        sql_start = opening + 1 + args[1]['start']
        sql_end = opening + 1 + args[1]['end']
        key = (call['start'], call['end'], sql_start, sql_end)
        if key in seen:
            continue
        seen.add(key)
        contracts.append(NativeQueryContract(
            source=source, sql=sql, result=result,
            source_expression=source_expression, sql_expression=sql_expression,
            source_reference=source_name, sql_reference=sql_name,
            result_reference=result.name if result else '',
            consumer_start=call['start'], consumer_end=call['end'],
            sql_argument_start=sql_start, sql_argument_end=sql_end,
            scope_start=scope_start, scope_end=scope_end,
            scope_bindings=effective_bindings,
        ))
    return contracts

def discover_native_sql_payloads(block: str, current_parameters: dict[str, str]):
    """Resolve and locate native SQL within the consumer's lexical M scope."""
    payloads = []
    seen_locations = set()
    for contract in _native_query_contracts(block):
        bindings = _contract_binding_directory(contract)
        sql_text = _evaluate_static_m_text(contract.sql_expression, bindings, current_parameters)
        if sql_text is None:
            continue
        if contract.sql is not None:
            target_expression = contract.sql.expression
            literal = re.fullmatch(r'\s*"((?:[^"]|"")*)"\s*', target_expression, re.DOTALL)
            encoded = literal.group(1) if literal else encode_m_sql_string(sql_text, '')
            location = {'start': contract.sql.expression_start, 'end': contract.sql.expression_end,
                        'mode': 'binding_expression', 'encoded': encoded}
            binding_name = contract.sql.name
            resolution = 'Lexical Source -> SQL binding -> Result contract'
        else:
            location = _discovered_sql_write_location(
                block, 'Value.NativeQuery', contract.sql_expression, '', sql_text,
                argument_start=contract.sql_argument_start, argument_end=contract.sql_argument_end)
            binding_name = ''
            resolution = 'Lexical Source -> inline SQL expression -> Result contract'
        loc_key = (location['start'], location['end'], location['mode'])
        if loc_key in seen_locations:
            continue
        seen_locations.add(loc_key)
        payloads.append(DiscoveredNativeSql(
            sql=sql_text,
            construct='Value.NativeQuery lexical Source -> SQL -> Result contract',
            argument_expression=contract.sql_expression,
            parameter_expression=(contract.sql.expression if contract.sql is not None else contract.sql_expression),
            binding_name=binding_name, resolution=resolution,
            write_start=location['start'], write_end=location['end'],
            write_mode=location['mode'], original_encoded_payload=location['encoded']))

    # Preserve explicit Sql.Database [Query=...] support when Value.NativeQuery is absent.
    outer_bindings = _stable_m_binding_spans(block)
    bindings = {key: span.expression.strip() for key, span in outer_bindings.items()}
    for call in _balanced_m_calls(block, ('Sql.Database',)):
        args = split_m_top_level(call['body'])
        if len(args) < 3:
            continue
        expression = _m_record_field_expression(args[2], 'Query')
        if not expression:
            continue
        sql_text = _evaluate_static_m_text(expression, bindings, current_parameters)
        if sql_text is None:
            continue
        binding_name = _m_identifier_name(expression)
        authoritative = outer_bindings.get(norm(binding_name)) if binding_name else None
        if authoritative is not None:
            location = {'start': authoritative.expression_start, 'end': authoritative.expression_end,
                        'mode': 'binding_expression', 'encoded': encode_m_sql_string(sql_text, '')}
        else:
            location = _discovered_sql_write_location(block, 'Sql.Database Query option', expression, binding_name, sql_text)
        loc_key = (location['start'], location['end'], location['mode'])
        if loc_key in seen_locations:
            continue
        seen_locations.add(loc_key)
        payloads.append(DiscoveredNativeSql(
            sql=sql_text, construct='Sql.Database Query option', argument_expression=expression,
            parameter_expression=(authoritative.expression if authoritative is not None else expression),
            binding_name=binding_name, resolution='Sql.Database Query option',
            write_start=location['start'], write_end=location['end'],
            write_mode=location['mode'], original_encoded_payload=location['encoded']))
    return payloads

def _persistent_sql_sources(sql_text: str):
    sources, function_audit, parse_errors = _sql_physical_sources(sql_text)
    persistent = [row for row in sources if not clean(row.get('Item')).startswith('#')]
    return persistent, function_audit, parse_errors

def _promote_unique_native_source(payloads):
    """Promote exactly one unique persistent SQL object to the query-level row."""
    records = {}
    errors = []
    for payload in payloads:
        sources, _audit, parse_errors = _persistent_sql_sources(payload.sql)
        errors.extend(parse_errors)
        for row in sources:
            key = (norm(row.get('Catalog')), norm(row.get('Schema')), norm(row.get('Item')))
            records[key] = row
    if len(records) == 1:
        row = next(iter(records.values()))
        return row.get('Physical Object', ''), row.get('Schema', ''), row.get('Item', ''), errors
    if len(records) > 1:
        return 'Multiple native SQL sources', '', '', errors
    return '', '', '', errors

def _sql_internal_object_names(statements) -> set[str]:
    """Return SQL-local relation names that are not external data sources.

    Includes CTE names and statement-created objects such as SELECT INTO targets.
    Temporary tables are also excluded independently by their # prefix.
    """
    names = set()
    for statement in statements:
        if statement is None:
            continue
        for cte in statement.find_all(exp.CTE):
            name = clean(cte.alias_or_name)
            if name:
                names.add(norm(name))
        for into in statement.find_all(exp.Into):
            target = into.this
            if isinstance(target, exp.Table):
                name = clean(target.name)
                if name:
                    names.add(norm(name))
        for create in statement.find_all(exp.Create):
            target = create.this
            if isinstance(target, exp.Table):
                name = clean(target.name)
                if name and (name.startswith('#') or not clean(target.db)):
                    names.add(norm(name))
    return names

def _sql_physical_sources(sql_text: str, dialect: str | None = None):
    """Return distinct external SQL dependencies found through SQLGlot.

    Query-local CTEs, derived relations, aliases, temporary tables, and objects
    created inside the SQL batch are excluded. Distinct source identity is
    catalog/schema/item; aliases are retained only as diagnostic metadata and do
    not produce extra Query & Source Inventory rows.
    """
    normalized_sql, function_audit = normalize_synapse_sql_functions(sql_text)
    records = {}
    parse_errors = []
    try:
        statements, detected_dialect, _detection = parse_detected_statements(normalized_sql) if dialect is None else (sqlglot.parse(normalized_sql, read=dialect), dialect, "caller")
    except ParseError as exc:
        return [], function_audit, [sanitize_diagnostic_text(exc)]

    internal_names = _sql_internal_object_names(statements)
    qualified_aliases = set()
    for statement in statements:
        if statement is None:
            continue
        for table in statement.find_all(exp.Table):
            if clean(table.db) or clean(table.catalog):
                alias = clean(table.alias)
                if alias:
                    qualified_aliases.add(norm(alias))

    for statement_index, statement in enumerate(statements, 1):
        if statement is None:
            continue
        try:
            scopes = list(traverse_scope(statement))
        except Exception as exc:
            parse_errors.append(f"Scope analysis failed for statement {statement_index}: {sanitize_diagnostic_text(exc)}")
            continue
        for scope_index, scope in enumerate(scopes, 1):
            for alias, source in scope.sources.items():
                if not isinstance(source, exp.Table):
                    continue
                catalog = clean(source.catalog)
                schema = clean(source.db)
                item = clean(source.name)
                item_key = norm(item)
                if not item or item.startswith('#'):
                    continue
                if not catalog and not schema and item_key in internal_names:
                    continue
                if not catalog and not schema and item_key in qualified_aliases:
                    continue
                alias_name = clean(alias) or clean(source.alias) or item
                physical = ".".join(part for part in (catalog, schema, item) if part)
                key = (norm(catalog), norm(schema), norm(item))
                record = records.setdefault(key, {
                    "Catalog": catalog,
                    "Schema": schema,
                    "Item": item,
                    "Alias": "",
                    "Aliases": set(),
                    "Physical Object": physical,
                    "Occurrences": 0,
                    "Scopes": set(),
                })
                if alias_name and norm(alias_name) != item_key:
                    record["Aliases"].add(alias_name)
                record["Occurrences"] += 1
                record["Scopes"].add(f"statement {statement_index}/scope {scope_index}")

    output = []
    for record in records.values():
        aliases = sorted(record.pop("Aliases"), key=str.casefold)
        record["Alias"] = ", ".join(aliases)
        record["Scopes"] = "; ".join(sorted(record["Scopes"]))
        output.append(record)
    output.sort(key=lambda row: (
        norm(row["Catalog"]), norm(row["Schema"]), norm(row["Item"]),
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
    current_parameters: dict[str, str] | None = None,
):
    """Create one inventory row per persistent physical source in resolved native SQL."""
    rows = []
    payloads = discover_native_sql_payloads(block, current_parameters or {})
    for payload_index, payload in enumerate(payloads, 1):
        physical_sources, function_audit, parse_errors = _persistent_sql_sources(payload.sql)
        for physical in physical_sources:
            notes = (
                f"Discovered from {payload.construct} payload {payload_index} using binding-aware "
                f"static M resolution and SQLGlot AST; alias={physical['Alias']}; "
                f"occurrences={physical['Occurrences']}; scopes={physical['Scopes']}"
            )
            rows.append({
                "Object Type": "Native SQL Source", "Query Name": query,
                "Inventory Level": _inventory_level(origin, query_status),
                "Connection Type": "Native SQL Table Reference",
                "Connection Expression": f"{payload.construct} payload {payload_index}",
                "Uses Parameters": "Yes" if current_parameters else "No",
                "Parameter Names": "", "Parameter Roles": "", "Resolved Parameter Values": "",
                "References Other Queries": "No", "Referenced Queries": "",
                "Discovery Method": "Binding-aware M resolution + SQLGlot AST scope analysis",
                "Query Type": "Native SQL Table Reference", "Source Type": "SQL",
                "Source": physical["Physical Object"], "Server": server, "Database": database,
                "Schema": physical["Schema"], "Item": physical["Item"], "Native Query": "Yes",
                "Query Status": query_status, "Origin": origin, "Source File": source_file,
                "Contains Join": contains_join, "Contains Merge": contains_merge,
                "Contains Append": contains_append, "Contains Unpivot": contains_unpivot,
                "Migration Complexity": "High", "Notes": notes,
            })
        if parse_errors:
            rows.append({
                "Object Type": "Native SQL Source Parse Error", "Query Name": query,
                "Inventory Level": _inventory_level(origin, query_status),
                "Connection Type": "Native SQL Table Reference",
                "Connection Expression": f"{payload.construct} payload {payload_index}",
                "Uses Parameters": "Yes" if current_parameters else "No",
                "Parameter Names": "", "Parameter Roles": "", "Resolved Parameter Values": "",
                "References Other Queries": "No", "Referenced Queries": "",
                "Discovery Method": "Binding-aware M resolution succeeded; SQLGlot analysis failed",
                "Query Type": "Native SQL Table Reference", "Source Type": "SQL", "Source": "",
                "Server": server, "Database": database, "Schema": "", "Item": "",
                "Native Query": "Yes", "Query Status": query_status, "Origin": origin,
                "Source File": source_file, "Contains Join": contains_join,
                "Contains Merge": contains_merge, "Contains Append": contains_append,
                "Contains Unpivot": contains_unpivot, "Migration Complexity": "High",
                "Notes": "Native SQL source extraction failed: " + " | ".join(parse_errors),
            })
    return rows

def _deduplicate_inventory_rows(rows):
    """Deduplicate inventory rows by physical source identity; aliases are diagnostic only."""
    seen = set()
    output = []
    for row in rows:
        key = tuple(norm(row.get(field)) for field in (
            "Inventory Level", "Object Type", "Query Name", "Server", "Database",
            "Schema", "Item", "Source", "Source File", "Notes",
        ))
        if key in seen:
            continue
        seen.add(key)
        output.append(row)
    return output

def _inventory_source_mapping(server: str, database: str, schema: str, item: str, sources: list[SourceMapping]):
    """Resolve inventory mapping without forcing model edits for ambiguous query-level rows.

    Exact schema/item identity is authoritative. When a parent query has only
    server/database coordinates, mappings that share one Databricks connection
    are treated as a connection-level match; table-specific target fields stay
    blank until the Native SQL Source/detail row supplies schema/item.
    """
    matches = []
    for source in sources:
        if norm(source.legacy_server) != norm(server):
            continue
        if norm(source.legacy_database) != norm(database):
            continue
        if schema and source.legacy_schema and norm(source.legacy_schema) != norm(schema):
            continue
        if item and source.legacy_item and norm(source.legacy_item) != norm(item):
            continue
        matches.append(source)
    if not matches:
        return None
    if schema or item:
        exact = [s for s in matches
                 if (not schema or not s.legacy_schema or norm(s.legacy_schema) == norm(schema))
                 and (not item or not s.legacy_item or norm(s.legacy_item) == norm(item))]
        candidates = exact or matches
        target_keys = {(norm(s.target_server), norm(s.target_http_path), norm(s.target_catalog),
                        norm(s.target_schema), norm(s.target_item)) for s in candidates}
        return candidates[0] if len(target_keys) == 1 else None
    connection_keys = {(norm(s.target_server), norm(s.target_http_path), norm(s.target_catalog)) for s in matches}
    return matches[0] if len(connection_keys) == 1 else None

def enrich_inventory_with_source_mappings(rows: list[dict], sources: list[SourceMapping]):
    for row in rows:
        source = _inventory_source_mapping(
            clean(row.get("Server")), clean(row.get("Database")),
            clean(row.get("Schema")), clean(row.get("Item")), sources,
        )
        row["Data Source Mapping Found"] = "Yes" if source else "No"
        row["Target Server"] = source.target_server if source else ""
        row["Target HTTP Path"] = source.target_http_path if source else ""
        row["Target Catalog"] = source.target_catalog if source else ""
        # Parent query rows can legitimately match only at the connection level.
        # Do not invent a target schema/item from an arbitrary table mapping.
        has_object_identity = bool(clean(row.get("Schema")) or clean(row.get("Item")))
        row["Target Schema"] = ((source.target_schema or clean(row.get("Schema"))) if source and has_object_identity else "")
        row["Target Item"] = ((source.target_item or clean(row.get("Item"))) if source and has_object_identity else "")
    return rows

def _m_identifier_name(value: str) -> str:
    value = clean(value)
    quoted = re.fullmatch(r'#"((?:[^"]|"")+)"', value, re.DOTALL)
    if quoted:
        return quoted.group(1).replace('""', '"')
    return value if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_ ]*', value) else ''

def _resolve_m_connector_argument(argument: str, current: dict[str, str]):
    """Resolve a connector argument as a literal or genuine report parameter."""
    argument = clean(argument)
    scalar, _data_type = _parameter_scalar(argument)
    if scalar is not None:
        return scalar, ''
    identifier = _m_identifier_name(argument)
    if not identifier:
        return '', ''
    actual = next((name for name in current if norm(name) == norm(identifier)), '')
    if not actual:
        return '', ''
    return current.get(actual, ''), actual
def _sql_database_connector(block: str, current: dict[str, str]):
    """Resolve Sql.Database arguments from literals or genuine parameters."""
    calls = _balanced_m_calls(block, ('Sql.Database',))
    if not calls:
        return {'server':'', 'database':'', 'parameters':[], 'expression':''}
    # Use the earliest connector call in the query.
    call = sorted(calls, key=lambda row: row['start'])[0]
    args = split_m_top_level(call['body'])
    if len(args) < 2:
        return {'server':'', 'database':'', 'parameters':[], 'expression':''}
    server, server_parameter = _resolve_m_connector_argument(args[0], current)
    database, database_parameter = _resolve_m_connector_argument(args[1], current)
    parameters = [name for name in (server_parameter, database_parameter) if name]
    return {
        'server': server, 'database': database,
        'parameters': parameters,
        'expression': 'Sql.Database(' + args[0].strip() + ', ' + args[1].strip() + ')',
    }

def _static_schema_item(block: str):
    """Extract a unique static Schema/Item navigation independently of connector resolution."""
    pairs = {
        (match.group(1).replace('""','"'), match.group(2).replace('""','"'))
        for match in re.finditer(
            r'(?is)\[\s*Schema\s*=\s*"((?:[^"]|"")*)"\s*,\s*Item\s*=\s*"((?:[^"]|"")*)"',
            block,
        )
    }
    return next(iter(pairs)) if len(pairs) == 1 else ('', '')

def _source_inventory_for_query(block: str, current: dict[str, str]):
    connector = _sql_database_connector(block, current)
    schema, item = _static_schema_item(block)
    if connector['server'] or connector['database'] or schema or item:
        source = '.'.join(part for part in (schema, item) if part)
        return 'SQL', source, connector['server'], connector['database'], schema, item, connector
    source_type, source, server, database, schema, item = detect_source(block)
    return source_type, source, server, database, schema, item, connector

def _query_parameter_directory(paths: Paths, definitions: dict[str, dict]):
    """Build direct connector-based query-to-parameter and reverse usage directories."""
    current = {name: definition['current'] for name, definition in definitions.items()}
    query_to_parameters = {}
    parameter_to_queries = {name: set() for name in definitions}
    files = sorted(paths.tables.glob('*.tmdl')) + ([paths.expressions] if paths.expressions.exists() else [])
    for path in files:
        for query_name, block in split_expression_blocks(path, read_text(path)):
            if path == paths.expressions and query_name in definitions:
                continue
            connector = _sql_database_connector(block, current)
            used = set(connector['parameters'])
            for param_name in definitions:
                if _source_role_for_reference(block, param_name):
                    used.add(param_name)
            for payload in discover_native_sql_payloads(block, current):
                expr = _payload_parameter_expression(payload)
                for param_name in _native_parameter_names(expr, current):
                    used.add(param_name)
            if used:
                query_to_parameters[query_name] = used
                # Prefer actual model tables. Helper expressions are retained only
                # when no model table directly uses the parameter.
                if path != paths.expressions:
                    for parameter_name in used:
                        parameter_to_queries.setdefault(parameter_name, set()).add(query_name)
    for parameter_name in definitions:
        if parameter_to_queries.get(parameter_name):
            continue
        for query_name, used in query_to_parameters.items():
            if parameter_name in used:
                parameter_to_queries.setdefault(parameter_name, set()).add(query_name)
    return query_to_parameters, parameter_to_queries

def _inventory_parameter_context(block: str, current: dict[str, str]):
    names, roles, values = [], [], []
    for name in sorted(current, key=str.casefold):
        if not re.search(_parameter_token(name), block, re.I):
            continue
        names.append(name)
        role = _source_role_for_reference(block, name) or "Referenced value"
        roles.append(f"{name}: {role}")
        value = current.get(name, "")
        values.append(f"{name}={value}" if value != "" else f"{name}=<blank>")
    return {
        "Uses Parameters": "Yes" if names else "No",
        "Parameter Names": "; ".join(names),
        "Parameter Roles": "; ".join(roles),
        "Resolved Parameter Values": "; ".join(values),
    }

def _known_query_directory(paths: Paths) -> dict[str, str]:
    directory = {}
    files = sorted(paths.tables.glob('*.tmdl')) + ([paths.expressions] if paths.expressions.exists() else [])
    for path in files:
        for name, _block in split_expression_blocks(path, read_text(path)):
            directory.setdefault(norm(name), name)
    return directory

def _inventory_query_references(block: str, known_queries: dict[str, str] | None = None):
    try:
        parsed = query_reference_names(block)
    except Exception:
        parsed = set()
    if known_queries is not None:
        names = sorted({known_queries[norm(name)] for name in parsed if norm(name) in known_queries}, key=str.casefold)
    else:
        names = sorted(parsed, key=str.casefold)
    return {
        "References Other Queries": "Yes" if names else "No",
        "Referenced Queries": "; ".join(names),
    }

def _inventory_level(origin: str, query_status: str) -> str:
    # User-facing load status: model table partitions are Active; expressions.tmdl
    # contains helper/disabled queries that are not loaded as model tables.
    return "Disabled" if origin == "expressions.tmdl" or query_status == "Disabled" else "Active"

def _connection_type(
    block: str,
    source_type: str,
    native: bool,
    uses_parameters: bool,
    referenced_queries: str,
    object_type: str,
) -> str:
    """Classify how a query obtains data using evidence in its M/TMDL block."""
    parameterized = uses_parameters
    if native:
        return "Parameterized Native SQL" if parameterized else "Native SQL"
    if re.search(r"(?i)\bSql\.Database\s*\(", block):
        return "Parameterized SQL Navigation" if parameterized else "SQL Navigation"
    if re.search(r"(?i)\bDatabricks\.Catalogs\s*\(", block):
        return "Parameterized Databricks Navigation" if parameterized else "Databricks Navigation"
    if re.search(r"(?i)\bExcel\.Workbook\s*\(", block):
        return "Excel Workbook"
    if re.search(r"(?i)\bCsv\.Document\s*\(", block):
        return "CSV File"
    if re.search(r"(?i)\bJson\.Document\s*\(", block):
        return "JSON"
    if re.search(r"(?i)\bFolder\.(?:Files|Contents)\s*\(", block):
        return "Folder"
    if re.search(r"(?i)\bSharePoint\.(?:Files|Contents|Tables)\s*\(", block):
        return "SharePoint"
    if re.search(r"(?i)\bWeb\.Contents\s*\(", block):
        return "Web"
    if re.search(r"(?i)\bOData\.Feed\s*\(", block):
        return "OData"
    if re.search(r"(?i)\b(?:Table\.FromRows|Table\.FromList)\s*\(|#table\s*\(", block):
        return "Calculated / Local Table"
    if re.search(r"(?s)\([^)]*\)\s*=>", block):
        return "Function"
    if re.search(r"(?i)\b(?:Table\.Invoke|Function\.Invoke)\s*\(", block):
        return "Function Invocation"
    if referenced_queries:
        return "Query Reference"
    if object_type == "Table" and re.search(r"(?mi)^\s*measure\s+", block) and not re.search(r"(?mi)^\s*partition\s+", block):
        return "Measures Table"
    if source_type and source_type not in {"Other", ""}:
        return source_type
    if object_type == "Expression":
        return "Parameter / Helper Query"
    return "Unknown / Unresolved"

def inventory(paths: Paths):
    rows, unpivots = [], []
    _parameter_defs, current_parameters = parameter_definitions(paths)
    known_queries = _known_query_directory(paths)
    files = sorted(paths.tables.glob("*.tmdl")) + ([paths.expressions] if paths.expressions.exists() else [])
    for path in files:
        text_value = read_text(path)
        for query, block in split_expression_blocks(path, text_value):
            resolved_block = resolve_parameters_for_inventory(block, current_parameters)
            source_type, source, server, database, schema, item, connector = _source_inventory_for_query(block, current_parameters)
            parameter_context = _inventory_parameter_context(block, current_parameters)
            # If connector parsing could not populate server/database, use the resolved
            # source-related parameter values already discovered from expressions.tmdl.
            if parameter_context["Uses Parameters"] == "Yes":
                role_pairs = [part.strip() for part in parameter_context.get("Parameter Roles", "").split(";") if part.strip()]
                for role_pair in role_pairs:
                    if ":" not in role_pair:
                        continue
                    parameter_name, role = [part.strip() for part in role_pair.split(":", 1)]
                    value = current_parameters.get(parameter_name, "")
                    if role.casefold() == "server" and not server and value:
                        server = value
                    elif role.casefold() == "database" and not database and value:
                        database = value
            reference_context = _inventory_query_references(block, known_queries)
            discovered_payloads = discover_native_sql_payloads(block, current_parameters)
            native = bool(discovered_payloads)
            if native:
                native_source, native_schema, native_item, native_parse_errors = _promote_unique_native_source(discovered_payloads)
                if native_source:
                    source = native_source
                    schema = native_schema
                    item = native_item
                    source_type = 'SQL'

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
                "Inventory Level": _inventory_level(origin, status),
                "Connection Type": _connection_type(
                    block, source_type, native,
                    parameter_context["Uses Parameters"] == "Yes",
                    reference_context["Referenced Queries"],
                    "Expression" if path == paths.expressions else "Table",
                ),
                "Connection Expression": connector.get("expression") or ("Parameterized M expression" if parameter_context["Uses Parameters"] == "Yes" else "Literal/static expression"),
                **parameter_context, **reference_context,
                "Discovery Method": "M connector and query analysis",
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
                    current_parameters=current_parameters,
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

def _parameter_expression_body(block: str) -> str:
    """Return the expression body after the TMDL ``expression <name> =`` prefix.

    Supports both same-line and multiline TMDL declarations without interpreting
    the expression value. This is intentionally limited to declaration parsing so
    existing connector, lineage, SQL, and migration behavior remains unchanged.
    """
    header = re.match(
        r'(?is)^\s*expression\s+(?:#["](?:[^"]|"")+["]|.+?)\s*=\s*',
        block,
    )
    return block[header.end():].strip() if header else block.strip()

def _parameter_expression_parts(block: str):
    body = _parameter_expression_body(block)
    parts = re.split(r'(?i)\s+meta\s*\[', body, maxsplit=1)
    expression = parts[0].strip()
    metadata = parts[1] if len(parts) > 1 else ''
    return expression, metadata

def _parameter_candidate(name: str, block: str):
    """Discover parameter-like expressions without restricting value type.

    Strong parameter metadata remains authoritative, but connector usage can also
    promote an expression as a parameter candidate. Function definitions are kept
    separate so helper-function arguments are never mistaken for model parameters.
    """
    expression, metadata = _parameter_expression_parts(block)
    is_function = bool(re.match(r'(?is)^\s*\([^)]*\)\s*=>', expression))
    marked = bool(re.search(r'(?i)\bIsParameterQuery\s*=\s*true\b', block))
    parameter_metadata = bool(re.search(
        r'(?i)\b(?:CurrentValue|DefaultValue|List|AllowedValues|SuggestedValues|IsParameterQueryRequired)\s*=',
        metadata,
    ))
    if is_function:
        return None
    if not (marked or parameter_metadata):
        return None
    type_match = re.search(r'(?i)\bType\s*=\s*"([^"]+)"', metadata)
    scalar, inferred = _parameter_scalar(expression)
    return {
        'name': name, 'expression': expression, 'metadata': metadata,
        'marked': marked, 'current': scalar, 'data_type': (type_match.group(1).strip() if type_match else inferred),
    }

def _parameter_definition(name: str, block: str, allow_connector_candidate: bool = False):
    """Parse only an explicitly marked Power Query parameter.

    The applied scalar expression before ``meta`` is authoritative for the current
    value. Metadata defaults and connector usage never promote helper expressions
    or function arguments into report parameters. ``allow_connector_candidate`` is
    retained for call-site compatibility.
    """
    if not re.search(r'(?i)\bIsParameterQuery\s*=\s*true\b', block):
        return None
    expression, metadata = _parameter_expression_parts(block)
    if re.match(r'(?is)^\s*\([^)]*\)\s*=>', expression):
        return None
    current, inferred = _parameter_scalar(expression)
    if current is None:
        return None
    type_match = re.search(r'(?i)\bType\s*=\s*"([^"]+)"', metadata)
    data_type = type_match.group(1).strip() if type_match else inferred
    values = []
    list_match = re.search(
        r'(?is)\b(?:List|AllowedValues|SuggestedValues)\s*=\s*\{(?P<body>[^{}]*)\}',
        metadata,
    )
    if list_match:
        tokens = re.findall(
            r'"(?:[^"]|"")*"|-?\d+(?:\.\d+)?|true|false|null',
            list_match.group('body'), re.I,
        )
        for token in tokens:
            parsed, token_type = _parameter_scalar(token)
            if parsed is not None and parsed not in values:
                values.append(parsed)
            if not data_type:
                data_type = token_type
    if current not in values:
        values.insert(0, current)
    return {'name': name, 'values': values or [current], 'current': current, 'data_type': data_type}
def _parameter_token(name: str) -> str:
    return rf'(?<![A-Za-z0-9_])(?:#"{re.escape(name)}"|{re.escape(name)})(?![A-Za-z0-9_])'

def _connector_argument_identifiers(block: str) -> set[str]:
    names = set()
    for call in _balanced_m_calls(block, ('Sql.Database', 'Databricks.Catalogs')):
        args = split_m_top_level(call['body'])
        for argument in args[:2]:
            name = _m_identifier_name(argument)
            if name:
                names.add(norm(name))
    return names

def parameter_definitions(paths: Paths):
    """Return genuine Power Query parameters and their applied current values."""
    if not paths.expressions.exists():
        return {}, {}
    definitions = {}
    for name, block in split_expression_blocks(paths.expressions, read_text(paths.expressions)):
        definition = _parameter_definition(name, block)
        if definition is not None:
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
    query_to_parameters, parameter_to_queries = _query_parameter_directory(paths, raw_definitions)
    rows, values_by_name, roles = [], {}, {}
    for name in sorted(raw_definitions, key=str.casefold):
        definition = raw_definitions[name]
        references = sorted(parameter_to_queries.get(name, set()), key=str.casefold)
        values_by_name[name] = []
        source_related = bool(references)
        for query_name in references:
            role = _source_role_for_reference(
                next((block for path in sorted(paths.tables.glob('*.tmdl')) + ([paths.expressions] if paths.expressions.exists() else [])
                      for found_name, block in split_expression_blocks(path, read_text(path))
                      if found_name == query_name), ''),
                name,
            ) or 'Connector argument'
            roles[(name, query_name)] = role
        for value in definition['values']:
            parameter = ParameterValue(
                name=name, value=value, is_current=(value == definition['current']),
                data_type=definition['data_type'], references=references,
                source_related=source_related,
            )
            values_by_name[name].append(parameter)
            rows.append({
                'Parameter Name': name,
                'Value': value,
                'Is Current': 'Yes' if parameter.is_current else 'No',
                'Data Type': parameter.data_type,
                'Query Tables Using Parameter': '; '.join(references),
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

def _parameter_connector_role(parameter_name: str, expressions_path: Path) -> str:
    """Infer connector role from usage, never from the parameter name."""
    definition = expressions_path.parent
    roles = set()
    for candidate in sorted(definition.rglob("*.tmdl")):
        role = _source_role_for_reference(read_text(candidate), parameter_name)
        if role in {"Server", "Database", "HTTP Path"}:
            roles.add("Database" if role == "HTTP Path" else role)
    return next(iter(roles)) if len(roles) == 1 else ""

def _parameter_target(parameter_name: str, value: str, sources: list[SourceMapping], role: str = "") -> str:
    """Map a current value by discovered connector role while preserving its name."""
    candidates = set()
    normalized_role = norm(role)
    for source in sources:
        if normalized_role == norm("Server"):
            pairs = ((source.legacy_server, source.target_server),)
        elif normalized_role in {norm("Database"), norm("HTTP Path")} :
            pairs = ((source.legacy_database, source.target_http_path),)
        else:
            pairs = ((source.legacy_schema, source.target_schema), (source.legacy_item, source.target_item))
        for old, new in pairs:
            if old and new and norm(old) == norm(value) and norm(old) != norm(new): candidates.add(new)
    return next(iter(candidates)) if len(candidates) == 1 else ''

def migrate_current_source_parameters(value: str, path: Path, definitions: dict[str,list[ParameterValue]], sources: list[SourceMapping], log, status: str) -> str:
    """Map current connection parameter literals in expressions.tmdl.

    Parameter names and metadata are preserved. For the Databricks connector,
    the legacy SqlServer value maps to Target Server and the legacy Database
    value maps to Target HTTP Path. Table/schema parameters retain the existing
    unique SourceMapping behavior.
    """
    if path.name.casefold() != 'expressions.tmdl':
        return value

    updated = value
    for name, values in definitions.items():
        current = next((v for v in values if v.is_current and v.source_related), None)
        if not current or not current.value:
            continue
        role = _parameter_connector_role(name, path)
        target = _parameter_target(name, current.value, sources, role)
        if not target:
            continue

        blocks = split_expression_blocks(path, updated)
        rebuilt = []
        changed = False
        escaped_old = re.escape(current.value.replace('"', '""'))
        # The previous pattern looked for a line containing only the literal.
        # TMDL stores the value on the declaration line:
        # expression SqlServer = "..." meta [...]
        declaration_pattern = re.compile(
            r'(?mi)^(\s*expression\s+(?:#"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'|[^=\r\n]+?)\s*=\s*)'
            r'"' + escaped_old + r'"'
            r'(?=\s*(?:meta\s*\[|$))'
        )
        new_literal = '"' + target.replace('"', '""') + '"'

        for block_name, block in blocks:
            if block_name != name:
                rebuilt.append(block)
                continue
            new_block, count = declaration_pattern.subn(
                lambda match: match.group(1) + new_literal,
                block,
                count=1,
            )
            if count:
                changed = True
                rebuilt.append(new_block)
                log['M Query Changes'].append({
                    'Input File': path.name,
                    'Query Name': name,
                    'M Step': 'Parameter',
                    'Change Type': 'Current source parameter mapping',
                    'Old Value': current.value,
                    'New Value': target,
                    'Result': status,
                })
            else:
                rebuilt.append(block)
        if changed:
            updated = ''.join(rebuilt)
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

def _m_call_argument_spans(text: str, opening: int):
    """Return top-level argument spans inside an M function call."""
    spans = []
    depth = 1
    bracket_depth = brace_depth = 0
    in_string = False
    start = opening + 1
    index = start
    while index < len(text) and depth:
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
            depth += 1
        elif char == ')':
            depth -= 1
            if depth == 0:
                spans.append((start, index))
                break
        elif char == '[':
            bracket_depth += 1
        elif char == ']':
            bracket_depth = max(0, bracket_depth - 1)
        elif char == '{':
            brace_depth += 1
        elif char == '}':
            brace_depth = max(0, brace_depth - 1)
        elif char == ',' and depth == 1 and bracket_depth == 0 and brace_depth == 0:
            spans.append((start, index))
            start = index + 1
        index += 1
    return spans

def native_sql_regions(text: str):
    """Find actual SQL payloads in Value.NativeQuery and Sql.Database [Query=]."""
    regions = []

    # Value.NativeQuery SQL is the second top-level argument. The first argument
    # may itself contain Databricks.Catalogs strings, so selecting the first quote
    # after the function name is incorrect.
    for marker in re.finditer(r"(?i)Value\.NativeQuery\s*\(", text):
        opening = text.find("(", marker.start(), marker.end())
        spans = _m_call_argument_spans(text, opening)
        if len(spans) < 2:
            continue
        arg_start, arg_end = spans[1]
        argument = text[arg_start:arg_end]
        leading = len(argument) - len(argument.lstrip())
        quote = arg_start + leading
        if quote >= arg_end or text[quote] != '"':
            continue
        end = _scan_m_string(text, quote)
        if end > quote and end <= arg_end:
            regions.append((quote + 1, end, text[quote + 1:end]))

    # Sql.Database Query option remains a named record field.
    for marker in re.finditer(r"(?i)\bQuery\s*=\s*", text):
        quote = text.find('"', marker.end())
        if quote < 0:
            continue
        end = _scan_m_string(text, quote)
        if end > quote:
            regions.append((quote + 1, end, text[quote + 1:end]))

    # Deduplicate exact spans while preserving source order.
    unique = {(start, end): payload for start, end, payload in regions}
    return [(start, end, unique[(start, end)]) for start, end in sorted(unique)]

def _unquote_identifier(value: str) -> str:
    """Remove supported TMDL/SQL identifier quoting."""
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

def infer_sql_dialects(source_maps: list[SourceMapping], sql_text: str = "") -> tuple[str | None, str]:
    """Detect the input dialect; migration direction fixes only the target."""
    source_dialect, _detection = detect_sql_dialect(sql_text)
    return source_dialect, TARGET_SQL_DIALECT

SQL_DIALECT_CANDIDATES = (
    "tsql", "databricks", "spark", "snowflake", "bigquery",
    "postgres", "mysql", "oracle", "redshift", "duckdb",
)

SQL_DIALECT_HINTS = {
    "tsql": ((r"(?i)\bTOP\s*(?:\(|\d)", 5), (r"(?i)\b(?:GETDATE|ISNULL|DATEADD|DATEDIFF|CONVERT|TRY_CONVERT)\s*\(", 3), (r"(?i)\b(?:NVARCHAR|DATETIME2|UNIQUEIDENTIFIER)\b", 3), (r"\[[^\]]+\]", 1)),
    "databricks": ((r"(?i)\b(?:QUALIFY|ILIKE|RLIKE)\b", 3), (r"(?i)\b(?:TIMESTAMPDIFF|DATE_TRUNC|TRY_CAST)\s*\(", 2)),
    "spark": ((r"(?i)\b(?:LATERAL\s+VIEW|RLIKE)\b", 4), (r"(?i)\b(?:COLLECT_LIST|POSEXPLODE|FROM_UNIXTIME)\s*\(", 3)),
    "snowflake": ((r"(?i)\bQUALIFY\b", 3), (r"(?i)\b(?:IFF|ZEROIFNULL|TRY_TO_DATE|TO_VARIANT)\s*\(", 4)),
    "bigquery": ((r"(?i)\bQUALIFY\b", 2), (r"(?i)\b(?:SAFE_CAST|UNNEST|STRUCT|ARRAY_AGG)\s*\(?", 4), (r"`[^`]+\.[^`]+`", 3)),
    "postgres": ((r"(?i)::\s*[A-Za-z]", 4), (r"(?i)\b(?:ILIKE|STRING_AGG|GENERATE_SERIES)\b", 3), (r"(?i)\bRETURNING\b", 2)),
    "mysql": ((r"(?i)\b(?:IFNULL|DATE_FORMAT|STR_TO_DATE|GROUP_CONCAT)\s*\(", 4), (r"(?i)\bLIMIT\s+\d+\s*,\s*\d+", 4)),
    "oracle": ((r"(?i)\b(?:NVL|DECODE|TO_CHAR|TO_DATE)\s*\(", 4), (r"(?i)\bFROM\s+DUAL\b", 5), (r"(?i)\bROWNUM\b", 4)),
    "redshift": ((r"(?i)\b(?:DISTKEY|SORTKEY|ENCODE)\b", 5),),
    "duckdb": ((r"(?i)\b(?:STRUCT_PACK|READ_PARQUET)\s*\(", 4),),
}

def detect_sql_dialect(sql_text: str) -> tuple[str | None, str]:
    """Detect source syntax. Portable SQL uses SQLGlot's generic parser."""
    text = clean(sql_text)
    if not text:
        raise ValueError("Cannot detect SQL dialect from an empty SQL payload")
    parsed = {}
    for dialect in SQL_DIALECT_CANDIDATES:
        try:
            parsed[dialect] = sqlglot.parse_one(text, read=dialect)
        except (ParseError, UnsupportedError, ValueError):
            pass
    if not parsed:
        try:
            sqlglot.parse_one(text)
            return None, "generic"
        except (ParseError, UnsupportedError, ValueError) as exc:
            raise ParseError(f"SQLGlot could not parse SQL in any supported dialect: {exc}") from exc
    scores = {dialect: 0 for dialect in parsed}
    for dialect, rules in SQL_DIALECT_HINTS.items():
        if dialect in scores:
            for pattern, weight in rules:
                if re.search(pattern, text):
                    scores[dialect] += weight
    best = max(scores.values())
    winners = [d for d, score in scores.items() if score == best]
    if best <= 0 or len(winners) != 1:
        # Generic detection is only a probe. A dialect-specific statement may be
        # valid even when SQLGlot's generic parser rejects it (for example,
        # Databricks backtick-qualified identifiers). Never let that probe abort
        # inventory. Prefer an unambiguous dialect-specific parse when available.
        try:
            sqlglot.parse_one(text)
            return None, "generic"
        except (ParseError, UnsupportedError, ValueError):
            preferred = [d for d in ("databricks", "spark", "tsql") if d in parsed]
            if len(preferred) == 1:
                return preferred[0], "dialect-fallback"
            if "`" in text:
                for dialect in ("databricks", "spark"):
                    if dialect in parsed:
                        return dialect, "dialect-fallback"
            if "tsql" in parsed:
                return "tsql", "dialect-fallback"
            if len(parsed) == 1:
                return next(iter(parsed)), "dialect-fallback"
            # Deterministic final fallback to a dialect that already parsed
            # successfully. This avoids re-running a known-failing generic parse.
            return next(d for d in SQL_DIALECT_CANDIDATES if d in parsed), "dialect-fallback"
    return winners[0], "syntax"

def parse_detected_sql(sql_text: str):
    dialect, method = detect_sql_dialect(sql_text)
    tree = sqlglot.parse_one(sql_text, read=dialect) if dialect else sqlglot.parse_one(sql_text)
    return tree, dialect, method

def parse_detected_statements(sql_text: str):
    dialect, method = detect_sql_dialect(sql_text)
    trees = sqlglot.parse(sql_text, read=dialect) if dialect else sqlglot.parse(sql_text)
    return trees, dialect, method

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
        legacy_object = ".".join(filter(None, [source_map.legacy_schema, source_map.legacy_item]))
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

def _ambiguous_unqualified_column_candidates(tables, column_name, mappings):
    """Return distinct mapped owners for an unqualified SQL column.

    Rewriting an unqualified column is safe only when exactly one physical source
    owns an applicable mapping. Multiple owners are a blocking ambiguity because
    choosing one can silently change join, filter, grouping, or projection semantics.
    """
    candidates = []
    seen = set()
    for alias, table in tables.items():
        mapping = _one_column_mapping(table, column_name, mappings)
        if mapping is None:
            continue
        identity = ".".join(
            part for part in (clean(table.catalog), clean(table.db), clean(table.name)) if part
        ) or clean(table.name)
        key = (
            norm(alias), _normalize_object_name(identity),
            norm(mapping.legacy_column), norm(mapping.target_column),
        )
        if key in seen:
            continue
        seen.add(key)
        candidates.append((alias, table, mapping, identity))
    return candidates

def _format_unqualified_column_ambiguity(column_name, candidates):
    owners = []
    for alias, _table, mapping, identity in candidates:
        qualifier = clean(alias) or identity or '<unaliased>'
        owners.append(
            f"{qualifier} ({identity or mapping.legacy_table}: "
            f"{mapping.legacy_column} -> {mapping.target_column})"
        )
    return (
        "AMBIGUOUS_UNQUALIFIED_SQL_COLUMN: "
        f"{column_name} matches multiple mapped physical sources: "
        + "; ".join(sorted(owners, key=str.casefold))
        + ". Qualify the source column in native SQL or remove the conflicting mapping."
    )

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
    resolved = candidates[0] if len(target_names) == 1 else None
    return apply_physical_column_compatibility_override(resolved) if resolved else None

def apply_physical_column_compatibility_override(
    mapping: MappingRow,
) -> MappingRow:
    """Apply an exact, opt-in physical-column override without changing lineage.

    The lookup requires the legacy table, legacy column, and resolved target
    table to match. If no exact rule exists, the original MappingRow object is
    returned unchanged. Semantic names, aliases, type-retention decisions, and
    all downstream migration logic therefore continue through the existing path.
    """
    key = (
        _normalize_object_name(mapping.legacy_table),
        norm(mapping.legacy_column),
        _normalize_object_name(mapping.target_table),
    )
    override = PHYSICAL_COLUMN_COMPATIBILITY_OVERRIDES.get(key)
    if not override:
        return mapping
    target_column = clean(override.get("target_column"))
    if not target_column or norm(target_column) == norm(mapping.target_column):
        return mapping
    return MappingRow(
        excel_row=mapping.excel_row,
        legacy_table=mapping.legacy_table,
        target_table=mapping.target_table,
        legacy_column=mapping.legacy_column,
        legacy_type=mapping.legacy_type,
        target_column=target_column,
        target_type=clean(override.get("target_type")) or mapping.target_type,
        selection=mapping.selection,
        notes=mapping.notes,
        risk_summary=mapping.risk_summary,
        system_recommendation=mapping.system_recommendation,
        dependency_notes=mapping.dependency_notes,
        dependencies=dict(mapping.dependencies),
    )

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
                    value = exp.cast(
                        value,
                        sqlglot_cast_type(m.legacy_type, TARGET_SQL_DIALECT),
                        dialect=TARGET_SQL_DIALECT,
                    )
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

def _flatten_add_expression(node: exp.Expression) -> list[exp.Expression]:
    if isinstance(node, exp.Add):
        return _flatten_add_expression(node.this) + _flatten_add_expression(node.expression)
    return [node]

def _text_like_expression(node: exp.Expression) -> bool:
    if isinstance(node, exp.Literal) and node.is_string:
        return True
    if isinstance(node, (exp.Cast, exp.TryCast, exp.Convert)):
        target = node.args.get("to") or node.args.get("this")
        target_sql = target.sql().casefold() if isinstance(target, exp.Expression) else ""
        return any(token in target_sql for token in ("char", "text", "string"))
    return isinstance(node, exp.Concat)

def normalize_databricks_ast(tree: exp.Expression):
    """Apply target-semantic rewrites that SQLGlot does not always infer from T-SQL."""
    audit = []

    # T-SQL uses + for string concatenation. Databricks uses CONCAT/||.
    for add in list(tree.find_all(exp.Add)):
        if isinstance(add.parent, exp.Add):
            continue
        parts = _flatten_add_expression(add)
        if not any(_text_like_expression(part) for part in parts):
            continue
        replacement = exp.Concat(expressions=[part.copy() for part in parts])
        add.replace(replacement)
        audit.append({
            "kind": "function", "old_value": "T-SQL string + operator",
            "new_value": "CONCAT(...)" , "occurrences": 1,
            "source_dialect": SOURCE_SQL_DIALECT,
            "target_dialect": TARGET_SQL_DIALECT,
        })

    # SQLGlot can retain the three-argument T-SQL DATEDIFF shape. Databricks
    # expects DATEDIFF(end, start) for days or TIMESTAMPDIFF(unit, start, end).
    for date_diff in list(tree.find_all(exp.DateDiff)):
        unit = date_diff.args.get("unit")
        end_value = date_diff.this
        start_value = date_diff.expression
        unit_name = clean(unit.name if isinstance(unit, exp.Expression) else unit).upper()
        if unit_name == "DAY":
            replacement = exp.Anonymous(
                this="DATEDIFF",
                expressions=[end_value.copy(), start_value.copy()],
            )
            new_value = "DATEDIFF(end, start)"
        else:
            replacement = exp.Anonymous(
                this="TIMESTAMPDIFF",
                expressions=[
                    exp.Var(this=unit_name or "DAY"),
                    start_value.copy(), end_value.copy(),
                ],
            )
            new_value = "TIMESTAMPDIFF(unit, start, end)"
        date_diff.replace(replacement)
        audit.append({
            "kind": "function", "old_value": "T-SQL DATEDIFF(unit, start, end)",
            "new_value": new_value, "occurrences": 1,
            "source_dialect": SOURCE_SQL_DIALECT,
            "target_dialect": TARGET_SQL_DIALECT,
        })
    return tree, audit

def _source_mapping_for_sql_table(
    table: exp.Table, source_maps: list[SourceMapping],
    legacy_server: str = "", legacy_database: str = "",
):
    """Resolve an exact Data Source Mapping for a native SQL physical table.

    This is a fallback for native SQL runs with no Table/Column Mapping rows. It
    requires an exact legacy schema/item match and, when available, the owning
    Sql.Database server/database. Existing target objects are never remapped.
    """
    actual_catalog, actual_schema, actual_item = table_parts(table)
    if not actual_schema or not actual_item:
        return None
    candidates = []
    for source in source_maps:
        if legacy_server and source.legacy_server and norm(source.legacy_server) != norm(legacy_server):
            continue
        if legacy_database and source.legacy_database and norm(source.legacy_database) != norm(legacy_database):
            continue
        if not source.legacy_schema or not source.legacy_item:
            continue
        if norm(source.legacy_schema) != norm(actual_schema):
            continue
        if norm(source.legacy_item) != norm(actual_item):
            continue
        target_identity = (source.target_catalog, source.target_schema, source.target_item)
        if all(target_identity) and not (
            norm(actual_catalog) == norm(source.target_catalog)
            and norm(actual_schema) == norm(source.target_schema)
            and norm(actual_item) == norm(source.target_item)
        ):
            candidates.append(source)
    target_keys = {
        (norm(x.target_catalog), norm(x.target_schema), norm(x.target_item))
        for x in candidates
    }
    return candidates[0] if len(target_keys) == 1 else None

def _source_mapping_target_table(source: SourceMapping) -> exp.Table:
    """Emit schema.item because the Databricks connector already selects catalog."""
    return exp.Table(
        this=exp.to_identifier(source.target_item),
        db=exp.to_identifier(source.target_schema) if source.target_schema else None,
        catalog=None,
    )

def _remaining_legacy_source_map_tables(
    tree: exp.Expression, source_maps: list[SourceMapping],
    legacy_server: str = "", legacy_database: str = "",
):
    problems = []
    for table in tree.find_all(exp.Table):
        source = _source_mapping_for_sql_table(
            table, source_maps, legacy_server, legacy_database
        )
        if source:
            current = ".".join(part for part in table_parts(table) if part)
            target = ".".join(
                part for part in (source.target_catalog, source.target_schema, source.target_item) if part
            )
            problems.append(
                f"Mapped legacy native SQL table remains: {current}; expected {target}"
            )
    return sorted(set(problems))

def _drop_intermediate_order_by(query: exp.Expression) -> exp.Expression:
    """Remove ORDER BY from intermediate CTE queries when it has no row-limiting role.

    T-SQL commonly allows ORDER BY in SELECT INTO staging statements. Databricks
    CTEs do not need that ordering, and Power BI must not rely on intermediate
    row order. The final query's ORDER BY is preserved.
    """
    candidate = query.copy()
    root = candidate.this if isinstance(candidate, exp.Subquery) else candidate
    if isinstance(root, exp.Select) and root.args.get("order") is not None:
        if root.args.get("limit") is None and root.args.get("offset") is None:
            root.set("order", None)
    return candidate

def _temporary_table_name(table: exp.Table | None) -> str:
    if not isinstance(table, exp.Table):
        return ""
    return clean(table.name).lstrip("#")

DATABRICKS_RESERVED_CTE_NAMES = {
    "all", "and", "as", "by", "case", "create", "cross", "delete",
    "distinct", "drop", "else", "end", "except", "exists", "false",
    "from", "full", "group", "having", "in", "inner", "insert",
    "intersect", "into", "is", "join", "left", "like", "limit", "not",
    "null", "on", "or", "order", "outer", "right", "select", "table",
    "then", "true", "union", "update", "values", "when", "where", "with",
}

def databricks_safe_cte_name(temp_name: str, used_names: set[str]) -> str:
    """Preserve the temp-table name as the CTE name whenever Databricks allows it.

    The old implementation prefixed every generated CTE with ``pbi_``. That was
    safe but over-corrected valid names and made the translated SQL harder to
    compare with the source. This allocator now changes a name only when required:
    invalid characters are normalized, a leading digit is protected, reserved
    words are prefixed, and collisions receive a deterministic numeric suffix.
    """
    original = clean(temp_name).lstrip("#")
    base = re.sub(r"[^A-Za-z0-9_]", "_", original)
    base = re.sub(r"_+", "_", base).strip("_") or "temp"
    if base[0].isdigit():
        base = "temp_" + base
    if norm(base) in DATABRICKS_RESERVED_CTE_NAMES:
        base = "temp_" + base
    candidate = base
    root = candidate
    ordinal = 2
    while norm(candidate) in used_names:
        candidate = f"{root}_{ordinal}"
        ordinal += 1
    used_names.add(norm(candidate))
    return candidate

def _validate_generated_cte_aliases(ctes: list[exp.CTE]):
    """Fail closed for missing, reserved, or duplicate generated CTE aliases."""
    errors = []
    aliases = []
    for cte in ctes:
        alias = clean(cte.alias_or_name)
        if not alias:
            errors.append("DATABRICKS_CTE_ALIAS_MISSING: generated CTE has no alias")
            continue
        aliases.append(norm(alias))
        if norm(alias) in DATABRICKS_RESERVED_CTE_NAMES:
            errors.append(f"DATABRICKS_CTE_ALIAS_RESERVED: {alias}")
    if len(aliases) != len(set(aliases)):
        errors.append("DATABRICKS_CTE_ALIAS_COLLISION: duplicate generated CTE aliases")
    return sorted(set(errors))

def _is_sql_separator_statement(statement: exp.Expression | None) -> bool:
    """Return True only for empty/semicolon parser artifacts, never executable SQL."""
    if statement is None:
        return True
    if clean(getattr(statement, "key", "")).casefold() == "semicolon":
        return True
    try:
        rendered = statement.sql(dialect=TARGET_SQL_DIALECT).strip()
    except Exception:
        rendered = ""
    return rendered in {"", ";"}

def _normalize_cte_temp_references(expression: exp.Expression, cte_names: dict[str, str]):
    """Rewrite #temp references in both table nodes and column/star qualifiers."""
    for table in expression.find_all(exp.Table):
        if clean(table.db) or clean(table.catalog):
            continue
        key = norm(clean(table.name).lstrip("#"))
        if key in cte_names:
            table.set("this", exp.to_identifier(cte_names[key]))
    for column in expression.find_all(exp.Column):
        qualifier = clean(column.table)
        if not qualifier:
            continue
        key = norm(qualifier.lstrip("#"))
        if key in cte_names:
            column.set("table", exp.to_identifier(cte_names[key]))
    return expression

def _remaining_temp_identifiers(expression: exp.Expression):
    remaining = set()
    for table in expression.find_all(exp.Table):
        if clean(table.name).startswith("#"):
            remaining.add(clean(table.name))
    for column in expression.find_all(exp.Column):
        if clean(column.table).startswith("#"):
            remaining.add(clean(column.table))
    return sorted(remaining, key=str.casefold)

def collapse_databricks_temp_table_batch(sql_text: str):
    """Convert a generated temp-table batch into one Databricks WITH query."""
    try:
        statements = sqlglot.parse(sql_text, read=TARGET_SQL_DIALECT)
    except (ParseError, UnsupportedError, ValueError) as exc:
        return sql_text, [], [f"Databricks temp-table batch parse failed: {sanitize_diagnostic_text(exc)}"]

    # Targeted CTE alias pass: allocate all names before rewriting any references.
    used_names = {
        norm(cte.alias_or_name)
        for statement in statements
        if isinstance(statement, exp.Expression)
        for cte in statement.find_all(exp.CTE)
        if clean(cte.alias_or_name)
    }
    temp_to_cte = {}
    for statement in statements:
        if not isinstance(statement, exp.Create):
            continue
        rendered_statement = statement.sql(dialect=TARGET_SQL_DIALECT).upper()
        temporary = bool(statement.args.get("temporary")) or "CREATE TEMP" in rendered_statement
        target = statement.this if isinstance(statement.this, exp.Table) else None
        temp_name = _temporary_table_name(target)
        if temporary and temp_name:
            key = norm(temp_name.lstrip("#"))
            if key not in temp_to_cte:
                temp_to_cte[key] = databricks_safe_cte_name(temp_name, used_names)

    ctes = []
    final_query = None
    ignored_drops = ignored_separators = 0
    for statement in statements:
        if _is_sql_separator_statement(statement):
            ignored_separators += 1
            continue
        if isinstance(statement, exp.Drop):
            ignored_drops += 1
            continue
        if isinstance(statement, exp.Create):
            rendered_statement = statement.sql(dialect=TARGET_SQL_DIALECT).upper()
            temporary = bool(statement.args.get("temporary")) or "CREATE TEMP" in rendered_statement
            target = statement.this if isinstance(statement.this, exp.Table) else None
            body = statement.expression
            name = _temporary_table_name(target)
            key = norm(name.lstrip("#"))
            if temporary and name and key in temp_to_cte and isinstance(body, exp.Expression):
                ctes.append(exp.CTE(
                    this=_drop_intermediate_order_by(body),
                    alias=exp.TableAlias(this=exp.to_identifier(temp_to_cte[key])),
                ))
                continue
        if isinstance(statement, (exp.Select, exp.Union, exp.Subquery)):
            final_query = statement.copy()
            continue
        return sql_text, [], [
            "DATABRICKS_NATIVE_BATCH_UNSUPPORTED: native SQL contains a non-query "
            f"statement that cannot be represented as a CTE: {statement.key}"
        ]

    if not ctes:
        return sql_text, [], []
    if final_query is None:
        return sql_text, [], ["DATABRICKS_NATIVE_BATCH_NO_FINAL_QUERY: temporary-table batch has no final SELECT"]

    alias_errors = _validate_generated_cte_aliases(ctes)
    if alias_errors:
        return sql_text, [], alias_errors

    # Use the same authoritative map for every active table and column qualifier.
    for cte in ctes:
        _normalize_cte_temp_references(cte.this, temp_to_cte)
    _normalize_cte_temp_references(final_query, temp_to_cte)

    remaining = set(_remaining_temp_identifiers(final_query))
    for cte in ctes:
        remaining.update(_remaining_temp_identifiers(cte.this))
    if remaining:
        return sql_text, [], [
            "DATABRICKS_TEMP_IDENTIFIER_REMAINS: " + ", ".join(sorted(remaining, key=str.casefold))
        ]

    final_query.set("with_", exp.With(expressions=ctes, recursive=False))
    rendered = final_query.sql(dialect=TARGET_SQL_DIALECT, pretty=True)
    audit = [{
        "kind": "function",
        "old_value": "T-SQL temp-table / SELECT INTO batch",
        "new_value": f"Single Databricks WITH query ({len(ctes)} CTEs)",
        "occurrences": len(ctes),
        "source_dialect": "tsql",
        "target_dialect": TARGET_SQL_DIALECT,
        "message": (
            f"Removed {ignored_drops} DROP statement(s); ignored {ignored_separators} "
            "empty/semicolon separator node(s); generated collision-safe preservation-first CTE aliases; "
            "preserved final result query"
        ),
    }]
    return rendered, audit, []

def _balanced_sql_function_calls(sql_text: str, function_name: str):
    """Return balanced function-call spans while respecting SQL strings."""
    pattern = re.compile(r"(?i)\b" + re.escape(function_name) + r"\s*\(")
    calls = []
    for match in pattern.finditer(sql_text):
        opening = sql_text.find("(", match.start(), match.end())
        depth, index = 1, opening + 1
        quote = ""
        while index < len(sql_text) and depth:
            char = sql_text[index]
            if quote:
                if char == quote:
                    if index + 1 < len(sql_text) and sql_text[index + 1] == quote:
                        index += 2
                        continue
                    quote = ""
                index += 1
                continue
            if char in {"'", '"', '`'}:
                quote = char
            elif char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
            index += 1
        if depth == 0:
            calls.append((match.start(), index, sql_text[opening + 1:index - 1]))
    return calls

def _sql_comment_spans(sql_text: str):
    spans = []
    index = 0
    quote = ""
    while index < len(sql_text):
        char = sql_text[index]
        nxt = sql_text[index + 1] if index + 1 < len(sql_text) else ""
        if quote:
            if char == quote:
                if index + 1 < len(sql_text) and sql_text[index + 1] == quote:
                    index += 2
                    continue
                quote = ""
            index += 1
            continue
        if char in {"'", '"', '`'}:
            quote = char
            index += 1
            continue
        if char == "-" and nxt == "-":
            end = sql_text.find("\n", index + 2)
            end = len(sql_text) if end < 0 else end
            spans.append((index, end))
            index = end
            continue
        if char == "/" and nxt == "*":
            end = sql_text.find("*/", index + 2)
            end = len(sql_text) if end < 0 else end + 2
            spans.append((index, end))
            index = end
            continue
        index += 1
    return spans

def normalize_databricks_sql_text(sql_text: str):
    """Normalize final active Databricks SQL without rewriting commented-out SQL."""
    output = sql_text
    audit = []
    protected = {}
    for ordinal, (start, end) in enumerate(reversed(_sql_comment_spans(output))):
        original = output[start:end]
        token = f"__PBI_SQL_COMMENT_{ordinal}_{sha(original)[:12]}__"
        protected[token] = original
        output = output[:start] + token + output[end:]

    replacements = []
    for function_name in ("DATEADD", "DATE_ADD"):
        for start, end, body in _balanced_sql_function_calls(output, function_name):
            args = split_m_top_level(body)
            if len(args) != 3:
                continue
            unit = clean(args[0]).upper()
            if unit not in {"YEAR", "QUARTER", "MONTH", "WEEK", "DAY", "HOUR", "MINUTE", "SECOND"}:
                continue
            replacements.append((start, end, f"TIMESTAMPADD({unit}, {args[1].strip()}, {args[2].strip()})"))
    for start, end, replacement in sorted(replacements, reverse=True):
        output = output[:start] + replacement + output[end:]
    if replacements:
        audit.append({
            "kind": "function", "old_value": "T-SQL DATEADD / invalid DATE_ADD(unit, value, date)",
            "new_value": "Databricks TIMESTAMPADD", "occurrences": len(replacements),
            "source_dialect": "tsql", "target_dialect": TARGET_SQL_DIALECT,
        })
    for token, original in protected.items():
        if output.count(token) != 1:
            raise ValueError("SQL comment protection token was altered during normalization")
        output = output.replace(token, original, 1)
    return output, audit

def validate_final_databricks_native_sql(sql_text: str):
    errors = []
    try:
        statements = [s for s in sqlglot.parse(sql_text, read=TARGET_SQL_DIALECT) if not _is_sql_separator_statement(s)]
    except (ParseError, UnsupportedError, ValueError) as exc:
        return [f"Generated Databricks SQL is invalid: {sanitize_diagnostic_text(exc)}"]
    if len(statements) != 1:
        errors.append(f"DATABRICKS_ADBC_REQUIRES_SINGLE_STATEMENT: generated {len(statements)} executable statements")
        return errors
    remaining = _remaining_temp_identifiers(statements[0])
    if remaining:
        errors.append("DATABRICKS_TEMP_IDENTIFIER_REMAINS: " + ", ".join(remaining))
    if re.search(r"(?i)\bDATE_ADD\s*\(\s*(?:YEAR|QUARTER|MONTH|WEEK|HOUR|MINUTE|SECOND)\s*,", sql_text):
        errors.append("DATABRICKS_INVALID_DATE_ADD_SIGNATURE")
    if re.search(r"(?i)\b(?:DROP|CREATE\s+(?:OR\s+REPLACE\s+)?(?:TEMPORARY\s+)?TABLE)\b", sql_text):
        errors.append("DATABRICKS_ADBC_NON_QUERY_STATEMENT_REMAINS")
    return errors

def rewrite_sql_ast(sql_text, mappings, source_maps, legacy_server="", legacy_database=""):
    source_dialect, detection_method = detect_sql_dialect(sql_text)
    normalized_sql, audit = normalize_synapse_sql_functions(sql_text) if source_dialect in {None, "tsql"} else (sql_text, [])
    errors = []
    try:
        tree = sqlglot.parse_one(normalized_sql, read=source_dialect) if source_dialect else sqlglot.parse_one(normalized_sql)
    except ParseError as exc:
        return sql_text, audit, [f"SQLGlot could not parse detected source SQL: {sanitize_diagnostic_text(exc)}"]
    stars = list(tree.find_all(exp.Star))
    if stars:
        audit.append({"kind":"star", "occurrences":len(stars), "source_dialect":source_dialect or "generic", "target_dialect":"databricks", "message":"SELECT * / alias.* preserved; mapped target table supplies columns at runtime"})
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

            if mapping is None and qualifier and qualifier in child_scopes:
                child_renames = scope_output_renames.get(id(child_scopes[qualifier]), {})
                if norm(column.name) in child_renames:
                    target_col_name, mapping = child_renames[norm(column.name)]
                    from_child_scope = True

            if mapping is None and not qualifier:
                candidates = _ambiguous_unqualified_column_candidates(
                    tables, column.name, mappings
                )
                if len(candidates) == 1:
                    alias, table, mapping, _identity = candidates[0]
                    target_col_name = mapping.target_column
                    if alias:
                        column.set("table", exp.to_identifier(alias))
                elif len(candidates) > 1:
                    errors.append(
                        _format_unqualified_column_ambiguity(column.name, candidates)
                    )
                    continue
                elif child_scopes:
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
            audit.append({"kind":"column", "legacy_table":m.legacy_table, "target_table":m.target_table, "legacy_column":m.legacy_column, "target_column":m.target_column, "aliases":", ".join(sorted(a for a in item["aliases"] if a)), "occurrences":item["occurrences"], "projection_aliases_added":alias_count, "casts_added":cast_count, "source_dialect":source_dialect or "generic", "target_dialect":TARGET_SQL_DIALECT})
    for table in list(tree.find_all(exp.Table)):
        conflicts = _table_mapping_conflicts(table, mappings)
        if conflicts:
            errors.append(
                f"Conflicting target table mappings for {table.sql(dialect=source_dialect or TARGET_SQL_DIALECT)}: {conflicts}"
            )
            continue
        mapping = _one_table_mapping(table, mappings)
        source_mapping = None if mapping is not None else _source_mapping_for_sql_table(
            table, source_maps, legacy_server, legacy_database
        )
        if mapping is None and source_mapping is None:
            continue
        alias = table.args.get("alias")
        implicit_alias_required = _implicit_legacy_table_alias_required(tree, table)
        legacy_catalog, legacy_schema, legacy_item = table_parts(table)
        legacy_object = ".".join(
            part for part in (legacy_catalog, legacy_schema, legacy_item) if part
        )
        legacy_object_name = clean(table.name)
        if mapping is not None:
            target_table = target_table_expression(mapping, source_maps)
            target_object = mapping.target_table
            mapping_source = "Table and Column Mapping"
        else:
            target_table = _source_mapping_target_table(source_mapping)
            target_object = ".".join(
                part for part in (
                    source_mapping.target_catalog, source_mapping.target_schema,
                    source_mapping.target_item,
                ) if part
            )
            mapping_source = "Data Source Mapping"
        if alias:
            target_table.set("alias", alias.copy())
        elif implicit_alias_required and legacy_object_name:
            target_table.set("alias", exp.TableAlias(this=exp.to_identifier(legacy_object_name)))
        table.replace(target_table)
        audit.append({
            "kind":"table",
            "legacy_table": mapping.legacy_table if mapping is not None else legacy_object,
            "target_table": target_object, "alias":target_table.alias_or_name,
            "occurrences":1, "source_dialect":source_dialect or "generic",
            "target_dialect":"databricks", "mapping_source": mapping_source,
            "alias_preservation": "Existing explicit alias preserved" if alias else "Implicit object-name qualifier preserved explicitly" if implicit_alias_required else "No alias required",
        })
    tree, databricks_audit = normalize_databricks_ast(tree)
    audit.extend(databricks_audit)
    errors = sorted(set(errors))
    try:
        rewritten = tree.sql(dialect="databricks", pretty=True)
        rewritten, batch_audit, batch_errors = collapse_databricks_temp_table_batch(rewritten)
        audit.extend(batch_audit)
        errors.extend(batch_errors)
        rewritten, text_audit = normalize_databricks_sql_text(rewritten)
        audit.extend(text_audit)
        errors.extend(validate_final_databricks_native_sql(rewritten))
        reparsed = sqlglot.parse_one(rewritten, read="databricks")
    except (ParseError, UnsupportedError, ValueError) as exc:
        return sql_text, audit, [f"Generated Databricks SQL is invalid: {sanitize_diagnostic_text(exc)}"]
    if re.search(r"(?i)\bCURDATE\s*\(", rewritten) or re.search(r"(?is)\{\s*fn\s+CURDATE", rewritten): errors.append("CURDATE remained after Synapse function normalization")
    errors.extend(_remaining_legacy_physical_refs(reparsed, mappings))
    errors.extend(_remaining_legacy_source_map_tables(
        reparsed, source_maps, legacy_server, legacy_database
    ))
    return rewritten, audit, errors

def replace_native_sql_payloads(
    value: str, mappings: list[MappingRow], source_maps: list[SourceMapping],
    current_parameters: dict[str, str] | None = None,
) -> tuple[str, list[dict], list[str]]:
    return _materialize_rewritten_native_bindings(
        value, mappings, source_maps, current_parameters or {}, databricks=True
    )

def _m_query_reference_tokens(value: str) -> dict[str, str]:
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
    # Quoted M identifiers, for example #"Customer Persona Source".
    candidates.extend(re.finditer(r'#"((?:[^"]|"")+)"', region))
    # Plain identifiers are protected only when they are the entire right-hand side
    # of a let binding. This captures Source = ExistingQuery without masking fields.
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
    """Update source-facing M references while preserving RenameColumns outputs.

    In Table.RenameColumns(table, {{"input", "output"}}), only the input name
    can represent an existing source column. The output name is a query-created
    schema contract and must not be changed by a coincidentally named physical
    column mapping (for example, Value -> Value_Count).
    """
    region, start, end = extract_m_region(value)
    if start < 0 or not legacy or not target or legacy == target:
        return value, 0

    # Protect external Power Query references first. Physical source-column
    # rewrites must never turn an existing query dependency into a target
    # schema/item-shaped M identifier.
    masked_value, query_references = _m_query_reference_tokens(value)
    region, start, end = extract_m_region(masked_value)
    # Temporarily mask only the destination names in static RenameColumns pairs.
    # Existing parsing/replacement behavior remains unchanged everywhere else.
    protected = {}
    masked_region = region
    pair_pattern = re.compile(
        r'(\{\s*"(?:[^"\r\n]|"")*"\s*,\s*)'
        r'("(?:[^"\r\n]|"")*")'
        r'(\s*\})',
        re.DOTALL,
    )

    # Protect Table.Combine branch-list arguments (e.g. forward-adapted Table.RenameColumns)
    # so downstream binding propagation never rewrites an unmigrated branch's legacy input names.
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

def _parse_tmdl_relationship_endpoint(value: str):
    """Parse TMDL Table.Column endpoints with either identifier independently quoted."""
    value = clean(value)
    if not value:
        return "", ""
    match = re.fullmatch(
        r"\s*(?:'(?P<qt>(?:[^']|'')+)'|(?P<pt>[^.\s][^.]*?))\s*\.\s*"
        r"(?:'(?P<qc>(?:[^']|'')+)'|(?P<pc>\S(?:.*?\S)?))\s*",
        value,
    )
    if not match:
        return "", ""
    table_name = match.group("qt") if match.group("qt") is not None else match.group("pt")
    column_name = match.group("qc") if match.group("qc") is not None else match.group("pc")
    if match.group("qt") is not None:
        table_name = table_name.replace("''", "'")
    if match.group("qc") is not None:
        column_name = column_name.replace("''", "'")
    return clean(table_name), clean(column_name)

def validate_relationships(paths, proposed_table_files: dict, proposed_relationship_files: dict | None = None):
    """Validate relationship endpoints against the complete semantic model.

    The declaration parser accepts either ``table <name>`` or ``table '<name>'``
    so semantic tables whose names contain spaces are not falsely reported as
    missing. Relationship endpoint parsing remains blocking for genuinely absent
    tables or columns.
    """
    complete_table_files = {
        path: read_text(path) for path in sorted(paths.tables.glob("*.tmdl"))
    }
    complete_table_files.update(proposed_table_files)

    declared = {}
    for path, value in complete_table_files.items():
        table_match = re.search(r"(?mi)^\s*table\s+(.+?)\s*$", value)
        if not table_match:
            continue
        table_name = _unquote_identifier(table_match.group(1))
        declared[norm(table_name)] = {
            norm(_unquote_identifier(match.group(1)))
            for match in re.finditer(r"(?mi)^\s*column\s+(.+?)\s*$", value)
        }

    def parse_relationship_endpoint(endpoint: str):
        return _parse_tmdl_relationship_endpoint(endpoint)
    errors = []
    endpoint_pattern = re.compile(
        r"(?mi)^\s*(?:fromColumn|toColumn)\s*:\s*(?P<endpoint>.+?)\s*$"
    )
    rel_overrides = proposed_relationship_files or {}
    for relationship_file in paths.relationships:
        relationship_text = rel_overrides.get(relationship_file, read_text(relationship_file))
        for match in endpoint_pattern.finditer(relationship_text):
            endpoint = match.group("endpoint")
            table_name, column_name = parse_relationship_endpoint(endpoint)
            if not table_name or not column_name:
                errors.append(
                    f"Unable to parse relationship endpoint: {endpoint} "
                    f"(referenced by {relationship_file.name})"
                )
                continue
            table_key = norm(table_name)
            column_key = norm(column_name)
            if table_key not in declared:
                errors.append(
                    f"Missing relationship table: {table_name} "
                    f"(referenced by {relationship_file.name})"
                )
            elif column_key not in declared[table_key]:
                errors.append(
                    f"Missing relationship column: {table_name}.{column_name} "
                    f"(referenced by {relationship_file.name})"
                )
    return sorted(set(errors))

def _tmdl_source_bindings(value: str):
    """Return normalized sourceColumn-to-semantic-column bindings."""
    bindings = {}
    starts = list(re.finditer(r"(?mi)^\s*column\s+(.+?)\s*$", value))
    for index, match in enumerate(starts):
        start = match.end()
        end = starts[index + 1].start() if index + 1 < len(starts) else len(value)
        block = value[start:end]
        source = re.search(r"(?mi)^\s*sourceColumn\s*:\s*(.+?)\s*$", block)
        if not source:
            continue
        source_name = _unquote_identifier(source.group(1))
        semantic_name = _unquote_identifier(match.group(1))
        bindings.setdefault(norm(source_name), set()).add(semantic_name)
    return bindings

def _unpivot_consumed_target_columns(value: str, expected_targets: set[str]):
    """Return mapped target columns intentionally consumed by a static unpivot.

    This is an additive terminal-state check for direct-navigation contracts.
    Table.Unpivot consumes its explicit input list. Table.UnpivotOtherColumns
    consumes mapped target columns not present in its retained-column list.
    Dynamic column-list expressions remain fail-closed and are not classified.
    """
    region, _, _ = extract_m_region(value)
    if not region or not expected_targets:
        return set()
    consumed = set()
    for call in _balanced_m_calls(
        region, ("Table.Unpivot", "Table.UnpivotOtherColumns")
    ):
        args = split_m_top_level(call["body"])
        if len(args) < 4:
            continue
        columns = _parse_static_text_list(args[1])
        if columns is None:
            continue
        column_keys = {norm(column) for column in columns}
        if call["function"].casefold() == "table.unpivot":
            consumed.update(expected_targets.intersection(column_keys))
        else:
            consumed.update(expected_targets.difference(column_keys))
    return consumed

def _static_select_columns_contract(value: str):
    """Return the latest static Table.SelectColumns contract, if one exists.

    The result is query-scoped and normalized. A static SelectColumns step is an
    authoritative statement that other physical source columns are intentionally
    not part of this query's output path. Dynamic lists remain fail-closed by
    returning None, which preserves the existing validation behavior.
    """
    region, _, _ = extract_m_region(value)
    if not region:
        return None
    calls = _balanced_m_calls(region, ("Table.SelectColumns",))
    if not calls:
        return None
    # _balanced_m_calls is reverse sorted, so the first call is the latest one
    # in the M expression and therefore the most restrictive surviving contract.
    args = split_m_top_level(calls[0]["body"])
    if len(args) < 2:
        return None
    columns = _parse_static_text_list(args[1])
    if columns is None:
        return None
    return {norm(column) for column in columns}

def validate_navigation_output_contract(value: str, mappings: list[MappingRow], current_parameters=None):
    """Validate non-native direct-navigation output contracts without imposing one strategy.

    A semantic column is valid when either the query restores the legacy output
    name with Table.RenameColumns, or TMDL sourceColumn is explicitly rebound to
    the physical target output. Table-only expressions have no TMDL binding and
    therefore require the legacy output adapter when a mapped name is exposed.
    """
    table_mapping = _direct_navigation_table_mapping(value, mappings, current_parameters)
    if table_mapping is None or is_native_query(value):
        return []
    selected_columns = _static_select_columns_contract(value)
    expected = [
        (clean(m.target_column), clean(m.legacy_column))
        for m in _column_mappings_for_table(table_mapping, mappings)
        if clean(m.target_column) and clean(m.legacy_column)
        and norm(m.target_column) != norm(m.legacy_column)
        # Query-scoped refinement: when the query has a static SelectColumns
        # contract, mappings outside that list are intentionally NOT_SELECTED.
        # Dynamic contracts remain fail-closed because selected_columns is None.
        and (
            selected_columns is None
            or norm(m.target_column) in selected_columns
            or norm(m.legacy_column) in selected_columns
        )
    ]
    if not expected:
        return []
    region, _, _ = extract_m_region(value)
    rename_pairs = set()
    for call in _balanced_m_calls(region, ("Table.RenameColumns",)):
        args = split_m_top_level(call["body"])
        if len(args) < 2:
            continue
        for match in re.finditer(
            r'\{\s*"((?:[^\"]|"")*)"\s*,\s*"((?:[^\"]|"")*)"\s*\}',
            args[1], re.DOTALL,
        ):
            rename_pairs.add((
                norm(match.group(1).replace('""', '"')),
                norm(match.group(2).replace('""', '"')),
            ))
    source_bindings = _tmdl_source_bindings(value)
    has_model_columns = bool(source_bindings)
    consumed_by_unpivot = _unpivot_consumed_target_columns(
        value, {norm(target) for target, _legacy in expected}
    )
    errors = []
    for target, legacy in expected:
        restored = (norm(target), norm(legacy)) in rename_pairs
        rebound = norm(target) in source_bindings
        # A legacy binding is also valid when the final adapter restored that name.
        legacy_bound = norm(legacy) in source_bindings
        if restored and (not has_model_columns or legacy_bound):
            continue
        if rebound:
            continue
        if ALIGN_TARGET_COLUMN_NAMES and not has_model_columns:
            continue
        # A mapped physical field may be a valid intermediate input that is
        # intentionally converted from a column into Attribute/Value rows.
        # Only the specific fields proven to be consumed by a static unpivot are
        # exempted; all surviving mapped fields retain the existing validation.
        if norm(target) in consumed_by_unpivot:
            continue
        errors.append(
            "OUTPUT_CONTRACT_MISMATCH: mapped physical column "
            f"{target} is neither restored to query output {legacy} nor used by "
            "a TMDL sourceColumn binding."
        )
    return sorted(set(errors))

def _target_metadata_requirements(mappings: list[MappingRow], sources: list[SourceMapping]):
    """Collapse target checks while retaining all compatible legacy provenance."""
    requirements = {}
    for mapping in mappings:
        if not clean(mapping.target_table):
            continue
        catalog, schema, item = mapping_object_parts(mapping.target_table)
        server = http_path = ""
        for source in sources:
            legacy_object = ".".join(filter(None, [source.legacy_schema, source.legacy_item]))
            if table_name_matches(mapping.legacy_table, legacy_object):
                server = source.target_server or server
                http_path = source.target_http_path or http_path
                catalog = source.target_catalog or catalog
                schema = source.target_schema or schema
                item = source.target_item or item
                break
        key = (norm(server), norm(http_path), norm(catalog), norm(schema), norm(item), norm(mapping.target_column))
        row = requirements.setdefault(key, {
            "server": server, "http_path": http_path, "catalog": catalog,
            "schema": schema, "item": item, "column": clean(mapping.target_column), "provenance": set(),
        })
        legacy = ".".join(part for part in (clean(mapping.legacy_table), clean(mapping.legacy_column)) if part)
        if legacy:
            row["provenance"].add(legacy)
    return list(requirements.values())

def validate_target_metadata(mappings: list[MappingRow], sources: list[SourceMapping]):
    """Validate planned Databricks objects/columns when a token is supplied.

    Set PBI_MIGRATOR_DATABRICKS_TOKEN. Host, HTTP path, and catalog continue to
    come from Data Source Mapping. Without a token, validation is reported as
    NOT_PERFORMED and migration behavior remains unchanged.
    """
    requirements = _target_metadata_requirements(mappings, sources)
    token = clean(os.getenv("PBI_MIGRATOR_DATABRICKS_TOKEN"))
    rows, errors, connections = [], [], {}
    for requirement in requirements:
        status, code, details = "NOT_PERFORMED", "TARGET_METADATA_NOT_VALIDATED", "Databricks token not supplied."
        if token:
            connection_key = (requirement["server"], requirement["http_path"])
            try:
                if connection_key not in connections:
                    from databricks import sql as databricks_sql
                    connections[connection_key] = databricks_sql.connect(
                        server_hostname=requirement["server"],
                        http_path=requirement["http_path"],
                        access_token=token,
                    )
                cursor = connections[connection_key].cursor()
                cursor.execute(
                    f"SELECT column_name FROM `{requirement['catalog']}`.information_schema.columns "
                    "WHERE table_schema = ? AND table_name = ?",
                    [requirement["schema"], requirement["item"]],
                )
                columns = {norm(row[0]) for row in cursor.fetchall()}
                cursor.close()
                if not columns:
                    status, code = "NOT_FOUND", "TARGET_OBJECT_NOT_FOUND"
                    details = f"Object {requirement['catalog']}.{requirement['schema']}.{requirement['item']} was not returned by information_schema.columns."
                elif requirement["column"] and norm(requirement["column"]) not in columns:
                    status, code = "NOT_FOUND", "TARGET_COLUMN_NOT_FOUND"
                    details = f"Column {requirement['column']} was not returned for {requirement['catalog']}.{requirement['schema']}.{requirement['item']}."
                else:
                    status, code, details = "VERIFIED", "", "Target object and mapped column were returned by target metadata."
            except Exception as exc:
                status, code, details = "ERROR", "TARGET_METADATA_CONNECTION_FAILED", str(exc)
        rows.append({
            "Platform": "Databricks", "Target Server": requirement["server"],
            "Target Database/Catalog": requirement["catalog"], "Target Schema": requirement["schema"],
            "Target Object": requirement["item"], "Target Column": requirement["column"],
            "Legacy Provenance": "; ".join(sorted(requirement["provenance"], key=str.casefold)),
            "Validation Status": status, "Error Code": code, "Details": details,
        })
        if status in {"NOT_FOUND", "ERROR"}:
            identity = ".".join(filter(None, [requirement["catalog"], requirement["schema"], requirement["item"], requirement["column"]]))
            errors.append(f"{code}: {identity} - {details}")
    for connection in connections.values():
        connection.close()
    return rows, sorted(set(errors))

def validate_mapping_convergence(mappings: list[MappingRow]):
    """Validate mappings without rejecting compatible many-to-one convergence.

    Multiple legacy tables may intentionally converge on one target table, and
    equivalent legacy columns may converge on one target column. Conflicts exist
    only when the same legacy owner or the same legacy owner/column resolves to
    multiple target identities. Target-only metadata rows are ignored here.
    """
    errors = []
    table_targets = {}
    column_targets = {}
    target_provenance = {}

    for mapping in mappings:
        legacy_table = clean(mapping.legacy_table)
        target_table = clean(mapping.target_table)
        legacy_column = clean(mapping.legacy_column)
        target_column = clean(mapping.target_column)
        if not legacy_table or not target_table:
            continue

        legacy_table_key = norm(legacy_table)
        target_table_key = norm(target_table)
        table_targets.setdefault(legacy_table_key, set()).add(target_table_key)

        if legacy_column and target_column:
            legacy_column_key = (legacy_table_key, norm(legacy_column))
            target_column_key = (target_table_key, norm(target_column))
            column_targets.setdefault(legacy_column_key, set()).add(target_column_key)
            target_provenance.setdefault(target_column_key, set()).add(legacy_column_key)

    for legacy_table_key, targets in sorted(table_targets.items()):
        if len(targets) > 1:
            errors.append(
                "LEGACY_TABLE_TARGET_CONFLICT: legacy table "
                f"{legacy_table_key} maps to multiple target tables: "
                + ", ".join(sorted(targets))
            )

    for (legacy_table_key, legacy_column_key), targets in sorted(column_targets.items()):
        if len(targets) > 1:
            rendered = ", ".join(
                f"{target_table}.{target_column}"
                for target_table, target_column in sorted(targets)
            )
            errors.append(
                "LEGACY_COLUMN_TARGET_CONFLICT: "
                f"{legacy_table_key}.{legacy_column_key} maps to multiple "
                f"target columns: {rendered}"
            )

    # target_provenance is intentionally many-valued. Multiple legacy owners
    # converging on the same target table/column is compatible and is preserved.
    return sorted(set(errors))

def validate_tmdl_m_indentation(path: Path, value: str):
    """Fail closed when a TMDL source expression escapes its indentation block."""
    if path.suffix.casefold() != ".tmdl":
        return []
    lines = value.splitlines()
    errors = []
    for source_index, line in enumerate(lines):
        if not re.match(r"^\s*source\s*=\s*$", line):
            continue
        source_indent = len(line) - len(line.lstrip("\t "))
        first_index = next((
            index for index in range(source_index + 1, len(lines))
            if lines[index].strip()
        ), None)
        if first_index is None:
            errors.append(
                f"TMDL_M_EXPRESSION_INDENTATION_INVALID: {path.name} source has no expression body."
            )
            continue
        first_line = lines[first_index]
        first_indent = len(first_line) - len(first_line.lstrip("\t "))
        if first_indent <= source_indent:
            errors.append(
                "TMDL_M_EXPRESSION_INDENTATION_INVALID: "
                f"{path.name} line {first_index + 1} is not indented beneath source =."
            )
            continue
        for index in range(first_index, len(lines)):
            current = lines[index]
            if not current.strip():
                continue
            indent = len(current) - len(current.lstrip("\t "))
            if index > first_index and indent <= source_indent:
                break
            if indent < first_indent:
                errors.append(
                    "TMDL_M_EXPRESSION_INDENTATION_INVALID: "
                    f"{path.name} line {index + 1} escapes the M expression block: "
                    f"{current.strip()[:120]}"
                )
    return sorted(set(errors))

def clean_preview_log(log):
    valid_files = {
        clean(row.get("Input File"))
        for tab in ("Table Changes", "Column Changes", "Native SQL Changes", "M Query Changes", "Connection Changes", "Semantic & Report Changes")
        for row in log.get(tab, [])
        if clean(row.get("Input File"))
    }
    error_files = {
        clean(row.get("Input File")) for row in log.get("Errors", [])
        if clean(row.get("Input File"))
    }
    for row in log.get("Files Processed", []):
        if (clean(row.get("Input File")) not in valid_files
                and clean(row.get("Input File")) not in error_files
                and row.get("Changed") == "Proposed"):
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
    "Table.DropColumns": "drop columns (invalid Power Query function)",
    "Table.ReorderColumns": "reorder columns",
    "Table.DuplicateColumn": "duplicate column",
    "Table.RenameColumns": "rename columns",
    "Table.TransformColumnTypes": "type conversion",
    "Table.Unpivot": "unpivot",
    "Table.UnpivotOtherColumns": "unpivot other columns",
}

def split_m_top_level_with_spans(text: str, delimiter: str = ",") -> list[dict]:
    """Split top-level M text and preserve exact trimmed source coordinates."""
    parts = []
    start = 0
    paren = brace = bracket = 0
    in_string = False
    line_comment = False
    block_comment = 0
    index = 0
    while index < len(text):
        char = text[index]
        nxt = text[index + 1] if index + 1 < len(text) else ""
        if line_comment:
            if char in "\r\n":
                line_comment = False
            index += 1
            continue
        if block_comment:
            if char == "/" and nxt == "*":
                block_comment += 1
                index += 2
                continue
            if char == "*" and nxt == "/":
                block_comment -= 1
                index += 2
                continue
            index += 1
            continue
        if in_string:
            if char == '"':
                if nxt == '"':
                    index += 2
                    continue
                in_string = False
            index += 1
            continue
        if char == "/" and nxt == "/":
            line_comment = True
            index += 2
            continue
        if char == "/" and nxt == "*":
            block_comment = 1
            index += 2
            continue
        if char == '"':
            in_string = True
        elif char == "(":
            paren += 1
        elif char == ")":
            paren = max(0, paren - 1)
        elif char == "{":
            brace += 1
        elif char == "}":
            brace = max(0, brace - 1)
        elif char == "[":
            bracket += 1
        elif char == "]":
            bracket = max(0, bracket - 1)
        elif char == delimiter and paren == 0 and brace == 0 and bracket == 0:
            raw = text[start:index]
            leading = len(raw) - len(raw.lstrip())
            trailing = len(raw.rstrip())
            parts.append({"text": raw.strip(), "start": start + leading, "end": start + trailing})
            start = index + 1
        index += 1
    raw = text[start:]
    leading = len(raw) - len(raw.lstrip())
    trailing = len(raw.rstrip())
    parts.append({"text": raw.strip(), "start": start + leading, "end": start + trailing})
    return parts

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
    """Return actual external query references, excluding local let-step names."""
    steps = parse_m_steps(m_text)
    local_steps = {norm(step.get('name', '').strip('#"')) for step in steps}
    names = set()
    def add(expression: str):
        name = _parse_query_identifier(expression)
        if name and norm(name) not in local_steps:
            names.add(name)
    for step in steps:
        add(step.get('expression', ''))
    for call in _balanced_m_calls(
        m_text, ('Table.NestedJoin', 'Table.Join', 'Table.FuzzyNestedJoin', 'Table.Combine')
    ):
        args = split_m_top_level(call['body'])
        if call['function'].casefold() == 'table.combine':
            if args:
                members = args[0].strip()
                if members.startswith('{') and members.endswith('}'):
                    for member in split_m_top_level(members[1:-1]):
                        add(member)
        else:
            if len(args) >= 1: add(args[0])
            if len(args) >= 3: add(args[2])
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
    # Mapping eligibility is branch-local. Dependency reachability and a
    # coincidental table/column name are impact signals only, never edit authority.
    # Generic M replacement is allowed only for a query that directly navigates
    # the exact mapped physical table. Unmapped and mixed-source derived queries
    # retain their existing output contracts.
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
    inner = value[1:-1].strip()
    if not inner:
        return []
    result = []
    for item in split_m_top_level(inner):
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
    """Restore legacy row values using a single-line, TMDL-safe M expression.

    A replacement may begin in the middle of an already-indented TMDL assignment.
    Rendering the wrapper on one line preserves the host document's indentation
    exactly and avoids introducing continuation lines outside the source block.
    """
    if not replacements:
        return unpivot_expression
    record = ", ".join(
        f'{_m_field(target)} = "{_m_escape(legacy)}"'
        for target, legacy in sorted(replacements.items(), key=lambda pair: norm(pair[0]))
    )
    transform = (
        f'{{{{"{_m_escape(attribute_column)}", '
        f"each Record.FieldOrDefault([{record}], _, _), type text}}}}"
    )
    return f"Table.TransformColumns({unpivot_expression}, {transform})"

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
        if not attribute_match or not value_match:
            errors.append(
                f"{input_file}: malformed {call['function']} output column arguments."
            )
            continue
        if columns is None:
            # Preserve runtime-generated column-list semantics. The migrator must
            # not enumerate or replace a dynamic Unpivot/UnpivotOtherColumns list.
            inventory.append({
                "Input File": input_file,
                "Query Name": query_name,
                "Unpivot Function": call["function"],
                "Attribute Column": attribute_match.group(1).replace('""', '"'),
                "Value Column": value_match.group(1).replace('""', '"'),
                "Legacy Attribute Value": "",
                "Target Source Column": "",
                "Input Reference Updated": "No",
                "Attribute Preservation Required": "No",
                "Attribute Preservation Method": "Dynamic column-list expression preserved",
                "Attribute Value Preserved": "Runtime semantics preserved",
                "Mappings Applied": "0",
                "Result": status,
            })
            continue

        attribute_column = attribute_match.group(1).replace('""', '"')
        value_column = value_match.group(1).replace('""', '"')
        original_columns = list(columns)

        # When ALIGN_TARGET_COLUMN_NAMES is True, native SQL also exposes target
        # column names, so post-native unpivots must map inputs to target names
        # and wrap Attribute values with Record.FieldOrDefault just like non-native.
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
        # If the outer SELECT kept an explicit author alias (so target is not exposed),
        # do not rebind downstream M/TMDL to an unexposed physical name.
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
            _defs, current_parameters = parameter_definitions(paths)
            for payload in discover_native_sql_payloads(block, current_parameters):
                sources, _, parse_errors = _persistent_sql_sources(payload.sql)
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
    """Match static M navigation records regardless of input step or field order."""
    return re.compile(
        r'(?is)(?P<input>#"(?:[^"]|"")+"|[A-Za-z_][A-Za-z0-9_ ]*)'
        r'\s*\{\s*\[(?P<fields>.*?)\]\s*\}\s*\[Data\]'
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

def _append_output_contract_rename(block: str, pairs: list[tuple[str, str]]):
    """Restore legacy exposed names after migrating one direct-navigation query."""
    unique, seen = [], set()
    for target, legacy in pairs:
        key = (norm(target), norm(legacy))
        if target and legacy and key[0] != key[1] and key not in seen:
            seen.add(key); unique.append((target, legacy))
    if not unique:
        return block, 0
    region, start, end = extract_m_region(block)
    if start < 0:
        return block, 0
    match = re.search(r'(?is)\bin\s+(?P<output>#"(?:[^"]|"")+"|[A-Za-z_][A-Za-z0-9_ ]*)\s*$', region)
    if not match:
        return block, 0
    output_step = match.group('output').strip()
    existing = {norm(step.get('name', '').strip('#"')) for step in parse_m_steps(region)}
    name, ordinal = 'Migration Output Contract', 2
    while norm(name) in existing:
        name = f'Migration Output Contract {ordinal}'; ordinal += 1
    step = '#"' + _m_escape(name) + '"'
    rendered = ', '.join('{{"' + _m_escape(target) + '", "' + _m_escape(legacy) + '"}}' for target, legacy in unique)
    indent_match = re.search(r'(?m)^(?P<indent>[ \t]*)[^\r\n]+\s*=\s*', region)
    indent = indent_match.group('indent') if indent_match else '    '
    adapter = output_step + ',\n' + indent + step + ' = Table.RenameColumns(' + output_step + ', {' + rendered + '}, MissingField.Ignore)\n' + indent + 'in\n' + indent + step
    new_region = region[:match.start()] + adapter + region[match.end():]
    return block[:start] + new_region + block[end:], len(unique)

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
    contract_pairs = []

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
                "",
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
                contract_pairs.append((column_mapping.target_column, column_mapping.legacy_column))
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
    # Physical names are local implementation details. Re-expose the original
    # query/model contract so downstream M, relationships, and TMDL sourceColumn
    # bindings continue to use established names.
    if contract_pairs and input_file.casefold() == "expressions.tmdl" and not ALIGN_TARGET_COLUMN_NAMES:
        updated, contract_count = _append_output_contract_rename(updated, contract_pairs)
        if contract_count:
            log["M Query Changes"].append({
                "Input File": input_file,
                "Query Name": query_name,
                "M Step": "Final output",
                "Change Type": "Legacy/model-facing output contract restoration",
                "Old Value": "; ".join(target for target, _legacy in contract_pairs),
                "New Value": "; ".join(legacy for _target, legacy in contract_pairs),
                "Result": f"{status}; {contract_count} output name(s) restored",
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

def repair_databricks_navigation_indentation(value: str):
    """Align generated Databricks navigation bindings to their Source binding.

    Navigation bindings are identified by Databricks Name/Kind records rather
    than fixed step names. This supports both Navigation 1-3 and collision-safe
    names such as Navigation Migrated 1 1-3.
    """
    region, start, end = extract_m_region(value)
    if start < 0 or "Databricks.Catalogs" not in region:
        return value, 0

    _, _, _, bindings = _m_let_bindings_with_spans(value)
    if not bindings:
        return value, 0

    bindings_by_name = {
        norm(binding["name_value"]): binding
        for binding in bindings
    }

    navigation_bindings = []

    for binding in bindings:
        match = _navigation_record_pattern().fullmatch(
            binding["expression"].strip()
        )
        if not match:
            continue

        fields = match.group("fields")
        node_kind = _m_record_field(fields, "Kind")
        node_name = _m_record_field(fields, "Name")

        if node_kind not in {"Database", "Schema", "Table", "View"}:
            continue
        if not node_name:
            continue

        input_name = (
            _parse_query_identifier(match.group("input"))
            or clean(match.group("input")).strip('#"')
        )

        navigation_bindings.append({
            "binding": binding,
            "input_name": input_name,
            "kind": node_kind,
        })

    repaired_region = region
    replacements = 0
    indent_changes = []

    for navigation in navigation_bindings:
        binding = navigation["binding"]
        input_binding = bindings_by_name.get(
            norm(navigation["input_name"])
        )

        if input_binding is None:
            continue

        expected_indent = input_binding.get("indent", "")
        current_indent = binding.get("indent", "")

        if current_indent == expected_indent:
            continue

        region_relative_start = (
            binding["start"] - start
            if binding["start"] >= start
            else binding["start"]
        )

        indent_changes.append(
            (
                region_relative_start,
                current_indent,
                expected_indent,
            )
        )

    for binding_start, current_indent, expected_indent in sorted(
        indent_changes,
        reverse=True,
    ):
        if repaired_region.startswith(current_indent, binding_start):
            repaired_region = (
                repaired_region[:binding_start]
                + expected_indent
                + repaired_region[
                    binding_start + len(current_indent):
                ]
            )
            replacements += 1

    return (
        value[:start] + repaired_region + value[end:],
        replacements,
    )

def validate_m_partition_indentation(path: Path, value: str):
    """Validate Databricks navigation binding indentation.

    Navigation bindings are identified by Databricks Name/Kind structure,
    not by fixed Navigation 1-3 step names.
    """
    errors = []
    region, _, _ = extract_m_region(value)

    if not region or "Databricks.Catalogs" not in region:
        return errors

    _, _, _, bindings = _m_let_bindings_with_spans(value)
    if not bindings:
        return errors

    bindings_by_name = {
        norm(binding["name_value"]): binding
        for binding in bindings
    }

    navigation_bindings = []

    for binding in bindings:
        match = _navigation_record_pattern().fullmatch(
            binding["expression"].strip()
        )
        if not match:
            continue

        fields = match.group("fields")
        node_kind = _m_record_field(fields, "Kind")
        node_name = _m_record_field(fields, "Name")

        if node_kind not in {"Database", "Schema", "Table", "View"}:
            continue
        if not node_name:
            continue

        input_name = (
            _parse_query_identifier(match.group("input"))
            or clean(match.group("input")).strip('#"')
        )

        navigation_bindings.append({
            "binding": binding,
            "input_name": input_name,
            "kind": node_kind,
            "node_name": node_name,
        })

    # Native Value.NativeQuery embeds Databricks.Catalogs inside Source and
    # therefore has no standalone navigation bindings to validate.
    if not navigation_bindings:
        return errors

    kinds = {
        norm(navigation["kind"])
        for navigation in navigation_bindings
    }

    for required_kind in ("Database", "Schema"):
        if norm(required_kind) not in kinds:
            errors.append(
                f"Missing generated Databricks {required_kind} "
                f"navigation in {path.name}"
            )

    if not {"table", "view"}.intersection(kinds):
        errors.append(
            f"Missing generated Databricks Table/View "
            f"navigation in {path.name}"
        )

    for navigation in navigation_bindings:
        binding = navigation["binding"]
        input_binding = bindings_by_name.get(
            norm(navigation["input_name"])
        )

        if input_binding is None:
            errors.append(
                "Databricks navigation input is not a declared M binding "
                f'in {path.name}: {binding["name"]} references '
                f'{navigation["input_name"]}'
            )
            continue

        expected_indent = input_binding.get("indent", "")
        actual_indent = binding.get("indent", "")

        if actual_indent != expected_indent:
            errors.append(
                "Databricks navigation indentation mismatch in "
                f'{path.name}: {binding["name"]} should align with '
                f'{input_binding["name"]}'
            )

    return sorted(set(errors))

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
            tree = sqlglot.parse_one(decode_m_sql_string(payload), read=TARGET_SQL_DIALECT)
        except ParseError as exc:
            errors.append(f"Native SQL output contract parse failed for payload {index}: {sanitize_diagnostic_text(exc)}")
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

def _direct_navigation_table_mapping(value: str, mappings: list[MappingRow], current_parameters=None):
    """Resolve the physical table directly navigated by one non-native query.

    Supports both legacy Synapse Schema/Item navigation and migrated
    Databricks Database/Schema/Table navigation, including parameterized
    Schema, Item, or Name expressions.
    """
    region, _, _ = extract_m_region(value)
    if not region:
        return None
    params = current_parameters or {}

    def _resolve_field(fields: str, field_name: str) -> str:
        val = _m_record_field(fields, field_name)
        if val:
            return val
        if params:
            expr = _m_record_field_expression(fields, field_name)
            if expr:
                resolved, _param = _resolve_m_connector_argument(expr, params)
                if resolved:
                    return resolved
        return ""

    # Legacy Synapse navigation:
    #   Source{[Schema="...", Item="..."]}[Data] (or parameterized Schema/Item)
    for match in _navigation_record_pattern().finditer(region):
        fields = match.group("fields")
        schema = _resolve_field(fields, "Schema")
        item = _resolve_field(fields, "Item")
        if not item:
            continue

        mapping = _table_mapping_for_navigation(schema, item, mappings)
        if mapping is not None:
            return mapping

        # Recovery/idempotency: navigation may already use the mapped target
        # schema/item while still retaining the legacy navigation shape.
        object_name = ".".join(part for part in (schema, item) if part)
        candidates = [
            pair for pair in _unique_table_pairs(mappings)
            if table_name_matches(pair.target_table, object_name)
        ]
        targets = {_normalize_object_name(pair.target_table) for pair in candidates}
        if len(targets) == 1:
            return candidates[0]

    # Migrated Databricks navigation:
    #   Navigation 1 -> Database
    #   Navigation 2 -> Schema (literal or parameter)
    #   Navigation 3 -> Table/View (literal or parameter)
    schema_nodes, item_nodes = [], []
    for match in _navigation_record_pattern().finditer(region):
        fields = match.group("fields")
        kind = _m_record_field(fields, "Kind")
        name_val = _resolve_field(fields, "Name")
        if not name_val:
            continue
        if norm(kind) == "schema":
            schema_nodes.append(name_val)
        elif norm(kind) in {"table", "view"}:
            item_nodes.append(name_val)

    if schema_nodes and item_nodes:
        schema = schema_nodes[-1]
        item = item_nodes[-1]
        object_name = ".".join(part for part in (schema, item) if part)
        for match_target in (True, False):
            candidates = [
                pair for pair in _unique_table_pairs(mappings)
                if table_name_matches(pair.target_table if match_target else pair.legacy_table, object_name)
            ]
            targets = {_normalize_object_name(pair.target_table) for pair in candidates}
            if len(targets) == 1:
                return candidates[0]
        # Fallback: when Item is a legacy parameter value (e.g., pOrdersTable = "FactOrders_Old")
        # paired with a target schema ("silver"), match by item name alone if unique.
        item_candidates = [
            pair for pair in _unique_table_pairs(mappings)
            if table_name_matches(pair.legacy_table, item) or table_name_matches(pair.target_table, item)
        ]
        item_targets = {_normalize_object_name(pair.target_table) for pair in item_candidates}
        if len(item_targets) == 1:
            return item_candidates[0]

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
    """Return legacy-to-current columns exposed by one direct navigation or native SQL query."""
    if is_native_query(value, current_parameters):
        if not ALIGN_TARGET_COLUMN_NAMES:
            return {}
        result, conflicts = {}, set()
        for payload in discover_native_sql_payloads(value, current_parameters or {}):
            sources, _audit, _errors = _persistent_sql_sources(payload.sql)
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
        for legacy in conflicts:
            result.pop(legacy, None)
        return result
    table_mapping = _direct_navigation_table_mapping(value, mappings, current_parameters)
    if table_mapping is None:
        return {}
    result, conflicts = {}, set()
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
    # Respect explicit same-query Table.RenameColumns overrides: if the query
    # itself renames a column to a custom alias (or back to its legacy name),
    # downstream queries see that explicit output name instead of the physical target.
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
                dst_col = m_pair.group(2).replace('""', '"')
                if norm(src_col) in result:
                    if norm(dst_col) == norm(src_col):
                        result.pop(norm(src_col), None)
                    else:
                        result.pop(norm(src_col), None)
    return result

def _parse_query_identifier(expression: str):
    value = expression.strip()
    quoted = re.fullmatch(r'#"((?:[^"]|"")+)"', value)
    if quoted:
        return quoted.group(1).replace('""', '"')
    plain = re.fullmatch(r'[A-Za-z_][A-Za-z0-9_ ]*', value)
    return plain.group(0).strip() if plain else ""

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
    """Reconcile Table.Combine members.

    When ALIGN_TARGET_COLUMN_NAMES is True (forward adaptation), migrated branches
    remain on target column names while unmigrated or partial branches are adapted
    forward with Table.RenameColumns(..., MissingField.Ignore). Fails closed on
    conflicting sibling targets or column collisions on an unmigrated branch.
    """
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

        if ALIGN_TARGET_COLUMN_NAMES:
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
            continue

        rewritten_members, local_changes = [], []
        for member in members:
            branch = _parse_query_identifier(member)
            if not branch:
                rewritten_members.append(member)
                continue
            binding_map, local_errors = resolved_query_output_bindings(
                branch, originals_by_query, mappings, binding_cache, None, current_parameters)
            errors.extend(local_errors)
            pairs = []
            for legacy_key, current in binding_map.items():
                legacy = next(
                    (m.legacy_column for m in mappings
                     if norm(m.legacy_column) == norm(legacy_key)
                     and norm(m.target_column) == norm(current)),
                    legacy_key,
                )
                if norm(current) != norm(legacy):
                    pairs.append((current, legacy))
            if not pairs:
                rewritten_members.append(member)
                continue
            rendered = ", ".join('{"' + _m_escape(current) + '", "' + _m_escape(legacy) + '"}'
                                  for current, legacy in pairs)
            adapted = "Table.RenameColumns(" + member.strip() + ", {" + rendered + "}, MissingField.Ignore)"
            rewritten_members.append(adapted)
            local_changes.append({"branch":branch, "pairs":pairs, "direction": "backward"})
        if not local_changes:
            continue
        args[0] = "{" + ", ".join(rewritten_members) + "}"
        rewritten = "Table.Combine(" + ", ".join(args) + ")"
        changed_region = changed_region[:call["start"]] + rewritten + changed_region[call["end"]:]
        changes.extend(local_changes)
    return value[:start] + changed_region + value[end:], changes, sorted(set(errors))

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
            if ALIGN_TARGET_COLUMN_NAMES:
                target_input = right_column_map.get(norm(input_name), input_name)
                target_output = target_input if norm(output_name) == norm(input_name) else output_name
                result[norm(output_name)] = target_output
            else:
                # ExpandTableColumn's fourth argument is the query output contract.
                # Physical changes on the right branch do not rename that output.
                result[norm(output_name)] = output_name
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
    """Propagate merged upstream target names into Table.AddColumn row formulas.

    Left-side row fields use the left query's physical mapping. Right-side fields
    are eligible only when explicitly introduced by Table.ExpandTableColumn.
    This prevents a same-named column on an unexpanded right query from stealing
    lineage from a left-side field such as Location.
    """
    # Resolve downstream row fields from the branch that supplies them.
    # Left-side fields inherit the migrated output of the merge's left query.
    # Right-side fields are eligible only when ExpandTableColumn explicitly exposes
    # them, preventing same-named right-side fields from stealing left lineage.
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
    reconciled = value[:start] + changed_region + value[end:]
    errors = []
    reconciled_region, _, _ = extract_m_region(reconciled)
    for call in _balanced_m_calls(reconciled_region, ("Table.AddColumn",)):
        args = split_m_top_level(call["body"])
        if len(args) < 3:
            continue
        formula = args[2]
        for legacy_norm, target in left_map.items():
            legacy_name = next((
                mapping.legacy_column for mapping in mappings
                if norm(mapping.legacy_column) == legacy_norm
                and norm(mapping.target_column) == norm(target)
            ), "")
            if not legacy_name:
                continue
            if re.search(r'\[#"' + re.escape(_m_escape(legacy_name)) + r'"\]|\[' + re.escape(legacy_name) + r'\]', formula, re.IGNORECASE):
                errors.append(
                    "DOWNSTREAM_M_COLUMN_NOT_RECONCILED: "
                    f"Table.AddColumn still references {legacy_name}; "
                    f"left query {left_query} exposes mapped column {target}."
                )
    return reconciled, changes, sorted(set(errors))

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
        right_map, local_errors = (
            resolved_query_output_bindings(right_query, originals_by_query, mappings, binding_cache, None, current_parameters)
            if right_query else ({}, [])
        )
        errors.extend(local_errors)
        if not right_map:
            continue
        new_inputs, new_outputs, local = [], [], {}
        for legacy_input, legacy_output in zip(inputs, outputs):
            target_input = right_map.get(norm(legacy_input), legacy_input)
            new_inputs.append(target_input)
            if ALIGN_TARGET_COLUMN_NAMES and norm(legacy_output) == norm(legacy_input):
                new_outputs.append(target_input)
                if norm(target_input) != norm(legacy_output):
                    local[norm(legacy_output)] = target_input
            else:
                new_outputs.append(legacy_output)
        if new_inputs == inputs and new_outputs == outputs:
            continue
        args[2] = _render_text_list(new_inputs)
        if len(args) >= 4:
            args[3] = _render_text_list(new_outputs)
        rewritten = "Table.ExpandTableColumn(" + ", ".join(args) + ")"
        changed_region = changed_region[:call["start"]] + rewritten + changed_region[call["end"]:]
        bindings.update(local)
        changes.append({"right_query":right_query,"old_inputs":inputs,"new_inputs":new_inputs,"old_outputs":outputs,"new_outputs":new_outputs})
    return value[:start]+changed_region+value[end:], changes, bindings, sorted(set(errors))

def _merged_query_output_bindings(value: str, originals_by_query: dict[str, str], mappings: list[MappingRow], binding_cache=None, current_parameters=None):
    """Resolve final model-bound names across single or chained merges and expansions."""
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

def _db_source_for(server, database, schema, item, sources):
    """Resolve the most-specific unambiguous authorized source mapping."""
    candidates = []
    for source in sources:
        if source.legacy_server and norm(source.legacy_server) != norm(server):
            continue
        if source.legacy_database and norm(source.legacy_database) != norm(database):
            continue
        if source.legacy_schema and schema and norm(source.legacy_schema) != norm(schema):
            continue
        if source.legacy_item and item and norm(source.legacy_item) != norm(item):
            continue
        if not (source.target_server and source.target_http_path and source.target_catalog):
            continue
        score = 2
        if source.legacy_schema and schema and norm(source.legacy_schema) == norm(schema):
            score += 1
        if source.legacy_item and item and norm(source.legacy_item) == norm(item):
            score += 2
        candidates.append((score, source))
    if not candidates:
        return None
    best_score = max(score for score, _ in candidates)
    best = [source for score, source in candidates if score == best_score]
    if schema or item:
        target_keys = {
            (norm(source.target_server), norm(source.target_http_path),
             norm(source.target_catalog), norm(source.target_schema), norm(source.target_item))
            for source in best
        }
        return best[0] if len(target_keys) == 1 else None
    connection_keys = {
        (norm(source.target_server), norm(source.target_http_path), norm(source.target_catalog))
        for source in best
    }
    return best[0] if len(connection_keys) == 1 else None

def _m_function_call_span(value: str, function_name: str):
    """Return the exact span of the first balanced M function call."""
    match = re.search(r'(?i)'+re.escape(function_name)+r'\s*\(', value)
    if not match:
        return None
    opening = value.find('(', match.start(), match.end())
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
    return (match.start(), index) if depth == 0 else None

def _m_let_bindings_with_spans(value: str):
    """Parse top-level M let bindings and preserve exact region-relative spans.

    This parser is connector-agnostic. It handles multiline expressions, quoted
    identifiers, comments, nested function calls, records, lists, and strings.
    It never evaluates M and it does not assume a binding is named Source.
    """
    region, region_start, region_end = extract_m_region(value)
    if region_start < 0:
        return region, region_start, region_end, []
    let_match = re.search(r'(?i)\blet\b', region)
    if not let_match:
        return region, region_start, region_end, []
    body_start = let_match.end()
    paren = brace = bracket = 0
    in_string = False
    line_comment = False
    block_comment = 0
    assignment_start = body_start
    assignments = []
    index = body_start
    in_start = -1
    while index < len(region):
        char = region[index]
        nxt = region[index + 1] if index + 1 < len(region) else ''
        if line_comment:
            if char in '\r\n':
                line_comment = False
            index += 1
            continue
        if block_comment:
            if char == '/' and nxt == '*':
                block_comment += 1; index += 2; continue
            if char == '*' and nxt == '/':
                block_comment -= 1; index += 2; continue
            index += 1
            continue
        if in_string:
            if char == '"':
                if nxt == '"':
                    index += 2; continue
                in_string = False
            index += 1
            continue
        if char == '/' and nxt == '/':
            line_comment = True; index += 2; continue
        if char == '/' and nxt == '*':
            block_comment = 1; index += 2; continue
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
        elif paren == 0 and brace == 0 and bracket == 0:
            if char == ',':
                assignments.append((assignment_start, index, index + 1))
                assignment_start = index + 1
            elif (char == 'i' or char == 'I') and re.match(r'(?i)in\b', region[index:]):
                before = region[index - 1] if index else ' '
                if not (before.isalnum() or before == '_'):
                    assignments.append((assignment_start, index, index))
                    in_start = index
                    break
        index += 1
    output = []
    name_pattern = re.compile(
        r'(?is)^(?P<leading>(?:\s+|//[^\r\n]*(?:\r?\n|$)|/\*.*?\*/\s*)*)'
        r'(?P<name>#"(?:[^"]|"")+"|[A-Za-z_][A-Za-z0-9_ ]*)'
        r'\s*=\s*(?P<expr>.*?)(?P<trailing>\s*)$'
    )
    for ordinal, (start, end, comma_end) in enumerate(assignments):
        raw = region[start:end]
        match = name_pattern.match(raw)
        if not match:
            continue
        leading = match.group('leading')
        last_line = leading.rsplit('\n', 1)[-1].rsplit('\r', 1)[-1]
        indent_match = re.match(r'[ \t]*$', last_line)
        indent = indent_match.group(0) if indent_match else ''
        output.append({
            'ordinal': ordinal,
            'name': match.group('name').strip(),
            'name_value': (_parse_query_identifier(match.group('name').strip())
                           or match.group('name').strip()),
            'expression': match.group('expr').strip(),
            'start': start, 'end': end, 'comma_end': comma_end,
            'leading': leading, 'indent': indent, 'trailing': match.group('trailing'),
            'expr_start': start + match.start('expr'),
            'expr_end': start + match.end('expr'),
        })
    return region, region_start, region_end, output

def _static_navigation_expression(expression: str, current_parameters=None):
    """Return input step/schema/item for one static or parameterized navigation expression."""
    match = _navigation_record_pattern().fullmatch(expression.strip())
    if not match:
        return None
    fields = match.group('fields')
    schema = _m_record_field(fields, 'Schema')
    item = _m_record_field(fields, 'Item')
    schema_alias = item_alias = ''
    if current_parameters:
        if not schema:
            schema_expr = _m_record_field_expression(fields, 'Schema')
            resolved_schema, param_name = _resolve_m_connector_argument(schema_expr, current_parameters)
            if resolved_schema:
                schema = resolved_schema
                schema_alias = schema_expr.strip() if param_name else ''
        if not item:
            item_expr = _m_record_field_expression(fields, 'Item')
            resolved_item, param_name = _resolve_m_connector_argument(item_expr, current_parameters)
            if resolved_item:
                item = resolved_item
                item_alias = item_expr.strip() if param_name else ''
    input_name = _parse_query_identifier(match.group('input')) or clean(match.group('input'))
    if not (schema and item):
        return None
    return {
        'input': input_name, 'schema': schema, 'item': item,
        'schema_alias': schema_alias, 'item_alias': item_alias,
    }

def _synapse_navigation_pairs(value: str, current_parameters=None):
    """Pair top-level Sql.Database bindings (literal or parameterized) with navigation bindings."""
    region, region_start, region_end, bindings = _m_let_bindings_with_spans(value)
    params = current_parameters or {}
    sources_by_name = {}
    for binding in bindings:
        expression = binding['expression']
        if not re.match(r'(?is)^Sql\.Database\s*\(', expression):
            continue
        connector = _sql_database_connector(expression, params)
        server, database = connector.get('server', ''), connector.get('database', '')
        if not server:
            server, database = _sql_db_values(expression)
        server_alias, database_alias = _legacy_connection_parameter_aliases(
            expression, params, server, database
        )
        if server:
            sources_by_name[norm(binding['name_value'])] = {
                **binding, 'server': server, 'database': database,
                'server_alias': server_alias, 'database_alias': database_alias,
            }
    pairs = []
    for binding in bindings:
        navigation = _static_navigation_expression(binding['expression'], params)
        if not navigation:
            continue
        source = sources_by_name.get(norm(navigation['input']))
        if not source:
            continue
        pairs.append({
            'source': source, 'navigation': binding,
            'schema': navigation['schema'], 'item': navigation['item'],
            'schema_alias': navigation.get('schema_alias', ''),
            'item_alias': navigation.get('item_alias', ''),
            'region_start': region_start, 'region_end': region_end,
        })
    return region, region_start, region_end, pairs

def _replace_m_identifier_after(text: str, old_name: str, new_name: str, start: int = 0):
    """Replace an M identifier token after a declaration, preserving unrelated text."""
    head, tail = text[:start], text[start:]
    if old_name.startswith('#"'):
        pattern = re.compile(re.escape(old_name))
    else:
        pattern = re.compile(r'(?<![A-Za-z0-9_])' + re.escape(old_name) + r'(?![A-Za-z0-9_])')
    return head + pattern.sub(lambda _m: new_name, tail)

def _legacy_navigation_binding(value: str):
    """Find a static Schema/Item navigation binding regardless of field order."""
    region, region_start, _ = extract_m_region(value)
    if region_start < 0:
        return None
    for match in _navigation_record_pattern().finditer(region):
        fields = match.group('fields')
        schema = _m_record_field(fields, 'Schema')
        item = _m_record_field(fields, 'Item')
        if not schema or not item:
            continue
        prefix = region[:match.start()]
        assignment = re.search(r'(?mi)(#"(?:[^"]|"")+"|[A-Za-z_][A-Za-z0-9_ ]*)\s*=\s*$', prefix)
        if not assignment:
            continue
        return {
            'step': assignment.group(1).strip(), 'schema': schema, 'item': item,
            'start': region_start + assignment.start(), 'end': region_start + match.end(),
        }
    return None

def _m_binding_name_before(value: str, expression_start: int):
    """Return the let-binding name whose expression starts at expression_start."""
    line_start = value.rfind('\n', 0, expression_start) + 1
    prefix = value[line_start:expression_start]
    match = re.search(r'(?is)^([ \t]*)(#"(?:[^"]|"")+"|[A-Za-z_][A-Za-z0-9_ ]*)\s*=\s*$', prefix)
    if not match:
        return None
    return {
        'indent': match.group(1), 'name': match.group(2).strip(),
        # Full binding start includes indentation. Atomic source/navigation replacement
        # must replace the whole binding, not start at the identifier after indentation.
        'start': line_start + match.start(0),
        'name_start': line_start + match.start(2),
        'expression_start': expression_start,
    }

def _legacy_navigation_bindings(value: str):
    """Return every static Synapse source/navigation pair in one M region."""
    region, region_start, _ = extract_m_region(value)
    if region_start < 0:
        return []
    sql_calls = []
    for match in re.finditer(r'(?i)Sql\.Database\s*\(', region):
        span = _m_function_call_span(region[match.start():], 'Sql.Database')
        if not span:
            continue
        absolute_start = region_start + match.start() + span[0]
        absolute_end = region_start + match.start() + span[1]
        binding = _m_binding_name_before(value, absolute_start)
        if not binding:
            continue
        server, database = _sql_db_values(value[absolute_start:absolute_end])
        if server:
            sql_calls.append({**binding, 'call_start': absolute_start, 'call_end': absolute_end,
                              'server': server, 'database': database})
    output = []
    for match in _navigation_record_pattern().finditer(region):
        fields = match.group('fields')
        schema = _m_record_field(fields, 'Schema')
        item = _m_record_field(fields, 'Item')
        input_step = clean(match.group('input'))
        if not schema or not item or not input_step:
            continue
        expression_start = region_start + match.start()
        nav_binding = _m_binding_name_before(value, expression_start)
        if not nav_binding:
            continue
        source_binding = next((row for row in sql_calls if norm(row['name'].strip('#"')) == norm(input_step.strip('#"'))), None)
        if not source_binding:
            continue
        output.append({
            **nav_binding, 'end': region_start + match.end(), 'schema': schema, 'item': item,
            'source_name': source_binding['name'], 'source_indent': source_binding['indent'],
            'source_binding_start': source_binding['start'],
            'source_call_start': source_binding['call_start'],
            'source_call_end': source_binding['call_end'], 'server': source_binding['server'],
            'database': source_binding['database'],
        })
    return output

def _native_connection_source(value: str, server: str, database: str, sources: list[SourceMapping]):
    """Resolve only the Databricks connection for native SQL.

    Schema/item belong to physical SQL sources and are intentionally not promoted
    into the Connection Changes row. All returned connection values originate
    from Data Source Mapping.
    """
    resolved = []
    parse_errors = []
    for _, _, payload in native_sql_regions(value):
        physical, _, errors = _sql_physical_sources(decode_m_sql_string(payload))
        parse_errors.extend(errors)
        for row in physical:
            source = _db_source_for(
                server, database, clean(row.get("Schema")), clean(row.get("Item")), sources
            )
            if source:
                resolved.append(source)
    if parse_errors:
        return None, [
            "NATIVE_CONNECTION_SOURCE_PARSE_FAILED: "
            + " | ".join(sorted(set(parse_errors)))
        ]
    if not resolved:
        fallback = _db_source_for(server, database, "", "", sources)
        if not fallback:
            return None, []
        return NativeConnectionTarget(
            target_server=fallback.target_server,
            target_http_path=fallback.target_http_path,
            target_catalog=fallback.target_catalog,
        ), []
    connection_keys = {
        (norm(x.target_server), norm(x.target_http_path), norm(x.target_catalog))
        for x in resolved
    }
    if len(connection_keys) != 1:
        return None, [
            "AMBIGUOUS_NATIVE_DATABRICKS_CONNECTION: physical SQL sources resolve "
            "to different host/warehouse/catalog values"
        ]
    first = resolved[0]
    return NativeConnectionTarget(
        target_server=first.target_server,
        target_http_path=first.target_http_path,
        target_catalog=first.target_catalog,
    ), []

def _sql_db_values(value):
    m=re.search(r'(?is)Sql\.Database\(\s*"((?:[^"]|"")*)"\s*,\s*"((?:[^"]|"")*)"',value)
    return (m.group(1).replace('""','"'),m.group(2).replace('""','"')) if m else ('','')

def _db_catalog(s, server_argument: str = 'SqlServer', http_path_argument: str = 'Database'):
    """Retain model parameters as Databricks host and HTTP-path arguments."""
    return (
        f'Databricks.Catalogs({server_argument}, {http_path_argument}, '
        f'[Catalog = "{_m_escape(s.target_catalog)}", Database = null, QueryTags = null, '
        f'EnableAutomaticProxyDiscovery = "enabled", '
        f'Implementation = "{_m_escape(DATABRICKS_IMPLEMENTATION or "2.0")}", Protocol = null])'
    )

def _m_region_for_validation(value: str):
    """Return executable M only. Connector validation must not scan unrelated TMDL metadata."""
    region, start, end = extract_m_region(value)
    return region if start >= 0 else value

def _token_locations(value: str, token: str, limit: int = 8):
    """Return line/column locations for actionable residual-connector diagnostics."""
    results = []
    if not token:
        return results
    start = 0
    while len(results) < limit:
        index = value.find(token, start)
        if index < 0:
            break
        line = value.count("\n", 0, index) + 1
        line_start = value.rfind("\n", 0, index) + 1
        column = index - line_start + 1
        line_end = value.find("\n", index)
        if line_end < 0:
            line_end = len(value)
        context = re.sub(r'\s+', ' ', value[line_start:line_end]).strip()
        results.append((line, column, context[:240]))
        start = index + max(1, len(token))
    return results

def _remaining_mapped_legacy_connectors(value: str, sources: list[SourceMapping], current_parameters=None):
    """Return only legacy connectors whose exact navigated object is mapped.

    A shared server/database is not enough. This prevents an unmapped table in the
    same Synapse database from being blocked because another table is mapped.
    Native SQL is resolved separately from its SQLGlot physical-source inventory.
    """
    region, _, _, pairs = _synapse_navigation_pairs(value, current_parameters)
    remaining = []
    for pair in pairs:
        source = pair['source']
        mapping = _db_source_for(
            source['server'], source['database'], pair['schema'], pair['item'], sources
        )
        if not mapping:
            continue
        expression = source['expression']
        line = region.count('\n', 0, source['start']) + 1
        remaining.append({
            'server': source['server'], 'database': source['database'],
            'schema': pair['schema'], 'item': pair['item'], 'line': line,
            'context': re.sub(r'\s+', ' ', expression).strip()[:240],
        })
    return remaining

def _databricks_navigation_names_by_kind(
    value: str,
    kind: str,
) -> list[str]:
    """Return Name values from static Databricks navigation records by Kind."""
    names = []

    for match in _navigation_record_pattern().finditer(value):
        fields = match.group("fields")
        node_kind = _m_record_field(fields, "Kind")
        node_name = _m_record_field(fields, "Name")

        if norm(node_kind) == norm(kind) and node_name:
            names.append(node_name)

    return names

def _validate_databricks_connection_rewrite(value, source_mapping, legacy_server,
                                             legacy_database, classification,
                                             parameter_aliases=('', '')):
    """Validate literals or preserved connection-parameter aliases."""
    errors = []
    scope = _m_region_for_validation(value)
    server_alias, http_alias = parameter_aliases or ('', '')
    required_literals = ['Databricks.Catalogs(', source_mapping.target_catalog]
    if classification == 'Navigation':
        required_literals.extend([source_mapping.target_schema, source_mapping.target_item])
        database_nodes = _databricks_navigation_names_by_kind(scope, "Database")
        schema_nodes = _databricks_navigation_names_by_kind(scope, "Schema")
        table_nodes = (
            _databricks_navigation_names_by_kind(scope, "Table")
            + _databricks_navigation_names_by_kind(scope, "View")
        )
        if norm(source_mapping.target_catalog) not in {norm(x) for x in database_nodes}:
            errors.append('DATABRICKS_CONNECTION_POST_VALIDATION_FAILED: target catalog navigation missing')
        if source_mapping.target_schema and norm(source_mapping.target_schema) not in {norm(x) for x in schema_nodes}:
            errors.append('DATABRICKS_CONNECTION_POST_VALIDATION_FAILED: target schema navigation missing')
        if source_mapping.target_item and norm(source_mapping.target_item) not in {norm(x) for x in table_nodes}:
            errors.append('DATABRICKS_CONNECTION_POST_VALIDATION_FAILED: target item navigation missing')
    for token in required_literals:
        if token and token not in scope:
            errors.append(f'DATABRICKS_CONNECTION_POST_VALIDATION_FAILED: missing {token}')
    if not ((server_alias and re.search(_parameter_token(_m_identifier_name(server_alias) or server_alias), scope, re.I))
            or source_mapping.target_server in scope):
        errors.append(f'DATABRICKS_CONNECTION_POST_VALIDATION_FAILED: missing {source_mapping.target_server}')
    if not ((http_alias and re.search(_parameter_token(_m_identifier_name(http_alias) or http_alias), scope, re.I))
            or source_mapping.target_http_path in scope):
        errors.append(f'DATABRICKS_CONNECTION_POST_VALIDATION_FAILED: missing {source_mapping.target_http_path}')
    if re.search(r'(?i)\bSql\.Database\s*\(', scope):
        errors.append('DATABRICKS_CONNECTION_POST_VALIDATION_FAILED: Sql.Database remains in executable M')
    return sorted(set(errors))

def _unique_connection_source_for_legacy_database(server: str, database: str, sources):
    """Resolve one Databricks connection from all mappings for a legacy database.

    SQL translation replaces legacy physical tables before connector conversion.
    At that point table-level lookup may legitimately find no legacy identifiers.
    Connection conversion only needs a unique target server/http-path/catalog tuple.
    """
    candidates = [source for source in sources
                  if norm(source.legacy_server) == norm(server)
                  and norm(source.legacy_database) == norm(database)]
    groups = {}
    for source in candidates:
        key = (norm(source.target_server), norm(source.target_http_path), norm(source.target_catalog))
        groups.setdefault(key, source)
    return next(iter(groups.values())) if len(groups) == 1 else None

def _inventory_target_coordinates_for_query(log, query_name: str) -> tuple[str, str]:
    """Reuse authoritative Native SQL Source mappings already stored in inventory."""
    rows = [
        row for row in log.get("Query & Source Inventory", [])
        if clean(row.get("Query Name")) == clean(query_name)
        and clean(row.get("Object Type")) == "Native SQL Source"
        and clean(row.get("Data Source Mapping Found")) == "Yes"
    ]
    target_schemas = sorted(
        {clean(row.get("Target Schema")) for row in rows if clean(row.get("Target Schema"))},
        key=str.casefold,
    )
    target_items = sorted(
        {clean(row.get("Target Item")) for row in rows if clean(row.get("Target Item"))},
        key=str.casefold,
    )
    return "; ".join(target_schemas), "; ".join(target_items)

def rewrite_synapse_connection_to_databricks(value, query_name, input_file, sources, log, status, current_parameters=None):
    """Convert every applicable Synapse connector in a query, atomically and idempotently."""
    original = value
    discovered_payloads = discover_native_sql_payloads(value, current_parameters or {})
    regions = native_sql_regions(value)
    if discovered_payloads:
        connector = _sql_database_connector(value, current_parameters or {})
        server, database = connector.get("server", ""), connector.get("database", "")
        if not server:
            return value, []
        sql_span = _m_function_call_span(value, 'Sql.Database')
        if not sql_span:
            return value, ['Unbalanced or missing Sql.Database expression']
        resolved_sources = []
        resolution_errors = []
        for discovered in discovered_payloads:
            physical, _audit, parse_errors = _persistent_sql_sources(discovered.sql)
            resolution_errors.extend(parse_errors)
            for row in physical:
                mapped = _db_source_for(server, database, clean(row.get("Schema")), clean(row.get("Item")), sources)
                if mapped:
                    resolved_sources.append(mapped)
        connection_keys = {(norm(x.target_server), norm(x.target_http_path), norm(x.target_catalog)) for x in resolved_sources}
        source = (resolved_sources[0] if len(connection_keys) == 1 else None)
        if source is None:
            source = _unique_connection_source_for_legacy_database(server, database, sources)
        if len(connection_keys) > 1:
            resolution_errors.append("AMBIGUOUS_NATIVE_DATABRICKS_CONNECTION: physical SQL sources resolve to different target connections")
        if resolution_errors:
            return value, resolution_errors
        if not source:
            return value, []
        # Rebuild the Databricks connection from the authorized target mapping.
        # Never carry Sql.Database arguments into Databricks.Catalogs. In
        # particular, a previously updated Sql.Database(host, http_path) call must
        # not survive or influence the generated connector expression.
        server_alias, database_alias = _legacy_connection_parameter_aliases(
            original, current_parameters or {}, server, database
        )
        target_server_argument = server_alias or f'"{_m_escape(source.target_server)}"'
        target_http_path_argument = database_alias or f'"{_m_escape(source.target_http_path)}"'
        database_node = (
            f'{_db_catalog(source, target_server_argument, target_http_path_argument)}'
            f'{{[Name = "{_m_escape(source.target_catalog)}", Kind = "Database"]}}[Data]'
        )
        value = value[:sql_span[0]] + database_node + value[sql_span[1]:]
        validation_errors = _validate_databricks_connection_rewrite(
            value, source, server, database, 'Native SQL',
            (server_alias, database_alias))
        if validation_errors:
            return original, validation_errors
        # Audit only: reuse the authoritative Native SQL Source mappings that
        # inventory already resolved. Do not rediscover or rematch SQL objects here.
        target_schema_log, target_item_log = _inventory_target_coordinates_for_query(log, query_name)
        log['Connection Changes'].append({
            'Input File': input_file, 'Query Name': query_name, 'Classification': 'Native SQL',
            'Legacy Connector': 'Sql.Database', 'Target Connector': 'Databricks.Catalogs + Value.NativeQuery',
            'Legacy Server': server, 'Legacy Database': database,
            'Target Server': source.target_server, 'Target HTTP Path': source.target_http_path,
            'Target Catalog': source.target_catalog,
            'Target Schema': target_schema_log,
            'Target Item': target_item_log,
            'Result': status,
        })
        return value, []

    region, region_start, region_end, pairs = _synapse_navigation_pairs(value, current_parameters)
    if not pairs:
        return value, []
    resolved_pairs = []
    for pair in pairs:
        source_binding = pair['source']
        mapping = _db_source_for(
            source_binding['server'], source_binding['database'],
            pair['schema'], pair['item'], sources
        )
        if mapping:
            resolved_pairs.append((pair, mapping))
    if not resolved_pairs:
        return value, []

    # Generate globally unique, deterministic navigation names for every converted
    # physical source in the same query. Single-source queries retain Navigation 1-3.
    existing_names = {
        norm(binding['name_value'])
        for binding in _m_let_bindings_with_spans(value)[3]
    }
    replacements = []
    changes = []
    for index, (pair, mapping) in enumerate(resolved_pairs, 1):
        source_binding = pair['source']
        nav_binding = pair['navigation']
        target_schema = mapping.target_schema or pair['schema']
        target_item = mapping.target_item or pair['item']
        base = '' if len(resolved_pairs) == 1 else f' {index}'
        nav_names = [f'#"Navigation{base} {step}"' for step in (1, 2, 3)]
        # Avoid collisions with pre-existing user step names.
        if any(norm(_parse_query_identifier(name)) in existing_names for name in nav_names):
            base = f' Migrated {index}'
            nav_names = [f'#"Navigation{base} {step}"' for step in (1, 2, 3)]
        nav1, nav2, nav3 = nav_names
        source_prefix = source_binding['leading']
        nav_prefix = nav_binding['leading']
        source_indent = source_binding.get('indent', '')
        nav_indent = nav_binding.get('indent', '')
        server_arg = source_binding.get('server_alias') or f'"{_m_escape(mapping.target_server)}"'
        http_arg = source_binding.get('database_alias') or f'"{_m_escape(mapping.target_http_path)}"'
        schema_arg = pair.get('schema_alias') or f'"{_m_escape(target_schema)}"'
        item_arg = pair.get('item_alias') or f'"{_m_escape(target_item)}"'
        source_expr = _db_catalog(mapping, server_arg, http_arg)
        source_assignment = (
            f'{source_prefix}{source_binding["name"]} = {source_expr}'
            f'{source_binding["trailing"]}'
        )
        nav_assignment = (
            f'{nav_prefix}{nav1} = {source_binding["name"]}'
            f'{{[Name = "{_m_escape(mapping.target_catalog)}", Kind = "Database"]}}[Data],\n'
            f'{nav_indent}{nav2} = {nav1}'
            f'{{[Name = {schema_arg}, Kind = "Schema"]}}[Data],\n'
            f'{nav_indent}{nav3} = {nav2}'
            f'{{[Name = {item_arg}, Kind = "{_m_escape(TARGET_ITEM_KIND)}"]}}[Data]'
            f'{nav_binding["trailing"]}'
        )
        replacements.extend([
            (source_binding['start'], source_binding['end'], source_assignment),
            (nav_binding['start'], nav_binding['end'], nav_assignment),
        ])
        changes.append((pair, mapping, target_schema, target_item, nav3))

    changed_region = region
    # Syntax-span replacements are independent of line layout and multiline calls.
    for start_offset, end_offset, replacement in sorted(replacements, reverse=True):
        changed_region = changed_region[:start_offset] + replacement + changed_region[end_offset:]

    # Rebind downstream references after structural replacements. Parse the new
    # region each time so offsets stay correct across multiple converted sources.
    for pair, mapping, target_schema, target_item, nav3 in changes:
        old_name = pair['navigation']['name']
        _, _, _, new_bindings = _m_let_bindings_with_spans(
            value[:region_start] + changed_region + value[region_end:]
        )
        declaration = next(
            (binding for binding in new_bindings if binding['name'] == nav3), None
        )
        declaration_end = declaration['end'] if declaration else 0
        changed_region = _replace_m_identifier_after(
            changed_region, old_name, nav3, declaration_end
        )

    candidate = value[:region_start] + changed_region + value[region_end:]
    # Validate exact mapped objects only. Unmapped sources on the same server/database
    # remain untouched and do not block the file.
    residual = _remaining_mapped_legacy_connectors(candidate, sources, current_parameters)
    if residual:
        detail = '; '.join(
            f"M line {row['line']}: {row['schema']}.{row['item']} - {row['context']}"
            for row in residual
        )
        return original, ['MAPPED_LEGACY_CONNECTOR_REMAINS: ' + detail]
    for pair, mapping, target_schema, target_item, nav3 in changes:
        source_binding = pair['source']
        required = (
            'Databricks.Catalogs(',
            source_binding.get('server_alias') or mapping.target_server,
            source_binding.get('database_alias') or mapping.target_http_path,
            mapping.target_catalog,
            pair.get('schema_alias') or target_schema,
            pair.get('item_alias') or target_item,
            nav3,
        )
        if not all(token in candidate for token in required if token):
            return original, [
                'NAVIGATION_RENDER_FAILED: generated Databricks navigation contract is incomplete'
            ]

    for pair, mapping, target_schema, target_item, nav3 in changes:
        source_binding = pair['source']
        old_name = pair['navigation']['name']
        log['Connection Changes'].append({
            'Input File': input_file, 'Query Name': query_name, 'Classification': 'Navigation',
            'Legacy Connector': 'Sql.Database', 'Target Connector': 'Databricks.Catalogs',
            'Legacy Server': source_binding['server'], 'Legacy Database': source_binding['database'],
            'Target Server': mapping.target_server, 'Target HTTP Path': mapping.target_http_path,
            'Target Catalog': mapping.target_catalog, 'Target Schema': target_schema,
            'Target Item': target_item, 'Result': status,
        })
        log['M Query Changes'].append({
            'Input File': input_file, 'Query Name': query_name,
            'M Step': 'Source/Navigation',
            'Change Type': 'Parsed Synapse source and navigation conversion',
            'Old Value': (
                f'{source_binding["name"]}=Sql.Database; '
                f'{old_name}={pair["schema"]}.{pair["item"]}'
            ),
            'New Value': (
                f'{source_binding["name"]}=Databricks.Catalogs; '
                f'{nav3}={mapping.target_catalog}.{target_schema}.{target_item}'
            ),
            'Result': status,
        })
    return candidate, []

def _verify_written_databricks_connections(path: Path, original: str, sources: list[SourceMapping]):
    """Reopen a written TMDL file and verify applicable mappings migrated fully."""
    actual = read_text(path)
    server, database = _sql_db_values(original)
    if not server:
        return []
    navigation = _legacy_navigation_binding(original)
    schema = navigation.get("schema", "") if navigation else ""
    item = navigation.get("item", "") if navigation else ""
    source = _db_source_for(server, database, schema, item, sources)
    if native_sql_regions(original):
        source, native_resolution_errors = _native_connection_source(original, server, database, sources)
        if native_resolution_errors:
            return native_resolution_errors
    if source is None:
        return []
    classification = "Native SQL" if native_sql_regions(original) else "Navigation"
    return _validate_databricks_connection_rewrite(
        actual, source, server, database, classification
    )

# =============================================================================
# SOURCE -> SQL -> RESULT STORAGE CONTRACT
# =============================================================================
def _value_native_query_consumers(value: str):
    """Return every parsed Value.NativeQuery consumer and its SQL argument span."""
    consumers = []
    for ordinal, contract in enumerate(_native_query_contracts(value), 1):
        consumers.append({
            'ordinal': ordinal,
            'call_start': contract.consumer_start,
            'call_end': contract.consumer_end,
            'argument_start': contract.sql_argument_start,
            'argument_end': contract.sql_argument_end,
            'argument_expression': contract.sql_expression,
            'binding_name': contract.sql.name if contract.sql else '',
            'source_binding_name': contract.source.name if contract.source else '',
            'result_binding_name': contract.result.name if contract.result else '',
        })
    return consumers

def enforce_native_sql_storage_contract(original: str, migrated: str, current_parameters=None):
    """Reparse and validate exact-span native SQL writes without reconstructing them."""
    current_parameters = current_parameters or {}
    original_contracts = _native_query_contracts(original)
    migrated_contracts = _native_query_contracts(migrated)
    if not original_contracts:
        return migrated, [], []
    if len(original_contracts) != len(migrated_contracts):
        return migrated, [], [
            'NATIVE_SQL_STORAGE_CONTRACT_CONSUMER_COUNT_CHANGED: '
            f'{len(original_contracts)} -> {len(migrated_contracts)}'
        ]
    errors = validate_native_sql_storage_contract(original, migrated, current_parameters)
    return migrated, [], errors

def _native_sql_parameter_contract_errors(original: str, migrated: str,
                                          current_parameters: dict[str, str]):
    """Ensure original native-SQL parameter references survive generated M."""
    errors = []
    before = _native_query_contracts(original)
    after = _native_query_contracts(migrated)
    for ordinal, (left, right) in enumerate(zip(before, after), 1):
        left_expression = left.sql.expression if left.sql is not None else left.sql_expression
        right_expression = right.sql.expression if right.sql is not None else right.sql_expression
        expected = set(_native_parameter_names(left_expression, current_parameters))
        actual = set(_native_parameter_names(right_expression, current_parameters))
        missing = sorted(expected - actual, key=str.casefold)
        if missing:
            errors.append(
                f'NATIVE_SQL_PARAMETER_CONTRACT_CHANGED[{ordinal}]: missing=' + ', '.join(missing)
            )
    return errors

def validate_native_sql_storage_contract(original: str, migrated: str, current_parameters=None):
    """Validate reparsed target contracts after authoritative exact-span write-back."""
    current_parameters = current_parameters or {}
    errors = _native_sql_parameter_contract_errors(
        original, migrated, current_parameters
    )
    original_contracts = _native_query_contracts(original)
    migrated_contracts = _native_query_contracts(migrated)
    if len(original_contracts) != len(migrated_contracts):
        return ['NATIVE_SQL_STORAGE_CONTRACT_COUNT_CHANGED: '
                f'{len(original_contracts)} -> {len(migrated_contracts)}']
    for ordinal, contract in enumerate(migrated_contracts, 1):
        bindings = _contract_binding_directory(contract)
        resolved = _evaluate_static_m_text(contract.sql_expression, bindings, current_parameters)
        if resolved is None and contract.sql is not None:
            resolved = _evaluate_static_m_text(
                contract.sql.expression, bindings, current_parameters
            )
        if resolved is None:
            errors.append(
                f'POST_WRITE_NATIVE_SQL_RESOLUTION_FAILED[{ordinal}]: '
                f'binding={contract.sql_reference or "<inline>"}; '
                f'expression={contract.sql_expression}'
            )
            continue
        errors.extend(f'native SQL {ordinal}: {error}'
                      for error in validate_final_databricks_native_sql(resolved))
        residue = _active_tsql_batch_residue(resolved)
        if residue:
            errors.append(
                f'NATIVE_SQL_STORAGE_CONTRACT_LEGACY_RESIDUE[{ordinal}]: '
                + ', '.join(residue)
            )
    return sorted(set(errors))

# =============================================================================
# AUTHORITATIVE NATIVE SQL STORAGE REPAIR
# =============================================================================
def _active_tsql_batch_residue(sql_text: str) -> list[str]:
    """Return active legacy batch constructs after masking SQL comments."""
    active = _mask_sql_comments(sql_text) if "_mask_sql_comments" in globals() else sql_text
    rules = (
        (r"(?i)\bDROP\s+TABLE\b", "DROP TABLE"),
        (r"(?i)\bCREATE\s+(?:TEMP(?:ORARY)?\s+)?TABLE\b", "CREATE TABLE"),
        (r"(?is)\bSELECT\b.*?\bINTO\s+(?:\[\s*#|#)", "SELECT INTO #temp"),
        (r"(?i)(?:\[\s*#[A-Za-z_][A-Za-z0-9_]*\s*\]|#[A-Za-z_][A-Za-z0-9_]*)", "#temp reference"),
    )
    return [label for pattern, label in rules if re.search(pattern, active)]

def repair_inlined_native_sql_binding(value: str, current_parameters=None):
    """Compatibility repair for partial migrations lacking a usable lexical contract.

    Valid lexical contracts are authoritative and must never be modified by the
    orphan-binding fallback. The legacy repair path is intentionally skipped
    whenever every Value.NativeQuery SQL argument resolves in its local scope.
    """
    current_parameters = current_parameters or {}
    contracts = _native_query_contracts(value)
    if contracts:
        all_resolved = all(
            _evaluate_static_m_text(
                contract.sql_expression,
                _contract_binding_directory(contract),
                current_parameters,
            ) is not None
            for contract in contracts
        )
        if all_resolved:
            return value, [], []
    # A consumer exists but is not resolvable. Do not guess among nearby SQL-like
    # bindings in active migration. Preserve the input and report a focused error.
    if contracts:
        unresolved = [
            contract.sql_expression for contract in contracts
            if _evaluate_static_m_text(
                contract.sql_expression,
                _contract_binding_directory(contract),
                current_parameters,
            ) is None
        ]
        return value, [], [
            'NATIVE_SQL_PARTIAL_MIGRATION_REPAIR_REQUIRES_REVIEW: '
            + '; '.join(unresolved)
        ]
    return value, [], []

def validate_no_orphaned_legacy_sql_bindings(value: str, current_parameters=None):
    """Block stale legacy SQL bindings after native SQL migration or repair."""
    current_parameters = current_parameters or {}
    _region, _start, _end, bindings = _m_let_bindings_with_spans(value)
    if not bindings:
        return []
    expressions = {norm(row['name_value']): row['expression'] for row in bindings}
    errors = []
    for binding in bindings:
        resolved = _evaluate_static_m_text(binding['expression'], expressions, current_parameters)
        if resolved is None:
            continue
        residue = _active_tsql_batch_residue(resolved)
        if not residue:
            continue
        # Limit this guard to SQL-like text bindings in a query that has an active
        # native SQL consumer, avoiding unrelated ordinary text parameters.
        if not re.search(r'(?is)\b(?:SELECT|WITH|DROP|CREATE)\b', resolved):
            continue
        errors.append(
            f"ORPHANED_LEGACY_NATIVE_SQL_BINDING: {binding['name_value']} retains "
            f"{', '.join(residue)} after migration"
        )
    return sorted(set(errors))

# =============================================================================
# BINDING-AWARE NATIVE SQL MIGRATION
# =============================================================================

def immutable_legacy_parameter_snapshot(definitions):
    """Return an immutable snapshot of applied legacy parameter values.

    The snapshot is captured before expressions.tmdl is edited. Connector
    discovery, native-SQL materialization, mapping resolution, and validation
    must all use this snapshot so an earlier parameter migration cannot change
    the coordinates used to resolve a later expression.
    """
    values = {}
    for name, entries in (definitions or {}).items():
        current = next((entry.value for entry in entries if getattr(entry, "is_current", False)), None)
        if current is not None:
            values[name] = current
    return MappingProxyType(dict(values))

def _validate_expression_migration(original_block, migrated_block, query_name, mappings, sources,
                                   legacy_parameters, native_before):
    """Run expression-local, fail-closed validation after migration."""
    errors = []
    pseudo_path = Path(f"{query_name or 'expression'}.tmdl")
    errors.extend(validate_protected_content(pseudo_path, original_block, migrated_block))
    errors.extend(validate_tmdl_m_indentation(pseudo_path, migrated_block))
    errors.extend(validate_navigation_output_contract(migrated_block, mappings, legacy_parameters))
    if native_before:
        original_contracts = _native_query_contracts(original_block)
        migrated_contracts = _native_query_contracts(migrated_block)
        payloads = discover_native_sql_payloads(migrated_block, legacy_parameters)
        if len(original_contracts) != len(migrated_contracts):
            errors.append(
                f"EXPRESSION_NATIVE_SQL_CONTRACT_CHANGED: {query_name}; "
                f"{len(original_contracts)} -> {len(migrated_contracts)}"
            )
        if not payloads:
            errors.append(f"EXPRESSION_NATIVE_SQL_LOST: {query_name}")
        for payload_index, payload in enumerate(payloads, 1):
            errors.extend(validate_final_databricks_native_sql(payload.sql))
        errors.extend(validate_no_orphaned_legacy_sql_bindings(migrated_block, legacy_parameters))
        errors.extend(validate_native_sql_storage_contract(
            original_block, migrated_block, legacy_parameters
        ))
        for contract in migrated_contracts:
            source_expression = (
                contract.source.expression if contract.source is not None
                else contract.source_expression
            )
            if re.search(r'(?i)\bSql\.Database\s*\(', source_expression):
                errors.append(
                    f"EXPRESSION_LEGACY_CONNECTOR_REMAINS: {query_name}; "
                    f"source={contract.source_reference or source_expression[:80]}"
                )
    for residual in _remaining_mapped_legacy_connectors(migrated_block, sources, legacy_parameters):
        errors.append(
            f"EXPRESSION_MAPPED_LEGACY_CONNECTOR_REMAINS: {query_name}; "
            f"M line {residual['line']}: {residual['context']}"
        )
    return sorted(set(errors))

def migrate_expressions_per_expression(value, mappings, sources, log, status, legacy_parameters):
    """Migrate and validate every expressions.tmdl expression independently.

    This prevents one expression's updated parameters, connector, SQL payload,
    or navigation steps from affecting discovery or write coordinates in another
    expression. Parameter declarations are also expressions, but only expressions
    with an applicable connector/native SQL/navigation are changed.
    """
    rebuilt = []
    errors = []
    for query_name, original_block in split_expression_blocks(Path("expressions.tmdl"), value):
        block, repair_changes, repair_errors = repair_inlined_native_sql_binding(
            original_block, legacy_parameters
        )
        errors.extend(f"{query_name}: {error}" for error in repair_errors)
        for change in repair_changes:
            log["M Query Changes"].append({
                "Input File": "expressions.tmdl", "Query Name": query_name,
                "M Step": change["binding_name"],
                "Change Type": "Restore translated SQL to authoritative named binding",
                "Old Value": change["old_sql"], "New Value": change["new_sql"],
                "Result": f"{status}; removed active residue: {change['residue']}",
            })
        native_before = is_native_query(block, legacy_parameters)
        if native_before:
            candidate, sql_audit, sql_errors = replace_native_sql_payloads(
                block, mappings, sources, legacy_parameters
            )
            errors.extend(f"{query_name}: {error}" for error in sql_errors)
            if not sql_errors:
                block = candidate
                for record in sql_audit:
                    if record.get("kind") == "sql_payload":
                        log["Native SQL Changes"].append({
                            "Input File": "expressions.tmdl", "Query Name": query_name,
                            "Change Type": "Per-expression complete SQL before and after SQLGlot",
                            "Old Value": record.get("old_value", ""),
                            "New Value": record.get("new_value", ""),
                            "Occurrences": record.get("occurrences", 1),
                            "Safe Replacement": "Parse validated",
                            "Message": f"{status}; immutable legacy parameter snapshot",
                        })
            if not sql_errors:
                block, connector_errors = rewrite_synapse_connection_to_databricks(
                    block, query_name, "expressions.tmdl", sources, log, status,
                    legacy_parameters,
                )
                errors.extend(f"{query_name}: {error}" for error in connector_errors)
        else:
            block, connector_errors = rewrite_synapse_connection_to_databricks(
                block, query_name, "expressions.tmdl", sources, log, status,
                legacy_parameters,
            )
            errors.extend(f"{query_name}: {error}" for error in connector_errors)
            block = migrate_m_navigation_per_query(
                Path("expressions.tmdl"), block, mappings, sources, log, status
            )
        block, _repairs = repair_databricks_navigation_indentation(block)
        block, storage_changes, storage_errors = enforce_native_sql_storage_contract(
            original_block, block, legacy_parameters
        )
        errors.extend(f"{query_name}: {error}" for error in storage_errors)
        for change in storage_changes:
            log["M Query Changes"].append({
                "Input File": "expressions.tmdl", "Query Name": query_name,
                "M Step": change["binding_name"],
                "Change Type": "Preserve named native SQL binding",
                "Old Value": change["old_binding_expression"],
                "New Value": change["new_binding_expression"],
                "Result": f"{status}; Value.NativeQuery restored to {change['new_consumer_argument']}",
            })
        block, _hardening_changes, hardening_errors = harden_missing_field_operations(
            block, query_name, "expressions.tmdl", log, status
        )
        errors.extend(f"{query_name}: {error}" for error in hardening_errors)
        local_errors = _validate_expression_migration(
            original_block, block, query_name, mappings, sources,
            legacy_parameters, native_before,
        )
        errors.extend(f"{query_name}: {error}" for error in local_errors)
        # Keep the generated candidate in the proposed expressions.tmdl text even
        # when this expression has blocking validation errors. The containing file
        # remains blocked by ``errors``; retaining the candidate ensures subsequent
        # validation examines the proposed Databricks M/SQL instead of silently
        # reintroducing the original Synapse block.
        rebuilt.append(block)
    return "".join(rebuilt), sorted(set(errors))

def _current_parameter_values(definitions):
    """Return applied current values from discover_parameters() output."""
    values = {}
    for name, entries in (definitions or {}).items():
        current = next((entry.value for entry in entries if getattr(entry, "is_current", False)), None)
        if current is not None:
            values[name] = current
    return values

def _binding_expression_span(value: str, binding_name: str):
    """Compatibility wrapper over the authoritative immutable binding directory."""
    binding = _stable_m_binding_spans(value).get(norm(binding_name))
    if binding is None:
        return None
    return {
        "start": binding.expression_start,
        "end": binding.expression_end,
        "expression": binding.expression,
        "name": binding.name,
    }

def _binding_aware_regions(value: str, current_parameters=None):
    """Return writable literal SQL spans, including SQL stored in a referenced binding."""
    regions = list(_literal_native_sql_regions(value))
    seen = {(start, end) for start, end, _payload in regions}
    for payload in discover_native_sql_payloads(value, current_parameters or {}):
        if not payload.binding_name:
            continue
        binding = _binding_expression_span(value, payload.binding_name)
        if not binding:
            continue
        expression = binding["expression"].strip()
        literal = re.fullmatch(r'"((?:[^"]|"")*)"', expression, re.DOTALL)
        if not literal:
            continue
        leading = len(binding["expression"]) - len(binding["expression"].lstrip())
        start = binding["start"] + leading + 1
        end = start + len(literal.group(1))
        if (start, end) not in seen:
            regions.append((start, end, literal.group(1)))
            seen.add((start, end))
    return sorted(regions)

# Preserve the established literal parser and broaden the public contract.
_literal_native_sql_regions = native_sql_regions

def native_sql_regions(text: str):
    return _binding_aware_regions(text, {})

def is_native_query(value: str, current_parameters=None) -> bool:
    return bool(discover_native_sql_payloads(value, current_parameters or {}))

def _payload_parameter_expression(payload: DiscoveredNativeSql) -> str:
    """Return the original M expression that contains parameter references."""
    return clean(payload.parameter_expression) or clean(payload.argument_expression)

def _identifier_parameter_role(parameter_name: str, current_value: str, source_maps):
    """Classify existing identifier parameters using the mapping template only."""
    roles = set()
    for source in source_maps or []:
        if current_value and norm(current_value) == norm(source.legacy_schema):
            roles.add('schema')
        if current_value and norm(current_value) == norm(source.legacy_item):
            roles.add('item')
        qualified = '.'.join(part for part in (source.legacy_schema, source.legacy_item) if part)
        if current_value and qualified and norm(current_value) == norm(qualified):
            roles.add('schema_item')
    return next(iter(roles)) if len(roles) == 1 else ''

def _legacy_connection_parameter_aliases(value: str, current_parameters: dict[str, str],
                                         server: str, database: str):
    """Preserve existing parameter identifiers used by Sql.Database arguments."""
    for call in _balanced_m_calls(value, ('Sql.Database',)):
        args = split_m_top_level(call['body'])
        if len(args) < 2:
            continue
        aliases = []
        for expression, resolved_value in zip(args[:2], (server, database)):
            name = _m_identifier_name(expression)
            parameter_name = next(
                (candidate for candidate in current_parameters
                 if norm(candidate) == norm(name)), ''
            ) if name else ''
            if parameter_name and norm(current_parameters.get(parameter_name)) == norm(resolved_value):
                aliases.append(expression.strip())
            else:
                aliases.append('')
        return tuple(aliases)
    return '', ''

def _native_parameter_names(expression: str, current_parameters: dict[str, str]) -> list[str]:
    """Return report parameters referenced by one SQL-producing M expression."""
    return [
        name for name in sorted(current_parameters, key=len, reverse=True)
        if re.search(_parameter_token(name), expression, re.I)
    ]

def _analysis_sql_with_value_markers(payload: DiscoveredNativeSql,
                                     current_parameters: dict[str, str], source_maps=None):
    """Materialize identifiers but protect value/list parameters with SQL markers.

    Parameter discovery uses the original SQL-producing M expression. Identifier
    parameters stay materialized during SQLGlot analysis and are restored after
    physical source mapping. Value/list parameters receive durable numeric markers.
    """
    sql_text = payload.sql
    markers = {}
    expression = _payload_parameter_expression(payload)
    names = _native_parameter_names(expression, current_parameters)
    for ordinal, name in enumerate(names, 1):
        value = clean(current_parameters.get(name))
        if not value:
            continue
        if _identifier_parameter_role(name, value, source_maps):
            continue
        marker = str(910000000000000000 + ordinal)
        count = sql_text.count(value)
        if count == 0:
            return sql_text, {}, names, (
                f'NATIVE_SQL_PARAMETER_PROVENANCE_NOT_FOUND: {name} current value was not found in materialized SQL'
            )
        sql_text = sql_text.replace(value, marker)
        markers[marker] = name
    return sql_text, markers, names, ''

def _replace_target_table_with_parameter_marker(sql_text: str, parameter_name: str,
                                                current_value: str, source_maps) -> tuple[str, str]:
    """Restore an existing schema/item parameter without removing static qualifiers."""
    role = _identifier_parameter_role(parameter_name, current_value, source_maps)
    if not role:
        return sql_text, ''
    marker = f'__PBI_M_IDENTIFIER_{re.sub(r"[^A-Za-z0-9_]", "_", parameter_name)}__'
    candidates = set()
    for source in source_maps:
        legacy_qualified = '.'.join(part for part in (source.legacy_schema, source.legacy_item) if part)
        if role == 'item' and norm(source.legacy_item) == norm(current_value):
            candidates.add(clean(source.target_item) or current_value)
        elif role == 'schema' and norm(source.legacy_schema) == norm(current_value):
            candidates.add(clean(source.target_schema) or current_value)
        elif role == 'schema_item' and norm(legacy_qualified) == norm(current_value):
            target = '.'.join(part for part in (source.target_schema, source.target_item) if part)
            if target:
                candidates.add(target)
    output = sql_text
    total = 0
    for candidate in sorted((c for c in candidates if c), key=len, reverse=True):
        escaped_parts = [re.escape(part) for part in candidate.split('.')]
        plain = r'(?<![A-Za-z0-9_])' + r'\.'.join(escaped_parts) + r'(?![A-Za-z0-9_])'
        backticked = r'`' + r'`\.`'.join(escaped_parts) + r'`'
        for pattern in (re.compile(backticked, re.I), re.compile(plain, re.I)):
            output, count = pattern.subn(marker, output)
            total += count
    return (output, marker) if total else (sql_text, '')

def _render_parameterized_sql_expression(rewritten_sql: str, payload: DiscoveredNativeSql,
                                         current_parameters: dict[str, str], source_maps,
                                         value_markers: dict[str, str]):
    """Render final M concatenation while retaining all original parameters."""
    parameter_markers = dict(value_markers)
    sql_text = rewritten_sql
    expression = _payload_parameter_expression(payload)
    names = _native_parameter_names(expression, current_parameters)
    for name in names:
        value = clean(current_parameters.get(name))
        if not _identifier_parameter_role(name, value, source_maps):
            continue
        sql_text, marker = _replace_target_table_with_parameter_marker(
            sql_text, name, value, source_maps
        )
        if not marker:
            return '', f'NATIVE_SQL_IDENTIFIER_PARAMETER_RESTORATION_FAILED: {name}'
        parameter_markers[marker] = name
    if not parameter_markers:
        return '"' + encode_m_sql_string(sql_text, payload.original_encoded_payload or expression) + '"', ''
    marker_pattern = re.compile('|'.join(
        re.escape(marker) for marker in sorted(parameter_markers, key=len, reverse=True)
    ))
    pieces = []
    cursor = 0
    found = set()
    for match in marker_pattern.finditer(sql_text):
        literal = sql_text[cursor:match.start()]
        if literal:
            pieces.append('"' + encode_m_sql_string(
                literal, payload.original_encoded_payload or expression
            ) + '"')
        marker = match.group(0)
        pieces.append(parameter_markers[marker])
        found.add(marker)
        cursor = match.end()
    tail = sql_text[cursor:]
    if tail:
        pieces.append('"' + encode_m_sql_string(
            tail, payload.original_encoded_payload or expression
        ) + '"')
    missing = sorted(set(parameter_markers).difference(found))
    if missing:
        return '', 'NATIVE_SQL_PARAMETER_RESTORATION_FAILED: ' + ', '.join(parameter_markers[m] for m in missing)
    return ' & '.join(pieces), ''

def _materialize_rewritten_native_bindings(value, mappings, source_maps, current_parameters, databricks=True):
    """Rewrite concrete analysis SQL, then restore the final parameterized M binding."""
    payloads = discover_native_sql_payloads(value, current_parameters or {})
    if not payloads:
        return value, [], []
    unresolved = [
        f"NATIVE_SQL_STORAGE_LOCATION_UNRESOLVED: {payload.construct} payload {index}; "
        f"binding={payload.binding_name or '<inline>'}; argument={payload.argument_expression[:120]}"
        for index, payload in enumerate(payloads, 1)
        if payload.write_start < 0 or payload.write_end <= payload.write_start or not payload.write_mode
    ]
    if unresolved:
        return value, [], unresolved
    output, audit_rows, errors = value, [], []
    connector = _sql_database_connector(value, current_parameters or {})
    replacements, occupied = [], set()
    for payload_index, payload in enumerate(payloads, 1):
        analysis_sql, value_markers, parameter_names, provenance_error = _analysis_sql_with_value_markers(
            payload, current_parameters or {}, source_maps
        )
        if provenance_error:
            errors.append(provenance_error)
            continue
        try:
            if databricks:
                rewritten, audit, local_errors = rewrite_sql_ast(
                    analysis_sql, mappings, source_maps,
                    connector.get("server", ""), connector.get("database", ""),
                )
            else:
                rewritten, audit, local_errors = rewrite_sql_ast(analysis_sql, mappings, source_maps)
        except TypeError:
            rewritten, audit, local_errors = rewrite_sql_ast(analysis_sql, mappings, source_maps)
        audit_rows.extend(audit)
        audit_rows.append({
            "kind": "sql_payload", "payload_index": payload_index,
            "construct": payload.construct, "binding_name": payload.binding_name,
            "old_value": payload.sql, "new_value": rewritten,
            "occurrences": 1, "changed": rewritten != analysis_sql,
            "parameters": "; ".join(parameter_names),
            "source_dialect": "detected", "target_dialect": "databricks",
        })
        errors.extend(local_errors)
        if local_errors:
            continue
        if payload.write_start < 0 or payload.write_end < payload.write_start or not payload.write_mode:
            errors.append(f"NATIVE_SQL_WRITE_POSITION_NOT_FOUND: {payload.construct} payload {payload_index}")
            continue
        coordinate = (payload.write_start, payload.write_end)
        if coordinate in occupied:
            errors.append(f"NATIVE_SQL_WRITE_POSITION_DUPLICATE: {payload.construct} payload {payload_index}")
            continue
        occupied.add(coordinate)
        if payload.write_mode in {"binding_expression", "argument_expression"}:
            replacement, restore_error = _render_parameterized_sql_expression(
                rewritten, payload, current_parameters or {}, source_maps, value_markers
            )
            if restore_error:
                errors.append(restore_error)
                continue
            audit_rows.append({
                "kind": "parameterized_m_sql", "payload_index": payload_index,
                "construct": payload.construct, "binding_name": payload.binding_name,
                "old_value": payload.argument_expression, "new_value": replacement,
                "occurrences": 1, "changed": True,
                "parameters": "; ".join(parameter_names),
                "source_dialect": "M", "target_dialect": "M",
            })
        else:
            encoded = encode_m_sql_string(rewritten, payload.original_encoded_payload or payload.argument_expression)
            replacement = encoded
        replacements.append((payload.write_start, payload.write_end, replacement))
    for start, end, replacement in sorted(replacements, reverse=True):
        output = output[:start] + replacement + output[end:]
    return output, audit_rows, sorted(set(errors))

def validate_runtime_resolution_regressions(paths: Paths, sources: list[SourceMapping]):
    """Validate consumer discovery, SQL resolution, and writable coordinates."""
    errors = []
    if '_ambiguous_unqualified_column_candidates' not in globals():
        errors.append('SQL_AMBIGUITY_GUARD_REGRESSION: unqualified native SQL ambiguity guard is unavailable')
    for required in ('_analysis_sql_with_value_markers', '_render_parameterized_sql_expression'):
        if required not in globals():
            errors.append(f'FINAL_M_ARCHITECTURE_REGRESSION: {required} is unavailable')
    _defs, current_parameters = parameter_definitions(paths)
    inventory_rows, _unpivots = inventory(paths)
    consumers = resolved = writable = 0
    files = sorted(paths.tables.glob('*.tmdl')) + ([paths.expressions] if paths.expressions.exists() else [])
    for path in files:
        for query_name, block in split_expression_blocks(path, read_text(path)):
            contracts = _native_query_contracts(block)
            payloads = discover_native_sql_payloads(block, current_parameters)
            consumers += len(contracts)
            resolved += len(payloads)
            writable_payloads = [p for p in payloads
                                 if p.write_start >= 0 and p.write_end > p.write_start and p.write_mode]
            writable += len(writable_payloads)
            if contracts and len(payloads) < len(contracts):
                errors.append(
                    f'NATIVE_QUERY_SQL_RESOLUTION_REGRESSION: {path.name}/{query_name}; '
                    f'consumers={len(contracts)} resolved={len(payloads)}'
                )
            if payloads and len(writable_payloads) < len(payloads):
                errors.append(
                    f'NATIVE_QUERY_WRITE_TARGET_REGRESSION: {path.name}/{query_name}; '
                    f'resolved={len(payloads)} writable={len(writable_payloads)}'
                )
    if consumers and resolved != consumers:
        errors.append(
            'NATIVE_QUERY_COUNT_REGRESSION: '
            f'consumers={consumers}; SQL expressions resolved={resolved}; writable={writable}'
        )
    if sources:
        enriched = enrich_inventory_with_source_mappings([dict(row) for row in inventory_rows], sources)
        discovered_sql = [row for row in enriched if clean(row.get('Server')) and clean(row.get('Database'))]
        matched = [row for row in discovered_sql if row.get('Data Source Mapping Found') == 'Yes']
        if discovered_sql and not matched:
            errors.append(f'DATA_SOURCE_MAPPING_MATCH_REGRESSION: {len(sources)} mappings loaded but zero discovered SQL sources matched')
    return sorted(set(errors))


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
    target_metadata_rows, target_metadata_errors = validate_target_metadata(mappings, sources)
    log["Target Metadata Validation"].extend(target_metadata_rows)
    blocking_errors.extend(target_metadata_errors)

    parameter_rows, parameter_definitions, parameter_roles = discover_parameters(paths)
    legacy_parameters = immutable_legacy_parameter_snapshot(parameter_definitions)
    current_parameters = legacy_parameters
    for path in files:
        updated = originals[path]
        updated = migrate_current_source_parameters(updated, path, parameter_definitions, sources, log, status)
        if path.name.casefold() == "expressions.tmdl":
            updated, expression_errors = migrate_expressions_per_expression(
                updated, mappings, sources, log, status, legacy_parameters
            )
            for expression_error in expression_errors:
                blocking_errors.append(f"{path.name}: {expression_error}")
                log["Errors"].append({
                    "Input File": path.name,
                    "Stage": "Per-expression migration and validation",
                    "Error Code": "EXPRESSION_MIGRATION_VALIDATION_FAILED",
                    "Error Message": expression_error,
                    "Existing Content Preserved": "Yes",
                    "Processing Continued": "No",
                })
            proposed[path] = updated
            continue
        query_name = tmdl_table_name(updated, path.stem)
        updated, repair_changes, repair_errors = repair_inlined_native_sql_binding(
            updated, current_parameters
        )
        for repair_error in repair_errors:
            blocking_errors.append(f"{path.name}: {repair_error}")
            log["Errors"].append({
                "Input File": path.name, "Stage": "Native SQL storage repair",
                "Error Code": "NATIVE_SQL_STORAGE_REPAIR_FAILED",
                "Error Message": repair_error, "Existing Content Preserved": "Yes",
                "Processing Continued": "No",
            })
        for change in repair_changes:
            log["M Query Changes"].append({
                "Input File": path.name, "Query Name": query_name,
                "M Step": change["binding_name"],
                "Change Type": "Restore translated SQL to authoritative named binding",
                "Old Value": change["old_sql"], "New Value": change["new_sql"],
                "Result": f"{status}; removed active residue: {change['residue']}",
            })
        native = is_native_query(updated, current_parameters)
        affected = set()

        # Regular Schema/Item navigation must be converted while the original
        # Synapse coordinates are still present. The generic navigation mapper
        # can change Item/Schema first, which prevents the connector converter
        # from resolving the legacy Data Source Mapping and can leave
        # Sql.Database behind. Native SQL remains on the later path because its
        # SQL payload must be rewritten before it is wrapped in Value.NativeQuery.
        early_connection_errors = []
        if not native:
            updated, join_key_changes, join_key_errors = migrate_join_keys_with_lineage(
                updated, query_name, original_queries, mappings, query_binding_cache, current_parameters)
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
                updated, query_name, original_queries, mappings, query_binding_cache, current_parameters)
            for change in append_changes:
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Table.Combine", "Change Type": "Append-local output contract reconciliation",
                    "Old Value": ", ".join(current for current, _legacy in change["pairs"]),
                    "New Value": ", ".join(legacy for _current, legacy in change["pairs"]),
                    "Result": f"{status}; branch={change['branch']}",
                })
            for error in append_errors:
                blocking_errors.append(f"{path.name}: {error}")
            updated, early_connection_errors = rewrite_synapse_connection_to_databricks(
                updated, query_name, path.name, sources, log, status, current_parameters
            )
            for connection_error in early_connection_errors:
                blocking_errors.append(f"{path.name}: {connection_error}")
                log["Errors"].append({
                    "Input File": path.name,
                    "Stage": "Synapse-to-Databricks connector conversion",
                    "Error Code": "DATABRICKS_CONNECTION_REWRITE_FAILED",
                    "Error Message": connection_error,
                    "Existing Content Preserved": "Yes",
                    "Processing Continued": "No",
                })

        # Apply table/column M navigation changes only after connector conversion.
        # Generated Databricks Name/Kind navigation does not match the legacy
        # Schema/Item pattern, so it is left intact.
        updated = migrate_m_navigation_per_query(
            path, updated, mappings, sources, log, status
        )

        # Source/navigation changes stay inside M. Mask native SQL first so
        # schema/item replacements cannot alter SQL identifiers or comments.
        early_masks = []
        if is_native_query(updated, current_parameters):
            updated, early_masks = mask_native_sql_payloads(updated)
        region, region_start, region_end = extract_m_region(updated)
        # Synapse-to-Databricks connector/navigation replacement is handled by
        # rewrite_synapse_connection_to_databricks(). Do not perform generic
        # source-field string replacement here because SourceMapping intentionally
        # contains only server/database/schema/item/catalog connection metadata.
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
            candidate, sql_audit, sql_errors = replace_native_sql_payloads(updated, mappings, sources, current_parameters)
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
                        "Stage": "SQLGlot Synapse-to-Databricks AST validation",
                        "Error Code": "NATIVE_SQL_AST_REWRITE_FAILED",
                        "Error Message": error,
                        "Existing Content Preserved": "Yes",
                        "Processing Continued": "No",
                    })
            for record in sql_audit:
                if record["kind"] == "sql_payload":
                    log["Native SQL Changes"].append({
                        "Input File": path.name, "Query Name": query_name,
                        "Change Type": "Complete SQL before and after SQLGlot",
                        "Old Value": record["old_value"], "New Value": record["new_value"],
                        "Occurrences": record["occurrences"],
                        "Safe Replacement": "Parse validated" if not sql_errors else "No",
                        "Message": (
                            f"{status}; payload={record['payload_index']}; "
                            f"construct={record['construct']}; binding={record['binding_name'] or '<direct literal>'}; "
                            f"before SHA-256={sha(record['old_value'])}; after SHA-256={sha(record['new_value'])}"
                        ),
                    })
                if record["kind"] == "parameterized_m_sql":
                    log["Native SQL Changes"].append({
                        "Input File": path.name, "Query Name": query_name,
                        "Change Type": "Final parameterized M SQL binding",
                        "Old Value": record["old_value"], "New Value": record["new_value"],
                        "Occurrences": 1, "Safe Replacement": "Yes",
                        "Message": f"{status}; binding={record.get('binding_name','')}; parameters={record.get('parameters','')}",
                    })
                elif record["kind"] == "function":
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
                        "Result": (
                            f"{status}; {record.get('alias_preservation', 'Alias preserved')}; "
                            f"source={record.get('mapping_source', 'Table and Column Mapping')}"
                        ),
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

            sql_payload_audit = {
                int(record.get("payload_index", 0)): record
                for record in sql_audit
                if record.get("kind") == "sql_payload"
            }
            final_payloads = discover_native_sql_payloads(candidate, current_parameters)
            parameterized_payload_audit = {
                int(record.get("payload_index", 0)): record
                for record in sql_audit
                if record.get("kind") == "parameterized_m_sql" and int(record.get("payload_index", 0)) > 0
            }
            for payload_index, final_payload in enumerate(final_payloads, 1):
                final_sql = final_payload.sql
                original_sql = clean(sql_payload_audit.get(payload_index, {}).get("old_value"))
                parameterized_record = parameterized_payload_audit.get(payload_index, {})
                parameterized_value = clean(parameterized_record.get("new_value")) if clean(parameterized_record.get("parameters")) else ""
                log["Native SQL Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "Change Type": "Final migrated SQL",
                    "Old Value": original_sql or "Original SQL payload unavailable",
                    "New Value": final_sql,
                    "New Value (With Parameters)": parameterized_value,
                    "Occurrences": 1, "Safe Replacement": "Parse validated",
                    "Message": f"{status}; payload={payload_index}; SHA-256={sha(final_sql)}",
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
            inherited_bindings, inherited_errors = resolved_query_output_bindings(
                query_name, original_queries, mappings, query_binding_cache, None, current_parameters)
            query_output_bindings.update(inherited_bindings)
            for inherited_error in inherited_errors:
                blocking_errors.append(f"{path.name}: {inherited_error}")
            updated, expand_changes, expand_bindings, expand_errors = migrate_join_expansions(
                updated, original_queries, mappings, query_binding_cache, current_parameters
            )
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
            query_output_bindings.update(_merged_query_output_bindings(
                updated, original_queries, mappings, query_binding_cache, current_parameters
            ))
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

            # Non-native joins, references, and appends are handled once by
            # branch-scoped recursive lineage. Global mapping-by-mapping replacement
            # can leak a left-branch mapping into the right branch.
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

            candidate, count = replace_non_native_m_column(updated, mapping.legacy_column, mapping.target_column)
            if count:
                updated, total_m = candidate, total_m + count
                log["M Query Changes"].append({
                    "Input File": path.name, "Query Name": query_name,
                    "M Step": "Existing source-facing transformations", "Change Type": "Mapped column references",
                    "Old Value": mapping.legacy_column, "New Value": mapping.target_column,
                    "Result": f"{status}; {count} occurrence(s)",
                })

            # Preserve the semantic/model column name while binding it to the
            # mapped physical output of this direct-navigation query.
            source_count = 0
            semantic_columns = []
            if not native and path.name.casefold() != "expressions.tmdl":
                updated, source_count, semantic_columns = replace_tmdl_source_bindings(
                    updated, mapping.legacy_column, mapping.target_column
                )
                if source_count:
                    for semantic_column in semantic_columns:
                        log["Column Changes"].append({
                            "Input File": path.name, "Model Table": query_name,
                            "Model Column": semantic_column,
                            "Old Source Column": mapping.legacy_column,
                            "New Source Column": mapping.target_column,
                            "Changed In": "TMDL sourceColumn from direct-navigation mapping",
                            "Model Name Changed": "No", "Result": status,
                        })

            if total_m or source_count:
                affected.add((mapping.legacy_table, mapping.legacy_column))

        # After upstream non-native queries adopt target column names,
        # rewrite Table.AddColumn row-context references only from resolved
        # left/right lineage. Do not let an unexpanded right-side mapping bleed
        # into the current query's row schema.
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
        # Native SQL connection conversion remains here, after SQLGlot has
        # produced the final Databricks SQL payload. Regular navigation queries
        # were converted before generic M navigation processing above.
        connection_errors = []
        if native:
            updated, connection_errors = rewrite_synapse_connection_to_databricks(
                updated, query_name, path.name, sources, log, status, current_parameters
            )
        updated, indentation_repairs = repair_databricks_navigation_indentation(updated)
        if indentation_repairs:
            log["M Query Changes"].append({
                "Input File": path.name,
                "Query Name": query_name,
                "M Step": "Databricks navigation",
                "Change Type": "TMDL-safe binding indentation repair",
                "Old Value": "Navigation binding indentation differed from Source",
                "New Value": "Navigation 1-3 aligned to Source binding indentation",
                "Result": f"{status}; {indentation_repairs} line(s) repaired",
            })
        for connection_error in connection_errors:
            blocking_errors.append(f"{path.name}: {connection_error}")
            log["Errors"].append({
                "Input File": path.name, "Stage": "Synapse-to-Databricks connector conversion",
                "Error Code": "DATABRICKS_CONNECTION_REWRITE_FAILED",
                "Error Message": connection_error, "Existing Content Preserved": "Yes",
                "Processing Continued": "No",
            })
        updated, _hardening_changes, hardening_errors = harden_missing_field_operations(
            updated, query_name, path.name, log, status
        )
        for hardening_error in hardening_errors:
            blocking_errors.append(f"{path.name}: {hardening_error}")
            log["Errors"].append({
                "Input File": path.name, "Stage": "M missing-field hardening",
                "Error Code": "M_MISSING_FIELD_HARDENING_FAILED",
                "Error Message": hardening_error,
                "Existing Content Preserved": "Yes", "Processing Continued": "No",
            })
        updated, storage_changes, storage_errors = enforce_native_sql_storage_contract(
            originals[path], updated, current_parameters
        )
        for storage_error in storage_errors:
            blocking_errors.append(f"{path.name}: {storage_error}")
            log["Errors"].append({
                "Input File": path.name, "Stage": "Native SQL storage contract",
                "Error Code": "NATIVE_SQL_STORAGE_CONTRACT_FAILED",
                "Error Message": storage_error, "Existing Content Preserved": "Yes",
                "Processing Continued": "No",
            })
        for change in storage_changes:
            log["M Query Changes"].append({
                "Input File": path.name, "Query Name": query_name,
                "M Step": change["binding_name"],
                "Change Type": "Preserve named native SQL binding",
                "Old Value": change["old_binding_expression"],
                "New Value": change["new_binding_expression"],
                "Result": f"{status}; Value.NativeQuery restored to {change['new_consumer_argument']}",
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

    per_file_errors = {path: [] for path in all_files}
    mapping_convergence_errors = validate_mapping_convergence(mappings)
    if mapping_convergence_errors:
        for path in files:
            per_file_errors[path].extend(mapping_convergence_errors)
    for error in blocking_errors:
        matched = False
        for path in files:
            prefix = path.name + ":"
            if error.startswith(prefix):
                per_file_errors[path].append(error[len(prefix):].strip())
                matched = True
                break
        if not matched:
            for path in files:
                per_file_errors[path].append(error)

    unpivot_rows_by_file = {}
    for row in log["Unpivot Inventory"]:
        unpivot_rows_by_file.setdefault(clean(row.get("Input File")), []).append(row)
    for path in files:
        per_file_errors[path].extend(validate_unpivot_inventory(proposed[path], unpivot_rows_by_file.get(path.name, [])))
        tbl_key = norm(tmdl_table_name(originals[path], path.stem))
        allowed_p3 = (p3_rename_plan.get(tbl_key, {}) if (ALIGN_TARGET_COLUMN_NAMES and RENAME_SEMANTIC_COLUMNS and p3_rename_plan) else None)
        per_file_errors[path].extend(validate_protected_content(path, originals[path], proposed[path], allowed_p3))
        per_file_errors[path].extend(validate_mapped_legacy_navigation(path, proposed[path], mappings))
        per_file_errors[path].extend(validate_navigation_output_contract(proposed[path], mappings, current_parameters))
        per_file_errors[path].extend(validate_tmdl_m_indentation(path, proposed[path]))
        if "__PBI_" in proposed[path]:
            per_file_errors[path].append("INTERNAL_PLACEHOLDER_REMAINS: generated TMDL contains an unresolved __PBI_ marker")
        for residual in _remaining_mapped_legacy_connectors(proposed[path], sources, current_parameters):
            per_file_errors[path].append(
                f"MAPPED_LEGACY_CONNECTOR_REMAINS: M line {residual['line']}: {residual['context']}"
            )
        per_file_errors[path].extend(validate_m_partition_indentation(path, proposed[path]))
        per_file_errors[path].extend(validate_fragile_m_column_operations(path, proposed[path]))
        per_file_errors[path].extend(validate_no_orphaned_legacy_sql_bindings(proposed[path], current_parameters))
        per_file_errors[path].extend(validate_native_sql_storage_contract(
            originals[path], proposed[path], current_parameters
        ))
        final_native_payloads = discover_native_sql_payloads(
            proposed[path], current_parameters
        )
        if final_native_payloads:
            original_server, original_database = _sql_db_values(originals[path])
            for payload_index, payload in enumerate(final_native_payloads, 1):
                try:
                    payload_tree = sqlglot.parse_one(
                        payload.sql, read=TARGET_SQL_DIALECT
                    )
                    per_file_errors[path].extend(
                        _remaining_legacy_source_map_tables(
                            payload_tree, sources, original_server, original_database
                        )
                    )
                except ParseError as exc:
                    per_file_errors[path].append(
                        f"Post-migration native SQL payload {payload_index} could not be validated: "
                        f"{sanitize_diagnostic_text(exc)}"
                    )

    relationship_errors = validate_relationships(
        paths,
        {path: proposed[path] for path in files if path.parent == paths.tables},
        {path: proposed[path] for path in p3_proposed if path in paths.relationships},
    )
    if relationship_errors:
        for relationship_error in relationship_errors:
            if not any(row.get("Input File") == "Semantic Model" and row.get("Error Message") == relationship_error for row in log["Errors"]):
                log["Errors"].append({
                    "Input File": "Semantic Model", "Stage": "Relationship validation",
                    "Error Code": "RELATIONSHIP_ENDPOINT_INVALID", "Error Message": relationship_error,
                    "Existing Content Preserved": "Yes", "Processing Continued": "No",
                })
        for path in files:
            if path.parent == paths.tables:
                per_file_errors[path].append("RELATIONSHIP_VALIDATION_BLOCK: see Semantic Model error rows")

    for path, errors in per_file_errors.items():
        for error in sorted(set(errors)):
            if error.startswith("RELATIONSHIP_VALIDATION_BLOCK:"):
                continue
            if not any(row.get("Input File") == path.name and row.get("Error Message") == error for row in log["Errors"]):
                log["Errors"].append({
                    "Input File": path.name, "Stage": "Pre-write validation",
                    "Error Code": "MODEL_INTEGRITY_VALIDATION_FAILED", "Error Message": error,
                    "Existing Content Preserved": "Yes", "Processing Continued": "No",
                })

    write_counts = {"proposed":0,"written":0,"verified":0,"blocked":0,"failed":0,"unchanged":0}
    for path in all_files:
        changed = proposed[path] != originals[path]
        query_name = tmdl_table_name(originals[path], path.stem)
        graph_node = graph.get(query_name.casefold())
        errors = sorted(set(per_file_errors[path]))
        write_status, action, message, changed_status = "Not Written", "Unchanged", "No applicable changes.", "No"
        generated_sha = sha(proposed[path]) if changed else ""
        if errors:
            write_counts["blocked"] += 1
            if changed:
                write_counts["proposed"] += 1
            changed_status = "Blocked"
            action, write_status, message = "Blocked", "Blocked by validation errors", "; ".join(errors)
        elif changed:
            write_counts["proposed"] += 1
            changed_status, action, write_status, message = "Proposed", "Would Modify", "Preview Only", "Synapse-to-Databricks migration previewed."
            if write_files:
                try:
                    ok, write_error = verified_atomic_write(path, proposed[path])
                except Exception as exc:
                    ok, write_error = False, sanitize_diagnostic_text(exc)
                if ok:
                    semantic_write_errors = _verify_written_databricks_connections(
                        path, originals[path], sources
                    )
                    if semantic_write_errors:
                        ok = False
                        write_error = "; ".join(semantic_write_errors)
                if ok:
                    changed_status, action, write_status = "Yes", "Modified", "Success - Verified"
                    message = "Synapse-to-Databricks migration applied and verified."
                    write_counts["written"] += 1; write_counts["verified"] += 1
                else:
                    write_counts["failed"] += 1
                    action, write_status, message = "Write Failed", "Write verification failed", write_error
                    log["Errors"].append({
                        "Input File": path.name, "Stage": "PBIP file write",
                        "Error Code": "PBIP_WRITE_VERIFICATION_FAILED", "Error Message": write_error,
                        "Existing Content Preserved": "No automatic backup created", "Processing Continued": "Yes",
                    })
        else:
            write_counts["unchanged"] += 1
        log["Files Processed"].append({
            "Run Timestamp": datetime.now().isoformat(timespec="seconds"), "Input File": path.name,
            "Input Path": str(path), "File Type": "TMDL",
            "Classification": classify_file_with_lineage(originals[path], graph_node),
            "Relevant Mapping Found": "Yes" if (changed or errors) else "No", "Changed": changed_status,
            "Action": action, "Write Status": write_status, "Message": message,
            "Input SHA-256": sha(originals[path]), "Generated SHA-256": generated_sha,
        })
    return write_counts

ANSI_ESCAPE_RE = re.compile(
    r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])"
)
EXCEL_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
EXCEL_CELL_TEXT_LIMIT = 32767

def sanitize_diagnostic_text(value) -> str:
    """Remove terminal formatting/control characters from diagnostic text."""
    if value is None:
        return ""
    text_value = str(value)
    text_value = ANSI_ESCAPE_RE.sub("", text_value)
    text_value = EXCEL_ILLEGAL_XML_RE.sub("", text_value)
    return text_value

def excel_safe_value(value):
    """Return an openpyxl-safe scalar without losing useful diagnostics."""
    if value is None or isinstance(value, (bool, int, float, datetime)):
        return value
    if isinstance(value, Path):
        value = str(value)
    elif not isinstance(value, str):
        value = str(value)
    value = sanitize_diagnostic_text(value)
    if len(value) > EXCEL_CELL_TEXT_LIMIT:
        suffix = "\n[Truncated for Excel cell limit]"
        value = value[: EXCEL_CELL_TEXT_LIMIT - len(suffix)] + suffix
    return value

def excel_safe_row(values):
    return [excel_safe_value(value) for value in values]

# =============================================================================
# RESILIENT M COLUMN OPERATIONS
# These rewrites are intentionally narrow. They add MissingField behavior only to
# Power Query functions that natively support it, preserve existing arguments,
# never modify native SQL payloads, and fail closed on invalid Table.DropColumns.
# =============================================================================
M_MISSING_FIELD_POLICIES = {
    "table.removecolumns": "MissingField.Ignore",
    "table.selectcolumns": "MissingField.UseNull",
    "table.renamecolumns": "MissingField.Ignore",
    "table.reordercolumns": "MissingField.Ignore",
}

def _m_function_has_missing_field_argument(function_name: str, args: list[str]) -> bool:
    """Return True when a supported function already specifies MissingField behavior."""
    lowered = function_name.casefold()
    if any(re.search(r"(?i)\bMissingField\.(?:Error|Ignore|UseNull)\b", arg) for arg in args):
        return True
    expected = {
        "table.removecolumns": 3,
        "table.selectcolumns": 3,
        "table.renamecolumns": 3,
        "table.reordercolumns": 3,
    }.get(lowered, 0)
    return bool(expected and len(args) >= expected)

def harden_missing_field_operations(value: str, query_name: str, input_file: str, log, status: str):
    """Add safe MissingField policies to static M column-list operations.

    Remove/Rename/Reorder use MissingField.Ignore because their requested action is
    destructive or positional. SelectColumns uses MissingField.UseNull to preserve
    the declared output contract when a mapped source column is temporarily absent.
    Native SQL is masked so SQL text/comments are never changed by this M pass.
    """
    masked, sql_payloads = mask_native_sql_payloads(value)
    region, start, end = extract_m_region(masked)
    if start < 0:
        restored, restore_errors = restore_native_sql_payloads(masked, sql_payloads)
        return restored, [], restore_errors
    functions = tuple(name for name in (
        "Table.RemoveColumns", "Table.SelectColumns",
        "Table.RenameColumns", "Table.ReorderColumns",
    ))
    changed_region = region
    changes = []
    errors = []
    for call in _balanced_m_calls(changed_region, functions):
        args = split_m_top_level(call["body"])
        function_name = call["function"]
        if len(args) < 2:
            errors.append(
                f"M_COLUMN_OPERATION_MALFORMED: {function_name} in {query_name} has fewer than two arguments"
            )
            continue
        if _m_function_has_missing_field_argument(function_name, args):
            continue
        policy = M_MISSING_FIELD_POLICIES[function_name.casefold()]
        args.append(policy)
        rewritten = function_name + "(" + ", ".join(args) + ")"
        changed_region = (
            changed_region[:call["start"]] + rewritten + changed_region[call["end"]:]
        )
        changes.append({
            "Input File": input_file,
            "Query Name": query_name,
            "M Step": function_name,
            "Change Type": "Missing-field resilience",
            "Old Value": call["function"] + "(" + call["body"] + ")",
            "New Value": rewritten,
            "Result": f"{status}; added {policy}",
        })
    updated = masked[:start] + changed_region + masked[end:]
    updated, restore_errors = restore_native_sql_payloads(updated, sql_payloads)
    errors.extend(restore_errors)
    for change in changes:
        log["M Query Changes"].append(change)
    return updated, changes, sorted(set(errors))

def validate_fragile_m_column_operations(path: Path, value: str):
    """Detect column operations that cannot be made safe without runtime schema evidence."""
    region, _, _ = extract_m_region(value)
    if not region:
        return []
    errors = []
    if re.search(r"(?i)\bTable\.DropColumns\s*\(", region):
        errors.append(
            "INVALID_POWER_QUERY_FUNCTION: Table.DropColumns is not a valid M function; use Table.RemoveColumns"
        )
    # Expand/Duplicate have no MissingField option. Dynamic rewrites could silently
    # alter schema, so retain them and require explicit review only when malformed.
    for function_name, minimum_args in (
        ("Table.ExpandTableColumn", 3),
        ("Table.ExpandRecordColumn", 3),
        ("Table.DuplicateColumn", 3),
        ("Table.Unpivot", 4),
        ("Table.UnpivotOtherColumns", 4),
    ):
        for call in _balanced_m_calls(region, (function_name,)):
            if len(split_m_top_level(call["body"])) < minimum_args:
                errors.append(
                    f"M_COLUMN_OPERATION_MALFORMED: {function_name} in {path.name} has fewer than {minimum_args} arguments"
                )
    return sorted(set(errors))

# =============================================================================
# FAIL-CLOSED DATABRICKS SQL VALIDATION AND AUDIT ENRICHMENT
# Wraps SQL AST rewriting with target-dialect validation and audit logging
# while preserving core lineage, mapping, connector, M, and semantic-model
# protection behavior.
# =============================================================================
_core_rewrite_sql_ast = rewrite_sql_ast

def _sql_sha256(value: str) -> str:
    return hashlib.sha256((value or "").encode("utf-8")).hexdigest()

def _mask_sql_comments(sql_text: str) -> str:
    """Mask -- and /* */ SQL comments without masking quoted literals.

    The masked text keeps the same length and newline positions so residue
    diagnostics remain stable. Comment markers inside single-quoted strings,
    double-quoted identifiers/strings, bracketed identifiers, and backtick
    identifiers are treated as data rather than comment delimiters.
    """
    chars = list(sql_text)
    index = 0
    state = "code"
    block_depth = 0
    while index < len(sql_text):
        char = sql_text[index]
        nxt = sql_text[index + 1] if index + 1 < len(sql_text) else ""
        if state == "line_comment":
            if char in "\r\n":
                state = "code"
            else:
                chars[index] = " "
            index += 1
            continue
        if state == "block_comment":
            if char == "/" and nxt == "*":
                chars[index] = chars[index + 1] = " "
                block_depth += 1
                index += 2
                continue
            if char == "*" and nxt == "/":
                chars[index] = chars[index + 1] = " "
                block_depth -= 1
                index += 2
                if block_depth <= 0:
                    state = "code"
                continue
            if char not in "\r\n":
                chars[index] = " "
            index += 1
            continue
        if state in {"single", "double", "bracket", "backtick"}:
            closing = {"single": "'", "double": '"', "bracket": "]", "backtick": "`"}[state]
            if char == closing:
                if nxt == closing:
                    index += 2
                    continue
                state = "code"
            index += 1
            continue
        if char == "-" and nxt == "-":
            chars[index] = chars[index + 1] = " "
            state = "line_comment"
            index += 2
            continue
        if char == "/" and nxt == "*":
            chars[index] = chars[index + 1] = " "
            state = "block_comment"
            block_depth = 1
            index += 2
            continue
        if char == "'":
            state = "single"
        elif char == '"':
            state = "double"
        elif char == "[":
            state = "bracket"
        elif char == "`":
            state = "backtick"
        index += 1
    return "".join(chars)

def _validate_databricks_sql_payload(original_sql: str, generated_sql: str):
    """Fail closed unless generated SQL parses as Databricks SQL.

    The core AST transformation remains authoritative. This validation adds a
    second independent parse of every generated statement, checks for common
    T-SQL-only residue, and records immutable old/new hashes for auditability.
    """
    errors = []
    details = []
    if not clean(generated_sql):
        errors.append("Generated Databricks SQL is empty")
        return errors, details
    try:
        statements = sqlglot.parse(generated_sql, read=TARGET_SQL_DIALECT)
        statements = [statement for statement in statements if statement is not None]
        if not statements:
            errors.append("Generated Databricks SQL contains no parseable statements")
        else:
            details.append(f"Databricks parse validated for {len(statements)} statement(s)")
            # Round-trip each statement independently so a later statement cannot
            # hide an unsupported construct in a multi-statement payload.
            for ordinal, statement in enumerate(statements, 1):
                rendered = statement.sql(dialect=TARGET_SQL_DIALECT, pretty=True)
                sqlglot.parse_one(rendered, read=TARGET_SQL_DIALECT)
                details.append(f"Statement {ordinal} round-trip validated")
    except (ParseError, UnsupportedError, ValueError) as exc:
        errors.append(
            "Generated Databricks SQL failed target-dialect validation: "
            + sanitize_diagnostic_text(exc)
        )
    residue_rules = (
        (r"(?i)\bGETDATE\s*\(", "GETDATE()"),
        (r"(?i)\bISNULL\s*\(", "ISNULL()"),
        (r"(?i)\bTRY_CONVERT\s*\(", "TRY_CONVERT()"),
        (r"(?i)\bCONVERT\s*\(", "CONVERT()"),
        (r"(?i)\bNVARCHAR\b", "NVARCHAR"),
        (r"(?i)\bDATETIME2\b", "DATETIME2"),
        (r"(?i)\bUNIQUEIDENTIFIER\b", "UNIQUEIDENTIFIER"),
        (r"(?i)\bTOP\s*(?:\(|\d)", "TOP"),
        (r"(?i)\bCROSS\s+APPLY\b|\bOUTER\s+APPLY\b", "APPLY"),
        (r"(?is)\{\s*fn\s+", "ODBC escape function"),
    )
    # Residue validation must inspect executable SQL only. SQLGlot preserves
    # comments, and legacy T-SQL examples such as GETDATE() may legitimately
    # remain inside -- or /* */ comments. Mask those spans before regex checks
    # while preserving length/newlines for stable diagnostics.
    active_sql = _mask_sql_comments(generated_sql)
    for pattern, label in residue_rules:
        if re.search(pattern, active_sql):
            errors.append(f"T-SQL-only residue remained after translation: {label}")
    if generated_sql == original_sql:
        details.append("SQL text unchanged after translation; target parse still validated")
    details.extend([
        f"Original SHA-256={_sql_sha256(original_sql)}",
        f"Generated SHA-256={_sql_sha256(generated_sql)}",
    ])
    return sorted(set(errors)), details

def rewrite_sql_ast(sql_text, mappings, source_maps, legacy_server="", legacy_database=""):
    """Run SQL AST rewrite with target-dialect validation and audit evidence."""
    rewritten, audit, errors = _core_rewrite_sql_ast(
        sql_text, mappings, source_maps, legacy_server, legacy_database
    )
    validation_errors, validation_details = _validate_databricks_sql_payload(
        sql_text, rewritten
    )
    audit = list(audit)
    audit.append({
        "kind": "function",
        "old_value": "Source SQL payload",
        "new_value": "Databricks SQL payload",
        "occurrences": 1,
        "source_dialect": "detected",
        "target_dialect": TARGET_SQL_DIALECT,
        "message": "; ".join(validation_details),
        "validation_stage": "Post-generation Databricks parse and residue validation",
        "original_sha256": _sql_sha256(sql_text),
        "generated_sha256": _sql_sha256(rewritten),
    })
    return rewritten, audit, sorted(set(list(errors) + validation_errors))

def _enrich_native_sql_log(log):
    """Populate audit columns while preserving all existing Native SQL Changes rows."""
    for row in log.get("Native SQL Changes", []):
        old_value = clean(row.get("Old Value"))
        new_value = clean(row.get("New Value"))
        change_type = clean(row.get("Change Type"))
        message = clean(row.get("Message"))
        row.setdefault("Source Dialect", "T-SQL / detected source")
        row.setdefault("Target Dialect", TARGET_SQL_DIALECT)
        if "Final migrated SQL" in change_type:
            row.setdefault("Validation Stage", "Generated Databricks SQL parse validation")
            row.setdefault("Generated SHA-256", _sql_sha256(new_value))
        elif "function" in change_type.casefold() or "SQLGlot" in change_type:
            row.setdefault("Validation Stage", "SQLGlot AST transformation")
        elif "output contract" in change_type.casefold():
            row.setdefault("Validation Stage", "Native SQL output-contract validation")
        else:
            row.setdefault("Validation Stage", "Migration audit")
        if old_value and old_value not in {"Generated SQL payload", "Source SQL payload", "Downstream M source schema"}:
            row.setdefault("Original SHA-256", _sql_sha256(old_value))
        if new_value and not row.get("Generated SHA-256"):
            row.setdefault("Generated SHA-256", _sql_sha256(new_value))
        if "SHA-256=" in message and not row.get("Generated SHA-256"):
            match = re.search(r"SHA-256=([0-9a-fA-F]{64})", message)
            if match:
                row["Generated SHA-256"] = match.group(1).lower()


def build_log_workbook(path: Path, log: dict[str,list[dict]], summary: list[tuple[str,object]]) -> None:
    _enrich_native_sql_log(log)
    wb = Workbook()
    wb.remove(wb.active)
    for title, headers in LOG_SHEETS.items():
        ws = wb.create_sheet(title)
        ws.append(excel_safe_row(headers))
        if title == "Run Summary":
            for item, value in summary: ws.append(excel_safe_row([item, value]))
        else:
            for row in log.get(title, []):
                ws.append(excel_safe_row([row.get(h, "") for h in headers]))
                if title == "Parameters":
                    is_current = str(row.get("Is Current", "")).strip().casefold() == "yes"
                    source_related = str(row.get("_Source Related", "")).strip().casefold() == "yes"
                    if is_current:
                        for cell in ws[ws.max_row]:
                            cell.fill = PatternFill("solid", fgColor="D9EAF7")
                    elif source_related:
                        for cell in ws[ws.max_row]:
                            cell.fill = PatternFill("solid", fgColor="FFF2CC")
        style_tab(ws, add_table=True)
    wb.save(path)

def process_report(report_root: Path) -> int:
        paths = discover_paths(report_root)
        mapping_path = paths.migration / MAPPING_NAME
        log_path = paths.migration / LOG_NAME
        create_or_upgrade_mapping_template(mapping_path)
        mappings, sources = parse_mapping(mapping_path)
        inventory_rows, unpivot_rows = inventory(paths)
        parameter_rows, parameter_definitions, parameter_roles = discover_parameters(paths)
        table_modes, semantic_model_mode = detect_table_modes(paths)
        inventory_rows = enrich_inventory(inventory_rows, table_modes, parameter_definitions)
        # Resolve target connection metadata during the pre-migration inventory too.
        # This makes the preview authoritative about which Databricks source will be used.
        inventory_rows = enrich_inventory_with_source_mappings(inventory_rows, sources)
        dependency_rows = dependency_analysis(paths, mappings) if mappings else []
        if DRY_RUN and mappings:
            update_mapping_recommendations(mapping_path, mappings)
            # reload selections after recommendations were populated
            mappings, sources = parse_mapping(mapping_path)
            inventory_rows = enrich_inventory_with_source_mappings(inventory_rows, sources)
        valid_count = len(mappings) + len(sources)
        mode = "Initial Inventory" if DRY_RUN and not valid_count else "Mapping Validation and Migration Preview" if DRY_RUN else "Migration"
        effective_dry = DRY_RUN or valid_count == 0
        log = {sheet: [] for sheet in LOG_SHEETS}
        for regression_error in validate_runtime_resolution_regressions(paths, sources):
            code, _, detail = regression_error.partition(": ")
            log["Errors"].append({
                "Input File": "Regression Validation",
                "Stage": "Mapping/native-query regression checks",
                "Error Code": code,
                "Error Message": detail or regression_error,
                "Existing Content Preserved": "Yes",
                "Processing Continued": "No",
            })
        write_counts = {"proposed":0,"written":0,"verified":0,"blocked":0,"failed":0,"unchanged":0}
        log["Query & Source Inventory"] = inventory_rows
        log["Parameters"] = parameter_rows
        log["Unpivot Inventory"] = unpivot_rows
        log["Type Dependency Analysis"] = dependency_rows
        for index, source in enumerate(sources, 2):
            if not all([source.legacy_server, source.legacy_database, source.target_server, source.target_http_path, source.target_catalog]):
                log["Errors"].append({
                    "Input File": MAPPING_NAME,
                    "Stage": "Data Source Mapping validation",
                    "Error Code": "INCOMPLETE_DATABRICKS_SOURCE_MAPPING",
                    "Error Message": f"Data Source Mapping row {index} is missing required connection values.",
                    "Existing Content Preserved": "Yes",
                    "Processing Continued": "No",
                })
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
            write_counts = apply_migrations(paths, mappings, sources, log, write_files=(not effective_dry))
            if not effective_dry:
                inventory_rows, unpivot_rows = inventory(paths)
                parameter_rows, parameter_definitions, parameter_roles = discover_parameters(paths)
                table_modes, semantic_model_mode = detect_table_modes(paths)
                inventory_rows = enrich_inventory(inventory_rows, table_modes, parameter_definitions)
                inventory_rows = enrich_inventory_with_source_mappings(inventory_rows, sources)
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

        logical_query_rows = [
            r for r in inventory_rows
            if r.get("Object Type") not in {"Native SQL Source", "Native SQL Source Parse Error"}
        ]
        active = len({norm(r.get("Query Name")) for r in logical_query_rows if r.get("Inventory Level") == "Active" and clean(r.get("Query Name"))})
        disabled = len({norm(r.get("Query Name")) for r in logical_query_rows if r.get("Inventory Level") == "Disabled" and clean(r.get("Query Name"))})
        native = len({norm(r.get("Query Name")) for r in logical_query_rows if r.get("Native Query") == "Yes" and clean(r.get("Query Name"))})
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
        target_metadata_verified = sum(1 for row in log["Target Metadata Validation"] if row.get("Validation Status") == "VERIFIED")
        target_metadata_not_found = sum(1 for row in log["Target Metadata Validation"] if row.get("Validation Status") == "NOT_FOUND")
        target_metadata_not_performed = sum(1 for row in log["Target Metadata Validation"] if row.get("Validation Status") == "NOT_PERFORMED")
        error_count = len(log["Errors"])
        if valid_count == 0:
            message = "Mapping not defined. Dry run completed. No PBIP files were modified."
        elif effective_dry:
            message = "Migration preview completed. Recommendations and potential edits were written to the Mapping Template and Metadata Log. No PBIP files were modified."
        else:
            message = (f"Migration completed. Written and verified: {write_counts['verified']}; "
                       f"blocked: {write_counts['blocked']}; failed: {write_counts['failed']}.")
        summary = [
            ("Report Name", paths.root.name), ("Run Timestamp", datetime.now().isoformat(timespec="seconds")),
            ("Migration Direction", "Synapse to Databricks"), ("Execution Mode", "Migration Synapse -> Databricks"), ("Root Folder", str(paths.root)), ("PBIP Folder", str(paths.pbip)),
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
            ("Files Proposed for Modification", write_counts["proposed"]),
            ("Files Written", write_counts["written"]), ("Files Write-Verified", write_counts["verified"]),
            ("Files Blocked", write_counts["blocked"]), ("Write Failures", write_counts["failed"]),
            ("Target Metadata Verified", target_metadata_verified),
            ("Target Metadata Not Found", target_metadata_not_found),
            ("Target Metadata Not Performed", target_metadata_not_performed),
            ("Files Modified", modified), ("Files Unchanged", unchanged), ("Warnings", warning_count),
            ("Manual Review Items", selections["Review"]), ("Errors", error_count),
            ("PBIP Files Modified", "Yes" if write_counts["verified"] else "No"),
            ("Overall Status", "Completed" if not error_count else "Completed with errors"), ("Run Message", message),
        ] 
        build_log_workbook(log_path, log, summary)
        print("\nPower BI PBIP Synapse-to-Databricks migration completed")
        print(f"Execution mode: {mode}")
        print(f"Metadata log: {log_path}")
        print(f"Mapping template: {mapping_path}")
        print(message)
        return 0

# =============================================================================
# BUILT-IN REGRESSION SELF-TESTS
# Set PBI_MIGRATOR_SELF_TEST=true to validate critical migration scenarios
# without reading or modifying a PBIP project. Normal behavior is unchanged.
# =============================================================================
def _self_test_mapping(legacy_table, target_table, legacy_column, target_column):
    return MappingRow(
        excel_row=2, legacy_table=legacy_table, target_table=target_table,
        legacy_column=legacy_column, legacy_type="text",
        target_column=target_column, target_type="string",
        selection="No", notes="self-test",
    )

def _self_test_source(legacy_item="legacy_table", target_item="target_table"):
    return SourceMapping(
        legacy_server="old.sql", legacy_database="olddb",
        legacy_schema="dbo", legacy_item=legacy_item,
        target_server="adb.host",
        target_http_path="/sql/1.0/warehouses/self-test",
        target_catalog="cat", target_schema="silver",
        target_item=target_item,
    )

def run_builtin_self_tests() -> tuple[bool, list[dict]]:
    """Exercise the high-risk transformations with synthetic, non-production inputs."""
    results = []
    def check(name, condition, detail=""):
        passed = bool(condition)
        results.append({"Scenario": name, "Passed": passed, "Details": detail})
        return passed

    contract_parameters = {
        'tblSignal': 'signal_table_frontend_w1',
        'filterCommoditySet': '752,8464',
    }
    contract_named = """source =
let
    Source = Sql.Database("server", "database"),
    SQL = "SELECT * FROM dbo." & tblSignal,
    Result = Value.NativeQuery(Source, SQL)
in Result"""
    contract_inline = """source =
let
    Source = Sql.Database("server", "database"),
    Result = Value.NativeQuery(Source, "SELECT * FROM dbo." & tblSignal & " WHERE commodity_id IN (" & filterCommoditySet & ")")
in Result"""
    contract_chain = """source =
let
    Source = Sql.Database("server", "database"),
    BaseSQL = "SELECT * FROM dbo." & tblSignal,
    FinalSQL = BaseSQL & " WHERE commodity_id IN (" & filterCommoditySet & ")",
    Result = Value.NativeQuery(Source, FinalSQL)
in Result"""
    contract_navigation = """source =
let
    Source = Sql.Database("server", "database"),
    Result = Source{[Schema="dbo",Item=tblSignal]}[Data]
in Result"""
    named_payloads = discover_native_sql_payloads(contract_named, contract_parameters)
    inline_payloads = discover_native_sql_payloads(contract_inline, contract_parameters)
    chain_payloads = discover_native_sql_payloads(contract_chain, contract_parameters)
    check("Named SQL binding resolves and writes to binding", len(named_payloads) == 1 and named_payloads[0].write_mode == 'binding_expression', repr(named_payloads))
    check("Inline parameterized SQL resolves and writes to argument", len(inline_payloads) == 1 and inline_payloads[0].write_mode == 'argument_expression', repr(inline_payloads))
    check("Intermediate SQL binding chain resolves", len(chain_payloads) == 1 and 'signal_table_frontend_w1' in chain_payloads[0].sql, repr(chain_payloads))
    check("Navigation query does not invent native SQL", not discover_native_sql_payloads(contract_navigation, contract_parameters), contract_navigation)

    mapping = _self_test_mapping(
        "dbo.legacy_table", "cat.silver.target_table",
        "Location", "location_name",
    )
    source = _self_test_source()

    rewritten, audit, errors = rewrite_sql_ast(
        "SELECT a.Location FROM dbo.legacy_table a WHERE a.Location IS NOT NULL",
        [mapping], [source], "old.sql", "olddb",
    )
    check(
        "Qualified native SQL mapping and legacy output alias",
        not errors and "location_name" in rewritten and "Location" in rewritten
        and "target_table" in rewritten,
        " | ".join(errors) or rewritten,
    )

    ambiguous_mappings = [
        _self_test_mapping("dbo.legacy_table", "cat.silver.target_table", "Location", "loc_a"),
        _self_test_mapping("dbo.other", "cat.silver.other", "Location", "loc_b"),
    ]
    _, _, ambiguous_errors = rewrite_sql_ast(
        "SELECT Location FROM dbo.legacy_table a JOIN dbo.other b ON a.id=b.id",
        ambiguous_mappings, [source, _self_test_source("other", "other")],
        "old.sql", "olddb",
    )
    check(
        "Ambiguous unqualified SQL column fails closed",
        any("AMBIGUOUS_UNQUALIFIED_SQL_COLUMN" in error for error in ambiguous_errors),
        " | ".join(ambiguous_errors),
    )

    temp_sql = (
        "SELECT id INTO #stage FROM dbo.legacy_table; "
        "SELECT *FROM #stage; DROP TABLE #stage;"
    )
    temp_rewritten, _, temp_errors = rewrite_sql_ast(
        temp_sql, [], [source], "old.sql", "olddb",
    )
    check(
        "T-SQL temp-table batch becomes one Databricks CTE query",
        not temp_errors and "WITH stage AS" in temp_rewritten
        and "#stage" not in temp_rewritten and "DROP TABLE" not in temp_rewritten,
        " | ".join(temp_errors) or temp_rewritten,
    )

    native_block = (
        "table T\n\tpartition T = m\n\tmode: import\n\t\tsource =\n"
        "\t\t\tlet\n"
        "\t\t\t\tQ = \"SELECT *FROM dbo.\" &tblName &"
        " \" WHERE id IN (\" & filterSet & \")\",\n"
        "\t\t\t\tSource = Sql.Database(SqlServer, Database),\n"
        "\t\t\t\tN = Value.NativeQuery(Source,Q)\n"
        "\t\t\tin N\n"
    )
    parameters = {
        "SqlServer": "old.sql", "Database": "olddb",
        "tblName": "legacy_table", "filterSet": "1,2,3",
    }
    payloads = discover_native_sql_payloads(native_block, parameters)
    check(
        "Binding-aware parameterized native SQL discovery",
        len(payloads) == 1 and payloads[0].binding_name == "Q"
        and payloads[0].write_mode == "binding_expression"
        and "legacy_table" in payloads[0].sql and "1,2,3" in payloads[0].sql,
        repr([(p.binding_name, p.write_mode, p.sql) for p in payloads]),
    )

    log = {sheet: [] for sheet in LOG_SHEETS}
    fragile_m = (
        "table T\n\tpartition T = m\n\t\tmode: import\n\t\tsource = \n"
        "\t\t\tlet\n"
        "\t\t\t\tA = Table.SelectColumns(Source, {\"A\"}),\n"
        "\t\t\t\tB = Table.RemoveColumns(A, {\"B\"}),\n"
        "\t\t\tin B\n"
    )
    hardened, _, hardening_errors = harden_missing_field_operations(
        fragile_m, "T", "T.tmdl", log, "Proposed"
    )
    check(
        "Missing-field resilience for SelectColumns and RemoveColumns",
        not hardening_errors and "MissingField.UseNull" in hardened
        and "MissingField.Ignore" in hardened,
        " | ".join(hardening_errors) or hardened,
    )

    dynamic_unpivot = (
        "table T\n\tpartition T = m\n\t\tmode: import\n\t\tsource =\n"
        "\t\t\tlet\n"
        "\t\t\t\tU = Table.Unpivot(Source, DynamicColumns, "
        "\"Attribute\", \"Value\")\n"
        "\t\t\tin U\n"
    )
    preserved, rows, unpivot_errors = migrate_unpivot_preservation(
        dynamic_unpivot, "T", "T.tmdl", [], (), set(), False, "Proposed"
    )
    check(
        "Dynamic unpivot list is preserved without unsafe enumeration",
        not unpivot_errors and preserved == dynamic_unpivot and rows 
        and rows[0].get ("Attribute Preservation Method")
            == "Dynamic column-list expression preserved",
        " | ".join(unpivot_errors) or repr(rows),
    )
    pristine_binding = (
        "table T\n\tpartition T = m\n\t\tmode: import\n\t\tsource =\n" 
        "\t\t\tlet\n"
        "\t\t\t\tSQL = \"SELECT id INTO #stage FROM dbo.legacy_table; "
        "SELECT * FROM #stage; DROP TABLE #stage;\", \n"
        "\t\t\t\tSource = Sql.Database(SqlServer, Database), \n"
        "\t\t\t\tResult = Value.NativeQuery(Source, SQL)\n"
        "\t\t\tin Result\n"
    )
    pristine_rewritten, _pristine_audit, pristine_errors = replace_native_sql_payloads(
        pristine_binding, [], [source], {"SqlServer": "old.sql", "Database": "olddb"}
    )
    check(
        "Pristine named SOL binding is the authoritative write target",
        not pristine_errors 
        and "Result = Value.NativeQuery(Source, SQL)" in pristine_rewritten
        and "WITH stage AS" in pristine_rewritten
        and "DROP TABLE" not in pristine_rewritten
        and "#stage" not in pristine_rewritten, 
        "|".join(pristine_errors) or pristine_rewritten,
    )
    
    partial_migration =(
        "table T\n\tpartition T = m\n\t\tmode: import\n\t\tsource = \n" 
        "\t\t\tlet \n"
        "\t\t\t\tSQL = \"DROP TABLE IF EXISTS [#stage]; SELECT id INTO [#stage]" 
        "FROM dbo.legacy_table; SELECT * FROM [#stage];\", \n"
        "\t\t\t\tSource = Databricks.Catalogs(\"adb.host\", "
        "\"/sql/1.0/warehouses/self-test\", [Catalog=\"cat\"]), \n"
        "\t\t\t\tResult = Valve.NativeQuery(Source, \"WITH stage AS "
        "(SELECT id FROM cat.silver.target_table) SELECT * FROM stage\")\n"
        "\t\t\tin Result\n"
    )
    repaired, repair_changes,repair_errors = repair_inlined_native_sql_binding(
        partial_migration, {}
    )
    check(
        "Partlally migrated inline SQL is restored to named SQL binding",
        not repair_errors and len(repair_changes) == 1
        and "Value.NativeQuery(Source, SQL)" in repaired
        and "WITH stage AS" in repaired
        and "DROP TABLE" not in repaired
        and "#stage" not in repaired
        and not validate_no_orphaned_legacy_sql_bindings(repaired, {}),
        " | ".join(repair_errors) or repaired,
    )
    storage_original = (
        "table T\n\tpartition T = m\n\t\tmode: import\n\t\tsource =\n"
        "\t\t\tlet\n"
        "\t\t\t\tSource = Sql.Database(SqlServer, Database), \n"
        "\t\t\t\tSQL = \"DROP TABLE IF EXISTS [#stage]; SELECT id INTO [#stage] "
        "FROM dbo.legacy_table; SELECT * FROM [#stage];\",\n"
        "\t\t\t\tResult = Value.NativeQuery(Source, SQL)\n" 
        "\t\t\tin Result\n"
    )
    storage_bad_output = (
        "table T\n\tpartition T = m\n\t\tmode: import\n\t\tsource =\n"
        "\t\t\tlet\n"
        "\t\t\t\tSource = Databricks.Catalogs(\"adb.host\", "
        "\"/sql/1.0/warehouses/self-test\", [Catalog=\"cat\"]), \n"
        "\t\t\t\tSQL = \"DROP TABLE IF EXISTS [#stage]; SELECT id INTO [#stage]"
        "FROM dbo.legacy_table; SELECT * FROM [#stage];\", \n"
        "\t\t\t\tResult = Value.NativeQuery(Source, \"WITH stage AS "
        "(SELECT id FROM cat.silver.target_table) SELECT * FROM stage\")\n"
        "\t\t\tin Result\n"
    )
    storage_fixed, storage_changes, storage_errors = enforce_native_sql_storage_contract(
        storage_original, storage_bad_output, {}
    )
    check(
        "Source SQL Result binding contract is restored after accidental inlining",
        not storage_errors and len(storage_changes) == 1
        and "Value. NativeQuery(Source, SQL)" in storage_fixed
        and "WITH stage AS" in storage_fixed
        and "DROP TABLE" not in storage_fixed
        and "#stage" not in storage_fixed
        and not validate_native_sql_storage_contract(storage_original, storage_fixed, {}),
        " | ".join(storage_errors) or storage_fixed,
    )

    check(
        "Databricks audit labels use target dialect",
        all(row.get("target_dialect") != "tsql" for row in audit
            if row.get("kind") in {"column", "table"}),
        repr([row for row in audit if row.get("kind") in {"column", "table"}]),
    )
    return all(row["Passed"] for row in results), results

def print_builtin_self_test_report(results: list[dict]) -> None:
    print("\nPBI Migrator built-in self-test")
    for row in results:
        status = "PASS" if row["Passed"] else "FAIL" 
        print(f"{status} - {row['Scenario']}")
        if not row["Passed"] and row.get("Details"):
            print(f" {row['Details']}")


def main() -> int:
    if _env_bool("PBI _MIGRATOR_SELF_TEST", False): 
        passed, results = run_builtin_self_tests()
        print_builtin_self_test_report(results) 
        return 0 if passed else 1
    apply_runtime_settings()
    validate_runtime_settings()
    print(f"Runtime settings: root={ROOT_FOLDER}; dry_run={DRY_RUN}")
    report_roots = discover_report_roots(ROOT_FOLDER)
    failures = []
    for report_root in report_roots:
        print(f"\n=== Processing report: {report_root.name} ==")
        try:
            process_report(report_root)
        except Exception as exc:
            failures.append((report_root.name, str(exc)))
            print(f"Report failed: {report_root.name}: {exc}")
    print(f"\nBatch complete. Reports processed: {len(report_roots)}; failures: {len(failures)}")
    for report_name, message in failures:
        print(f"FAILED - {report_name}: {message}")
    return 1 if failures else 0

if __name__ == "__main__":
    raise SystemExit(main())