"""
Enterprise Finance Reconciliation Tool
======================================

Purpose:
--------
Performs automated end-to-end reconciliation of financial records between an Enterprise
System of Record (ERP) and an External Source (Bank / Payment Gateway / Partner).

Key Features:
-------------
1. Multi-format Ingestion: Ingests CSV and XLSX files defensively with SHA-256 hash tracking.
2. Data Sanitization & Collision Prevention: Normalizes headers to snake_case and detects header collisions.
3. Explicit Schema & Typed Rules: Replaces substring searching with typed comparison rules
   (Money, Text, Date) and exact Decimal representation for financial amounts.
4. Strict Key Quarantine: Isolates missing primary keys (MISSING_PRIMARY_KEY) and ambiguous duplicate
   keys (DUPLICATE_KEY_IN_ERP, DUPLICATE_KEY_IN_EXTERNAL, DUPLICATE_KEY_IN_BOTH) prior to matching.
5. One-to-One Vectorized Reconciliation: Merges clean datasets under strict 1-to-1 validation constraints.
6. Comprehensive Audit Reporting: Multi-sheet Excel workbook with executive summary KPIs, currency-grouped
   exposures, run metadata, formula injection protection, and source row traceability.
"""

import argparse
import hashlib
import logging
import os
import re
import sys
import uuid
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Set up logger
logger = logging.getLogger("reconciliation")

SUPPORTED_EXTENSIONS = {".csv", ".xlsx"}
DEFAULT_TOLERANCE = Decimal("0.01")
DEFAULT_KEY_COLUMN = "invoice_id"

# Machine-readable status taxonomy constants
STATUS_MATCH = "MATCH"
STATUS_MISMATCH = "MISMATCH"
STATUS_MISSING_IN_EXTERNAL = "MISSING_IN_EXTERNAL"
STATUS_MISSING_IN_ERP = "MISSING_IN_ERP"
STATUS_MISSING_PRIMARY_KEY = "MISSING_PRIMARY_KEY"
STATUS_DUPLICATE_IN_ERP = "DUPLICATE_KEY_IN_ERP"
STATUS_DUPLICATE_IN_EXTERNAL = "DUPLICATE_KEY_IN_EXTERNAL"
STATUS_DUPLICATE_IN_BOTH = "DUPLICATE_KEY_IN_BOTH"
STATUS_INVALID_FIELD_VALUE = "INVALID_FIELD_VALUE"

INVALID_VALUE_SENTINEL = "<INVALID_FIELD_VALUE>"

CURRENCY_PATTERN = re.compile(
    r"[$₹€£¥]|Rs\.?|\b(USD|EUR|INR|GBP|CAD|AUD|JPY|CHF|CNY|HKD|NZD)\b",
    re.IGNORECASE,
)


class ReconciliationError(Exception):
    """Base exception for reconciliation failures."""
    pass


class ConfigurationError(ReconciliationError):
    """Raised when configuration parameters are invalid."""
    pass


class SchemaError(ReconciliationError):
    """Raised when dataset schemas or header validations fail."""
    pass


class IngestionError(ReconciliationError):
    """Raised when file reading or ingestion fails."""
    pass


def calculate_file_hash(file_path: str) -> str:
    """Calculates SHA-256 hash of a file for run auditability."""
    sha256 = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            sha256.update(chunk)
    return sha256.hexdigest()


def read_spreadsheet(file_path: str, source_name: str = "Source") -> pd.DataFrame:
    """
    Reads a CSV or XLSX file into a pandas DataFrame.
    Attaches source row traceability metadata (_source_row_<source_name>).
    Header is row 1, first data row is row 2.
    """
    file_path = file_path.strip().strip('"').strip("'")
    if not os.path.exists(file_path):
        raise IngestionError(f"{source_name} file not found: '{file_path}'")

    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        raise IngestionError(
            f"Unsupported file extension '{ext}' for {source_name}. Supported types: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    try:
        if ext == ".csv":
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        raise IngestionError(f"Failed to read {source_name} file '{file_path}': {e}") from e

    # Source row traceability: header is row 1, data rows start at row 2
    row_col_name = f"_source_row_{source_name.lower().replace(' ', '_')}"
    df[row_col_name] = range(2, len(df) + 2)

    return df


def sanitize_column_name(col: Any) -> str:
    """
    Normalizes column headers to lowercase snake_case.
    Handles spaces, hyphens, punctuation, and leading/trailing separators.
    Example: 'Invoice-ID!' -> 'invoice_id'
    """
    s = str(col).strip().lower()
    s = re.sub(r"[^\w]+", "_", s)
    s = re.sub(r"_+", "_", s)
    s = s.strip("_")
    return s


def preprocess_data(df: pd.DataFrame, source_name: str = "Source") -> pd.DataFrame:
    """
    Cleans and standardizes the DataFrame:
    - Normalizes column names and checks for collision errors.
    - Trims whitespace from string cells.
    - Drops completely empty rows (preserving source traceability).
    """
    df = df.copy()

    # Column normalization and collision detection
    sanitized_map: Dict[str, List[str]] = {}
    new_columns = []

    for orig_col in df.columns:
        if str(orig_col).startswith("_source_row_"):
            new_columns.append(orig_col)
            continue
        clean_col = sanitize_column_name(orig_col)
        new_columns.append(clean_col)
        sanitized_map.setdefault(clean_col, []).append(str(orig_col))

    collisions = {k: v for k, v in sanitized_map.items() if len(v) > 1}
    if collisions:
        collision_details = "; ".join([f"'{k}': {v}" for k, v in collisions.items()])
        raise SchemaError(
            f"Header normalization collision detected in {source_name}: {collision_details}"
        )

    df.columns = new_columns

    # Drop rows where all business columns (excluding source row metadata) are null
    biz_cols = [c for c in df.columns if not c.startswith("_source_row_")]
    df = df.dropna(how="all", subset=biz_cols)

    # Trim leading/trailing string whitespace
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    return df


def parse_decimal(val: Any) -> Union[Decimal, None, str]:
    """
    Safely parses financial amounts into decimal.Decimal.

    Supports:
    - 1500, 1500.00, $1,500.00, ₹1,500.00, €1,500.00, USD 1500.00
    - Negative amounts: -1500.00, (500.00)

    Returns:
    - Decimal object for valid monetary values
    - None for missing/null/blank values
    - INVALID_VALUE_SENTINEL for malformed non-financial values (e.g. 'ERROR123', '12.3.4')
    """
    if val is None or pd.isna(val):
        return None

    if isinstance(val, Decimal):
        return val

    if isinstance(val, (int, float)):
        if pd.isna(val):
            return None
        try:
            return Decimal(str(val))
        except InvalidOperation:
            return INVALID_VALUE_SENTINEL

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("nan", "none", "null", "<na>"):
        return None

    is_parenthesized_negative = False
    if val_str.startswith("(") and val_str.endswith(")"):
        is_parenthesized_negative = True
        val_str = val_str[1:-1].strip()

    # Strip recognized currency symbols/codes, commas, and spaces
    cleaned = CURRENCY_PATTERN.sub("", val_str)
    cleaned = cleaned.replace(",", "").replace(" ", "")

    if is_parenthesized_negative and not cleaned.startswith("-"):
        cleaned = f"-{cleaned}"

    # Strict numeric validation regex: allows optional leading minus and single decimal point
    if not re.match(r"^-?\d+(\.\d+)?$", cleaned):
        return INVALID_VALUE_SENTINEL

    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return INVALID_VALUE_SENTINEL


def is_valid_key(val: Any) -> bool:
    """Validates if a primary key value is present and non-blank."""
    if val is None or pd.isna(val):
        return False
    s = str(val).strip()
    return bool(s) and s.lower() not in ("nan", "none", "null", "<na>")


def sanitize_excel_cell_value(val: Any) -> Any:
    """
    Prevents Excel formula injection by prefixing text starting with '=', '+', '-', '@'
    with a single quote, unless it is a valid numeric/date/boolean value.
    """
    if isinstance(val, str) and val:
        if val[0] in ("=", "+", "-", "@"):
            # Do not escape valid negative numbers
            if val[0] == "-" and re.match(r"^-?\d+(\.\d+)?$", val):
                return val
            return f"'{val}"
    return val


def reconcile_data(
    df_erp: pd.DataFrame,
    df_external: pd.DataFrame,
    key_column: str = DEFAULT_KEY_COLUMN,
    amount_tolerance: Union[float, Decimal] = DEFAULT_TOLERANCE,
    field_rules: Optional[Dict[str, Dict[str, Any]]] = None,
    source_erp_name: str = "ERP",
    source_external_name: str = "External",
) -> Dict[str, pd.DataFrame]:
    """
    Core Financial Reconciliation Engine.
    """
    run_id = str(uuid.uuid4())
    run_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if isinstance(amount_tolerance, (float, int, str)):
        try:
            amount_tolerance = Decimal(str(amount_tolerance))
        except InvalidOperation:
            raise ConfigurationError(f"Invalid amount tolerance value: {amount_tolerance}")

    if amount_tolerance < Decimal("0"):
        raise ConfigurationError(f"Amount tolerance cannot be negative: {amount_tolerance}")

    key_col = sanitize_column_name(key_column)

    # 1. Key Column Presence Validation
    if key_col not in df_erp.columns:
        raise SchemaError(
            f"Key column '{key_col}' not found in {source_erp_name}. Available columns: {[c for c in df_erp.columns if not c.startswith('_source_row_')]}"
        )
    if key_col not in df_external.columns:
        raise SchemaError(
            f"Key column '{key_col}' not found in {source_external_name}. Available columns: {[c for c in df_external.columns if not c.startswith('_source_row_')]}"
        )

    # 2. Configured Field Rules Setup & Validation
    if field_rules is None:
        common_cols = [
            c for c in df_erp.columns
            if c in df_external.columns and c != key_col and not c.startswith("_source_row_")
        ]
        field_rules = {}
        for c in common_cols:
            if any(kw in c for kw in ["amount", "price", "fee", "total", "val", "cost", "sum"]):
                field_rules[c] = {"type": "money", "tolerance": amount_tolerance}
            elif "date" in c:
                field_rules[c] = {"type": "date", "format": "%Y-%m-%d"}
            else:
                field_rules[c] = {"type": "text", "case_sensitive": False}

    comparison_fields = []
    for col, rules in field_rules.items():
        clean_c = sanitize_column_name(col)
        if clean_c == key_col:
            continue
        if clean_c in df_erp.columns and clean_c in df_external.columns:
            comparison_fields.append(clean_c)
        else:
            missing_in = []
            if clean_c not in df_erp.columns:
                missing_in.append(source_erp_name)
            if clean_c not in df_external.columns:
                missing_in.append(source_external_name)
            raise SchemaError(
                f"Required comparison field '{clean_c}' is missing in dataset(s): {', '.join(missing_in)}"
            )

    if not comparison_fields:
        raise SchemaError("Zero valid comparison fields were supplied for reconciliation.")

    logger.info(
        f"Starting reconciliation run {run_id}. Key: '{key_col}'. Compared fields: {comparison_fields}"
    )

    # -------------------------------------------------------------
    # 3. Primary Key Null / Invalid Quarantine
    # -------------------------------------------------------------
    erp_valid_key_mask = df_erp[key_col].apply(is_valid_key)
    ext_valid_key_mask = df_external[key_col].apply(is_valid_key)

    df_erp_invalid_keys = df_erp[~erp_valid_key_mask].copy()
    df_ext_invalid_keys = df_external[~ext_valid_key_mask].copy()

    df_erp_valid = df_erp[erp_valid_key_mask].copy()
    df_ext_valid = df_external[ext_valid_key_mask].copy()

    invalid_key_records = []
    if not df_erp_invalid_keys.empty:
        df_erp_invalid_keys["reconciliation_status"] = STATUS_MISSING_PRIMARY_KEY
        df_erp_invalid_keys["status_reason"] = f"Missing or blank primary key in {source_erp_name}"
        df_erp_invalid_keys["source_system"] = source_erp_name
        invalid_key_records.append(df_erp_invalid_keys)

    if not df_ext_invalid_keys.empty:
        df_ext_invalid_keys["reconciliation_status"] = STATUS_MISSING_PRIMARY_KEY
        df_ext_invalid_keys["status_reason"] = f"Missing or blank primary key in {source_external_name}"
        df_ext_invalid_keys["source_system"] = source_external_name
        invalid_key_records.append(df_ext_invalid_keys)

    # -------------------------------------------------------------
    # 4. Duplicate Key Quarantine Logic
    # -------------------------------------------------------------
    erp_counts = df_erp_valid[key_col].value_counts()
    ext_counts = df_ext_valid[key_col].value_counts()

    erp_dupe_keys = set(erp_counts[erp_counts > 1].index)
    ext_dupe_keys = set(ext_counts[ext_counts > 1].index)
    all_dupe_keys = erp_dupe_keys.union(ext_dupe_keys)

    df_erp_dupes = df_erp_valid[df_erp_valid[key_col].isin(all_dupe_keys)].copy()
    df_ext_dupes = df_ext_valid[df_ext_valid[key_col].isin(all_dupe_keys)].copy()

    df_erp_clean = df_erp_valid[~df_erp_valid[key_col].isin(all_dupe_keys)].copy()
    df_ext_clean = df_ext_valid[~df_ext_valid[key_col].isin(all_dupe_keys)].copy()

    dupe_records = []
    for k in all_dupe_keys:
        in_erp = k in erp_dupe_keys
        in_ext = k in ext_dupe_keys

        if in_erp and in_ext:
            status = STATUS_DUPLICATE_IN_BOTH
            reason = f"Duplicate primary key '{k}' detected in both {source_erp_name} and {source_external_name}"
        elif in_erp:
            status = STATUS_DUPLICATE_IN_ERP
            reason = f"Duplicate primary key '{k}' detected in {source_erp_name}"
        else:
            status = STATUS_DUPLICATE_IN_EXTERNAL
            reason = f"Duplicate primary key '{k}' detected in {source_external_name}"

        rows_erp = df_erp_dupes[df_erp_dupes[key_col] == k].copy()
        if not rows_erp.empty:
            rows_erp["reconciliation_status"] = status
            rows_erp["status_reason"] = reason
            rows_erp["source_system"] = source_erp_name
            dupe_records.append(rows_erp)

        rows_ext = df_ext_dupes[df_ext_dupes[key_col] == k].copy()
        if not rows_ext.empty:
            rows_ext["reconciliation_status"] = status
            rows_ext["status_reason"] = reason
            rows_ext["source_system"] = source_external_name
            dupe_records.append(rows_ext)

    # -------------------------------------------------------------
    # 5. One-to-One Vector Outer Merge on Clean Records
    # -------------------------------------------------------------
    merged = df_erp_clean.merge(
        df_ext_clean,
        on=key_col,
        how="outer",
        suffixes=(f"_{source_erp_name.lower()}", f"_{source_external_name.lower()}"),
        indicator="_merge",
        validate="one_to_one",
    )

    erp_suffix = f"_{source_erp_name.lower()}"
    ext_suffix = f"_{source_external_name.lower()}"

    missing_records = merged[merged["_merge"] != "both"].copy()
    missing_records["reconciliation_status"] = missing_records["_merge"].map(
        {
            "left_only": STATUS_MISSING_IN_EXTERNAL,
            "right_only": STATUS_MISSING_IN_ERP,
        }
    )
    missing_records["status_reason"] = missing_records["_merge"].map(
        {
            "left_only": f"Present in {source_erp_name} only (Missing in {source_external_name})",
            "right_only": f"Present in {source_external_name} only (Missing in {source_erp_name})",
        }
    )
    missing_records["source_system"] = missing_records["_merge"].map(
        {
            "left_only": source_erp_name,
            "right_only": source_external_name,
        }
    )

    # -------------------------------------------------------------
    # 6. Field-Level Typed Comparison for Common Records
    # -------------------------------------------------------------
    common = merged[merged["_merge"] == "both"].copy()

    common["reconciliation_status"] = STATUS_MATCH
    common["status_reason"] = ""
    common["mismatch_fields"] = ""
    common["mismatch_count"] = 0
    common["amount_difference"] = None

    statuses = []
    status_reasons = []
    mismatch_fields_list = []
    mismatch_counts = []
    amount_diffs = []

    for idx, row in common.iterrows():
        row_mismatched_fields = []
        row_reasons = []
        row_has_invalid = False
        row_amount_diff: Optional[Decimal] = None

        for col in comparison_fields:
            rule = field_rules.get(col, {"type": "text"})
            f_type = rule.get("type", "text")

            val_erp_raw = row.get(f"{col}{erp_suffix}")
            val_ext_raw = row.get(f"{col}{ext_suffix}")

            if f_type == "money":
                col_tol = rule.get("tolerance", amount_tolerance)
                dec_erp = parse_decimal(val_erp_raw)
                dec_ext = parse_decimal(val_ext_raw)

                if dec_erp == INVALID_VALUE_SENTINEL or dec_ext == INVALID_VALUE_SENTINEL:
                    row_has_invalid = True
                    row_reasons.append(
                        f"{col} has invalid monetary format ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')"
                    )
                elif dec_erp is None and dec_ext is None:
                    pass
                elif dec_erp is None or dec_ext is None:
                    row_mismatched_fields.append(col)
                    row_reasons.append(
                        f"{col} missing in one dataset ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')"
                    )
                else:
                    diff = abs(dec_erp - dec_ext)
                    if diff > col_tol:
                        row_mismatched_fields.append(col)
                        row_amount_diff = diff
                        row_reasons.append(
                            f"{col} diff ({source_erp_name}: {dec_erp}, {source_external_name}: {dec_ext}, diff: {diff})"
                        )

            elif f_type == "text":
                case_sens = rule.get("case_sensitive", False)
                str_erp = str(val_erp_raw).strip() if val_erp_raw is not None and not pd.isna(val_erp_raw) else None
                str_ext = str(val_ext_raw).strip() if val_ext_raw is not None and not pd.isna(val_ext_raw) else None

                if str_erp is None and str_ext is None:
                    pass
                elif str_erp is None or str_ext is None:
                    row_mismatched_fields.append(col)
                    row_reasons.append(
                        f"{col} missing in one dataset ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')"
                    )
                else:
                    c_erp = str_erp if case_sens else str_erp.lower()
                    c_ext = str_ext if case_sens else str_ext.lower()
                    if c_erp != c_ext:
                        row_mismatched_fields.append(col)
                        row_reasons.append(
                            f"{col} mismatch ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')"
                        )

            elif f_type == "date":
                fmt = rule.get("format", None)
                str_erp = str(val_erp_raw).strip() if val_erp_raw is not None and not pd.isna(val_erp_raw) else None
                str_ext = str(val_ext_raw).strip() if val_ext_raw is not None and not pd.isna(val_ext_raw) else None

                if str_erp is None and str_ext is None:
                    pass
                elif str_erp is None or str_ext is None:
                    row_mismatched_fields.append(col)
                    row_reasons.append(
                        f"{col} date missing in one dataset ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')"
                    )
                else:
                    dt_erp = pd.to_datetime(str_erp, format=fmt, errors="coerce")
                    dt_ext = pd.to_datetime(str_ext, format=fmt, errors="coerce")

                    if pd.isna(dt_erp) or pd.isna(dt_ext):
                        if str_erp != str_ext:
                            row_mismatched_fields.append(col)
                            row_reasons.append(
                                f"{col} unparseable date mismatch ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')"
                            )
                    elif dt_erp != dt_ext:
                        row_mismatched_fields.append(col)
                        row_reasons.append(
                            f"{col} date mismatch ({source_erp_name}: {dt_erp.strftime('%Y-%m-%d')}, {source_external_name}: {dt_ext.strftime('%Y-%m-%d')})"
                        )

        if row_has_invalid:
            statuses.append(STATUS_INVALID_FIELD_VALUE)
        elif row_mismatched_fields:
            statuses.append(STATUS_MISMATCH)
        else:
            statuses.append(STATUS_MATCH)

        status_reasons.append("; ".join(row_reasons))
        mismatch_fields_list.append(", ".join(row_mismatched_fields))
        mismatch_counts.append(len(row_mismatched_fields))
        amount_diffs.append(row_amount_diff)

    common["reconciliation_status"] = statuses
    common["status_reason"] = status_reasons
    common["mismatch_fields"] = mismatch_fields_list
    common["mismatch_count"] = mismatch_counts
    common["amount_difference"] = amount_diffs

    # -------------------------------------------------------------
    # 7. Aggregate Unified Category DataFrames
    # -------------------------------------------------------------
    df_matches = common[common["reconciliation_status"] == STATUS_MATCH].copy()
    df_mismatches = common[common["reconciliation_status"] == STATUS_MISMATCH].copy()

    data_issue_components = []
    if invalid_key_records:
        data_issue_components.extend(invalid_key_records)
    if dupe_records:
        data_issue_components.extend(dupe_records)

    df_invalid_fields = common[common["reconciliation_status"] == STATUS_INVALID_FIELD_VALUE].copy()
    if not df_invalid_fields.empty:
        data_issue_components.append(df_invalid_fields)

    if data_issue_components:
        df_data_issues = pd.concat(data_issue_components, ignore_index=True)
    else:
        df_data_issues = pd.DataFrame(
            columns=[key_col, "reconciliation_status", "status_reason", "source_system"]
        )

    df_missing = missing_records.copy()

    # -------------------------------------------------------------
    # 8. Build All Records Complete Dataset
    # -------------------------------------------------------------
    all_records_list = [common, df_missing]
    if invalid_key_records:
        all_records_list.extend(invalid_key_records)
    if dupe_records:
        all_records_list.extend(dupe_records)

    df_all_records = pd.concat(all_records_list, ignore_index=True)
    df_all_records = df_all_records.drop(columns=["_merge"], errors="ignore")

    primary_cols = [key_col, "reconciliation_status", "status_reason", "mismatch_fields", "amount_difference"]
    other_cols = [c for c in df_all_records.columns if c not in primary_cols]
    df_all_records = df_all_records[primary_cols + other_cols]

    # -------------------------------------------------------------
    # 9. Summary KPIs & Financial Exposure Calculation
    # -------------------------------------------------------------
    total_erp_ingested = len(df_erp)
    total_ext_ingested = len(df_external)
    eligible_count = len(common)
    matches_count = len(df_matches)
    mismatches_count = len(df_mismatches)
    missing_ext_count = len(df_missing[df_missing["reconciliation_status"] == STATUS_MISSING_IN_EXTERNAL])
    missing_erp_count = len(df_missing[df_missing["reconciliation_status"] == STATUS_MISSING_IN_ERP])
    data_issues_count = len(df_data_issues)

    match_rate = (matches_count / eligible_count * 100) if eligible_count > 0 else 0.0
    total_records = total_erp_ingested + total_ext_ingested
    unreconciled_count = total_records - (matches_count * 2)
    exception_rate = (unreconciled_count / total_records * 100) if total_records > 0 else 0.0

    summary_rows = [
        {"Metric / Indicator": "Total ERP Records Ingested", "Value": total_erp_ingested},
        {"Metric / Indicator": "Total External Records Ingested", "Value": total_ext_ingested},
        {"Metric / Indicator": "Eligible 1-to-1 Reconciliation Population", "Value": eligible_count},
        {"Metric / Indicator": "Fully Matched Records (MATCH)", "Value": matches_count},
        {"Metric / Indicator": "Field Discrepancies (MISMATCH)", "Value": mismatches_count},
        {"Metric / Indicator": "Missing in External (MISSING_IN_EXTERNAL)", "Value": missing_ext_count},
        {"Metric / Indicator": "Missing in ERP (MISSING_IN_ERP)", "Value": missing_erp_count},
        {"Metric / Indicator": "Data Quality Exceptions (Duplicates & Missing Keys)", "Value": data_issues_count},
        {"Metric / Indicator": "Match Rate (%)", "Value": f"{match_rate:.2f}%"},
        {"Metric / Indicator": "Overall Exception Rate (%)", "Value": f"{exception_rate:.2f}%"},
    ]

    if "currency" in comparison_fields:
        curr_erp_col = f"currency_{source_erp_name.lower()}"
        curr_ext_col = f"currency_{source_external_name.lower()}"

        if not df_mismatches.empty and "amount_difference" in df_mismatches.columns:
            currencies = df_mismatches[curr_erp_col].fillna(df_mismatches[curr_ext_col]).unique()
            for curr in currencies:
                if pd.isna(curr) or not curr:
                    continue
                curr_mask = (df_mismatches[curr_erp_col] == curr) | (df_mismatches[curr_ext_col] == curr)
                exposure = df_mismatches[curr_mask]["amount_difference"].dropna().sum()
                summary_rows.append(
                    {
                        "Metric / Indicator": f"Mismatched Exposure ({curr})",
                        "Value": f"{exposure:.2f}",
                    }
                )

        if not df_missing.empty:
            amt_erp_col = f"amount_{source_erp_name.lower()}"
            amt_ext_col = f"amount_{source_external_name.lower()}"

            missing_ext_df = df_missing[df_missing["reconciliation_status"] == STATUS_MISSING_IN_EXTERNAL]
            if not missing_ext_df.empty and amt_erp_col in missing_ext_df.columns:
                for curr in missing_ext_df[curr_erp_col].dropna().unique():
                    c_df = missing_ext_df[missing_ext_df[curr_erp_col] == curr]
                    sum_amt = c_df[amt_erp_col].apply(parse_decimal).map(lambda x: x if isinstance(x, Decimal) else Decimal('0')).sum()
                    summary_rows.append(
                        {
                            "Metric / Indicator": f"Missing in External Financial Exposure ({curr})",
                            "Value": f"{sum_amt:.2f}",
                        }
                    )

            missing_erp_df = df_missing[df_missing["reconciliation_status"] == STATUS_MISSING_IN_ERP]
            if not missing_erp_df.empty and amt_ext_col in missing_erp_df.columns:
                for curr in missing_erp_df[curr_ext_col].dropna().unique():
                    c_df = missing_erp_df[missing_erp_df[curr_ext_col] == curr]
                    sum_amt = c_df[amt_ext_col].apply(parse_decimal).map(lambda x: x if isinstance(x, Decimal) else Decimal('0')).sum()
                    summary_rows.append(
                        {
                            "Metric / Indicator": f"Missing in ERP Financial Exposure ({curr})",
                            "Value": f"{sum_amt:.2f}",
                        }
                    )

    df_summary = pd.DataFrame(summary_rows)

    df_metadata = pd.DataFrame(
        [
            {"Parameter": "Run ID", "Value": run_id},
            {"Parameter": "Run Timestamp", "Value": run_timestamp},
            {"Parameter": "Tool Version", "Value": "1.0.0"},
            {"Parameter": "Primary Key Column", "Value": key_col},
            {"Parameter": "Amount Tolerance Threshold", "Value": str(amount_tolerance)},
            {"Parameter": "Compared Fields", "Value": ", ".join(comparison_fields)},
            {"Parameter": f"{source_erp_name} Total Ingested Rows", "Value": total_erp_ingested},
            {"Parameter": f"{source_external_name} Total Ingested Rows", "Value": total_ext_ingested},
        ]
    )

    results = {
        "summary": df_summary,
        "run_metadata": df_metadata,
        "matches": df_matches.drop(columns=["_merge"], errors="ignore"),
        "mismatches": df_mismatches.drop(columns=["_merge"], errors="ignore"),
        "missing": df_missing.drop(columns=["_merge"], errors="ignore"),
        "data_issues": df_data_issues.drop(columns=["_merge"], errors="ignore"),
        "all_records": df_all_records,
    }

    return results


def write_output(
    results: Dict[str, pd.DataFrame],
    output_folder: str,
    output_filename: Optional[str] = None,
    metadata_info: Optional[Dict[str, str]] = None,
) -> str:
    """
    Writes reconciliation results into a multi-sheet Excel workbook.
    """
    output_folder = output_folder.strip().strip('"').strip("'")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder, exist_ok=True)

    if not output_filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"reconciliation_output_{timestamp}.xlsx"

    output_path = os.path.join(output_folder, output_filename)

    if metadata_info and "run_metadata" in results:
        meta_df = results["run_metadata"].copy()
        for k, v in metadata_info.items():
            meta_df = pd.concat([meta_df, pd.DataFrame([{"Parameter": k, "Value": str(v)}])], ignore_index=True)
        results["run_metadata"] = meta_df

    sheet_order = ["summary", "run_metadata", "matches", "mismatches", "missing", "data_issues", "all_records"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name in sheet_order:
            if sheet_name in results:
                df = results[sheet_name].copy()

                for col in df.columns:
                    if df[col].dtype == object:
                        df[col] = df[col].apply(sanitize_excel_cell_value)

                df.to_excel(writer, sheet_name=sheet_name, index=False)

    wb = openpyxl.load_workbook(output_path)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.border = thin_border
                if isinstance(cell.value, (int, float, Decimal)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 50)

    wb.save(output_path)
    logger.info(f"Reconciliation workbook successfully saved to '{output_path}'")
    return output_path


def parse_cli_args(args: Optional[List[str]] = None) -> argparse.Namespace:
    """Configures CLI argument parsing."""
    parser = argparse.ArgumentParser(
        description="Enterprise Finance Reconciliation Engine CLI",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--erp", type=str, help="Path to ERP / System of Record dataset (CSV/XLSX)")
    parser.add_argument("--external", type=str, help="Path to External / Bank dataset (CSV/XLSX)")
    parser.add_argument("--output", type=str, default="output", help="Output directory path")
    parser.add_argument("--key", type=str, default=DEFAULT_KEY_COLUMN, help="Primary key column name")
    parser.add_argument(
        "--amount-tolerance",
        type=float,
        default=0.01,
        help="Monetary comparison tolerance threshold (e.g. 0.01)",
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Force interactive mode prompts for missing arguments",
    )
    return parser.parse_args(args)


def main():
    """CLI Entry Point."""
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    print("=" * 70)
    print(" Enterprise Finance Reconciliation Tool ")
    print("=" * 70)

    args = parse_cli_args()

    erp_path = args.erp
    external_path = args.external
    output_folder = args.output
    key_column = args.key
    tolerance = Decimal(str(args.amount_tolerance))

    if not erp_path or not external_path or args.interactive:
        erp_path = input("Enter ERP / System of Record file path: ").strip()
        external_path = input("Enter External / Bank file path: ").strip()
        output_folder_input = input(f"Enter output directory path [Default: '{output_folder}']: ").strip()
        if output_folder_input:
            output_folder = output_folder_input
        key_input = input(f"Enter Key Column name [Default: '{key_column}']: ").strip()
        if key_input:
            key_column = key_input

    try:
        logger.info("Ingesting datasets...")
        erp_hash = calculate_file_hash(erp_path)
        ext_hash = calculate_file_hash(external_path)

        df_erp_raw = read_spreadsheet(erp_path, source_name="ERP")
        df_ext_raw = read_spreadsheet(external_path, source_name="External")

        df_erp = preprocess_data(df_erp_raw, source_name="ERP")
        df_ext = preprocess_data(df_ext_raw, source_name="External")

        logger.info("Reconciling datasets...")
        results = reconcile_data(
            df_erp,
            df_ext,
            key_column=key_column,
            amount_tolerance=tolerance,
            source_erp_name="ERP",
            source_external_name="External",
        )

        metadata_info = {
            "ERP File Path": erp_path,
            "ERP File SHA-256": erp_hash,
            "External File Path": external_path,
            "External File SHA-256": ext_hash,
        }

        logger.info("Generating formatted Excel workbook...")
        output_file = write_output(results, output_folder, metadata_info=metadata_info)

        print("\n[SUCCESS] Reconciliation Completed Successfully!")
        print("\n[SUMMARY] Executive KPI Dashboard:")
        for _, row in results["summary"].iterrows():
            print(f"   - {row['Metric / Indicator']}: {row['Value']}")
        print(f"\n[OUTPUT] Auditable Workbook Saved To:\n   {output_file}\n")

    except ReconciliationError as re_err:
        logger.error(f"Reconciliation Validation Error: {re_err}")
        print(f"\n[ERROR] Reconciliation Error: {re_err}")
        sys.exit(1)
    except Exception as e:
        logger.exception(f"Unexpected System Error: {e}")
        print(f"\n[ERROR] Critical Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
