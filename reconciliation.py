"""
Enterprise Finance Reconciliation Tool
======================================

Purpose:
--------
Performs automated, deterministic, rule-driven financial reconciliation between an
Enterprise System of Record (ERP) and an External Source (Bank / Payment Gateway / Partner).

Key Principles:
---------------
1. Explicit Configuration: Zero substring-based field type inference. Field rules (money, text, date)
   are explicitly defined, validated, and serialized for auditability.
2. Raw Identifier Preservation: Preserves leading zeros and raw strings during ingestion.
3. Decimal Financial Precision: Uses decimal.Decimal and strict monetary syntax rules.
4. Strict Key & Duplicate Quarantine: Isolates invalid keys (MISSING_PRIMARY_KEY) and ambiguous duplicate
   keys (DUPLICATE_KEY_IN_ERP, DUPLICATE_KEY_IN_EXTERNAL, DUPLICATE_KEY_IN_BOTH) deterministically.
5. One-to-One Vector Outer Merge: Merges clean datasets under pandas validate="one_to_one" constraints.
6. Multi-Money Field & Currency Exposure: Computes per-field monetary variances and isolates cross-currency
   mismatches without invalid cross-currency variance summing.
7. Auditable Reporting: Multi-sheet Excel workbook with UTC timestamps, SHA-256 file/config hashes,
   source row lineage, formula injection hardening, and collision-resistant output paths.
"""

import argparse
import hashlib
import json
import logging
import os
import re
import sys
import uuid
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from importlib.metadata import PackageNotFoundError, version
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger("reconciliation")

SUPPORTED_EXTENSIONS = {".csv", ".xlsx"}
DEFAULT_TOLERANCE = Decimal("0.01")
DEFAULT_KEY_COLUMN = "invoice_id"
INTERNAL_PREFIX = "_recon_"

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

# Allowed currency codes and symbols for money parsing
CURRENCY_PATTERN = re.compile(
    r"[$₹€£¥]|Rs\.?|\b(USD|EUR|INR|GBP|CAD|AUD|JPY|CHF|CNY|HKD|NZD)\b",
    re.IGNORECASE,
)


class ReconciliationError(Exception):
    """Base exception for reconciliation failures."""

    pass


class ConfigurationError(ReconciliationError):
    """Raised when configuration parameters or field rules are invalid."""

    pass


class SchemaError(ReconciliationError):
    """Raised when dataset schemas or header validations fail."""

    pass


class IngestionError(ReconciliationError):
    """Raised when file reading or ingestion fails."""

    pass


class OutputError(ReconciliationError):
    """Raised when output file writing or overwrite checks fail."""

    pass


def get_tool_version() -> str:
    """Returns dynamic package version or fallback."""
    try:
        return version("enterprise-finance-reconciliation-tool")
    except PackageNotFoundError:
        return "1.0.0"
    except Exception:
        return "1.0.0"


def calculate_file_hash(file_path: str) -> str:
    """Calculates SHA-256 hash of a file for auditability."""
    sha256 = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            sha256.update(chunk)
    return sha256.hexdigest()


def read_spreadsheet(file_path: str, source_name: str = "Source") -> pd.DataFrame:
    """
    Reads a CSV or XLSX file into a pandas DataFrame.
    Preserves raw string identifiers (keep_default_na=False, dtype=str for CSV).
    Validates that no reserved internal '_recon_' columns pre-exist in input file.
    Attaches source row traceability metadata (_recon_source_row_<source_name>).
    """
    file_path = file_path.strip().strip('"').strip("'")
    if not os.path.exists(file_path):
        raise IngestionError(f"{source_name} file not found: '{file_path}'")

    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext not in SUPPORTED_EXTENSIONS:
        raise IngestionError(f"Unsupported file extension '{ext}' for {source_name}. Supported types: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")

    try:
        if ext == ".csv":
            df = pd.read_csv(file_path, dtype=str, keep_default_na=False)
        else:
            df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        raise IngestionError(f"Failed to read {source_name} file '{file_path}': {e}") from e

    # Reserved internal namespace check
    for col in df.columns:
        if str(col).startswith(INTERNAL_PREFIX):
            raise SchemaError(f"Input dataset {source_name} contains reserved internal column prefix '{INTERNAL_PREFIX}': '{col}'")

    row_col_name = f"{INTERNAL_PREFIX}source_row_{source_name.lower().replace(' ', '_')}"
    df[row_col_name] = range(2, len(df) + 2)

    return df


def sanitize_column_name(col: Any) -> str:
    """
    Normalizes column headers to lowercase snake_case.
    Example: 'Invoice-ID!' -> 'invoice_id'
    """
    s = str(col).strip().lower()
    s = re.sub(r"[^\w]+", "_", s)
    s = re.sub(r"_+", "_", s)
    s = s.strip("_")
    return s


def preprocess_data(df: pd.DataFrame, source_name: str = "Source") -> pd.DataFrame:
    """
    Standardizes the DataFrame:
    - Normalizes column names and checks for header collision errors.
    - Trims whitespace from string cells.
    - Drops completely empty rows (preserving source traceability).
    """
    df = df.copy()

    sanitized_map: Dict[str, List[str]] = {}
    new_columns = []

    for orig_col in df.columns:
        if str(orig_col).startswith(INTERNAL_PREFIX):
            new_columns.append(orig_col)
            continue
        clean_col = sanitize_column_name(orig_col)
        new_columns.append(clean_col)
        sanitized_map.setdefault(clean_col, []).append(str(orig_col))

    collisions = {k: v for k, v in sanitized_map.items() if len(v) > 1}
    if collisions:
        collision_details = "; ".join([f"'{k}': {v}" for k, v in collisions.items()])
        raise SchemaError(f"Header normalization collision detected in {source_name}: {collision_details}")

    df.columns = new_columns

    biz_cols = [c for c in df.columns if not c.startswith(INTERNAL_PREFIX)]
    df = df.dropna(how="all", subset=biz_cols)
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    return df


def parse_decimal(val: Any) -> Union[Decimal, None, str]:
    """
    Safely parses financial amounts into decimal.Decimal.

    Supports:
    - 1500, 1500.00, 1,500.00, $1,500.00, ₹1,500.00, €1,500.00, USD 1500.00
    - Negative amounts: -1500.00, (500.00), ($1,500.00)
    - Indian grouping: 1,23,456.78

    Returns:
    - Decimal object for valid monetary values
    - None for missing/null/blank values
    - INVALID_VALUE_SENTINEL for malformed non-financial values (ERROR123, ABC500XYZ, 12.3.4, 1,2,3)
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

    # Strip currency tokens
    without_currency = CURRENCY_PATTERN.sub("", val_str).strip()

    # Validate numeric string syntax BEFORE removing commas
    # Western grouping: 1,500.00 or 1500.00 | Indian grouping: 1,23,456.78 | Simple: 1500
    valid_syntax = (
        re.match(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$", without_currency)
        or re.match(r"^-?\d{1,3}(,\d{2})*(,\d{3})+(\.\d+)?$", without_currency)
        or re.match(r"^-?\d+(\.\d+)?$", without_currency)
    )

    if not valid_syntax:
        return INVALID_VALUE_SENTINEL

    cleaned = without_currency.replace(",", "")
    if is_parenthesized_negative and not cleaned.startswith("-"):
        cleaned = f"-{cleaned}"

    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return INVALID_VALUE_SENTINEL


def parse_date(val: Any, fmt: Optional[str] = "%Y-%m-%d") -> Union[pd.Timestamp, None, str]:
    """
    Safely parses dates into normalized pd.Timestamp.

    Returns:
    - pd.Timestamp (normalized to midnight) for valid dates
    - None for missing/null/blank values
    - INVALID_VALUE_SENTINEL for unparseable invalid date values
    """
    if val is None or pd.isna(val):
        return None

    if isinstance(val, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(val).normalize()

    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("nan", "none", "null", "<na>"):
        return None

    try:
        if fmt:
            parsed = pd.to_datetime(val_str, format=fmt, errors="coerce")
        else:
            parsed = pd.to_datetime(val_str, errors="coerce")

        if pd.isna(parsed):
            return INVALID_VALUE_SENTINEL
        return parsed.normalize()
    except Exception:
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
    with a single quote, unless it is a valid numeric value.
    """
    if isinstance(val, str) and val:
        if val[0] in ("=", "+", "-", "@"):
            if val[0] == "-" and re.match(r"^-?\d+(\.\d+)?$", val):
                return val
            return f"'{val}"
    return val


def get_default_field_rules(tolerance: Decimal) -> Dict[str, Dict[str, Any]]:
    """Returns explicit default field rules."""
    return {
        "amount": {"type": "money", "tolerance": tolerance},
        "currency": {"type": "text", "case_sensitive": False},
        "customer_name": {"type": "text", "case_sensitive": False},
        "date": {"type": "date", "format": "%Y-%m-%d"},
    }


def normalize_and_validate_field_rules(field_rules: Dict[str, Dict[str, Any]], key_column: str) -> Dict[str, Dict[str, Any]]:
    """
    Normalizes configured field names to snake_case exactly once.
    Validates rule types, tolerances, and parameter validity.
    """
    key_norm = sanitize_column_name(key_column)
    normalized_rules: Dict[str, Dict[str, Any]] = {}

    for orig_name, rule in field_rules.items():
        if not isinstance(rule, dict):
            raise ConfigurationError(f"Field rule for '{orig_name}' must be a dictionary.")

        norm_name = sanitize_column_name(orig_name)
        if norm_name == key_norm:
            raise ConfigurationError(f"Primary key column '{key_column}' cannot be included in comparison field_rules.")
        if norm_name in normalized_rules:
            raise ConfigurationError(f"Duplicate field rule after normalization: '{orig_name}' and another field both normalize to '{norm_name}'")

        f_type = str(rule.get("type", "")).strip().lower()
        if f_type not in ("money", "text", "date"):
            raise ConfigurationError(f"Invalid field type '{rule.get('type')}' for field '{orig_name}'. Allowed types: 'money', 'text', 'date'.")

        validated_rule: Dict[str, Any] = {"type": f_type}

        if f_type == "money":
            raw_tol = rule.get("tolerance", DEFAULT_TOLERANCE)
            try:
                tol_dec = Decimal(str(raw_tol))
            except (InvalidOperation, ValueError, TypeError):
                raise ConfigurationError(f"Invalid monetary tolerance '{raw_tol}' for field '{orig_name}'.")
            if tol_dec < Decimal("0"):
                raise ConfigurationError(f"Monetary tolerance for field '{orig_name}' cannot be negative: {tol_dec}")
            validated_rule["tolerance"] = tol_dec

        elif f_type == "text":
            case_sens = rule.get("case_sensitive", False)
            if not isinstance(case_sens, bool):
                raise ConfigurationError(f"Parameter 'case_sensitive' for field '{orig_name}' must be a boolean.")
            validated_rule["case_sensitive"] = case_sens

        elif f_type == "date":
            fmt = rule.get("format", "%Y-%m-%d")
            if fmt is not None and not isinstance(fmt, str):
                raise ConfigurationError(f"Parameter 'format' for field '{orig_name}' must be a string format.")
            validated_rule["format"] = fmt

        normalized_rules[norm_name] = validated_rule

    if not normalized_rules:
        raise ConfigurationError("Zero valid comparison field rules were provided.")

    return normalized_rules


def load_config_file(config_path: str) -> Tuple[Dict[str, Any], str, str]:
    """Loads configuration JSON file and returns (config_dict, key_column, config_hash)."""
    config_path = config_path.strip().strip('"').strip("'")
    if not os.path.exists(config_path):
        raise ConfigurationError(f"Configuration file not found: '{config_path}'")

    config_hash = calculate_file_hash(config_path)
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        raise ConfigurationError(f"Failed to parse JSON config file '{config_path}': {e}") from e

    return data, data.get("key_column", DEFAULT_KEY_COLUMN), config_hash


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
    run_timestamp_utc = datetime.now(timezone.utc).isoformat()

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
            f"Key column '{key_col}' not found in {source_erp_name}. Available columns: {[c for c in df_erp.columns if not c.startswith(INTERNAL_PREFIX)]}"
        )
    if key_col not in df_external.columns:
        raise SchemaError(
            f"Key column '{key_col}' not found in {source_external_name}. Available columns: {[c for c in df_external.columns if not c.startswith(INTERNAL_PREFIX)]}"
        )

    # 2. Field Rules Setup & Mandatory Validation
    if field_rules is None:
        all_defaults = get_default_field_rules(amount_tolerance)
        field_rules = {
            col: rule for col, rule in all_defaults.items() if sanitize_column_name(col) in df_erp.columns and sanitize_column_name(col) in df_external.columns
        }
        if not field_rules:
            common_cols = [c for c in df_erp.columns if c in df_external.columns and c != key_col and not c.startswith(INTERNAL_PREFIX)]
            field_rules = {c: {"type": "text"} for c in common_cols}

    normalized_rules = normalize_and_validate_field_rules(field_rules, key_col)

    # Mandatory Schema Validation: Ensure every configured rule field exists in both datasets
    comparison_fields = []
    for clean_c in normalized_rules.keys():
        if clean_c in df_erp.columns and clean_c in df_external.columns:
            comparison_fields.append(clean_c)
        else:
            missing_in = []
            if clean_c not in df_erp.columns:
                missing_in.append(source_erp_name)
            if clean_c not in df_external.columns:
                missing_in.append(source_external_name)
            raise SchemaError(f"Configured comparison field '{clean_c}' is missing in dataset(s): {', '.join(missing_in)}")

    if not comparison_fields:
        raise SchemaError("Zero valid comparison fields exist for reconciliation.")

    logger.info(f"Starting reconciliation run {run_id}. Key: '{key_col}'. Compared fields: {comparison_fields}")

    # 3. Attach Canonical Reconciliation Key (_recon_key)
    df_erp[f"{INTERNAL_PREFIX}key"] = df_erp[key_col].astype(str).str.strip()
    df_external[f"{INTERNAL_PREFIX}key"] = df_external[key_col].astype(str).str.strip()

    # 4. Primary Key Null / Invalid Quarantine
    if df_erp.empty:
        df_erp_valid = df_erp.copy()
        df_erp_invalid_keys = df_erp.copy()
    else:
        erp_valid_key_mask = df_erp[key_col].apply(is_valid_key)
        df_erp_invalid_keys = df_erp[~erp_valid_key_mask].copy()
        df_erp_valid = df_erp[erp_valid_key_mask].copy()

    if df_external.empty:
        df_ext_valid = df_external.copy()
        df_ext_invalid_keys = df_external.copy()
    else:
        ext_valid_key_mask = df_external[key_col].apply(is_valid_key)
        df_ext_invalid_keys = df_external[~ext_valid_key_mask].copy()
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

    # 5. Deterministic Duplicate Key Quarantine Logic
    if df_erp_valid.empty or key_col not in df_erp_valid.columns:
        erp_counts = pd.Series(dtype=int)
    else:
        erp_counts = df_erp_valid[key_col].value_counts()

    if df_ext_valid.empty or key_col not in df_ext_valid.columns:
        ext_counts = pd.Series(dtype=int)
    else:
        ext_counts = df_ext_valid[key_col].value_counts()

    erp_dupe_keys = set(erp_counts[erp_counts > 1].index)
    ext_dupe_keys = set(ext_counts[ext_counts > 1].index)
    all_dupe_keys_set = erp_dupe_keys.union(ext_dupe_keys)

    # Sort duplicate keys deterministically
    sorted_dupe_keys = sorted(list(all_dupe_keys_set), key=str)

    if df_erp_valid.empty or key_col not in df_erp_valid.columns:
        df_erp_dupes = pd.DataFrame(columns=df_erp.columns)
        df_erp_clean = df_erp_valid.copy()
    else:
        df_erp_dupes = df_erp_valid[df_erp_valid[key_col].isin(all_dupe_keys_set)].copy()
        df_erp_clean = df_erp_valid[~df_erp_valid[key_col].isin(all_dupe_keys_set)].copy()

    if df_ext_valid.empty or key_col not in df_ext_valid.columns:
        df_ext_dupes = pd.DataFrame(columns=df_external.columns)
        df_ext_clean = df_ext_valid.copy()
    else:
        df_ext_dupes = df_ext_valid[df_ext_valid[key_col].isin(all_dupe_keys_set)].copy()
        df_ext_clean = df_ext_valid[~df_ext_valid[key_col].isin(all_dupe_keys_set)].copy()

    dupe_records = []
    for k in sorted_dupe_keys:
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

    # 6. One-to-One Vector Outer Merge on Clean Records
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

    # 7. Field-Level Typed Comparison for Common Records
    common = merged[merged["_merge"] == "both"].copy()

    common["reconciliation_status"] = STATUS_MATCH
    common["status_reason"] = ""
    common["mismatch_fields"] = ""
    common["mismatch_count"] = 0

    # Initialize per-field variance columns for monetary fields
    for col in comparison_fields:
        rule = normalized_rules[col]
        if rule["type"] == "money":
            common[f"{col}_variance"] = None
            common[f"{col}_abs_difference"] = None

    statuses = []
    status_reasons = []
    mismatch_fields_list = []
    mismatch_counts = []
    currency_mismatch_flags = []

    for idx, row in common.iterrows():
        row_mismatched_fields = []
        row_reasons = []
        row_has_invalid = False
        is_currency_mismatch = False

        for col in comparison_fields:
            rule = normalized_rules[col]
            f_type = rule["type"]

            val_erp_raw = row.get(f"{col}{erp_suffix}")
            val_ext_raw = row.get(f"{col}{ext_suffix}")

            if f_type == "money":
                col_tol = rule.get("tolerance", amount_tolerance)
                dec_erp = parse_decimal(val_erp_raw)
                dec_ext = parse_decimal(val_ext_raw)

                if dec_erp == INVALID_VALUE_SENTINEL or dec_ext == INVALID_VALUE_SENTINEL:
                    row_has_invalid = True
                    row_reasons.append(f"{col} has invalid monetary format ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')")
                elif dec_erp is None and dec_ext is None:
                    common.at[idx, f"{col}_variance"] = Decimal("0")
                    common.at[idx, f"{col}_abs_difference"] = Decimal("0")
                elif dec_erp is None or dec_ext is None:
                    row_mismatched_fields.append(col)
                    row_reasons.append(f"{col} missing in one dataset ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')")
                else:
                    variance = dec_ext - dec_erp  # External - ERP
                    abs_diff = abs(variance)
                    common.at[idx, f"{col}_variance"] = variance
                    common.at[idx, f"{col}_abs_difference"] = abs_diff

                    if abs_diff > col_tol:
                        row_mismatched_fields.append(col)
                        row_reasons.append(f"{col} diff ({source_erp_name}: {dec_erp}, {source_external_name}: {dec_ext}, variance: {variance})")

            elif f_type == "text":
                case_sens = rule.get("case_sensitive", False)
                str_erp = str(val_erp_raw).strip() if val_erp_raw is not None and not pd.isna(val_erp_raw) else None
                str_ext = str(val_ext_raw).strip() if val_ext_raw is not None and not pd.isna(val_ext_raw) else None

                if str_erp is None and str_ext is None:
                    pass
                elif str_erp is None or str_ext is None:
                    row_mismatched_fields.append(col)
                    row_reasons.append(f"{col} missing in one dataset ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')")
                else:
                    c_erp = str_erp if case_sens else str_erp.lower()
                    c_ext = str_ext if case_sens else str_ext.lower()
                    if c_erp != c_ext:
                        row_mismatched_fields.append(col)
                        if col == "currency":
                            is_currency_mismatch = True
                        row_reasons.append(f"{col} mismatch ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')")

            elif f_type == "date":
                fmt = rule.get("format", "%Y-%m-%d")
                dt_erp = parse_date(val_erp_raw, fmt=fmt)
                dt_ext = parse_date(val_ext_raw, fmt=fmt)

                if dt_erp == INVALID_VALUE_SENTINEL or dt_ext == INVALID_VALUE_SENTINEL:
                    row_has_invalid = True
                    row_reasons.append(f"{col} has unparseable date format ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')")
                elif dt_erp is None and dt_ext is None:
                    pass
                elif dt_erp is None or dt_ext is None:
                    row_mismatched_fields.append(col)
                    row_reasons.append(f"{col} date missing in one dataset ({source_erp_name}: '{val_erp_raw}', {source_external_name}: '{val_ext_raw}')")
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
        currency_mismatch_flags.append(is_currency_mismatch)

    common["reconciliation_status"] = statuses
    common["status_reason"] = status_reasons
    common["mismatch_fields"] = mismatch_fields_list
    common["mismatch_count"] = mismatch_counts
    common["is_currency_mismatch"] = currency_mismatch_flags

    # 8. Aggregate Unified Category DataFrames
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
        df_data_issues = pd.DataFrame(columns=[key_col, "reconciliation_status", "status_reason", "source_system"])

    df_missing = missing_records.copy()

    # 9. Build Master Consolidated View (reconciliation_results / all_records)
    all_records_list = [common, df_missing]
    if invalid_key_records:
        all_records_list.extend(invalid_key_records)
    if dupe_records:
        all_records_list.extend(dupe_records)

    df_all_records = pd.concat(all_records_list, ignore_index=True)
    df_all_records = df_all_records.drop(columns=["_merge"], errors="ignore")

    primary_cols = [key_col, "reconciliation_status", "status_reason", "mismatch_fields", "mismatch_count"]
    other_cols = [c for c in df_all_records.columns if c not in primary_cols]
    df_all_records = df_all_records[primary_cols + other_cols]

    # 10. Compute Explicit KPIs and Currency-Isolated Exposure
    total_erp_ingested = len(df_erp)
    total_ext_ingested = len(df_external)
    eligible_count = len(common)
    matches_count = len(df_matches)
    mismatches_count = len(df_mismatches)
    missing_ext_count = len(df_missing[df_missing["reconciliation_status"] == STATUS_MISSING_IN_EXTERNAL])
    missing_erp_count = len(df_missing[df_missing["reconciliation_status"] == STATUS_MISSING_IN_ERP])

    data_issue_rows_count = len(df_data_issues)
    distinct_issue_keys_count = df_data_issues[key_col].nunique() if not df_data_issues.empty else 0
    currency_mismatch_count = common["is_currency_mismatch"].sum() if "is_currency_mismatch" in common.columns else 0

    match_rate = (matches_count / eligible_count * 100) if eligible_count > 0 else 0.0
    total_records = total_erp_ingested + total_ext_ingested
    unreconciled_count = total_records - (matches_count * 2)
    exception_rate = (unreconciled_count / total_records * 100) if total_records > 0 else 0.0

    summary_rows = [
        {"Metric / Indicator": "Total ERP Source Rows", "Value": total_erp_ingested},
        {"Metric / Indicator": "Total External Source Rows", "Value": total_ext_ingested},
        {"Metric / Indicator": "Eligible 1-to-1 Reconciliation Units", "Value": eligible_count},
        {"Metric / Indicator": "Fully Matched Reconciliation Units (MATCH)", "Value": matches_count},
        {"Metric / Indicator": "Mismatch Reconciliation Units (MISMATCH)", "Value": mismatches_count},
        {"Metric / Indicator": "Missing in External Units (MISSING_IN_EXTERNAL)", "Value": missing_ext_count},
        {"Metric / Indicator": "Missing in ERP Units (MISSING_IN_ERP)", "Value": missing_erp_count},
        {"Metric / Indicator": "Data Quality Exception Rows", "Value": data_issue_rows_count},
        {"Metric / Indicator": "Distinct Data Quality Exception Keys", "Value": distinct_issue_keys_count},
        {"Metric / Indicator": "Currency Mismatch Count", "Value": int(currency_mismatch_count)},
        {"Metric / Indicator": "Eligible Match Rate (%)", "Value": f"{match_rate:.2f}%"},
        {"Metric / Indicator": "Reconciliation Unit Exception Rate (%)", "Value": f"{exception_rate:.2f}%"},
    ]

    # Calculate monetary exposure ONLY when currencies are identical
    if "currency" in comparison_fields:
        curr_erp_col = f"currency_{source_erp_name.lower()}"
        curr_ext_col = f"currency_{source_external_name.lower()}"

        if not df_mismatches.empty:
            same_currency_mismatches = df_mismatches[df_mismatches[curr_erp_col].str.lower() == df_mismatches[curr_ext_col].str.lower()]
            if not same_currency_mismatches.empty and "amount_abs_difference" in same_currency_mismatches.columns:
                for curr in same_currency_mismatches[curr_erp_col].dropna().unique():
                    c_df = same_currency_mismatches[same_currency_mismatches[curr_erp_col] == curr]
                    exposure = c_df["amount_abs_difference"].dropna().sum()
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
                    sum_amt = c_df[amt_erp_col].apply(parse_decimal).map(lambda x: x if isinstance(x, Decimal) else Decimal("0")).sum()
                    summary_rows.append(
                        {
                            "Metric / Indicator": f"Missing in External Exposure ({curr})",
                            "Value": f"{sum_amt:.2f}",
                        }
                    )

            missing_erp_df = df_missing[df_missing["reconciliation_status"] == STATUS_MISSING_IN_ERP]
            if not missing_erp_df.empty and amt_ext_col in missing_erp_df.columns:
                for curr in missing_erp_df[curr_ext_col].dropna().unique():
                    c_df = missing_erp_df[missing_erp_df[curr_ext_col] == curr]
                    sum_amt = c_df[amt_ext_col].apply(parse_decimal).map(lambda x: x if isinstance(x, Decimal) else Decimal("0")).sum()
                    summary_rows.append(
                        {
                            "Metric / Indicator": f"Missing in ERP Exposure ({curr})",
                            "Value": f"{sum_amt:.2f}",
                        }
                    )

    df_summary = pd.DataFrame(summary_rows)

    # 11. Run Metadata Sheet
    serialized_rules = json.dumps(
        {k: {rk: str(rv) for rk, rv in v.items()} for k, v in normalized_rules.items()},
        indent=2,
    )

    df_metadata = pd.DataFrame(
        [
            {"Parameter": "Run ID", "Value": run_id},
            {"Parameter": "Run Timestamp UTC", "Value": run_timestamp_utc},
            {"Parameter": "Tool Version", "Value": get_tool_version()},
            {"Parameter": "Primary Key Column", "Value": key_col},
            {"Parameter": "Amount Tolerance Threshold", "Value": str(amount_tolerance)},
            {"Parameter": "Compared Fields", "Value": ", ".join(comparison_fields)},
            {"Parameter": "Serialized Field Rules", "Value": serialized_rules},
            {"Parameter": f"{source_erp_name} Total Ingested Rows", "Value": total_erp_ingested},
            {"Parameter": f"{source_external_name} Total Ingested Rows", "Value": total_ext_ingested},
            {"Parameter": "Eligible 1-to-1 Reconciliation Units", "Value": eligible_count},
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
    overwrite: bool = False,
) -> str:
    """
    Writes reconciliation results into a multi-sheet Excel workbook.
    Prevents silent file overwrites unless overwrite=True.
    """
    output_folder = output_folder.strip().strip('"').strip("'")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder, exist_ok=True)

    if not output_filename:
        timestamp_str = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")
        short_id = str(uuid.uuid4())[:8]
        output_filename = f"reconciliation_{timestamp_str}_{short_id}.xlsx"

    output_path = os.path.join(output_folder, output_filename)

    if os.path.exists(output_path) and not overwrite:
        raise OutputError(f"Output file '{output_path}' already exists. Use --overwrite flag to explicitly overwrite.")

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

                # Apply formula injection protection to string cells regardless of pandas dtype
                for col in df.columns:
                    df[col] = df[col].map(lambda value: sanitize_excel_cell_value(value) if isinstance(value, str) else value)

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
    parser.add_argument("--config", type=str, help="Path to JSON reconciliation configuration file")
    parser.add_argument("--output", type=str, default="output", help="Output directory path")
    parser.add_argument("--key", type=str, default=DEFAULT_KEY_COLUMN, help="Primary key column name")
    parser.add_argument(
        "--amount-tolerance",
        type=float,
        default=0.01,
        help="Monetary comparison tolerance threshold (e.g. 0.01)",
    )
    parser.add_argument("--overwrite", action="store_true", help="Overwrite output file if it already exists")
    parser.add_argument("--interactive", action="store_true", help="Force interactive mode prompts for missing arguments")
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
    config_path = args.config
    output_folder = args.output
    key_column = args.key
    tolerance = Decimal(str(args.amount_tolerance))
    field_rules = None
    config_hash = None

    if config_path:
        config_data, key_column_from_config, config_hash = load_config_file(config_path)
        if "key_column" in config_data:
            key_column = config_data["key_column"]
        if "field_rules" in config_data:
            field_rules = config_data["field_rules"]

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
            field_rules=field_rules,
            source_erp_name="ERP",
            source_external_name="External",
        )

        metadata_info = {
            "ERP File Path": erp_path,
            "ERP File SHA-256": erp_hash,
            "External File Path": external_path,
            "External File SHA-256": ext_hash,
        }
        if config_path:
            metadata_info["Config File Path"] = config_path
            metadata_info["Config File SHA-256"] = config_hash

        logger.info("Generating formatted Excel workbook...")
        output_file = write_output(results, output_folder, metadata_info=metadata_info, overwrite=args.overwrite)

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
