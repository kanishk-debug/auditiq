# ============================================================
# AUDITIQ — PHASE 5A: GSTR-2B PARSER
# Production-ready SaaS component
# ============================================================

import pandas as pd
import hashlib
import re as _re
from pathlib import Path

GSTR2B_SCHEMA = [
    "supplier_gstin",        # 0
    "supplier_name",         # 1
    "invoice_number",        # 2
    "invoice_type",          # 3
    "invoice_date",          # 4
    "invoice_value",         # 5
    "place_of_supply",       # 6
    "rcm_flag",              # 7
    "taxable_value",         # 8
    "igst_amount",           # 9
    "cgst_amount",           # 10
    "sgst_amount",           # 11
    "cess_amount",           # 12
    "gstr1_period",          # 13
    "gstr1_filing_date",     # 14
    "itc_availability",      # 15
    "reason",                # 16
    "applicable_tax_rate",   # 17
    "source",                # 18
    "irn",                   # 19
    "irn_date",              # 20
    "ca_remarks_raw",        # 21
]

NUMERIC_COLS_2B = [
    "invoice_value", "taxable_value",
    "igst_amount", "cgst_amount", "sgst_amount", "cess_amount"
]

CA_REMARK_MAP = {
    "done": "MATCHED",
    "ineligible": "INELIGIBLE",
    "rcm": "RCM",
}


# ── LAYER 1 ───────────────────────────────────────────────────
def _2b_load_file(filepath: str, sheet_name: str = "B2B") -> pd.DataFrame:
    path = Path(filepath)
    if not path.exists():
        raise ValueError(f"File not found: {filepath}")
    try:
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Cannot open file or sheet '{sheet_name}': {e}")
    return df


# ── LAYER 2 ───────────────────────────────────────────────────
def _2b_detect_structure(df: pd.DataFrame, sheet_name: str) -> tuple:
    format_warnings = []
    headers = df.iloc[4].tolist()
    data = df.iloc[6:].reset_index(drop=True)
    actual_cols = len(headers)
    # Minimum 17 — cols 0-16 are actual GST data.
    # Cols 17-21 (IRN, Source, Remarks) are optional.
    # Remarks (col 21) only exists if CA manually added it —
    # which they won't when using AuditIQ as their reconciliation engine.
    if actual_cols < 17:
        raise ValueError(
            f"CRITICAL: Expected ≥17 columns, found {actual_cols}. "
            "File may be wrong format or corrupted."
        )
    if actual_cols > 22:
        format_warnings.append(
            f"FORMAT_WARNING: File has {actual_cols} columns, expected 21-22. Extra columns ignored."
        )
    return data, format_warnings


# ── LAYER 3 ───────────────────────────────────────────────────
def _2b_standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Pad to 22 columns if optional columns are missing.
    # Fresh portal download = 21 cols (no Remarks).
    # ca_remarks_raw will be None → _map_ca_remark returns UNCLASSIFIED → correct.
    while df.shape[1] < 22:
        df[f"_pad_{df.shape[1]}"] = None
    df = df.iloc[:, :22].copy()
    df.columns = GSTR2B_SCHEMA
    df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")], errors="ignore")
    df = df.reset_index(drop=True)
    df["source_row"] = df.index + 7
    return df


# ── LAYER 4 ───────────────────────────────────────────────────
def _2b_normalize_invoice_number(inv) -> str:
    if pd.isna(inv) or inv is None:
        return ""
    return str(inv).lower().replace(" ", "").replace("/", "").replace("-", "").replace(".", "")


def _2b_normalize_gstin(gstin) -> str:
    if pd.isna(gstin) or gstin is None:
        return ""
    return str(gstin).strip().upper()


def _map_ca_remark(raw) -> str:
    if pd.isna(raw) or str(raw).strip() == "":
        return "UNCLASSIFIED"
    normalized = str(raw).strip().lower()
    if normalized in CA_REMARK_MAP:
        return CA_REMARK_MAP[normalized]
    if "not in books" in normalized:
        return "MISSING_IN_BOOKS"
    if "previous month" in normalized:
        return "PREVIOUS_MONTH"
    return "UNCLASSIFIED"


def _2b_validate_rows(df: pd.DataFrame) -> tuple:
    error_records = []
    flags_list = []
    severity_list = []
    seen_invoices = {}

    for idx, row in df.iterrows():
        flags = []
        source_row = row["source_row"]
        supplier = row.get("supplier_name", "")
        inv_raw = row.get("invoice_number", "")
        gstin_raw = row.get("supplier_gstin", "")

        # 1. GSTIN
        if pd.isna(gstin_raw) or str(gstin_raw).strip() == "":
            flags.append("MISSING_GSTIN")
            error_records.append({
                "source_row": source_row, "supplier_name": supplier,
                "invoice_number": inv_raw, "flag": "MISSING_GSTIN",
                "issue": "Supplier GSTIN is missing", "severity": "WARNING"
            })

        # 2. Invoice number
        if pd.isna(inv_raw) or str(inv_raw).strip() == "":
            flags.append("MISSING_INVOICE_NUMBER")
            error_records.append({
                "source_row": source_row, "supplier_name": supplier,
                "invoice_number": inv_raw, "flag": "MISSING_INVOICE_NUMBER",
                "issue": "Invoice number is missing", "severity": "WARNING"
            })

        # 3. Date
        date_val = row.get("invoice_date")
        if pd.isna(date_val) or str(date_val).strip() == "":
            flags.append("MISSING_DATE")
            error_records.append({
                "source_row": source_row, "supplier_name": supplier,
                "invoice_number": inv_raw, "flag": "MISSING_DATE",
                "issue": "Invoice date is missing", "severity": "WARNING"
            })
        else:
            try:
                pd.to_datetime(date_val, dayfirst=True)
            except Exception:
                flags.append("BAD_DATE")
                error_records.append({
                    "source_row": source_row, "supplier_name": supplier,
                    "invoice_number": inv_raw, "flag": "BAD_DATE",
                    "issue": f"Cannot parse date: {date_val}", "severity": "WARNING"
                })

        # 4. Tax numeric check
        for tax_col in ["igst_amount", "cgst_amount", "sgst_amount"]:
            val = row.get(tax_col)
            converted = pd.to_numeric(val, errors="coerce")
            if pd.notna(val) and str(val).strip() != "" and pd.isna(converted):
                flags.append("NON_NUMERIC_VALUE")
                error_records.append({
                    "source_row": source_row, "supplier_name": supplier,
                    "invoice_number": inv_raw, "flag": "NON_NUMERIC_VALUE",
                    "issue": f"Non-numeric tax value in {tax_col}: {val}", "severity": "WARNING"
                })

        # 5. Zero tax check
        igst = pd.to_numeric(row.get("igst_amount"), errors="coerce") or 0
        cgst = pd.to_numeric(row.get("cgst_amount"), errors="coerce") or 0
        sgst = pd.to_numeric(row.get("sgst_amount"), errors="coerce") or 0
        inv_type = str(row.get("invoice_type", "")).lower()
        rcm = str(row.get("rcm_flag", "")).strip().upper()
        if igst == 0 and cgst == 0 and sgst == 0:
            if "credit note" not in inv_type and rcm != "YES":
                flags.append("ZERO_TAX")
                error_records.append({
                    "source_row": source_row, "supplier_name": supplier,
                    "invoice_number": inv_raw, "flag": "ZERO_TAX",
                    "issue": "All tax values are zero (not a credit note or RCM)",
                    "severity": "WARNING"
                })

        # 6. CA Remarks
        raw_remark = row.get("ca_remarks_raw")
        ca_status = _map_ca_remark(raw_remark)
        if not (pd.isna(raw_remark) or str(raw_remark).strip() == "") and ca_status == "UNCLASSIFIED":
            flags.append("UNRECOGNIZED_REMARK")
            error_records.append({
                "source_row": source_row, "supplier_name": supplier,
                "invoice_number": inv_raw, "flag": "UNRECOGNIZED_REMARK",
                "issue": f"Unrecognized CA remark: '{raw_remark}'", "severity": "WARNING"
            })

        # 7. Duplicate detection
        inv_norm = _2b_normalize_invoice_number(inv_raw)
        gstin_norm = _2b_normalize_gstin(gstin_raw)
        dup_key = f"{gstin_norm}||{inv_norm}"
        if inv_norm and dup_key in seen_invoices:
            flags.append("DUPLICATE_INVOICE")
            error_records.append({
                "source_row": source_row, "supplier_name": supplier,
                "invoice_number": inv_raw, "flag": "DUPLICATE_INVOICE",
                "issue": f"Duplicate of row {seen_invoices[dup_key]}", "severity": "WARNING"
            })
        elif inv_norm:
            seen_invoices[dup_key] = source_row

        if not flags:
            severity = "OK"
        else:
            severity = "ERROR" if "MISSING_INVOICE_NUMBER" in flags else "WARNING"

        flags_list.append(flags)
        severity_list.append(severity)

    df = df.copy()
    df["validation_flags"] = flags_list
    df["severity"] = severity_list
    error_log = pd.DataFrame(error_records)
    return df, error_log


# ── LAYER 5 ───────────────────────────────────────────────────
def _2b_clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["supplier_gstin"] = df["supplier_gstin"].apply(_2b_normalize_gstin)
    str_cols = ["supplier_name", "invoice_number", "invoice_type",
                "place_of_supply", "itc_availability", "reason",
                "source", "irn", "ca_remarks_raw"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
    for col in NUMERIC_COLS_2B:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for date_col in ["invoice_date", "gstr1_filing_date", "irn_date"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
    df["rcm_flag"] = df["rcm_flag"].apply(
        lambda x: True if str(x).strip().upper() == "YES" else False
    )
    return df


# ── LAYER 6 ───────────────────────────────────────────────────
def _2b_enrich(df: pd.DataFrame, filepath: str) -> pd.DataFrame:
    df = df.copy()
    df["invoice_number_norm"]  = df["invoice_number"].apply(_2b_normalize_invoice_number)
    df["ca_status"]            = df["ca_remarks_raw"].apply(_map_ca_remark)
    df["source_file"]          = str(filepath)
    df["is_footer_row"]        = False
    df["source"]               = "2B"
    df["supplier_name_norm"]   = df["supplier_name"].apply(
        lambda x: " ".join(str(x).lower().split()) if pd.notna(x) else ""
    )
    df["invoice_number_tail"]  = df["invoice_number_norm"].apply(
        lambda n: (_re.findall(r'\d{2,}', str(n)) or [""])[-1]
    )
    return df


# ── LAYER 7 ───────────────────────────────────────────────────
def _2b_build_report(df: pd.DataFrame, error_log: pd.DataFrame,
                     filepath: str, format_warnings: list) -> dict:
    from collections import Counter
    total_rows   = len(df)
    clean_rows   = int((df["severity"] == "OK").sum())
    flagged_rows = total_rows - clean_rows
    flag_counts  = Counter()
    for flags in df["validation_flags"]:
        flag_counts.update(flags)
    ca_status_summary = df["ca_status"].value_counts().to_dict() if "ca_status" in df.columns else {}
    rcm_count = int(df["rcm_flag"].sum()) if "rcm_flag" in df.columns else 0
    credit_note_count = int(
        df["invoice_type"].str.lower().str.contains("credit note", na=False).sum()
    ) if "invoice_type" in df.columns else 0
    return {
        "format_version":   "GSTR2B_v1.0",
        "file":             str(filepath),
        "sheet":            "B2B",
        "total_rows":       total_rows,
        "clean_rows":       clean_rows,
        "flagged_rows":     flagged_rows,
        "dropped_rows":     0,
        "summary_warnings": dict(flag_counts),
        "format_warnings":  format_warnings,
        "ca_status_summary":ca_status_summary,
        "rcm_count":        rcm_count,
        "credit_note_count":credit_note_count,
    }


# ── ENTRY POINT ────────────────────────────────────────────────
def parse_gstr2b(filepath: str, sheet_name: str = "B2B") -> dict:
    raw_df                     = _2b_load_file(filepath, sheet_name)
    data_df, format_warnings   = _2b_detect_structure(raw_df, sheet_name)
    data_df                    = _2b_standardize_columns(data_df)
    data_df, error_log         = _2b_validate_rows(data_df)
    data_df                    = _2b_clean_data(data_df)
    data_df                    = _2b_enrich(data_df, filepath)
    parse_report               = _2b_build_report(data_df, error_log, filepath, format_warnings)
    return {"data": data_df, "error_log": error_log, "parse_report": parse_report}

# removed quick test block
