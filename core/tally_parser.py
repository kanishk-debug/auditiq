# ============================================================
# AUDITIQ — PHASE 5B: TALLY EXCEL PARSER  (v1.3 — RCM sheet fix)
# Changes from v1.2:
#   RCM-BOOKS: _load_sheets now READS RCM payable sheets (was just naming them)
#   RCM-BOOKS: _parse_rcm_sheet() added — parses liability journal entries
#   RCM-BOOKS: parse_tally() merges RCM rows into output with bucket=RCM_BOOKS
# All v1.2 changes retained unchanged
# ============================================================
 
import re
import hashlib
import pandas as pd
from pathlib import Path
from datetime import datetime, timezone
from collections import Counter
 
# ── Output schema ──────────────────────────────────────────────
EXPECTED_OUTPUT_COLUMNS = [
    "source_row", "source_sheet", "source_file", "source",
    "invoice_date", "vendor_name", "invoice_number", "narration",
    "total_amount", "original_amount", "igst", "cgst", "sgst",
    "invoice_number_norm",
    "invoice_number_tail",       # FIX 2: numeric tail for fuzzy matching
    "inv_from_narration",        # FIX 2: flag for 5C to know extraction source
    "vendor_name_norm",
    "ca_tally_remark",           # FIX 3: CA's own pre-classification from Tally
    "ca_tally_status",           # FIX 3: normalized enum of above
    "is_itc_claim_entry",        # FIX 1: Razorpay-style ITC-only rows
    "is_rcm", "rcm_hint",
    "is_footer_row", "is_empty_row",
    "validation_flags", "severity", "file_hash",
]
 
_INTERNAL_COLS = ["_inv_ambiguous"]
 
# Column map — priority order, more specific first
COLUMN_MAP_PRIORITY = [
    ("voucher no.",   "invoice_number"),
    ("voucher no",    "invoice_number"),
    ("particulars",   "vendor_name"),
    ("gross total",   "total_amount"),
    ("igst input",    "igst"),
    ("cgst input",    "cgst"),
    ("sgst input",    "sgst"),
    ("narration",     "narration"),
    ("date",          "invoice_date"),
    ("remarks",       "ca_tally_remark"),   # FIX 3
    ("igst",          "igst"),
    ("cgst",          "cgst"),
    ("sgst",          "sgst"),
]
 
FOOTER_LABELS = {"total", "grand total"}
NUMERIC_COLS  = ["total_amount", "igst", "cgst", "sgst"]
 
INV_PATTERN = re.compile(
    r'(?:invoice\s*no\.?|inv\.?\s*no\.?|inv\.?)\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9/\-]{3,})',
    re.IGNORECASE
)
FALLBACK_PATTERN = re.compile(r'\b([A-Za-z0-9]{2,}[/\-][A-Za-z0-9/\-]{2,})\b')
 
REJECT_TOKENS = {
    "jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec",
    "2024","2025","2026","2027",
    "input","gst","rcm","tds","igst","cgst","sgst",
}
 
RCM_NARRATION_KEYWORDS = ["rcm", "reverse charge", "reverse charg"]
 
# FIX 3 — CA Tally remark normalization (mirrors 5A logic)
CA_TALLY_REMARK_MAP = {
    "done":              "MATCHED",
    "ineligible":        "INELIGIBLE",
    "rcm":               "RCM",
}
def _map_ca_tally_remark(raw) -> str:
    if pd.isna(raw) or str(raw).strip() == "":
        return "UNCLASSIFIED"
    n = str(raw).strip().lower()
    if n in CA_TALLY_REMARK_MAP:
        return CA_TALLY_REMARK_MAP[n]
    if "not in gstr" in n or "not in 2b" in n:
        return "MISSING_IN_2B"
    if "not in books" in n:
        return "MISSING_IN_BOOKS"
    if "previous month" in n:
        return "PREVIOUS_MONTH"
    return "UNCLASSIFIED"
 
 
# ── File identity ──────────────────────────────────────────────
def _compute_file_hash(filepath: str) -> str:
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()
 
 
# ── Helpers ────────────────────────────────────────────────────
def _normalize_invoice_number(inv) -> str:
    if pd.isna(inv) or str(inv).strip() in ("", "nan"):
        return ""
    return str(inv).lower().replace(" ","").replace("/","").replace("-","").replace(".","")
 
def _extract_invoice_tail(inv_norm: str) -> str:
    """
    FIX 2: Extract trailing numeric portion for fuzzy matching.
    e.g. 'gt2526 1176' → '1176', 'cipl fb26cyg002' → '002'
    Lets 5C match 'GT/25-26/1176' (books) with '1176' (2B).
    """
    if not inv_norm:
        return ""
    digits = re.findall(r'\d{2,}', inv_norm)
    return digits[-1] if digits else ""
 
def _normalize_vendor_name(name) -> str:
    if pd.isna(name) or str(name).strip() in ("", "nan"):
        return ""
    return " ".join(str(name).lower().split())
 
def _is_valid_invoice_token(token: str) -> bool:
    token = token.strip()
    if len(token) < 5:
        return False
    if token.isdigit():
        return False
    if token.lower() in REJECT_TOKENS:
        return False
    return True
 
def _extract_invoice_from_narration(narration) -> tuple[str, bool]:
    if pd.isna(narration) or str(narration).strip() == "":
        return "", False
    text = str(narration)
    primary = [t for t in INV_PATTERN.findall(text) if _is_valid_invoice_token(t)]
    if len(primary) == 1:
        return primary[0].strip(), False
    if len(primary) > 1:
        return max(primary, key=len).strip(), True
    fallback = [t for t in FALLBACK_PATTERN.findall(text) if _is_valid_invoice_token(t)]
    if len(fallback) == 1:
        return fallback[0].strip(), False
    if len(fallback) > 1:
        return max(fallback, key=len).strip(), True
    return "", False
 
 
# ── LAYER 1: File Loading ──────────────────────────────────────
# CHANGED: now returns 4 values. RCM sheets are READ, not just named.
def _load_sheets(filepath: str) -> tuple[dict, list, list, dict]:
    """
    Returns:
      sheets_loaded    — IGST/CGST ITC input sheets (go through normal pipeline)
      sheets_skipped   — irrelevant sheets (Summary etc.)
      rcm_sheet_names  — names of RCM sheets found (for reporting)
      rcm_sheets_data  — {sheet_name: raw_df} for RCM sheets (NOW ACTUALLY READ)
    """
    path = Path(filepath)
    if not path.exists():
        raise ValueError(f"File not found: {filepath}")
    try:
        xl = pd.ExcelFile(filepath, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Cannot open file: {e}")
 
    all_names      = xl.sheet_names
    sheets_loaded  = {}
    sheets_skipped = []
    rcm_sheet_names= []
    rcm_sheets_data= {}
 
    for sheet in all_names:
        s = sheet.lower()
        if "rcm" in s:
            # READ the sheet — these are RCM liability entries
            rcm_sheet_names.append(sheet)
            rcm_sheets_data[sheet] = pd.read_excel(
                filepath, sheet_name=sheet, header=None, engine="openpyxl"
            )
        elif "igst" in s or "cgst" in s:
            sheets_loaded[sheet] = pd.read_excel(
                filepath, sheet_name=sheet, header=None, engine="openpyxl"
            )
        else:
            sheets_skipped.append(sheet)
 
    if not sheets_loaded:
        raise ValueError(
            f"CRITICAL: No valid IGST/CGST sheets found. Available: {all_names}"
        )
    return sheets_loaded, sheets_skipped, rcm_sheet_names, rcm_sheets_data
 
 
# ── NEW: RCM Sheet Parser ──────────────────────────────────────
def _parse_rcm_sheet(sheet_name: str, raw_df: pd.DataFrame,
                     filepath: str, file_hash: str) -> pd.DataFrame:
    """
    Parses Tally RCM payable sheet (e.g. 'IGST @ 18% RCM Payable').
    These are RCM LIABILITY entries — company must pay this GST to govt.
    They are NOT ITC input entries — never run through matching engine.
    Output goes directly to RCM_BOOKS bucket in reconcile().
 
    Uses the same auto-detect header logic as _detect_structure so it
    handles any row offset without hardcoding.
    """
    recognizable = {"date", "voucher", "particulars", "igst", "gross"}
 
    header_row_idx = None
    for i in range(min(6, len(raw_df))):
        row_lower = [str(v).lower().strip() for v in raw_df.iloc[i].tolist()]
        score = sum(1 for cell in row_lower if any(k in cell for k in recognizable))
        if score >= 2:
            header_row_idx = i
            break
 
    if header_row_idx is None:
        print(f"  ⚠️  Could not detect header in RCM sheet '{sheet_name}' — skipping")
        return pd.DataFrame()
 
    headers = [
        str(v).strip() if pd.notna(v) else f"col_{i}"
        for i, v in enumerate(raw_df.iloc[header_row_idx])
    ]
    data = raw_df.iloc[header_row_idx + 1:].copy()
    data.columns = headers
    data = data.reset_index(drop=True)
 
    # Map to standard column names using same priority logic as IGST/CGST sheets
    col_map = {}
    mapped = set()
    for raw_col in data.columns:
        c = str(raw_col).lower().strip()
        for key, target in COLUMN_MAP_PRIORITY:
            if key in c and target not in mapped:
                col_map[raw_col] = target
                mapped.add(target)
                break
    data = data.rename(columns=col_map)
 
    # Ensure all required columns exist
    for col in ["invoice_date", "vendor_name", "invoice_number",
                "narration", "total_amount", "igst", "cgst", "sgst"]:
        if col not in data.columns:
            data[col] = None
 
    # Drop footer / grand total rows
    data = data[~data["vendor_name"].astype(str).str.lower().str.strip().isin(
        {"grand total", "total", "nan", ""}
    )].copy()
 
    # Drop fully empty rows
    data = data[~(
        data["invoice_date"].isna() &
        data["vendor_name"].isna() &
        data["invoice_number"].isna()
    )].copy()
 
    if data.empty:
        print(f"  ⚠️  RCM sheet '{sheet_name}' — no data rows after filtering")
        return pd.DataFrame()
 
    # Clean numeric and date columns
    data["igst"]         = pd.to_numeric(data["igst"],         errors="coerce").fillna(0)
    data["cgst"]         = pd.to_numeric(data["cgst"],         errors="coerce").fillna(0)
    data["sgst"]         = pd.to_numeric(data["sgst"],         errors="coerce").fillna(0)
    data["total_amount"] = pd.to_numeric(data["total_amount"], errors="coerce").fillna(0).abs()
    data["invoice_date"] = pd.to_datetime(data["invoice_date"], dayfirst=True, errors="coerce")
 
    # Tag all rows — these BYPASS the matching engine entirely
    data["source"]             = "BOOKS_RCM"
    data["source_sheet"]       = sheet_name
    data["source_file"]        = str(filepath)
    data["file_hash"]          = file_hash
    data["bucket"]             = "RCM_BOOKS"   # 5C _preprocess will read this
    data["is_rcm"]             = True
    data["rcm_hint"]           = True
    data["is_footer_row"]      = False
    data["is_empty_row"]       = False
    data["is_itc_claim_entry"] = False
    data["inv_from_narration"] = False
    data["validation_flags"]   = [[] for _ in range(len(data))]
    data["severity"]           = "OK"
    data["original_amount"]    = data["total_amount"]
    data["narration"]          = data.get("narration", pd.Series([None] * len(data)))
    data["ca_tally_remark"]    = None
    data["ca_tally_status"]    = "RCM"
    data["invoice_number_norm"]= data["invoice_number"].apply(_normalize_invoice_number)
    data["invoice_number_tail"]= data["invoice_number_norm"].apply(_extract_invoice_tail)
    data["vendor_name_norm"]   = data["vendor_name"].apply(_normalize_vendor_name)
    data["source_row"]         = data.index + header_row_idx + 2
 
    total_igst = data["igst"].sum()
    print(f"  📋 RCM sheet '{sheet_name}': {len(data)} entries, "
          f"total IGST ₹{total_igst:,.2f}")
    return data
 
 
# ── LAYER 2: Structure Detection ──────────────────────────────
def _detect_structure(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    recognizable = {k for k, _ in COLUMN_MAP_PRIORITY}
    best_row_idx, best_score = None, 0
    for i in range(min(5, len(df))):
        row_vals = [str(v).lower().strip() for v in df.iloc[i].tolist()]
        score = sum(1 for cell in row_vals if any(key in cell for key in recognizable))
        if score > best_score:
            best_score, best_row_idx = score, i
    if best_row_idx is None or best_score == 0:
        raise ValueError(f"CRITICAL: Could not detect header row in sheet '{sheet_name}'.")
    headers = [str(v).strip() if pd.notna(v) else f"col_{i}"
               for i, v in enumerate(df.iloc[best_row_idx])]
    data = df.iloc[best_row_idx + 1:].reset_index(drop=True)
    data.columns = headers
    col_str = " ".join(data.columns).lower()
    if "voucher no" not in col_str:
        raise ValueError(
            f"CRITICAL: 'Voucher No.' not found after header detection in '{sheet_name}'. "
            f"Columns: {list(data.columns)}"
        )
    data["source_sheet"] = sheet_name
    data["source_row"]   = data.index + best_row_idx + 2
    return data
 
 
# ── LAYER 3: Column Standardization ───────────────────────────
def _standardize_columns(df: pd.DataFrame, sheet_name: str) -> tuple[pd.DataFrame, list]:
    warnings      = []
    rename_map    = {}
    mapped_targets= set()
    for raw_col in df.columns:
        if raw_col in ("source_sheet", "source_row"):
            continue
        col_normalized = str(raw_col).lower().strip()
        for key, target in COLUMN_MAP_PRIORITY:
            if key in col_normalized and target not in mapped_targets:
                rename_map[raw_col] = target
                mapped_targets.add(target)
                break
    df = df.rename(columns=rename_map)
    required_defaults = {
        "invoice_date": None, "vendor_name": None,
        "invoice_number": None, "narration": None,
        "total_amount": 0, "igst": 0, "cgst": 0, "sgst": 0,
        "ca_tally_remark": None,   # FIX 3 — default None if no remarks col
    }
    for col, default in required_defaults.items():
        if col not in df.columns:
            df[col] = default
            if col != "ca_tally_remark":  # silence remark warning, it's optional
                warnings.append(
                    f"MISSING_COLUMN: '{col}' not in '{sheet_name}', defaulted to {repr(default)}"
                )
    if "vendor_name" not in mapped_targets and "total_amount" not in mapped_targets:
        raise ValueError(
            f"CRITICAL: Both vendor_name and total_amount missing after mapping in '{sheet_name}'. "
            f"Raw cols: {list(df.columns)}"
        )
    return df, warnings
 
 
# ── Special Row Detection ──────────────────────────────────────
def _detect_special_rows(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
 
    # Footer rows
    df["is_footer_row"] = df["vendor_name"].apply(
        lambda x: str(x).strip().lower() in FOOTER_LABELS if pd.notna(x) else False
    )
 
    # FIX 1 — ITC_CLAIM_ENTRY: vendor is blank BUT invoice number present AND tax > 0
    def _is_itc_claim(row):
        vendor_blank = pd.isna(row.get("vendor_name")) or str(row.get("vendor_name","")).strip() in ("","nan")
        inv_present  = pd.notna(row.get("invoice_number")) and str(row.get("invoice_number","")).strip() not in ("","nan")
        igst_val     = pd.to_numeric(row.get("igst"), errors="coerce") or 0
        cgst_val     = pd.to_numeric(row.get("cgst"), errors="coerce") or 0
        sgst_val     = pd.to_numeric(row.get("sgst"), errors="coerce") or 0
        has_tax      = (igst_val + cgst_val + sgst_val) > 0
        return vendor_blank and inv_present and has_tax
 
    df["is_itc_claim_entry"] = df.apply(_is_itc_claim, axis=1)
 
    # Empty rows
    def _is_empty(row):
        if row.get("is_itc_claim_entry", False):
            return False
        inv_empty    = pd.isna(row.get("invoice_number")) or str(row.get("invoice_number","")).strip() in ("","nan")
        vendor_empty = pd.isna(row.get("vendor_name"))    or str(row.get("vendor_name","")).strip() in ("","nan")
        amount       = pd.to_numeric(row.get("total_amount"), errors="coerce")
        amount_zero  = pd.isna(amount) or amount == 0
        has_narration= pd.notna(row.get("narration")) and str(row.get("narration","")).strip() not in ("","nan")
        return inv_empty and vendor_empty and amount_zero and not has_narration
 
    df["is_empty_row"] = df.apply(_is_empty, axis=1)
 
    # RCM flags
    sheet_is_rcm  = df["source_sheet"].str.lower().str.contains("rcm", na=False)
    narration_rcm = df["narration"].apply(
        lambda x: any(kw in str(x).lower() for kw in RCM_NARRATION_KEYWORDS) if pd.notna(x) else False
    )
    df["is_rcm"]   = sheet_is_rcm
    df["rcm_hint"] = narration_rcm
 
    return df
 
 
# ── LAYER 4: Invoice Number Extraction ────────────────────────
def _extract_invoice_numbers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["inv_from_narration"] = False
    df["_inv_ambiguous"]     = False
 
    for idx, row in df.iterrows():
        inv = row.get("invoice_number")
        if pd.notna(inv) and str(inv).strip() not in ("", "nan"):
            continue
        extracted, ambiguous = _extract_invoice_from_narration(row.get("narration"))
        df.at[idx, "invoice_number"]    = extracted if extracted else None
        df.at[idx, "inv_from_narration"]= bool(extracted)
        df.at[idx, "_inv_ambiguous"]    = ambiguous
 
    return df
 
 
# ── LAYER 5: Validation Engine ────────────────────────────────
def _validate_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    error_records = []
    flags_list    = []
    severity_list = []
    seen_invoices = {}
 
    for idx, row in df.iterrows():
        flags        = []
        source_row   = row.get("source_row")
        source_sheet = row.get("source_sheet", "")
        vendor       = row.get("vendor_name", "")
        inv_raw      = row.get("invoice_number", "")
 
        def _log(flag, issue, severity="WARNING"):
            flags.append(flag)
            error_records.append({
                "source_row": source_row, "source_sheet": source_sheet,
                "vendor_name": vendor, "invoice_number": inv_raw,
                "flag": flag, "issue": issue, "severity": severity,
            })
 
        if row.get("is_footer_row", False):
            _log("FOOTER_ROW", "Grand Total / summary row — not an invoice")
            flags_list.append(flags); severity_list.append("WARNING"); continue
 
        if row.get("is_empty_row", False):
            _log("EMPTY_ROW", "Row has no vendor, invoice, or amount")
            flags_list.append(flags); severity_list.append("WARNING"); continue
 
        if row.get("is_itc_claim_entry", False):
            _log("ITC_CLAIM_ENTRY",
                 f"ITC-only journal entry (no vendor name in Tally) — "
                 f"invoice {inv_raw} has tax but no vendor. "
                 f"Match by invoice number only in 5C.")
        else:
            if pd.isna(vendor) or str(vendor).strip() in ("", "nan"):
                _log("MISSING_VENDOR_NAME", "Vendor name is missing")
 
        if pd.isna(inv_raw) or str(inv_raw).strip() in ("", "nan"):
            _log("MISSING_INVOICE_NUMBER",
                 "Invoice number missing and narration extraction found nothing")
 
        if row.get("inv_from_narration", False):
            _log("NARRATION_EXTRACTED",
                 f"Invoice number '{inv_raw}' extracted from narration (not Voucher No.). "
                 f"5C will use tail matching as fallback.")
 
        if row.get("_inv_ambiguous", False):
            _log("AMBIGUOUS_INVOICE_NUMBER",
                 "Multiple invoice tokens in narration — longest selected, verify manually")
 
        date_val = row.get("invoice_date")
        if pd.isna(date_val) or str(date_val).strip() == "":
            _log("MISSING_DATE", "Invoice date is missing")
        else:
            try:
                pd.to_datetime(date_val, dayfirst=True)
            except Exception:
                _log("BAD_DATE", f"Cannot parse date: {date_val}")
 
        for num_col in ["igst", "cgst", "sgst", "total_amount"]:
            val = row.get(num_col)
            if pd.notna(val) and str(val).strip() not in ("", "nan", "0"):
                if pd.isna(pd.to_numeric(val, errors="coerce")):
                    _log("NON_NUMERIC_VALUE", f"Non-numeric in {num_col}: '{val}' → will be 0")
 
        if not row.get("is_itc_claim_entry", False):
            total = pd.to_numeric(row.get("total_amount"), errors="coerce")
            if pd.isna(total) or total == 0:
                _log("ZERO_TOTAL", "total_amount is zero or missing")
 
        inv_norm    = _normalize_invoice_number(inv_raw)
        vendor_norm = _normalize_vendor_name(vendor)
        amount      = abs(pd.to_numeric(row.get("total_amount"), errors="coerce") or 0)
        dup_key     = f"{inv_norm}||{vendor_norm}||{amount}"
        if inv_norm and dup_key in seen_invoices:
            _log("DUPLICATE_INVOICE", f"Duplicate of row {seen_invoices[dup_key]} in {source_sheet}")
        elif inv_norm:
            seen_invoices[dup_key] = source_row
 
        severity = "OK" if not flags else "WARNING"
        flags_list.append(flags)
        severity_list.append(severity)
 
    df = df.copy()
    df["validation_flags"] = flags_list
    df["severity"]         = severity_list
    return df, pd.DataFrame(error_records)
 
 
# ── LAYER 6: Data Cleaning ────────────────────────────────────
def _clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in ["vendor_name", "invoice_number", "narration", "ca_tally_remark"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["original_amount"] = df["total_amount"]
    df["total_amount"]    = df["total_amount"].abs()
    if "invoice_date" in df.columns:
        df["invoice_date"] = pd.to_datetime(df["invoice_date"], dayfirst=True, errors="coerce")
    return df
 
 
# ── LAYER 7: Enrichment ───────────────────────────────────────
def _enrich(df: pd.DataFrame, filepath: str, file_hash: str) -> pd.DataFrame:
    df = df.copy()
    df["invoice_number_norm"] = df["invoice_number"].apply(_normalize_invoice_number)
    df["invoice_number_tail"] = df["invoice_number_norm"].apply(_extract_invoice_tail)
    df["vendor_name_norm"]    = df["vendor_name"].apply(_normalize_vendor_name)
    df["ca_tally_status"]     = df["ca_tally_remark"].apply(_map_ca_tally_remark)
    df["source_file"]         = str(filepath)
    df["source"]              = "BOOKS"
    df["file_hash"]           = file_hash
    return df
 
 
# ── Merge Safety ──────────────────────────────────────────────
def _safe_merge(sheet_dfs: list) -> pd.DataFrame:
    if not sheet_dfs:
        raise ValueError("No sheet data to merge.")
    all_cols = set()
    for df in sheet_dfs:
        all_cols.update(df.columns)
    aligned = []
    for df in sheet_dfs:
        df = df.copy()
        for col in all_cols:
            if col not in df.columns:
                df[col] = None
        aligned.append(df[sorted(all_cols)])
    return pd.concat(aligned, ignore_index=True).reset_index(drop=True)
 
 
# ── Output Schema Lock ────────────────────────────────────────
def _enforce_output_schema(df: pd.DataFrame) -> pd.DataFrame:
    df = df.drop(columns=[c for c in _INTERNAL_COLS if c in df.columns])
    missing = [c for c in EXPECTED_OUTPUT_COLUMNS if c not in df.columns]
    if missing:
        raise RuntimeError(f"OUTPUT SCHEMA VIOLATION — missing: {missing}. Parser bug.")
    return df[EXPECTED_OUTPUT_COLUMNS]
 
 
# ── Reporting ─────────────────────────────────────────────────
def _build_report(df, error_log, filepath, file_hash,
                  sheets_processed, sheets_skipped, rcm_sheets, format_warnings) -> dict:
    flag_counts = Counter()
    for flags in df["validation_flags"]:
        flag_counts.update(flags)
    total_rows = len(df)
    clean_rows = int((df["severity"] == "OK").sum())
    rows_per_sheet = df.groupby("source_sheet").size().to_dict() if "source_sheet" in df.columns else {}
    ca_status_summary = df["ca_tally_status"].value_counts().to_dict() if "ca_tally_status" in df.columns else {}
 
    return {
        "format_version":           "TALLY_5B_v1.3",
        "file":                     str(filepath),
        "file_hash":                file_hash,
        "processed_at":             datetime.now(timezone.utc).isoformat(),
        "total_rows":               total_rows,
        "clean_rows":               clean_rows,
        "flagged_rows":             total_rows - clean_rows,
        "dropped_rows":             0,
        "summary_warnings":         dict(flag_counts),
        "sheet_names_processed":    sheets_processed,
        "sheets_skipped":           sheets_skipped,
        "rcm_sheets_detected":      rcm_sheets,
        "rows_per_sheet":           rows_per_sheet,
        "missing_invoice_count":    flag_counts.get("MISSING_INVOICE_NUMBER", 0),
        "narration_extracted_count":flag_counts.get("NARRATION_EXTRACTED", 0),
        "ambiguous_invoice_count":  flag_counts.get("AMBIGUOUS_INVOICE_NUMBER", 0),
        "itc_claim_entry_count":    flag_counts.get("ITC_CLAIM_ENTRY", 0),
        "duplicate_count":          flag_counts.get("DUPLICATE_INVOICE", 0),
        "footer_rows":              flag_counts.get("FOOTER_ROW", 0),
        "empty_rows":               flag_counts.get("EMPTY_ROW", 0),
        "ca_tally_status_summary":  ca_status_summary,
        "rcm_hint_count":           int(df.get("rcm_hint", pd.Series(False)).sum()),
        "format_warnings":          format_warnings,
        "source":                   "BOOKS",
    }
 
 
# ── ENTRY POINT ───────────────────────────────────────────────
# CHANGED: _load_sheets now returns 4 values. RCM sheets parsed and merged.
def parse_tally(filepath: str) -> dict:
    file_hash           = _compute_file_hash(filepath)
    all_format_warnings = []
 
    # CHANGED: unpack 4 values — rcm_sheets_data is new
    raw_sheets, sheets_skipped, rcm_sheet_names, rcm_sheets_data = \
        _load_sheets(filepath)
    sheets_processed = list(raw_sheets.keys())
 
    # Normal ITC input sheets — unchanged pipeline
    processed_dfs = []
    for sheet_name, raw_df in raw_sheets.items():
        df, warnings = _standardize_columns(
            _detect_structure(raw_df, sheet_name), sheet_name
        )
        all_format_warnings.extend(warnings)
        df = _detect_special_rows(df)
        df = _extract_invoice_numbers(df)
        processed_dfs.append(df)
 
    merged_df            = _safe_merge(processed_dfs)
    merged_df, error_log = _validate_rows(merged_df)
    merged_df            = _clean_data(merged_df)
    merged_df            = _enrich(merged_df, filepath, file_hash)
 
    # NEW: Parse RCM payable sheets and append to output
    rcm_dfs = []
    for sheet_name, raw_df in rcm_sheets_data.items():
        rcm_df = _parse_rcm_sheet(sheet_name, raw_df, filepath, file_hash)
        if not rcm_df.empty:
            rcm_dfs.append(rcm_df)
 
    if rcm_dfs:
        rcm_combined = pd.concat(rcm_dfs, ignore_index=True)
        # Align columns to EXPECTED_OUTPUT_COLUMNS before concat
        for col in EXPECTED_OUTPUT_COLUMNS:
            if col not in rcm_combined.columns:
                rcm_combined[col] = None
        rcm_combined = rcm_combined[EXPECTED_OUTPUT_COLUMNS]
        merged_df = pd.concat([merged_df, rcm_combined], ignore_index=True)
        print(f"  ✅ Merged {len(rcm_combined)} RCM payable rows into books output")
 
    parse_report = _build_report(
        merged_df, error_log, filepath, file_hash,
        sheets_processed, sheets_skipped, rcm_sheet_names, all_format_warnings
    )
    merged_df = _enforce_output_schema(merged_df)
 
    return {"data": merged_df, "error_log": error_log, "parse_report": parse_report}
 
 
