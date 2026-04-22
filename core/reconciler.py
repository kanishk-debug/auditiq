# ============================================================
# AUDITIQ — PHASE 5C: RECONCILIATION MATCHING ENGINE (v1.5)
# Changes from v1.4:
#   RCM-BOOKS: _preprocess now separates RCM_BOOKS rows from books before matching
#   RCM-BOOKS: reconcile() handles books_rcm_payable → RCM bucket
#              (these come from "IGST @ 18% RCM Payable" sheet parsed in 5B v1.3)
# All v1.4 changes retained unchanged
# ============================================================

import pandas as pd
from rapidfuzz import fuzz
from datetime import datetime, timezone
from collections import Counter

# ── Constants ──────────────────────────────────────────────────
DATE_TOLERANCE_DAYS  = 45
AMOUNT_FLOOR         = 1.0
TAIL_AMOUNT_PCT      = 0.01
POSSIBLE_AMOUNT_PCT  = 0.02
VENDOR_FUZZY_TAIL    = 80
VENDOR_FUZZY_POSSIBLE= 75   # lowered from 85 → Dwija scores 79, now passes

ADJUSTMENT_KEYWORDS = [
    "bank charges", "bank charge", "bank fee", "bank service",
    "interest paid", "bank interest", "late payment", "gst input claimed",
]

BUCKET_MATCHED         = "MATCHED"
BUCKET_PROBABLE        = "PROBABLE_MATCH"
BUCKET_REVIEW          = "REVIEW"
BUCKET_INELIGIBLE      = "INELIGIBLE"
BUCKET_RCM             = "RCM"
BUCKET_MISSING_IN_2B   = "MISSING_IN_2B"
BUCKET_MISSING_IN_BOOKS= "MISSING_IN_BOOKS"
BUCKET_PREVIOUS_MONTH  = "PREVIOUS_MONTH"
BUCKET_ADJUSTMENT      = "ADJUSTMENT"

MATCH_EXACT      = "EXACT"
MATCH_SUBSTRING  = "SUBSTRING_MATCH"
MATCH_ITC        = "ITC_CLAIM_MATCH"
MATCH_TAIL       = "TAIL_MATCH"
MATCH_POSSIBLE   = "POSSIBLE_MATCH"
MATCH_NONE       = "UNMATCHED"

CONF_HIGH   = "HIGH"
CONF_MEDIUM = "MEDIUM"
CONF_LOW    = "LOW"

OUTPUT_COLUMNS = [
    "source", "vendor_name", "invoice_number", "invoice_date",
    "total_amount", "igst", "cgst", "sgst",
    "bucket", "match_type", "confidence", "match_reasoning",
    "matched_invoice_2b", "matched_vendor_2b", "matched_amount_2b",
    "ca_tally_status", "ca_remark_conflict",
    "is_itc_claim_entry", "inv_from_narration",
    "source_sheet", "source_row",
]


# ── Helpers ────────────────────────────────────────────────────
def _within_amount(a: float, b: float, pct: float) -> bool:
    return abs(a - b) <= max(AMOUNT_FLOOR, max(a, b) * pct)

# ── Cross-column amount extractors ────────────────────────────
def _get_tax_amount_2b(row) -> float:
    igst = pd.to_numeric(row.get('igst_amount', 0), errors='coerce') or 0
    cgst = pd.to_numeric(row.get('cgst_amount', 0), errors='coerce') or 0
    if igst > 0:
        return round(float(igst), 2)
    if cgst > 0:
        return round(float(cgst), 2)
    return 0.0

def _get_tax_amount_books(row) -> float:
    tax = pd.to_numeric(row.get('tax_amount', 0), errors='coerce') or 0
    if tax > 0:
        return round(float(tax), 2)
    igst = pd.to_numeric(row.get('igst', 0), errors='coerce') or 0
    cgst = pd.to_numeric(row.get('cgst', 0), errors='coerce') or 0
    return round(float(igst or cgst), 2)

def _amounts_match_cross(books_row, b2b_row, tolerance: float = 2.0) -> bool:
    books_amt = _get_tax_amount_books(books_row)
    b2b_amt   = _get_tax_amount_2b(b2b_row)
    if books_amt == 0 and b2b_amt == 0:
        return False
    return abs(books_amt - b2b_amt) <= tolerance

# ── Substring invoice match ────────────────────────────────────
def _invoice_substring_match(norm_a: str, norm_b: str) -> bool:
    if not norm_a or not norm_b:
        return False
    shorter = norm_a if len(norm_a) <= len(norm_b) else norm_b
    longer  = norm_b if len(norm_a) <= len(norm_b) else norm_a
    if len(shorter) < 3:
        return False
    return shorter in longer

# ── Shared keyword vendor match ────────────────────────────────
def _shared_keyword_match(name_a: str, name_b: str, min_word_len: int = 5) -> bool:
    noise = {
        'india', 'private', 'limited', 'pvt', 'ltd', 'services',
        'technologies', 'solutions', 'enterprises', 'company',
        'group', 'foods', 'works', 'trade', 'trading'
    }
    def keywords(name):
        return set(
            w.lower().strip('().,&-/')
            for w in str(name).split()
            if len(w.strip('().,&-/')) >= min_word_len
            and w.lower().strip('().,&-/') not in noise
        )
    words_a = keywords(name_a)
    words_b = keywords(name_b)
    return bool(words_a & words_b)

# ── Original helpers ───────────────────────────────────────────
def _within_date(d1, d2) -> bool:
    try:
        if pd.isna(d1) or pd.isna(d2):
            return True
        diff = abs((pd.Timestamp(d1) - pd.Timestamp(d2)).days)
        return diff <= DATE_TOLERANCE_DAYS
    except Exception:
        return True

def _vendor_score(a: str, b: str) -> int:
    if not a or not b:
        return 0
    return fuzz.token_sort_ratio(str(a), str(b))

def _tail_suffix_match(books_tail: str, b2b_tail: str) -> bool:
    if not books_tail or not b2b_tail:
        return False
    if len(books_tail) < 3 or len(b2b_tail) < 3:
        return False
    return books_tail.endswith(b2b_tail) or b2b_tail.endswith(books_tail)

def _is_adjustment(row) -> bool:
    inv_blank = pd.isna(row.get("invoice_number")) or \
                str(row.get("invoice_number", "")).strip() in ("", "nan")
    vendor = str(row.get("vendor_name", "")).lower()
    return inv_blank and any(kw in vendor for kw in ADJUSTMENT_KEYWORDS)

def _norm_inv(inv) -> str:
    if pd.isna(inv) or str(inv).strip() in ("", "nan"):
        return ""
    return str(inv).lower().replace(" ","").replace("/","").replace("-","").replace(".","")

def _tail(inv_norm: str) -> str:
    if not inv_norm:
        return ""
    import re as _re
    digits = _re.findall(r'\d{2,}', inv_norm)
    return digits[-1] if digits else ""

def _empty_result(brow) -> dict:
    return {
        "source":             "BOOKS",
        "vendor_name":        brow.get("vendor_name"),
        "invoice_number":     brow.get("invoice_number"),
        "invoice_date":       brow.get("invoice_date"),
        "total_amount":       brow.get("total_amount", 0),
        "igst":               brow.get("igst", 0),
        "cgst":               brow.get("cgst", 0),
        "sgst":               brow.get("sgst", 0),
        "bucket":             None,
        "match_type":         None,
        "confidence":         None,
        "match_reasoning":    "",
        "matched_invoice_2b": None,
        "matched_vendor_2b":  None,
        "matched_amount_2b":  None,
        "ca_tally_status":    brow.get("ca_tally_status"),
        "ca_remark_conflict": False,
        "is_itc_claim_entry": brow.get("is_itc_claim_entry", False),
        "inv_from_narration": brow.get("inv_from_narration", False),
        "source_sheet":       brow.get("source_sheet"),
        "source_row":         brow.get("source_row"),
    }

def _empty_2b_result(row) -> dict:
    """Standard 2B-sourced result dict — used by RCM, INELIGIBLE, MISSING helpers."""
    return {
        "source":             "2B",
        "vendor_name":        row.get("supplier_name"),
        "invoice_number":     row.get("invoice_number"),
        "invoice_date":       row.get("invoice_date"),
        "total_amount":       row.get("taxable_value", 0),
        "igst":               row.get("igst_amount", 0),
        "cgst":               row.get("cgst_amount", 0),
        "sgst":               row.get("sgst_amount", 0),
        "bucket":             None,
        "match_type":         MATCH_NONE,
        "confidence":         CONF_HIGH,
        "match_reasoning":    "",
        "matched_invoice_2b": None,
        "matched_vendor_2b":  None,
        "matched_amount_2b":  None,
        "ca_tally_status":    None,
        "ca_remark_conflict": False,
        "is_itc_claim_entry": False,
        "inv_from_narration": False,
        "source_sheet":       "B2B",
        "source_row":         row.get("source_row"),
    }


# ── Pre-processing ─────────────────────────────────────────────
# v1.5: Now also separates RCM_BOOKS rows from books before matching
def _preprocess(gstr2b_df, books_df):
    g = gstr2b_df.copy()
    if "supplier_name_norm" not in g.columns:
        g["supplier_name_norm"] = g["supplier_name"].apply(
            lambda x: " ".join(str(x).lower().split()) if pd.notna(x) else ""
        )
    if "invoice_number_tail" not in g.columns:
        g["invoice_number_norm"] = g["invoice_number"].apply(_norm_inv)
        g["invoice_number_tail"] = g["invoice_number_norm"].apply(_tail)

    b2b_all = g[~g["is_footer_row"].fillna(False)].copy().reset_index(drop=True)

    # ── Separate pre-classified rows BEFORE matching ──────────
    b2b_prev = b2b_all[b2b_all["ca_status"].fillna("") == "PREVIOUS_MONTH"].copy()

    rcm_ca_mask  = b2b_all["ca_status"].fillna("") == "RCM"
    rcm_col_mask = b2b_all["rcm_flag"].fillna(False) == True
    b2b_rcm      = b2b_all[rcm_ca_mask | rcm_col_mask].copy()

    b2b_inelig_2b = b2b_all[b2b_all["ca_status"].fillna("") == "INELIGIBLE"].copy()

    skip_mask = (
        b2b_all["ca_status"].fillna("").isin(["PREVIOUS_MONTH", "RCM", "INELIGIBLE"]) |
        (b2b_all["rcm_flag"].fillna(False) == True)
    )
    b2b_active = b2b_all[~skip_mask].copy().reset_index(drop=True)

    # ── Books side ────────────────────────────────────────────
    bk = books_df[~books_df["is_footer_row"].fillna(False)].copy()

    # FIX: "bucket" column only exists when 5B v1.3 added RCM rows from the
    # "IGST @ 18% RCM Payable" sheet.  Older 5B versions never set this column
    # on normal Tally rows, so bk["bucket"] raises KeyError on the whole df.
    # Guard with a column-existence check before filtering.
    if "bucket" in bk.columns:
        rcm_books_mask = bk["bucket"].fillna("") == "RCM_BOOKS"
    else:
        rcm_books_mask = pd.Series([False] * len(bk), index=bk.index)
    books_rcm_payable = bk[rcm_books_mask].copy()
    bk                = bk[~rcm_books_mask].copy()

    books_itc   = bk[bk["is_itc_claim_entry"].fillna(False)].copy()
    bk          = bk[~bk["is_itc_claim_entry"].fillna(False)].copy()
    adj_mask    = bk.apply(_is_adjustment, axis=1)
    books_adj   = bk[adj_mask].copy()
    bk          = bk[~adj_mask].copy()
    inelig_mask = bk["ca_tally_status"].fillna("") == "INELIGIBLE"
    books_inelig= bk[inelig_mask].copy()
    books_active= bk[~inelig_mask].copy().reset_index(drop=True)

    return (b2b_active, b2b_prev, b2b_rcm, b2b_inelig_2b,
            books_active, books_itc, books_adj, books_inelig,
            books_rcm_payable)


# ── Core matching (UNCHANGED from v1.4) ───────────────────────
def _match_books_to_2b(books_active, b2b_active):
    results      = []
    b2b_used_idx = set()

    b2b_by_inv = {}
    for i, row in b2b_active.iterrows():
        key = row.get("invoice_number_norm", "") or _norm_inv(row.get("invoice_number",""))
        if key:
            b2b_by_inv.setdefault(key, []).append(i)

    for _, brow in books_active.iterrows():
        b_inv_norm  = str(brow.get("invoice_number_norm", "") or "")
        b_inv_tail  = str(brow.get("invoice_number_tail", "") or "")
        b_vendor    = str(brow.get("vendor_name_norm") or brow.get("vendor_name", "") or "")
        b_amount    = float(brow.get("total_amount", 0) or 0)
        b_date      = brow.get("invoice_date")
        b_narration = bool(brow.get("inv_from_narration", False))
        matched     = False
        result      = _empty_result(brow)

        # ── LEVEL 1: Exact normalized invoice match ────────────
        if b_inv_norm and b_inv_norm in b2b_by_inv:
            candidates = [i for i in b2b_by_inv[b_inv_norm] if i not in b2b_used_idx]
            if len(candidates) >= 1:
                b2b_row = b2b_active.loc[candidates[0]]
                b2b_used_idx.add(candidates[0])
                conf = CONF_HIGH if len(candidates) == 1 else CONF_MEDIUM
                note = "" if len(candidates) == 1 else \
                    f" Note: {len(candidates)} identical invoice numbers in 2B — verify."
                result.update({
                    "bucket":           BUCKET_MATCHED,
                    "match_type":       MATCH_EXACT,
                    "confidence":       conf,
                    "match_reasoning":  (
                        f"Exact invoice match. "
                        f"Books: {brow.get('invoice_number')} = "
                        f"2B: {b2b_row.get('invoice_number')} "
                        f"({b2b_row.get('supplier_name')}).{note}"
                    ),
                    "matched_invoice_2b": b2b_row.get("invoice_number"),
                    "matched_vendor_2b":  b2b_row.get("supplier_name"),
                    "matched_amount_2b":  b2b_row.get("taxable_value"),
                })
                matched = True

        # ── LEVEL 1.5: Substring invoice match ─────────────────
        if not matched and b_inv_norm and len(b_inv_norm) >= 3:
            for i, b2b_row in b2b_active.iterrows():
                if i in b2b_used_idx:
                    continue
                b2b_inv_n = str(
                    b2b_row.get("invoice_number_norm","") or
                    _norm_inv(str(b2b_row.get("invoice_number","")))
                )
                if not _invoice_substring_match(b_inv_norm, b2b_inv_n):
                    continue
                if not _amounts_match_cross(brow, b2b_row, tolerance=5.0):
                    continue
                b2b_used_idx.add(i)
                result.update({
                    "bucket":          BUCKET_MATCHED,
                    "match_type":      MATCH_SUBSTRING,
                    "confidence":      CONF_HIGH,
                    "match_reasoning": (
                        f"Invoice substring match: "
                        f"Books '{brow.get('invoice_number')}' (norm: '{b_inv_norm}') "
                        f"contains/in 2B '{b2b_row.get('invoice_number')}' "
                        f"(norm: '{b2b_inv_n}') "
                        f"from {b2b_row.get('supplier_name')}. "
                        f"Amount verified. HIGH confidence."
                    ),
                    "matched_invoice_2b": b2b_row.get("invoice_number"),
                    "matched_vendor_2b":  b2b_row.get("supplier_name"),
                    "matched_amount_2b":  b2b_row.get("taxable_value"),
                })
                matched = True
                break

        # ── LEVEL 2: Tail suffix match (narration rows only) ───
        if not matched and b_narration and b_inv_tail:
            tail_candidates = []
            for i, b2b_row in b2b_active.iterrows():
                if i in b2b_used_idx:
                    continue
                b2b_tail_val = str(
                    b2b_row.get("invoice_number_tail", "") or
                    _tail(_norm_inv(str(b2b_row.get("invoice_number",""))))
                )
                b2b_vendor = str(
                    b2b_row.get("supplier_name_norm", "") or
                    " ".join(str(b2b_row.get("supplier_name","")).lower().split())
                )
                b2b_date = b2b_row.get("invoice_date")
                if not _tail_suffix_match(b_inv_tail, b2b_tail_val):
                    continue
                vscore = _vendor_score(b_vendor, b2b_vendor)
                if vscore < VENDOR_FUZZY_TAIL:
                    continue
                if not _amounts_match_cross(brow, b2b_row, tolerance=5.0):
                    continue
                if not _within_date(b_date, b2b_date):
                    continue
                tail_candidates.append((i, b2b_row, vscore))

            if len(tail_candidates) == 1:
                idx, b2b_row, vscore = tail_candidates[0]
                b2b_used_idx.add(idx)
                result.update({
                    "bucket":          BUCKET_PROBABLE,
                    "match_type":      MATCH_TAIL,
                    "confidence":      CONF_MEDIUM,
                    "match_reasoning": (
                        f"Invoice from narration ({brow.get('invoice_number')}). "
                        f"Tail suffix matches 2B invoice "
                        f"'{b2b_row.get('invoice_number')}' "
                        f"({b2b_row.get('supplier_name')}). "
                        f"Vendor similarity {vscore}%. Amount within tolerance. "
                        f"MEDIUM confidence — please verify invoice number."
                    ),
                    "matched_invoice_2b": b2b_row.get("invoice_number"),
                    "matched_vendor_2b":  b2b_row.get("supplier_name"),
                    "matched_amount_2b":  b2b_row.get("taxable_value"),
                })
                matched = True
            elif len(tail_candidates) > 1:
                names = ", ".join(
                    f"{r.get('invoice_number')} ({r.get('supplier_name')})"
                    for _, r, _ in tail_candidates[:3]
                )
                result.update({
                    "bucket":          BUCKET_REVIEW,
                    "match_type":      "AMBIGUOUS_TAIL",
                    "confidence":      CONF_LOW,
                    "match_reasoning": (
                        f"Tail '{b_inv_tail}' matches multiple 2B invoices: {names}. "
                        f"Cannot auto-match. Manual review required."
                    ),
                })
                matched = True

        # ── LEVEL 3: Fuzzy vendor + amount match ───────────────
        if not matched:
            possible = []
            for i, b2b_row in b2b_active.iterrows():
                if i in b2b_used_idx:
                    continue
                b2b_inv_n = str(
                    b2b_row.get("invoice_number_norm","") or
                    _norm_inv(str(b2b_row.get("invoice_number","")))
                )
                if b2b_inv_n == b_inv_norm and b_inv_norm:
                    continue
                b2b_vendor = str(
                    b2b_row.get("supplier_name_norm","") or
                    " ".join(str(b2b_row.get("supplier_name","")).lower().split())
                )
                b2b_amount_display = _get_tax_amount_2b(b2b_row)
                vscore = _vendor_score(b_vendor, b2b_vendor)
                vendor_name_raw = str(brow.get("vendor_name", "") or "")
                b2b_vendor_raw  = str(b2b_row.get("supplier_name", "") or "")
                vendor_ok = (
                    vscore >= VENDOR_FUZZY_POSSIBLE or
                    _shared_keyword_match(vendor_name_raw, b2b_vendor_raw)
                )
                if not vendor_ok:
                    continue
                if not _amounts_match_cross(brow, b2b_row, tolerance=2.0):
                    continue
                possible.append((i, b2b_row, vscore, b2b_amount_display))

            if len(possible) == 1:
                idx, b2b_row, vscore, b2b_amt_display = possible[0]
                b2b_used_idx.add(idx)
                books_amt_display = _get_tax_amount_books(brow)
                result.update({
                    "bucket":          BUCKET_REVIEW,
                    "match_type":      MATCH_POSSIBLE,
                    "confidence":      CONF_LOW,
                    "match_reasoning": (
                        f"Same vendor ({b2b_row.get('supplier_name')}, {vscore}% similarity) "
                        f"and matching tax amount "
                        f"(Books ₹{books_amt_display:,.2f} vs "
                        f"2B ₹{b2b_amt_display:,.2f}) "
                        f"but invoice numbers differ: "
                        f"Books '{brow.get('invoice_number')}' vs "
                        f"2B '{b2b_row.get('invoice_number')}'. "
                        f"LOW confidence — verify with vendor."
                    ),
                    "matched_invoice_2b": b2b_row.get("invoice_number"),
                    "matched_vendor_2b":  b2b_row.get("supplier_name"),
                    "matched_amount_2b":  b2b_row.get("taxable_value"),
                })
                matched = True

        if not matched:
            result.update({
                "bucket":          BUCKET_MISSING_IN_2B,
                "match_type":      MATCH_NONE,
                "confidence":      CONF_HIGH,
                "match_reasoning": (
                    f"No match found in GSTR-2B for invoice "
                    f"'{brow.get('invoice_number')}' "
                    f"from {brow.get('vendor_name')}. "
                    f"Supplier may not have filed or filed in wrong period."
                ),
            })

        results.append(result)
    return results, b2b_used_idx


# ── ITC entry matching (UNCHANGED) ────────────────────────────
def _match_itc_entries(books_itc, b2b_active, b2b_used_idx):
    results    = []
    b2b_by_inv = {}
    for i, row in b2b_active.iterrows():
        key = row.get("invoice_number_norm","") or _norm_inv(str(row.get("invoice_number","")))
        if key:
            b2b_by_inv.setdefault(key, []).append(i)

    for _, brow in books_itc.iterrows():
        b_inv_norm = str(brow.get("invoice_number_norm","") or "")
        result     = _empty_result(brow)
        candidates = [i for i in b2b_by_inv.get(b_inv_norm, [])
                      if i not in b2b_used_idx]
        if candidates:
            b2b_row = b2b_active.loc[candidates[0]]
            b2b_used_idx.add(candidates[0])
            result.update({
                "bucket":          BUCKET_MATCHED,
                "match_type":      MATCH_ITC,
                "confidence":      CONF_HIGH,
                "match_reasoning": (
                    f"ITC claim entry — no vendor in Tally. "
                    f"Invoice {brow.get('invoice_number')} matched to "
                    f"{b2b_row.get('supplier_name')} in 2B by invoice number only. "
                    f"HIGH confidence."
                ),
                "matched_invoice_2b": b2b_row.get("invoice_number"),
                "matched_vendor_2b":  b2b_row.get("supplier_name"),
                "matched_amount_2b":  b2b_row.get("taxable_value"),
            })
        else:
            result.update({
                "bucket":          BUCKET_MISSING_IN_2B,
                "match_type":      MATCH_NONE,
                "confidence":      CONF_MEDIUM,
                "match_reasoning": (
                    f"ITC claim entry — invoice {brow.get('invoice_number')} "
                    f"not found in GSTR-2B."
                ),
            })
        results.append(result)
    return results


# ── Unmatched 2B rows (UNCHANGED) ─────────────────────────────
def _build_unmatched_2b(b2b_active, b2b_used_idx, b2b_prev):
    results = []
    for i, row in b2b_active.iterrows():
        if i in b2b_used_idx:
            continue
        ca_status = str(row.get("ca_status","") or "")
        if ca_status == "PREVIOUS_MONTH":
            bucket    = BUCKET_PREVIOUS_MONTH
            reasoning = (
                f"CA marked 'In Books of Previous Month'. "
                f"Invoice '{row.get('invoice_number')}' already handled."
            )
        else:
            bucket    = BUCKET_MISSING_IN_BOOKS
            reasoning = (
                f"Invoice '{row.get('invoice_number')}' from "
                f"{row.get('supplier_name')} in GSTR-2B but not found in Books. "
                f"Chase client or check Tally."
            )
        r = _empty_2b_result(row)
        r.update({"bucket": bucket, "match_reasoning": reasoning})
        results.append(r)

    for _, row in b2b_prev.iterrows():
        r = _empty_2b_result(row)
        r.update({
            "bucket":          BUCKET_PREVIOUS_MONTH,
            "match_reasoning": (
                f"CA marked 'In Books of Previous Month'. "
                f"Invoice '{row.get('invoice_number')}' "
                f"({row.get('supplier_name')}) already handled."
            ),
        })
        results.append(r)
    return results


# ── CA conflict checker (UNCHANGED) ───────────────────────────
def _check_ca_conflicts(results):
    bucket_to_status = {
        BUCKET_MATCHED:          {"MATCHED"},
        BUCKET_PROBABLE:         {"MATCHED"},
        BUCKET_REVIEW:           {"MATCHED"},
        BUCKET_MISSING_IN_2B:    {"MISSING_IN_2B"},
        BUCKET_MISSING_IN_BOOKS: {"MISSING_IN_BOOKS"},
        BUCKET_INELIGIBLE:       {"INELIGIBLE"},
        BUCKET_RCM:              {"RCM"},
        BUCKET_PREVIOUS_MONTH:   {"PREVIOUS_MONTH"},
        BUCKET_ADJUSTMENT:       {"UNCLASSIFIED", "MATCHED"},
    }
    for r in results:
        ca = str(r.get("ca_tally_status","") or "")
        if ca in ("UNCLASSIFIED","","None"):
            r["ca_remark_conflict"] = False
            continue
        expected = bucket_to_status.get(r.get("bucket",""), set())
        if ca not in expected:
            r["ca_remark_conflict"] = True
            r["match_reasoning"] += (
                f" ⚠️ CA REMARK CONFLICT: CA marked '{ca}' "
                f"but engine says '{r.get('bucket')}'. Verify."
            )
        else:
            r["ca_remark_conflict"] = False
    return results


# ── Summary builder (UNCHANGED) ───────────────────────────────
def _build_summary(results, period=""):
    df      = pd.DataFrame(results)
    buckets = df["bucket"].value_counts().to_dict()

    def itc_sum(mask):
        sub = df[mask]
        return {
            "igst": round(sub["igst"].fillna(0).astype(float).sum(), 2),
            "cgst": round(sub["cgst"].fillna(0).astype(float).sum(), 2),
            "sgst": round(sub["sgst"].fillna(0).astype(float).sum(), 2),
        }

    claimable = df["bucket"].isin([BUCKET_MATCHED, BUCKET_PROBABLE])
    return {
        "period":               period,
        "processed_at":         datetime.now(timezone.utc).isoformat(),
        "total_rows":           len(df),
        "books_rows":           int((df["source"] == "BOOKS").sum()),
        "b2b_rows":             int((df["source"] == "2B").sum()),
        "bucket_counts":        buckets,
        "claimable_itc":        itc_sum(claimable),
        "review_required":      int(df["bucket"].isin([BUCKET_PROBABLE, BUCKET_REVIEW]).sum()),
        "ca_remark_conflicts":  int(df["ca_remark_conflict"].fillna(False).sum()),
        "confidence_breakdown": df["confidence"].value_counts().to_dict(),
        "match_type_breakdown": df["match_type"].value_counts().to_dict(),
    }


# ── ENTRY POINT ────────────────────────────────────────────────
def reconcile(gstr2b_data, books_data, period=""):
    (b2b_active, b2b_prev, b2b_rcm, b2b_inelig_2b,
     books_active, books_itc, books_adj, books_inelig,
     books_rcm_payable) = _preprocess(gstr2b_data, books_data)

    # Books-side RCM from normal IGST/CGST sheets (is_rcm=True rows)
    books_rcm_mask    = books_active["is_rcm"].fillna(False) == True
    books_rcm_entries = books_active[books_rcm_mask].copy()
    books_active      = books_active[~books_rcm_mask].copy().reset_index(drop=True)

    main_results, b2b_used_idx = _match_books_to_2b(books_active, b2b_active)
    itc_results  = _match_itc_entries(books_itc, b2b_active, b2b_used_idx)

    adj_results = []
    for _, brow in books_adj.iterrows():
        r = _empty_result(brow)
        r.update({
            "bucket":          BUCKET_ADJUSTMENT,
            "match_type":      MATCH_NONE,
            "confidence":      CONF_HIGH,
            "match_reasoning": (
                f"Bank charges/adjustment — no invoice. "
                f"Vendor: {brow.get('vendor_name')}. "
                f"CGST ₹{brow.get('cgst',0):,.2f} + SGST ₹{brow.get('sgst',0):,.2f}. "
                f"CA review required."
            ),
        })
        adj_results.append(r)

    inelig_results = []
    for _, brow in books_inelig.iterrows():
        r = _empty_result(brow)
        r.update({
            "bucket":          BUCKET_INELIGIBLE,
            "match_type":      MATCH_NONE,
            "confidence":      CONF_HIGH,
            "match_reasoning": (
                f"Pre-classified INELIGIBLE by CA in Tally. "
                f"Blocked under Sec 17(5). "
                f"Invoice: {brow.get('invoice_number')} from {brow.get('vendor_name')}."
            ),
        })
        inelig_results.append(r)

    inelig_2b_results = []
    for _, row in b2b_inelig_2b.iterrows():
        r = _empty_2b_result(row)
        r.update({
            "bucket":          BUCKET_INELIGIBLE,
            "match_reasoning": (
                f"CA marked 'Ineligible' in GSTR-2B col 21. "
                f"ITC blocked under Sec 17(5). "
                f"Invoice: {row.get('invoice_number')} from {row.get('supplier_name')}."
            ),
        })
        inelig_2b_results.append(r)

    books_rcm_results = []
    for _, brow in books_rcm_entries.iterrows():
        r = _empty_result(brow)
        r.update({
            "bucket":          BUCKET_RCM,
            "match_type":      MATCH_NONE,
            "confidence":      CONF_HIGH,
            "match_reasoning": (
                f"RCM invoice from Books — not expected in GSTR-2B. "
                f"Pay GST directly to govt and claim ITC in same period. "
                f"Invoice: {brow.get('invoice_number')} from {brow.get('vendor_name')}."
            ),
        })
        books_rcm_results.append(r)

    rcm_results = []
    for _, row in b2b_rcm.iterrows():
        source_label = (
            "CA marked RCM in col 21"
            if str(row.get("ca_status","")).upper() == "RCM"
            else "Govt portal col H = Yes"
        )
        r = _empty_2b_result(row)
        r.update({
            "bucket":          BUCKET_RCM,
            "match_reasoning": (
                f"Reverse Charge Mechanism — {source_label}. "
                f"Pay GST directly to govt, then claim ITC in same period. "
                f"Invoice: {row.get('invoice_number')} "
                f"from {row.get('supplier_name')}."
            ),
        })
        rcm_results.append(r)

    books_rcm_payable_results = []
    for _, brow in books_rcm_payable.iterrows():
        r = _empty_result(brow)
        r.update({
            "bucket":          BUCKET_RCM,
            "match_type":      MATCH_NONE,
            "confidence":      CONF_HIGH,
            "match_reasoning": (
                f"RCM payable entry (Tally liability journal). "
                f"Company owes this GST to govt — no 2B match expected. "
                f"Claim ITC after payment is made. "
                f"Voucher: {brow.get('invoice_number')} | "
                f"Vendor: {brow.get('vendor_name')} | "
                f"IGST ₹{float(brow.get('igst', 0) or 0):,.2f}"
            ),
        })
        books_rcm_payable_results.append(r)

    unmatched_2b = _build_unmatched_2b(b2b_active, b2b_used_idx, b2b_prev)

    all_results = (
        main_results              +
        itc_results               +
        adj_results               +
        inelig_results            +
        inelig_2b_results         +
        rcm_results               +
        books_rcm_results         +
        books_rcm_payable_results +
        unmatched_2b
    )
    all_results = _check_ca_conflicts(all_results)

    final_df = pd.DataFrame(all_results)
    for col in OUTPUT_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = None
    final_df = final_df[OUTPUT_COLUMNS].reset_index(drop=True)
    summary  = _build_summary(all_results, period)

    return {"data": final_df, "summary_report": summary, "match_log": final_df}

