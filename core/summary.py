# ============================================================
# AUDITIQ — PHASE 5D: RECONCILIATION SUMMARY & DISPLAY
# Takes reconcile() output and produces:
#   1. Formatted CA-facing console summary (for Colab)
#   2. Per-bucket DataFrames ready for 5E Excel output
#   3. ITC position statement with claimable amounts
# ============================================================

import pandas as pd

# ── Bucket display config ──────────────────────────────────────
BUCKET_META = {
    "MATCHED": {
        "label":   "Matched — Claim ITC",
        "emoji":   "✅",
        "action":  "Claim ITC this month",
        "color":   "green",
    },
    "PROBABLE_MATCH": {
        "label":   "Probable Match — Verify",
        "emoji":   "✅",
        "action":  "Verify invoice number then claim ITC",
        "color":   "green",
    },
    "INELIGIBLE": {
        "label":   "Ineligible — Do Not Claim",
        "emoji":   "🚫",
        "action":  "ITC blocked under Sec 17(5) — do not claim",
        "color":   "red",
    },
    "RCM": {
        "label":   "Reverse Charge",
        "emoji":   "🔄",
        "action":  "Pay GST to govt directly, then claim ITC",
        "color":   "orange",
    },
    "MISSING_IN_2B": {
        "label":   "Missing in GSTR-2B",
        "emoji":   "⏳",
        "action":  "Supplier hasn't filed — hold, check next month",
        "color":   "yellow",
    },
    "MISSING_IN_BOOKS": {
        "label":   "Missing in Books",
        "emoji":   "📋",
        "action":  "Chase client for invoice or check Tally",
        "color":   "blue",
    },
    "PREVIOUS_MONTH": {
        "label":   "Previous Month",
        "emoji":   "📅",
        "action":  "Already handled in prior month",
        "color":   "purple",
    },
    "REVIEW": {
        "label":   "Needs Review",
        "emoji":   "⚠️",
        "action":  "Invoice numbers differ — verify with vendor",
        "color":   "yellow",
    },
    "ADJUSTMENT": {
        "label":   "Adjustment / Bank Charges",
        "emoji":   "🏦",
        "action":  "CA review required",
        "color":   "gray",
    },
}

# ── Helpers ───────────────────────────────────────────────────
def _fmt_inr(val) -> str:
    try:
        v = float(val or 0)
        return f"₹{v:>12,.2f}"
    except Exception:
        return f"{'₹0.00':>13}"

def _tax_totals(df: pd.DataFrame) -> tuple:
    igst = round(df["igst"].fillna(0).astype(float).sum(), 2)
    cgst = round(df["cgst"].fillna(0).astype(float).sum(), 2)
    sgst = round(df["sgst"].fillna(0).astype(float).sum(), 2)
    return igst, cgst, sgst

def _divider(char="─", width=68) -> str:
    return char * width

# ── Core function ─────────────────────────────────────────────
def display_reconciliation_summary(recon_result: dict, period: str = "") -> dict:
    """
    Takes reconcile() output dict.
    Prints formatted CA-facing summary to console.
    Returns dict of per-bucket DataFrames for 5E.
    """
    df      = recon_result["data"].copy()
    summary = recon_result["summary_report"]

    print("\n" + _divider("═"))
    print(f"  AUDITIQ — GST RECONCILIATION REPORT")
    print(f"  Period : {period or summary.get('period','')}")
    print(f"  Date   : {summary.get('processed_at','')[:10]}")
    print(_divider("═"))

    # ── Section 1: ITC Position ───────────────────────────────
    print("\n  ITC POSITION SUMMARY")
    print(_divider())

    claimable_mask = df["bucket"].isin(["MATCHED", "PROBABLE_MATCH"])
    c_igst, c_cgst, c_sgst = _tax_totals(df[claimable_mask])
    c_total = round(c_igst + c_cgst + c_sgst, 2)

    blocked_mask = df["bucket"] == "INELIGIBLE"
    b_igst, b_cgst, b_sgst = _tax_totals(df[blocked_mask])
    b_total = round(b_igst + b_cgst + b_sgst, 2)

    hold_mask = df["bucket"].isin(["MISSING_IN_2B", "REVIEW"])
    h_igst, h_cgst, h_sgst = _tax_totals(df[hold_mask])
    h_total = round(h_igst + h_cgst + h_sgst, 2)

    print(f"  {'Claimable ITC (Matched):':<32} IGST {_fmt_inr(c_igst)}  |  CGST+SGST {_fmt_inr(c_cgst + c_sgst)}")
    print(f"  {'  Total claimable:':<32} {_fmt_inr(c_total)}")
    print()
    print(f"  {'Blocked ITC (Ineligible):':<32} IGST {_fmt_inr(b_igst)}  |  CGST+SGST {_fmt_inr(b_cgst + b_sgst)}")
    print(f"  {'  Total blocked:':<32} {_fmt_inr(b_total)}")
    print()
    print(f"  {'On Hold (Missing in 2B + Review):':<32} IGST {_fmt_inr(h_igst)}  |  CGST+SGST {_fmt_inr(h_cgst + h_sgst)}")
    print(f"  {'  Total on hold:':<32} {_fmt_inr(h_total)}")

    # ── Section 2: Bucket Breakdown ───────────────────────────
    print("\n\n  BUCKET BREAKDOWN")
    print(_divider())
    print(f"  {'Bucket':<30} {'Count':>6}  {'IGST':>12}  {'CGST':>12}  {'SGST':>12}  Action")
    print(_divider())

    BUCKET_ORDER = [
        "MATCHED", "PROBABLE_MATCH", "INELIGIBLE", "RCM",
        "MISSING_IN_2B", "MISSING_IN_BOOKS", "PREVIOUS_MONTH",
        "REVIEW", "ADJUSTMENT"
    ]

    bucket_dfs = {}
    for bucket in BUCKET_ORDER:
        sub = df[df["bucket"] == bucket]
        if sub.empty:
            continue
        bucket_dfs[bucket] = sub
        meta    = BUCKET_META.get(bucket, {"label": bucket, "emoji": "·", "action": ""})
        igst, cgst, sgst = _tax_totals(sub)
        label   = f"{meta['emoji']} {meta['label']}"
        print(f"  {label:<30} {len(sub):>6}  {_fmt_inr(igst)}  {_fmt_inr(cgst)}  {_fmt_inr(sgst)}  {meta['action']}")

    print(_divider())
    total_igst, total_cgst, total_sgst = _tax_totals(df)
    print(f"  {'TOTAL':<30} {len(df):>6}  {_fmt_inr(total_igst)}  {_fmt_inr(total_cgst)}  {_fmt_inr(total_sgst)}")

    # ── Section 3: Action Items ───────────────────────────────
    action_buckets = ["MISSING_IN_BOOKS", "MISSING_IN_2B", "REVIEW", "PROBABLE_MATCH"]
    action_rows    = df[df["bucket"].isin(action_buckets)]

    if not action_rows.empty:
        print(f"\n\n  ACTION ITEMS  ({len(action_rows)} entries require CA attention)")
        print(_divider())
        for bucket in action_buckets:
            sub = action_rows[action_rows["bucket"] == bucket]
            if sub.empty:
                continue
            meta = BUCKET_META.get(bucket, {"emoji": "·", "label": bucket})
            print(f"\n  {meta['emoji']} {meta['label'].upper()} — {len(sub)} entries")
            for _, row in sub.iterrows():
                vendor  = str(row.get("vendor_name","") or "")[:38]
                inv     = str(row.get("invoice_number","") or "")[:22]
                igst    = float(row.get("igst",0) or 0)
                cgst    = float(row.get("cgst",0) or 0)
                tax_str = f"IGST ₹{igst:,.2f}" if igst else f"CGST ₹{cgst:,.2f}"
                print(f"    {vendor:<38}  {inv:<22}  {tax_str}")
                reasoning = str(row.get("match_reasoning",""))
                if reasoning:
                    # Print first sentence only — keeps it readable
                    first = reasoning.split(".")[0] + "."
                    print(f"      → {first[:90]}")

    # ── Section 4: CA Remark Conflicts ────────────────────────
    conflicts = df[df["ca_remark_conflict"].fillna(False) == True]
    print(f"\n\n  CA REMARK CONFLICTS: {len(conflicts)}")
    if len(conflicts) > 0:
        print(_divider())
        for _, row in conflicts.iterrows():
            print(f"  ⚠️  {row.get('vendor_name','')} | {row.get('invoice_number','')}")
            print(f"     {row.get('match_reasoning','')[:100]}")
    else:
        print(f"  {_divider()}")
        print(f"  ✅ None — tool output matches all CA remarks")

    # ── Section 5: Confidence Breakdown ──────────────────────
    conf = df["confidence"].value_counts()
    high = conf.get("HIGH", 0)
    low  = conf.get("LOW", 0)
    med  = conf.get("MEDIUM", 0)
    pct  = round(high / len(df) * 100) if len(df) else 0
    print(f"\n\n  CONFIDENCE: HIGH {high} ({pct}%)  |  MEDIUM {med}  |  LOW {low}")
    print(_divider("═"))
    print(f"  ✅ Phase 5D complete — {len(df)} rows classified across {len(bucket_dfs)} buckets")
    print(_divider("═") + "\n")

    return bucket_dfs

