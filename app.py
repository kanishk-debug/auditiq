import streamlit as st
import tempfile
import os
import sys
from pathlib import Path
from collections import Counter

sys.path.insert(0, str(Path(__file__).parent))

from core.gstr2b_parser import parse_gstr2b
from core.tally_parser  import parse_tally
from core.reconciler    import reconcile
from core.excel_output  import generate_reconciliation_excel

from core.invoice_processing.vision         import extract_invoice_from_bytes
from core.invoice_processing.cleaner        import clean_invoice
from core.invoice_processing.validator      import validate_invoice, generate_invoice_id
from core.invoice_processing.itc_classifier import generate_audit_report
from core.invoice_processing.invoice_excel  import generate_invoice_excel

st.set_page_config(page_title="AuditIQ", page_icon="📊", layout="centered")

page = st.sidebar.radio("Navigation", ["Reconciliation", "Invoice Processing"], index=0)

if page == "Reconciliation":
    st.title("📊 AuditIQ")
    st.subheader("GST Reconciliation Engine")
    st.write("Upload your Tally export and GSTR-2B Excel to get a complete reconciliation report.")
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Tally Export**")
        tally_file = st.file_uploader("IGST Input + CGST Input + RCM sheets", type=["xlsx"], key="tally")
    with col2:
        st.markdown("**GSTR-2B**")
        gstr2b_file = st.file_uploader("Downloaded from GST Portal", type=["xlsx"], key="gstr2b")

    st.divider()
    col3, col4 = st.columns(2)
    with col3:
        client_name = st.text_input("Client Name", placeholder="e.g. ABC Pvt Ltd")
    with col4:
        period = st.text_input("Period", placeholder="e.g. February 2026")

    st.divider()
    if not (tally_file and gstr2b_file):
        st.caption("Upload both files above to enable reconciliation.")

    if st.button("Run Reconciliation", type="primary", disabled=not (tally_file and gstr2b_file), use_container_width=True):
        try:
            with st.status("Parsing GSTR-2B...", expanded=True):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(gstr2b_file.read()); gstr2b_path = tmp.name
                gstr2b_result = parse_gstr2b(gstr2b_path)
                st.write(f"GSTR-2B: {len(gstr2b_result['data'])} rows parsed")

            with st.status("Parsing Tally books...", expanded=True):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(tally_file.read()); tally_path = tmp.name
                books_result = parse_tally(tally_path)
                st.write(f"Tally: {len(books_result['data'])} rows parsed")

            with st.status("Running reconciliation...", expanded=True):
                recon = reconcile(gstr2b_data=gstr2b_result["data"], books_data=books_result["data"], period=period or "Current Period")
                st.write("Reconciliation complete")

            with st.status("Generating Excel report...", expanded=True):
                output_path = tempfile.mktemp(suffix=".xlsx")
                generate_reconciliation_excel(recon_result=recon, output_path=output_path, period=period or "Current Period")
                with open(output_path, "rb") as f: excel_bytes = f.read()
                os.unlink(output_path)
                st.write("Excel ready")

            st.divider()
            st.subheader("Reconciliation Summary")
            summary = recon["summary_report"]
            buckets = summary.get("bucket_counts", {})
            itc = summary.get("claimable_itc", {})

            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Matched", buckets.get("MATCHED",0))
            c2.metric("Missing in Books", buckets.get("MISSING_IN_BOOKS",0))
            c3.metric("Ineligible", buckets.get("INELIGIBLE",0))
            c4.metric("RCM", buckets.get("RCM",0))

            c5,c6,c7 = st.columns(3)
            c5.metric("Previous Month", buckets.get("PREVIOUS_MONTH",0))
            c6.metric("Missing in 2B", buckets.get("MISSING_IN_2B",0))
            c7.metric("Needs Review", buckets.get("REVIEW",0)+buckets.get("PROBABLE_MATCH",0))

            igst=float(itc.get("igst",0)); cgst=float(itc.get("cgst",0)); sgst=float(itc.get("sgst",0))
            st.success(f"**Claimable ITC this month: Rs {igst+cgst+sgst:,.2f}**  (IGST Rs {igst:,.2f} | CGST Rs {cgst:,.2f} | SGST Rs {sgst:,.2f})")

            st.divider()
            safe = (client_name or period or "Report").replace(" ","_")
            st.download_button(label="Download Reconciliation Excel", data=excel_bytes, file_name=f"AuditIQ_{safe}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            try: os.unlink(gstr2b_path); os.unlink(tally_path)
            except: pass

        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.exception(e)

elif page == "Invoice Processing":
    st.title("Invoice Processing")
    st.subheader("Extract, validate and classify invoices")
    st.write("Upload invoice images. AuditIQ extracts all GST fields and checks ITC eligibility.")
    st.divider()

    uploaded_files = st.file_uploader("Upload invoices (JPG, PNG)", type=["jpg","jpeg","png"], accept_multiple_files=True)
    client_name_inv = st.text_input("Client Name", placeholder="e.g. ABC Pvt Ltd", key="inv_client")

    st.divider()
    if not uploaded_files:
        st.caption("Upload at least one invoice image above.")

    if st.button("Extract & Classify", type="primary", disabled=not uploaded_files, use_container_width=True):
        audit_reports=[]; errors=[]
        progress=st.progress(0); status_text=st.empty()

        for i, f in enumerate(uploaded_files):
            status_text.text(f"Processing {f.name}  ({i+1} of {len(uploaded_files)})...")
            try:
                raw=extract_invoice_from_bytes(f.read())
                cleaned=clean_invoice(raw)
                validation=validate_invoice(cleaned)
                inv_id=generate_invoice_id(cleaned)
                report=generate_audit_report(cleaned, validation)
                report["invoice_id"]=inv_id; report["filename"]=f.name
                audit_reports.append(report)
            except Exception as e:
                errors.append(f"{f.name}: {e}")
            progress.progress((i+1)/len(uploaded_files))

        status_text.text("Done.")

        if audit_reports:
            statuses=Counter(r["itc_status"] for r in audit_reports)
            st.divider(); st.subheader("Results")
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Total", len(audit_reports))
            c2.metric("Eligible", statuses.get("ELIGIBLE",0))
            c3.metric("Blocked", statuses.get("BLOCKED",0))
            c4.metric("Review", statuses.get("REVIEW_NEEDED",0)+statuses.get("RCM",0))

            out=tempfile.mktemp(suffix=".xlsx")
            generate_invoice_excel(audit_reports, out)
            with open(out,"rb") as f: xb=f.read()
            os.unlink(out)
            safe=(client_name_inv or "Invoices").replace(" ","_")
            st.download_button(label="Download Invoice Audit Excel", data=xb, file_name=f"AuditIQ_{safe}_Invoices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

        if errors:
            st.warning(f"{len(errors)} file(s) had issues:")
            for e in errors: st.caption(f"- {e}")