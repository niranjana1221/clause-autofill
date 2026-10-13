import io
import pandas as pd
import streamlit as st
from autofill_clauses import fill

st.set_page_config(page_title="Clause Auto-Fill", page_icon="🛠️", layout="wide")
st.title("🛠️ Engine Template – Clause Auto-Fill")
st.caption("Upload the engine Table PDF and the Excel template. The 'Table CLs. no' columns are filled automatically.")

with st.sidebar:
    st.header("Settings")
    auto = st.checkbox("Auto-detect main table pages", value=True,
                       help="Main table = pages up to 'Sheet No : N of N'. Untick to set it yourself.")
    pages = None if auto else st.number_input("Number of main-table pages", 1, 300, 16)
    st.markdown("**Colours in the result**\n\n- No colour: confident match\n- 🟨 Yellow: check in PDF\n- 🟥 Red: not found in table (see annexures)")

c1, c2 = st.columns(2)
pdf_file = c1.file_uploader("1) Engine Table PDF", type=["pdf"])
xls_file = c2.file_uploader("2) Excel template", type=["xlsx"])

if st.button("▶ Fill clauses", type="primary", disabled=not (pdf_file and xls_file)):
    try:
        with st.spinner("Reading PDF table and matching clauses..."):
            wb, report, n_clauses = fill(io.BytesIO(xls_file.getvalue()), io.BytesIO(pdf_file.getvalue()), pages)
            buf = io.BytesIO(); wb.save(buf)
        st.session_state["result"] = (buf.getvalue(), report, n_clauses, xls_file.name)
    except Exception as e:
        st.error(f"Could not process the files: {e}")

if "result" in st.session_state:
    data, report, n_clauses, name = st.session_state["result"]
    df = pd.DataFrame(report, columns=["Row", "Column", "Clause(s)", "Status"])
    ok = df["Status"].isin(["ok", "override"]).sum()
    chk = (df["Status"] == "check").sum()
    nf = df["Status"].isin(["NOT FOUND", "empty"]).sum()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Clauses read from PDF", n_clauses)
    m2.metric("Confident", int(ok)); m3.metric("Check (yellow)", int(chk)); m4.metric("Not found (red)", int(nf))
    if n_clauses < 50:
        st.warning("Very few clauses were read from the PDF. Is this the right PDF, and is the page count correct?")
    st.download_button("⬇ Download filled Excel", data, file_name=name.replace(".xlsx", "_filled.xlsx"),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.subheader("Match report")
    st.dataframe(df, use_container_width=True, hide_index=True)