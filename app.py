import os
import tempfile
from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="BSNL FTTH Warangal Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("📊 BSNL FTTH Warangal Dashboard")
st.caption("Upload the OLT Wise Provisioning Excel file and generate the dashboard.")

uploaded = st.file_uploader(
    "Upload OLT_Wise_Provisioning.xlsx",
    type=["xlsx"],
    help="Select the Excel report exported from the BSNL portal.",
)

if uploaded:
    st.success(f"File selected: {uploaded.name}")

    if st.button("🚀 Generate Dashboard", type="primary", use_container_width=True):
        with st.spinner("Processing report… please wait."):
            try:
                with tempfile.TemporaryDirectory() as td:
                    work = Path(td)
                    input_file = work / "OLT_Wise_Provisioning.xlsx"
                    output_xlsx = work / "FTTH_Warangal_Dashboard.xlsx"
                    output_html = work / "FTTH_Warangal_Dashboard.html"

                    input_file.write_bytes(uploaded.getbuffer())

                    from report_processor import run_report
                    xlsx_path, html_path, log = run_report(
                        input_file, output_xlsx, output_html
                    )

                    st.session_state["xlsx_bytes"] = xlsx_path.read_bytes()
                    st.session_state["html_bytes"] = html_path.read_bytes()
                    st.session_state["log"] = log

                st.session_state["generated"] = True
                st.success("✅ Dashboard generated successfully.")

            except Exception as e:
                st.session_state["generated"] = False
                st.error("❌ Report generation failed.")
                st.exception(e)

if st.session_state.get("generated"):
    st.subheader("Download")
    c1, c2 = st.columns(2)

    with c1:
        st.download_button(
            "⬇️ Download Excel Dashboard",
            data=st.session_state["xlsx_bytes"],
            file_name="FTTH_Warangal_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with c2:
        st.download_button(
            "⬇️ Download HTML Dashboard",
            data=st.session_state["html_bytes"],
            file_name="FTTH_Warangal_Dashboard.html",
            mime="text/html",
            use_container_width=True,
        )

    with st.expander("Processing log"):
        st.code(st.session_state.get("log", "Completed."))

st.divider()
st.info(
    "No Python installation is required on the user's PC when this app is deployed "
    "to Streamlit Community Cloud. The PC only needs a browser and internet access."
)
