import tempfile
from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="BSNL FTTH Warangal Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("📊 BSNL FTTH Warangal Dashboard")
st.caption(
    "Upload any supported Excel report from any folder/location on your computer. "
    "The original filename is not required to be OLT_Wise_Provisioning.xlsx."
)

uploaded = st.file_uploader(
    "📂 Select FTTH Excel Report",
    type=["xlsx", "xlsm", "xltx", "xltm"],
    help=(
        "You can select the Excel file from ANY folder/location accessible "
        "through your computer's file picker. The file may have any filename."
    ),
)

if uploaded:
    st.success(f"✅ File selected: {uploaded.name}")

    if st.button(
        "🚀 Generate FTTH Warangal Dashboard",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("Processing uploaded Excel report… please wait."):
            try:
                with tempfile.TemporaryDirectory() as td:
                    work = Path(td)

                    # The user's original filename is retained only for logging.
                    # Internally, the uploaded bytes are saved to a temporary file.
                    input_file = work / uploaded.name
                    output_xlsx = work / "FTTH_Warangal_Dashboard.xlsx"
                    output_html = work / "FTTH_Warangal_Dashboard.html"

                    input_file.write_bytes(uploaded.getbuffer())

                    from report_processor import run_report

                    xlsx_path, html_path, log = run_report(
                        input_file,
                        output_xlsx,
                        output_html,
                    )

                    st.session_state["xlsx_bytes"] = xlsx_path.read_bytes()
                    st.session_state["html_bytes"] = html_path.read_bytes()
                    st.session_state["log"] = log
                    st.session_state["source_filename"] = uploaded.name

                st.session_state["generated"] = True
                st.success("✅ Dashboard generated successfully.")

            except Exception as e:
                st.session_state["generated"] = False
                st.error("❌ Report generation failed.")
                st.exception(e)

if st.session_state.get("generated"):
    st.subheader("📥 Download Dashboard")

    st.info(
        f"Source file: **{st.session_state.get('source_filename', 'Uploaded Excel')}**"
    )

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
    "💡 The user does NOT need to create an uploads folder or rename the Excel file. "
    "When the dashboard is deployed, the browser's file picker lets the user choose "
    "the Excel file from any accessible folder on their computer. The selected file "
    "is uploaded securely to the Streamlit session and processed temporarily."
)
