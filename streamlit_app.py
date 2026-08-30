import tempfile
from datetime import datetime, timedelta

import requests
import streamlit as st

import report_builder as core
import kpi_fetch

st.set_page_config(page_title="Hutch KPI Auto Report", page_icon="📊")

st.title("Hutch KPI Auto Report")
st.caption("Fetches KPI data directly from the RAN portal and builds the Excel report — no downloads, no manual exports.")

with st.form("kpi_form"):
    col1, col2 = st.columns(2)
    email = col1.text_input("Username / Email")
    password = col2.text_input("Password", type="password")

    region = st.selectbox("Region", list(kpi_fetch.DISTRICTS.keys()))

    today = datetime.now()
    col3, col4 = st.columns(2)
    start_date = col3.date_input("Start Date", value=today - timedelta(days=90))
    end_date = col4.date_input("End Date", value=today)

    include_network = st.checkbox("Include Network totals (BSS tab)", value=True)

    submitted = st.form_submit_button("Fetch & Generate Report", type="primary")

st.caption("Your credentials are used only for this request and are never stored.")

if submitted:
    if not email or not password:
        st.error("Enter your username and password.")
        st.stop()

    log_area = st.empty()
    log_lines = []

    def log(msg):
        log_lines.append(msg)
        log_area.code("\n".join(log_lines))

    try:
        with st.spinner("Working..."):
            session = requests.Session()
            log("Logging in...")
            user = kpi_fetch.login(session, email, password)
            log(f"Logged in as {user['FirstName']} {user['LastName']}")

            with tempfile.TemporaryDirectory() as tmp_dir:
                kpi_fetch.fetch_region_report(
                    session, region,
                    start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"),
                    tmp_dir, include_network, log,
                )

                log("Building Excel report...")
                output_path = core.build_excel_report(tmp_dir, region)
                log("Done!")

                with open(output_path, "rb") as f:
                    report_bytes = f.read()

        st.success("Report generated.")
        st.download_button(
            "Download Excel Report",
            data=report_bytes,
            file_name=f"{region}_Region_Final_Sales_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Failed: {e}")
