import os
import tempfile
from datetime import datetime, timedelta

import requests
import streamlit as st

import report_builder as core
import kpi_fetch
import pptx_report
from email_report import send_report_email

st.set_page_config(page_title="Sales Report automation"", page_icon="📊")

st.title("Sales Report automation")
st.caption("Fetches KPI data directly from the RAN portal and builds the Excel report + presentation for the region")

with st.form("kpi_form"):
    col1, col2 = st.columns(2)
    email = col1.text_input("RANDORS Username / Email")
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
                excel_path = core.build_excel_report(tmp_dir, region)
                with open(excel_path, "rb") as f:
                    st.session_state["excel_bytes"] = f.read()
                log("Excel done.")

                st.session_state["pptx_bytes"] = None
                template_path = os.path.join(os.path.dirname(__file__), "templates", f"{region}.pptx")
                if os.path.exists(template_path):
                    log("Updating presentation charts...")
                    pptx_out_path = os.path.join(tmp_dir, "report.pptx")
                    pptx_report.build_pptx_report(template_path, excel_path, pptx_out_path, log=log)
                    with open(pptx_out_path, "rb") as f:
                        st.session_state["pptx_bytes"] = f.read()
                    log("Presentation done.")
                else:
                    log(f"No presentation template configured for {region} yet — skipping.")

        st.session_state["region"] = region
        st.session_state["recipient_email"] = email
        st.success("Report generated.")
    except Exception as e:
        st.error(f"Failed: {e}")

if st.session_state.get("excel_bytes"):
    region = st.session_state["region"]
    st.download_button(
        "Download Excel Report",
        data=st.session_state["excel_bytes"],
        file_name=f"{region}_Region_Final_Sales_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if st.session_state.get("pptx_bytes"):
        st.download_button(
            "Download Presentation",
            data=st.session_state["pptx_bytes"],
            file_name=f"{region}_Region_Monthly_Review.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    st.divider()
    recipient = st.session_state["recipient_email"]
    if st.button(f"Email these files to {recipient}"):
        gmail_email = st.secrets.get("gmail_email")
        gmail_app_password = st.secrets.get("gmail_app_password")
        if not gmail_email or not gmail_app_password:
            st.error("Email sending isn't configured yet — the app owner needs to add gmail_email and gmail_app_password in Streamlit Cloud's app Secrets.")
        else:
            attachments = [(
                f"{region}_Region_Final_Sales_Report.xlsx",
                st.session_state["excel_bytes"],
                "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )]
            if st.session_state.get("pptx_bytes"):
                attachments.append((
                    f"{region}_Region_Monthly_Review.pptx",
                    st.session_state["pptx_bytes"],
                    "application", "vnd.openxmlformats-officedocument.presentationml.presentation",
                ))
            try:
                with st.spinner("Sending..."):
                    send_report_email(
                        gmail_email, gmail_app_password, recipient,
                        subject=f"{region} Region KPI Report",
                        body=f"Attached: the {region} region KPI report" + (" and presentation." if st.session_state.get("pptx_bytes") else "."),
                        attachments=attachments,
                    )
                st.success(f"Sent to {recipient}.")
            except Exception as e:
                st.error(f"Failed to send email: {e}")
