import streamlit as st
import os
from email_parser_reporting import connect_to_email, fetch_attachments, parse_files, save_summary

st.title("📧 Email Parser for Analytics Reporting")
st.markdown("Automate the extraction of Excel/CSV attachments from emails and generate summary reports.")

if st.button("Run Email Parser"):
    server = connect_to_email()
    st.info("Connected to email...")

    files = fetch_attachments(server)
    if not files:
        st.warning("No files found in the last 7 days.")
    else:
        st.success(f"Found {len(files)} files.")
        df = parse_files(files)
        save_summary(df)
        st.download_button("Download Summary Report", open("Summary_Report.xlsx", "rb"))
        st.dataframe(df.head())
