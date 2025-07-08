The Email Parser for Analytics Reporting is a Python-based Streamlit application designed to automate the extraction of data from email attachments (Excel or CSV), parse them, and generate summary reports in Excel format.

Key functionality includes:

1) Secure Email Connection: Connects to an email server using credentials stored in a .env file.

2) Attachment Extraction: Automatically scans recent emails (within the last 7 days) to find and download relevant file attachments.

3) Data Parsing: Reads structured data from attachments using pandas and processes it into a unified DataFrame.

4) Summary Report Generation: Aggregates and summarizes the parsed data, exporting it to a file called Summary_Report.xlsx.

5) Download Option: Users can download the final summary report directly through a button in the Streamlit interface.

This tool is ideal for automating analytics workflows involving recurring email reports, especially in finance, sales, or operations dashboards.
