import os
import io
import datetime
import pandas as pd
from imapclient import IMAPClient
import pyzmail
from dotenv import load_dotenv

load_dotenv()
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
HOST = 'imap.gmail.com'
SAVE_DIR = './downloads/'
REPORT_FILE = 'Summary_Report.xlsx'
os.makedirs(SAVE_DIR, exist_ok=True)

def connect_to_email():
    server = IMAPClient(HOST, ssl=True)
    server.login(EMAIL, PASSWORD)
    server.select_folder('INBOX', readonly=True)
    return server

def fetch_attachments(server):
    since_date = datetime.date.today() - datetime.timedelta(days=30)
    messages = server.search(['SINCE', since_date])
    response = server.fetch(messages, ['BODY[]'])
    files = []
    for msgid, data in response.items():
        msg = pyzmail.PyzMessage.factory(data[b'BODY[]'])
        for part in msg.mailparts:
            if part.filename and (part.filename.endswith('.xlsx') or part.filename.endswith('.csv')):
                filepath = os.path.join(SAVE_DIR, part.filename)
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload())
                files.append(filepath)
    return files

def parse_files(file_list):
    summary_df = pd.DataFrame()
    for file in file_list:
        try:
            if file.endswith('.xlsx'):
                df = pd.read_excel(file)
            elif file.endswith('.csv'):
                df = pd.read_csv(file)
            df['Source File'] = os.path.basename(file)
            summary_df = pd.concat([summary_df, df], ignore_index=True)
        except Exception as e:
            print(f"Error parsing {file}: {e}")
    return summary_df

def save_summary(df):
    df.to_excel(REPORT_FILE, index=False)
    print(f"Summary saved to {REPORT_FILE}")

def main():
    server = connect_to_email()
    files = fetch_attachments(server)
    if not files:
        print("No files found.")
        return
    summary_df = parse_files(files)
    save_summary(summary_df)

if __name__ == '__main__':
    main()
