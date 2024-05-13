import os
import re
import imaplib
import email
import argparse
from datetime import datetime

from tqdm import tqdm
import pandas as pd

from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

EMAIL = os.environ.get("EMAIL")
APP_PASSWORD = os.environ.get("APP_PASSWORD")


def login_to_gmail(email: str, app_password: str) -> imaplib.IMAP4_SSL:
    try:
        connection = imaplib.IMAP4_SSL("imap.gmail.com")
        connection.login(email, app_password)
        print("Login successful")
        return connection
    except imaplib.IMAP4.error:
        raise Exception(
            f"Login failed with credentials: email-id: {EMAIL} and application password {APP_PASSWORD}"
        )


def get_part_body(part):
    """
    Extracts text from a message part.
    """
    try:
        return part.get_payload(decode=True).decode()
    except (UnicodeDecodeError, AttributeError):
        return


def bodyText_sentFiles(msg):
    body_texts = []
    sent_files = []
    if msg.is_multipart():
        for part in msg.walk():
            if "attachment" not in str(part.get("Content-Disposition")):
                text = get_part_body(part)
                if text:
                    body_texts.append(text)
            else:
                sent_files.append(part["Content-Type"])
    else:
        text = get_part_body(msg)
        if text:
            body_texts.append(text)
    return "\n".join(body_texts), ",".join(sent_files)


def get_email_meta(email_id_list, connection):
    meta = []
    for entry in tqdm(email_id_list):
        status, msg_data = connection.fetch(entry, "(RFC822)")
        if status == "OK":
            msg = email.message_from_bytes(msg_data[0][1])
            bodyText, sentFiles = bodyText_sentFiles(msg)
            meta.append(
                {
                    "subject": msg.get("Subject"),
                    "from": msg.get("From"),
                    "body": bodyText,
                    "sent_files": sentFiles,
                    "date": msg.get("Date"),
                }
            )
    return meta


def format_date(date):
    subbed = re.sub(r"\([A-Z]+\)", "", date).strip()
    try:
        pattern = "%a, %d %b %Y %H:%M:%S %z"
        return datetime.strptime(subbed, pattern)
    except ValueError:
        pattern = "%a, %d %b %Y %H:%M:%S"
        return datetime.strptime(" ".join(subbed.split()[:-1]), pattern)


def mailbox_to_df(email_id_list, connection):
    emails_meta = get_email_meta(email_id_list, connection)
    df = pd.DataFrame(emails_meta)
    # convert date to datetime type

    df["date"] = df["date"].apply(lambda v: format_date(v))
    return df


def parse_arguments():
    """
    Parses command-line arguments using argparse.
    """
    parser = argparse.ArgumentParser(
        description="Ingest emails from a given mailbox, strip out meta data and save in tabular format."
    )
    parser.add_argument(
        "-m", "--mailbox", type=str, help="Name of the mailbox to select"
    )
    parser.add_argument(
        "-o", "--out_path", type=str, help="Path to csv file for dumping df output."
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_arguments()

    connection = login_to_gmail(EMAIL, APP_PASSWORD)
    status, _ = connection.select(args.mailbox)
    if status != "OK":
        raise Exception("Mailbox could not be selected")

    # get all mails from mailbox folder
    status, email_ids = connection.search(None, "ALL")
    if status != "OK":
        raise Exception("Emails from the mailbox folder could not be fetched.")
    email_id_list = email_ids[0].split()

    emails_df = mailbox_to_df(email_id_list, connection)
    emails_df.to_csv(args.out_path, index=False)
    print(f"Out file save to path: {args.out_path}")

    connection.close()
    connection.logout()
