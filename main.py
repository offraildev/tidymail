import os
from prefect import flow, task
from ingest_emails import login_to_gmail
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

EMAIL = os.environ.get("EMAIL")
APP_PASSWORD = os.environ.get("APP_PASSWORD")


JOB_ALERTS = {"linkedin": '"LinkedIn Job Alerts"', "iimjobs": '"info@iimjobs.com"'}

@task(log_prints=True)
def mark_emails_for_deletion(connection, email_ids):
    for email_id in email_ids:
        status, _ = connection.store(email_id, "+FLAGS", "\\Deleted")
        if status != "OK":
            print(f"Error marking emails for deletion, emails: {email_ids}")

@task(log_prints=True, retries=3)
def mark_job_alert_emails(connection):
    status, email_ids_linkedin = connection.search(
        None, f'(FROM {JOB_ALERTS["linkedin"]})'
    )
    print(f"{status}")
    status, email_ids_iimjobs = connection.search(
        None, f'(FROM {JOB_ALERTS["iimjobs"]})'
    )
    print(f"{status}")
    mark_emails_for_deletion(connection, email_ids_linkedin[0].split())
    mark_emails_for_deletion(connection, email_ids_iimjobs[0].split())


@flow(name="rm_job_alerts")
def rm_job_alerts():
    connection = login_to_gmail(EMAIL, APP_PASSWORD)
    mailbox = "INBOX"

    status, _ = connection.select(mailbox)
    if status == "OK":
        print(f"{mailbox} Selected")
    else:
        print(f"Error Selecting {mailbox} ")

    job_alert_on = False
    if not job_alert_on:
        mark_job_alert_emails(connection)

    status, _ = connection.expunge()
    if status != "OK":
        print(f"Error expunging deleted emails")
    print("Email deletion process completed.")

    connection.close()
    connection.logout()


if __name__ == "__main__":
    rm_job_alerts.serve(name="rm_job_alert_deployment")
