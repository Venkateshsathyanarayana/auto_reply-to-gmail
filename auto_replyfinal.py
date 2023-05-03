import imaplib
import smtplib
import email
import random
import time

# User defined parameters
GMAIL_USERNAME = '********@gmail.com'
GMAIL_PASSWORD = '****************'
VACATION_SUBJECT = 'Out of office'
VACATION_MESSAGE = 'Thank you for your email. I am currently out of the office and will not be able to respond until [date].'
LABEL_NAME = 'Vacation_Autoresponse'


def get_unread_emails():
    imap_server = imaplib.IMAP4_SSL("imap.gmail.com")
    imap_server.login(GMAIL_USERNAME, GMAIL_PASSWORD)
    imap_server.select("inbox")

    _, search_data = imap_server.search(None, "UNSEEN")
    email_ids = search_data[0].split()

    unread_emails = []
    for email_id in email_ids:
        _, data = imap_server.fetch(email_id, "(RFC822)")
        raw_email = data[0][1].decode("utf-8")
        email_message = email.message_from_string(raw_email)
        unread_emails.append(email_message)

    imap_server.close()
    imap_server.logout()

    return unread_emails, email_ids


def send_auto_response(email_from):
    email_to = email_from
    email_subject = VACATION_SUBJECT
    email_body = VACATION_MESSAGE.replace("[date]", time.strftime("%Y-%m-%d"))

    message = f"""From: {GMAIL_USERNAME}
To: {email_to}
Subject: {email_subject}

{email_body}
"""
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(GMAIL_USERNAME, GMAIL_PASSWORD)
    server.sendmail(GMAIL_USERNAME, email_to, message)
    server.quit()


def mark_as_replied(email_id):
    imap_server = imaplib.IMAP4_SSL("imap.gmail.com")
    imap_server.login(GMAIL_USERNAME, GMAIL_PASSWORD)
    imap_server.select("inbox")

    imap_server.store(email_id, "+FLAGS", "\\Seen \\Flagged")

    # Check if label exists before creating it
    rv, data = imap_server.list()
    if f"\\HasNoChildren {LABEL_NAME}" not in data:
        imap_server.create(LABEL_NAME)
    imap_server.store(email_id, "+X-GM-LABELS", LABEL_NAME)

    imap_server.close()
    imap_server.logout()


def run_auto_responder():
    print("Starting vacation autoresponder...")
    while True:
        unread_emails, email_ids = get_unread_emails()

        for i, email_message in enumerate(unread_emails):
            print(f"Unseen email #{i + 1}:")
            print(f"From: {email_message['From']}")
            print(f"Subject: {email_message['Subject']}")
            print(f"Body: {email_message.get_payload()}")
            print()

            email_from = email_message['From']
            email_id = email_ids[i]

            # Check if the email has already been replied to
            imap_server = imaplib.IMAP4_SSL("imap.gmail.com")
            imap_server.login(GMAIL_USERNAME, GMAIL_PASSWORD)
            imap_server.select("inbox")

            _, search_data = imap_server.search(None, f"HEADER Message-ID {email_id}")
            email_ids = search_data[0].split()

            if len(email_ids) == 0:
                # Email has not been replied to
                send_auto_response(email_from)
                mark_as_replied(email_id)

            imap_server.close()
            imap_server.logout()

        # Wait for random interval between 45 and 120 seconds
        wait_time = random.randint(45, 120)
        time.sleep(wait_time)


if __name__ == '__main__':
    run_auto_responder()

