import os
import imaplib
import base64
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow

load_dotenv()

def connect_IMAP_login():
    SCOPES = ['https://mail.google.com/']
    
    creds_path = os.getenv('CREDENTIALS_PATH')
    user_email = os.getenv('USER_EMAIL')

    flow = InstalledAppFlow.from_client_secrets_file(
        creds_path,
        SCOPES
    )

    creds = flow.run_local_server(port=0)   # opens a browser for authentication
    
    # 1. Connect to the Gmail IMAP server
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    
    # 2. Generate the OAuth2 string
    # Gmail expects: "user=" + email + "\1auth=Bearer " + token + "\1\1"
    auth_string = f"user={user_email}\1auth=Bearer {creds.token}\1\1"
    
    # 3. Authenticate
    try:
        imap.authenticate('XOAUTH2', lambda x: auth_string.encode('utf-8'))
        print("Successfully authenticated with IMAP!")
        return imap
    except Exception as e:
        print(f"Failed to connect: {e}")
        return None

'''
test function to list all mailboxes/folders in the account

def check_inbox(imap):
    # This lists all available folders/labels
    status, mailboxes = imap.list()
    if status == 'OK':
        print("\n--- Available Mailboxes ---")
        for mailbox in mailboxes:
            print(mailbox.decode())
    else:
        print("Failed to retrieve mailboxes.")
'''

def fetch_latest_glowforge_email(imap):
    """
    searches for most recent email from Glowforge and returns raw message data
    """
    # 1. Select inbox (readonly=True ensures we don't mark as read automatically)
    imap.select("INBOX", readonly=True)

    # 2. Search for emails from Glowforge
    # You can change 'FROM "Glowforge"' to specific subjects like 'SUBJECT "Appointment"'
    status, messages = imap.search(None, 'FROM "Glowforge Sign-up"')

    if status != 'OK' or not messages[0]:
        print("No Glowforge emails found.")
        return None

    # 3. Get the ID of the latest email (the last one in the list)
    message_ids = messages[0].split()
    latest_id = message_ids[-1]

    # 4. Fetch the raw RFC822 data
    status, data = imap.fetch(latest_id, '(RFC822)')

    if status == 'OK':
        print(f"Successfully fetched email ID: {latest_id.decode()}")
        return data
    
    return None

def create_notification(raw_data):
    """
    parses raw IMAP fetch data (tuple) to extract key info 
    and returns a formatted Discord payload without using the email library.
    """
    # raw_data is usually the list returned by imap.fetch: [(b'1 (RFC822 {1234}', b'RawContent'), b')']
    # We grab the actual content (the second element of the first tuple)
    email_content = raw_data[0][1].decode('utf-8', errors='ignore')
    
    # Simple manual header parsing
    lines = email_content.split('\r\n')
    subject = "Unknown Subject"
    sender = "Unknown Sender"
    body_start_index = 0

    for i, line in enumerate(lines):
        if line.startswith("Subject: "):
            subject = line.replace("Subject: ", "")
        if line.startswith("From: "):
            sender = line.replace("From: ", "")
        # The body starts after the first completely empty line
        if not line.strip() and body_start_index == 0:
            body_start_index = i + 1
            break

    # join the lines after header to get the body
    raw_body = "\n".join(lines[body_start_index:])
    # clean up simple whitespace and truncate for Discord
    clean_body = raw_body.strip()[:400] + "..." if len(raw_body) > 400 else raw_body.strip()

    # Structure the Discord Webhook Payload
    notification = {
        "username": "Glowforge Monitor",
        "embeds": [{
            "title": "------New Appointment Notification------",
            "description": f"**Subject:** {subject}\n**From:** {sender}",
            "fields": [
                {
                    "name": "Message Snippet",
                    "value": clean_body if clean_body else "No plain text content found.",
                    "inline": False
                }
            ],
            "color": 12960 # GV Blue
        }]
    }

    return notification

if __name__ == "__main__":
    session = connect_IMAP_login()
    if session:
        data = fetch_latest_glowforge_email(session)
    if data:
        notification = create_notification(data)
        print(notification) # what we POST to the webhook
    session.logout()
        