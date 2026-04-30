import os
import imaplib
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials 
from google.auth.transport.requests import Request
import requests
import json
import email
from email.policy import default

load_dotenv()

def connect_IMAP_login():
    SCOPES = ['https://mail.google.com/']
    
    creds_path = os.getenv('CREDENTIALS_PATH')
    user_email = os.getenv('USER_EMAIL')
    token_path = 'token.json'

    flow = InstalledAppFlow.from_client_secrets_file(
        creds_path,
        SCOPES
    )

    creds = None

    # Check if token.json exists to skip login
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            # Refresh access token if expired 
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                creds_path,
                SCOPES
            )
            # open a browser window for the user to log in and authorize access
            creds = flow.run_local_server(port=0)
            
        # Save the credentials for the next run
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
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
    Parses raw IMAP data using the email library and returns a 
    Discord-ready dictionary payload.
    """
    # raw_data[0][1] contains the actual bytes of the email
    raw_bytes = raw_data[0][1]
    
    # Use the email library to parse the bytes into a message object
    msg = email.message_from_bytes(raw_bytes, policy=default)

    subject = msg['subject'] or "No Subject"
    sender = msg['from'] or "Unknown Sender"
    
    # Extract the body
    # This looks for the plain text version of the email specifically
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_content()
                break
    else:
        body = msg.get_content()

    # Clean up and truncate for Discord (1024 char limit for field values)
    clean_body = body.strip()
    if len(clean_body) > 500:
        clean_body = clean_body[:500] + "..."

    # Structure the Discord Webhook Payload
    notification = {
        "username": "Glowforge Monitor",
        "embeds": [{
            "title": f"{subject}",
            "color": 12960,  # GV Blue
            "fields": [
                {
                    "name": "From",
                    "value": sender,
                    "inline": True
                },
                {
                    "name": "Message Snippet",
                    "value": clean_body if clean_body else "No plain text content found.",
                    "inline": False
                }
            ],
            "footer": {
                "text": "Glowforge Real-time Updates"
            }
        }]
    }

    return notification

def send_to_discord(payload):
    webhook_url = os.getenv('DISCORD_WEBHOOK_URL')
    response = requests.post(
        webhook_url, 
        data=json.dumps(payload),
        headers={'Content-Type': 'application/json'}
    )
    
    if response.status_code == 204:
        print("Notification sent successfully!")
    else:
        print(f"Failed to send: {response.status_code}, {response.text}")

if __name__ == "__main__":
    session = connect_IMAP_login()
    if session:
        data = fetch_latest_glowforge_email(session)
    if data:
        notification = create_notification(data)
        send_to_discord(notification) # what we POST to the webhook
    session.logout()
        