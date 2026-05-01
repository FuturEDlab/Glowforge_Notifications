import os
import re
import imaplib
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials 
from google.auth.transport.requests import Request
import requests
import json
import email
from email.policy import default
from datetime import datetime, timedelta, timezone
import email.utils

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

def fetch_latest_glowforge_email(imap):
    """
    searches for most recent email from Glowforge and returns raw message data
    """
    # 1. Select inbox (readonly=True ensures we don't mark as read automatically)
    imap.select("INBOX", readonly=True)

    # 2. Search for emails from Glowforge
    # You can change 'FROM "Glowforge"' to specific subjects like 'SUBJECT "Appointment"'
    status, messages = imap.search(None, 'SUBJECT "New booking:"')

    if status != 'OK' or not messages[0]:
        print("No Glowforge emails found.")
        return None

    # 3. Get the ID of the latest email (the last one in the list)
    message_ids = messages[0].split()
    latest_id = message_ids[-1]

    # 4. Fetch the raw RFC822 data
    status, data = imap.fetch(latest_id, '(RFC822)')

    if status == 'OK':
        # --- NEW TIME-CHECK LOGIC ---
        raw_bytes = data[0][1]
        msg = email.message_from_bytes(raw_bytes, policy=default)
        
        # Parse the "Date" header from the email
        date_str = msg.get('Date')
        if date_str:
            mail_date = email.utils.parsedate_to_datetime(date_str)
            now = datetime.now(timezone.utc)
            
            # Check if the email is newer than 3 minutes
            if now - mail_date < timedelta(minutes=3):
                print(f"New email detected (Arrived: {mail_date}). Sending to Discord.")
                return data
            else:
                print(f"Latest unread email is from {mail_date}. Skipping to avoid duplicate notification.")
                return None
        # ----------------------------
    
    return None

def create_notification(raw_data):
    raw_bytes = raw_data[0][1]
    msg = email.message_from_bytes(raw_bytes, policy=default)

    subject = msg['subject'] or "No Subject"
    sender = msg['from'] or "Unknown Sender"
    
    body = ""
    # Look for content in both plain text and HTML
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            # If plain text exists, we use it
            if ctype == "text/plain":
                body = part.get_content()
                break
            # Otherwise, we grab the HTML as a backup
            elif ctype == "text/html":
                body = part.get_content()
    else:
        body = msg.get_content()

    # If it's HTML, we need to strip tags to make it readable
    if "<html" in body.lower():
        # Simple regex to remove HTML tags and clean up whitespace
        clean_body = re.sub('<[^<]+?>', ' ', body)
        clean_body = re.sub(r'\s+', ' ', clean_body).strip()
    else:
        clean_body = body.strip()

    # --- EXTRACT BOOKING DETAILS ---
    # Based on the screenshot, we want date and time range
    # Pattern for Date (e.g., Monday, May 4, 2026)
    date_match = re.search(r'([A-Z][a-z]+, [A-Z][a-z]+ \d{1,2}, \d{4})', clean_body)
    
    # Pattern for time range (e.g., 12:30 PM - 2:30 PM)
    time_match = re.search(r'(\d{1,2}:\d{2}\s?[AP]M\s?-\s?\d{1,2}:\d{2}\s?[AP]M)', clean_body)

    appt_date = date_match.group(0) if date_match else "Date not found"
    appt_time = time_match.group(0) if time_match else "Time not found"

    # Structure the Discord Webhook Payload with extra fields
    notification = {
        "username": "Glowforge Monitor",
        "embeds": [{
            "title": f"{subject}",
            "color": 12960,
            "fields": [
                {"name": "From", "value": sender, "inline": False},
                {"name": "Booking Date", "value": appt_date, "inline": False},
                {"name": "Booking Time", "value": appt_time, "inline": False}
            ],
            "footer": {"text": "Glowforge Real-time Updates"}
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
        