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

def check_inbox(imap):
    # This lists all available folders/labels
    status, mailboxes = imap.list()
    if status == 'OK':
        print("\n--- Available Mailboxes ---")
        for mailbox in mailboxes:
            print(mailbox.decode())
    else:
        print("Failed to retrieve mailboxes.")

if __name__ == "__main__":
    session = connect_IMAP_login()
    if session:
        check_inbox(session)
        # Always logout when finished
        session.logout()
        