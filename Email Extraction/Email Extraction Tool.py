
# Library
import os
import csv
import datetime
import time
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Set up the scope and credentials
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'
CSV_FILE = 'extracted-email.csv'

# Check if credentials and token exists for authentication
# Take note gmail api authentication should have been setup
def get_gmail_service():
    creds = None
    # Check if token file exists and load the credentials
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    # Build the Gmail service
    service = build('gmail', 'v1', credentials=creds)
    print('\nAPI Connected\n')
    time.sleep(1)
    return service

# Timestamp convertion method
def convert_timestamp_to_datetime(timestamp):
    return datetime.datetime.fromtimestamp(int(timestamp) / 1000.0)


def find_specific_email_by_address(email_address):
    service = get_gmail_service()
    try:
        # Search for emails based on the provided email address
        search_query = f'from:{email_address} OR to:{email_address}'
        response = service.users().messages().list(userId='me', q=search_query).execute()
        messages = response.get('messages', [])

        if not messages:
            # Notify if exact mail doesn't exist
            print(f"No matching emails found for '{email_address}'.")
        else:
            # Count matched emails
            for count in range(len(messages)):
                print(f"Matched Email #, {count+1} decoded!...\n")
                time.sleep(.5)

            # Create CSV file
            with open(CSV_FILE, 'w', newline='', encoding='utf-8') as csvfile:
                csv_writer = csv.writer(csvfile)
                # Crate the title on the CSV FILE
                csv_writer.writerow(['SENDER', 'DATE', 'SUBJECT', 'EMAIL BODY', 'historyID', 'threadId'])
                print(f"\nList of matching emails for '{email_address}':")
                print("_" * 150)
                for message in messages:
                    # Check each searched mails and extract data
                    msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
                    sender = next((hdr['value'] for hdr in msg['payload']['headers'] if hdr['name'] == 'From'), 'N/A')
                    # Call method -> convert_timestamp_to_datetime
                    date_sent = convert_timestamp_to_datetime(msg['internalDate'])
                    subject = next((hdr['value'] for hdr in msg['payload']['headers'] if hdr['name'] == 'Subject'),
                                   'N/A')
                    emailBody = msg['snippet'] if 'snippet' in msg else 'N/A'
                    historyId = msg['historyId'] if 'historyId' in msg else 'N/A'
                    threadId = msg['threadId'] if 'threadId' in msg else 'N/A'

                    # Write the information in a row of each email searched
                    csv_writer.writerow([sender, date_sent, subject, emailBody, historyId, threadId])

                    time.sleep(.5)
                    # Display Information
                    print(f"SENDER: {sender}")
                    print(f"DATE: {date_sent}")
                    print(f"SUBJECT: {subject}")
                    print(f"MAIL BODY: {emailBody}")
                    print(f"HISTORY ID: {historyId}")
                    print(f"THREAD ID: {threadId}")
                    print("_" * 150)

                time.sleep(2)
                print('\nCreating Dataframe...\nDataframe Created\nDataframe exported to csv')
                time.sleep(1)
                print('\nEmail Extraction Done!...')


    except Exception as e:
        # Notify if error occurred
        print(f"Error: {e}")


if __name__ == '__main__':
    # Enter Specific Email to search
    email_address = 'emailtosearch@gmail.com'

    # Call the method (find_specific_email_by_address) to find the sent email
    find_specific_email_by_address(email_address)
