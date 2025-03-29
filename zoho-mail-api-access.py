import requests
import csv
import time
import re
import html
import os
import sys
from datetime import datetime


# Replace these values with your Zoho credentials
CLIENT_ID = ""
CLIENT_SECRET = ""
REFRESH_TOKEN = ""
ACCOUNT_ID = ""
FOLDER_ID = ""
OUTPUT_CSV = ""
MESSAGE_ID = ""
FROM_EMAIL = ""
SUBJECT = ""


# Zoho API Endpoints
TOKEN_URL = "https://accounts.zoho.in/oauth/v2/token"
MAILS_API_URL = f"https://mail.zoho.in/api/accounts/{ACCOUNT_ID}/messages/view?folderId={FOLDER_ID}&includeto=true"
DELETE_MAIL_API_URL = f"https://mail.zoho.in/api/accounts/{ACCOUNT_ID}/folders/{FOLDER_ID}/messages"
SEND_MAIL_API_URL = f"https://mail.zoho.in/api/accounts/{ACCOUNT_ID}/messages"
UPLOAD_ATTACHMENT_URL = f"https://mail.zoho.in/api/accounts/{ACCOUNT_ID}/messages/attachments"

# Function to get a new access token
def get_access_token():
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": REFRESH_TOKEN,
        "grant_type": "refresh_token"
    }
    response = requests.post(TOKEN_URL, data=payload)
    data = response.json()
    return data.get("access_token")

# Function to fetch emails in a folder
def fetch_emails(access_token):
    sent_to_list = set()
    email_ids = []  # Store message IDs for deletion
    has_more = True
    page = 1
    limit = 200  # Fetch 200 emails per request
    start = 1

    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}",
        "Content-Type": "application/json"
    }

    while has_more:
        params = {
            "start": start,
            "limit": limit
        }

        response = requests.get(MAILS_API_URL, headers=headers, params=params)
        if response.status_code != 200:
            print(f"Error fetching emails: {response.json()}")
            break

        response_data = response.json()

        data = response_data["data"]
        if not data:
            break

        # Extract sent-to email addresses
        for email in data:
            sent_to = email["toAddress"].lower()
            sent_to = html.unescape(sent_to)
            sent_to = sent_to.replace(">","")
            sent_to = sent_to.replace("<","")
            sent_to = re.sub(r'".*?"', '', sent_to)
            sent_to_list.add(sent_to)
            email_ids.append(email["messageId"])

        time.sleep(1)  # Avoid hitting rate limits
        start = start + 200
    return list(sent_to_list), email_ids

# Function to delete emails
def delete_emails(access_token, email_ids):
    if not email_ids:
        return

    headers = {
        "Authorization": f"Zoho-oauthtoken {access_token}",
        "Content-Type": "application/json"
    }

    deleted_count = 0
    failed_count = 0

    for message_id in email_ids:
        response = requests.delete(f"{DELETE_MAIL_API_URL}/{message_id}", headers=headers)
        
        if response.status_code == 200:
            print(f"Successfully deleted email with messageId: {message_id}")
            deleted_count += 1
        else:
            print(f"Error deleting email {message_id}: {response.json()}")
            failed_count += 1
    print(f"Total emails deleted: {deleted_count}")
    print(f"Total emails failed to delete: {failed_count}")
    return deleted_count, failed_count


# Function to load existing email IDs from the CSV file
def load_existing_emails():
    existing_emails = set()
    if os.path.exists(OUTPUT_CSV):
        with open(OUTPUT_CSV, mode="r", newline="", encoding="utf-8") as file:
            reader = csv.reader(file)
            next(reader, None)  # Skip header row
            for row in reader:
                if row:
                    existing_emails.add(row[0].strip().lower())
    return existing_emails

# Function to upload attachment
def upload_attachment(access_token):
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}

    with open(OUTPUT_CSV, "rb") as file:
        file_data = file.read()
    
    params = {"fileName": OUTPUT_CSV}
    response = requests.post(UPLOAD_ATTACHMENT_URL, headers=headers, params=params, data=file_data)

    if response.status_code == 200:
        return response.json()["data"]["attachmentName"], response.json()["data"]["storeName"], response.json()["data"]["attachmentPath"]
    else:
        print(f"Failed to upload attachment: {response.json()}")
        return None

# Function to send email using Zoho Mail API
def send_email_report(access_token, to_email, total_unique, total_new, deleted_count, failed_count):
    
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}", "Content-Type": "application/json", "Accept": "application/json"}
    
    body_text = (f"Total number of unique emails in the file: {total_unique}\n\n"
                 f"Total number of new emails added: {total_new}\n\n"
                 f"Total number of emails deleted: {deleted_count}\n\n"
                 f"Total number of emails failed to delete: {failed_count}")
    if total_new > 0:
        attachment_name, store_name, attachment_path = upload_attachment(access_token)
        if not attachment_name:
            print("Attachment upload failed. Sending email without attachment.")
            attachment_name = ""
        payload = {
            "fromAddress": FROM_EMAIL,
            "toAddress": to_email,
            "subject": SUBJECT,
            "content": body_text,
            "attachments": [{
                "storeName": store_name,
                "attachmentName":attachment_name,
                "attachmentPath": attachment_path
                }
                ]
        }
    else:
        payload = {
            "fromAddress": FROM_EMAIL,
            "toAddress": to_email,
            "subject": SUBJECT,
            "content": body_text
        }
    
    response = requests.post(SEND_MAIL_API_URL, headers=headers, json=payload)
    if response.status_code == 200:
        print(f"Report email sent to {to_email}")
    else:
        print(f"Failed to send email: {response.json()}")

# Function to save new email IDs to CSV
def save_to_csv(sent_to_list):
    existing_emails = load_existing_emails()
    new_emails = set(sent_to_list) - existing_emails  # Only keep new email IDs
    
    if new_emails:
        with open(OUTPUT_CSV, mode="a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            for email in sorted(new_emails):  # Sorting to ensure consistent order
                writer.writerow([email])
        print(f"Added {len(new_emails)} new unique email addresses to {OUTPUT_CSV}")
    else:
        print("No new email addresses found.")
    return len(existing_emails) + len(new_emails), len(new_emails)

# Main script execution
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py recipient_email")
        sys.exit(1)
    
    to_email = sys.argv[1]
    
    token = get_access_token()
    if token:
        print("Access token obtained successfully.")
        sent_to_list, email_ids = fetch_emails(token)
        if sent_to_list:
            print(f"Total unique email addresses fetched: {len(sent_to_list)}")
            total_unique, total_new = save_to_csv(sent_to_list)
            deleted_count, failed_count = delete_emails(token, email_ids)  # Delete emails after processing
            send_email_report(token, to_email, total_unique, total_new, deleted_count, failed_count)
        else:
            print("No emails found.")
    else:
        print("Failed to obtain access token.")
