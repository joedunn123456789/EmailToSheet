"""
Python Script: Export Outlook Emails to Excel via Microsoft Graph API (OAuth2)
This script uses Modern Authentication instead of app passwords

What this script does:
- Uses OAuth2 to authenticate with Microsoft
- Connects to Outlook via Microsoft Graph API
- Exports emails to an Excel file

Requirements:
- Python 3.6 or higher
- Libraries: openpyxl, python-dotenv, msal, requests
"""

import os
import json
from datetime import datetime
from openpyxl import Workbook
from dotenv import load_dotenv
import msal
import requests
import webbrowser

# Load environment variables
load_dotenv()

# Microsoft Graph API Configuration
CLIENT_ID = os.getenv("CLIENT_ID", "your-client-id-here")
TENANT_ID = os.getenv("TENANT_ID", "common")  # "common" works for personal accounts
FOLDER = os.getenv("FOLDER", "inbox")
MAX_EMAILS = int(os.getenv("MAX_EMAILS", "100"))
OUTPUT_FILE = os.getenv("OUTPUT_FILE", "outlook_emails.xlsx")

# Microsoft Graph API endpoints
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/Mail.Read"]
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"


def get_access_token():
    """
    Get access token using device code flow (user-friendly for personal use)
    """
    print("\n🔐 Authenticating with Microsoft...")
    print("-" * 50)

    # Create a public client application
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
    )

    # Try to get token from cache first
    accounts = app.get_accounts()
    if accounts:
        print("Found cached credentials, attempting to use them...")
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result:
            print("✅ Using cached authentication")
            return result["access_token"]

    # If no cache, use device code flow
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        raise ValueError("Failed to create device flow")

    print("\n" + "=" * 60)
    print("AUTHENTICATION REQUIRED")
    print("=" * 60)
    print(flow["message"])
    print("\nOpening browser for authentication...")
    print("=" * 60)

    # Try to open browser automatically
    try:
        webbrowser.open(flow["verification_uri"])
    except:
        pass

    # Wait for user to authenticate
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        print("\n✅ Authentication successful!")
        return result["access_token"]
    else:
        print(f"\n❌ Authentication failed: {result.get('error_description')}")
        return None


def get_folder_id(access_token, folder_name):
    """
    Get the folder ID for a given folder name
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Get all mail folders
    url = f"{GRAPH_API_ENDPOINT}/me/mailFolders"
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"Error getting folders: {response.status_code}")
        return None

    folders = response.json().get("value", [])

    # Find the folder
    folder_name_lower = folder_name.lower()
    for folder in folders:
        if folder["displayName"].lower() == folder_name_lower:
            return folder["id"]

    # If not found, return inbox
    print(f"⚠️  Folder '{folder_name}' not found, using inbox")
    for folder in folders:
        if folder["displayName"].lower() == "inbox":
            return folder["id"]

    return None


def get_emails(access_token, folder_id, max_count):
    """
    Get emails from the specified folder
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Get messages
    url = f"{GRAPH_API_ENDPOINT}/me/mailFolders/{folder_id}/messages"
    params = {
        "$top": min(max_count, 100),  # Max 100 per request
        "$select": "receivedDateTime,from,subject,bodyPreview,isRead",
        "$orderby": "receivedDateTime DESC"
    }

    all_emails = []

    while url and len(all_emails) < max_count:
        response = requests.get(url, headers=headers, params=params)

        if response.status_code != 200:
            print(f"Error getting emails: {response.status_code}")
            print(response.text)
            break

        data = response.json()
        emails = data.get("value", [])
        all_emails.extend(emails)

        # Get next page URL
        url = data.get("@odata.nextLink")
        params = None  # params are in the nextLink URL

        print(f"  Retrieved {len(all_emails)} emails...")

    return all_emails[:max_count]


def export_emails():
    """
    Main function to export emails to Excel
    """
    print("=" * 60)
    print("  OUTLOOK TO EXCEL EXPORTER (OAuth2)")
    print("=" * 60)

    # Check configuration
    if CLIENT_ID == "your-client-id-here":
        print("\n❌ CLIENT_ID not configured!")
        print("\nFor personal Microsoft accounts, you can use this public client ID:")
        print("CLIENT_ID=d3590ed6-52b3-4102-aeff-aad2292ab01c")
        print("\nAdd this to your .env file and try again.")
        return

    # Get access token
    access_token = get_access_token()
    if not access_token:
        return

    print(f"\n📧 Exporting emails from folder: {FOLDER}")
    print(f"📊 Will create Excel file: {OUTPUT_FILE}")
    print("-" * 50)

    # Get folder ID
    print(f"\n📁 Finding folder: {FOLDER}")
    folder_id = get_folder_id(access_token, FOLDER)
    if not folder_id:
        print("❌ Could not find folder")
        return

    print(f"✅ Found folder")

    # Get emails
    print(f"\n📥 Retrieving up to {MAX_EMAILS} emails...")
    emails = get_emails(access_token, folder_id, MAX_EMAILS)

    if not emails:
        print("❌ No emails found")
        return

    print(f"✅ Retrieved {len(emails)} emails")

    # Create Excel workbook
    print(f"\n📝 Creating Excel file...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Email Export"

    # Create headers
    headers = ["Date Received", "From", "Subject", "Body Preview", "Status"]
    ws.append(headers)

    # Make headers bold
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)

    # Process each email
    print(f"\n📥 Processing emails...")
    for idx, email in enumerate(emails, 1):
        try:
            # Extract email information
            received_date = email.get("receivedDateTime", "")
            if received_date:
                # Parse and format date
                date_obj = datetime.fromisoformat(received_date.replace("Z", "+00:00"))
                date_formatted = date_obj.strftime("%Y-%m-%d %H:%M:%S")
            else:
                date_formatted = "Unknown Date"

            from_data = email.get("from", {})
            from_addr = from_data.get("emailAddress", {}).get("address", "Unknown")

            subject = email.get("subject", "No Subject")
            body_preview = email.get("bodyPreview", "")[:500]  # Limit to 500 chars
            status = "Read" if email.get("isRead") else "Unread"

            # Add row to Excel
            row_data = [date_formatted, from_addr, subject, body_preview, status]
            ws.append(row_data)

            if idx % 10 == 0:
                print(f"  Processed {idx}/{len(emails)} emails...")

        except Exception as e:
            print(f"  ⚠️  Error processing email {idx}: {e}")
            continue

    # Auto-resize columns
    for column in ws.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = min(max_length + 2, 50)  # Max width of 50
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

    # Save the Excel file
    wb.save(OUTPUT_FILE)

    print(f"\n✅ Success! Exported {len(emails)} emails to {OUTPUT_FILE}")
    print(f"📊 File saved in the current directory")


if __name__ == "__main__":
    export_emails()
    print("\n" + "=" * 60)
    print("Script finished!")
    print("=" * 60)
