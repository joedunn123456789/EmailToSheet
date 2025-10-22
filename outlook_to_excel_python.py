"""
Python Script: Export Outlook Emails to Excel via IMAP
This script connects to your Outlook account and exports emails to an Excel file

What this script does:
- Connects to Outlook using IMAP (a way to access email)
- Searches for emails in a specific folder
- Extracts important information from each email
- Saves everything to an Excel file

Requirements:
- Python 3.6 or higher
- Libraries: openpyxl, python-dotenv
"""

import imaplib  # Library that connects to email servers
import email  # Library that reads email content
from email.header import decode_header  # Helps decode email headers
from openpyxl import Workbook  # Library to create Excel files
from datetime import datetime  # Helps format dates
import re  # Regular expressions for text processing
import os  # Access operating system features (for environment variables)
from dotenv import load_dotenv  # Load environment variables from .env file

# ============================================
# LOAD CONFIGURATION FROM .ENV FILE
# ============================================

# Load environment variables from .env file
# This reads the .env file and makes the variables available
load_dotenv()

# Get configuration from environment variables
# If a variable isn't set, use a default value
EMAIL = os.getenv("EMAIL", "your-email@outlook.com")
PASSWORD = os.getenv("PASSWORD", "your-app-password-here")
FOLDER = os.getenv("FOLDER", "Job Hunting")
MAX_EMAILS = int(os.getenv("MAX_EMAILS", "100"))
OUTPUT_FILE = os.getenv("OUTPUT_FILE", "outlook_emails.xlsx")

# ============================================
# FUNCTIONS - The code that does the work
# ============================================

def connect_to_outlook():
    """
    Connects to Outlook's IMAP server
    This is like opening a door to your email account
    
    Returns:
        mail: The connection to your email account
    """
    print("Connecting to Outlook...")
    
    # Outlook's IMAP server address
    IMAP_SERVER = "outlook.office365.com"
    IMAP_PORT = 993
    
    try:
        # Create connection to the server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        
        # Log in with your credentials
        mail.login(EMAIL, PASSWORD)
        
        print("✅ Successfully connected to Outlook!")
        return mail
    
    except imaplib.IMAP4.error as e:
        print(f"❌ Login failed: {e}")
        print("\nCommon issues:")
        print("1. Make sure you're using an APP PASSWORD, not your regular password")
        print("2. Check that your email address is correct")
        print("3. Make sure IMAP is enabled in your Outlook settings")
        return None
    except Exception as e:
        print(f"❌ Connection error: {e}")
        return None


def decode_email_subject(subject):
    """
    Decodes the email subject line
    Email subjects can be encoded in weird ways, this fixes that
    
    Args:
        subject: The raw subject from the email
    
    Returns:
        The decoded, readable subject
    """
    if subject is None:
        return "No Subject"
    
    # Decode the subject
    decoded_parts = decode_header(subject)
    decoded_subject = ""
    
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            # If it's bytes, decode it
            decoded_subject += part.decode(encoding or "utf-8", errors="ignore")
        else:
            decoded_subject += part
    
    return decoded_subject


def clean_email_address(address):
    """
    Cleans up email addresses
    Sometimes emails come with names and extra characters, this extracts just the email
    
    Args:
        address: The raw email address string
    
    Returns:
        Clean email address
    """
    if not address:
        return "Unknown"
    
    # Try to extract just the email address using regex
    match = re.search(r'[\w\.-]+@[\w\.-]+', address)
    if match:
        return match.group(0)
    
    return address


def get_email_body(msg):
    """
    Extracts the body text from an email
    Emails can have multiple parts, this finds the text part
    
    Args:
        msg: The email message object
    
    Returns:
        The email body text (limited to 500 characters)
    """
    body = ""
    
    # Check if email has multiple parts
    if msg.is_multipart():
        # Loop through all parts
        for part in msg.walk():
            content_type = part.get_content_type()
            
            # We want plain text
            if content_type == "text/plain":
                try:
                    # Get the text and decode it
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
                except:
                    continue
    else:
        # Single part email
        try:
            body = msg.get_payload(decode=True).decode(errors="ignore")
        except:
            body = "Could not decode email body"
    
    # Limit to 500 characters so Excel doesn't get overwhelmed
    body = body.replace('\r', '').replace('\n', ' ')  # Remove line breaks
    body = body[:500]  # First 500 characters only
    
    return body


def export_emails():
    """
    Main function that exports emails to Excel
    This is the function that does all the work
    """
    print(f"\n📧 Starting email export from folder: {FOLDER}")
    print(f"📊 Will create Excel file: {OUTPUT_FILE}")
    print("-" * 50)
    
    # Connect to Outlook
    mail = connect_to_outlook()
    if not mail:
        return
    
    try:
        # Select the folder (mailbox)
        print(f"\n📁 Opening folder: {FOLDER}")
        status, messages = mail.select(FOLDER)
        
        if status != "OK":
            print(f"❌ Could not open folder '{FOLDER}'")
            print("Available folders:")
            status, folders = mail.list()
            for folder in folders:
                print(f"  - {folder.decode()}")
            return
        
        # Search for all emails in the folder
        print(f"🔍 Searching for emails...")
        status, message_ids = mail.search(None, "ALL")
        
        if status != "OK":
            print("❌ Could not search emails")
            return
        
        # Get list of email IDs
        email_ids = message_ids[0].split()
        total_emails = len(email_ids)
        
        print(f"✅ Found {total_emails} emails")
        
        # Limit to MAX_EMAILS
        if total_emails > MAX_EMAILS:
            print(f"⚠️  Limiting to {MAX_EMAILS} most recent emails")
            email_ids = email_ids[-MAX_EMAILS:]  # Get the most recent ones
        
        # Create Excel workbook
        print(f"\n📝 Creating Excel file...")
        wb = Workbook()
        ws = wb.active
        ws.title = "Email Export"
        
        # Create headers
        headers = ["Date Received", "From", "Subject", "Body Preview", "Folder"]
        ws.append(headers)
        
        # Make headers bold
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        
        # Process each email
        print(f"\n📥 Processing emails...")
        processed = 0
        
        for email_id in email_ids:
            try:
                # Fetch the email
                status, msg_data = mail.fetch(email_id, "(RFC822)")
                
                if status != "OK":
                    continue
                
                # Parse the email
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
                
                # Extract email information
                subject = decode_email_subject(msg.get("Subject"))
                from_addr = clean_email_address(msg.get("From"))
                date_str = msg.get("Date")
                
                # Parse the date
                try:
                    date_tuple = email.utils.parsedate_tz(date_str)
                    if date_tuple:
                        date_obj = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                        date_formatted = date_obj.strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        date_formatted = date_str
                except:
                    date_formatted = date_str or "Unknown Date"
                
                # Get email body
                body = get_email_body(msg)
                
                # Add row to Excel
                row_data = [date_formatted, from_addr, subject, body, FOLDER]
                ws.append(row_data)
                
                processed += 1
                
                # Show progress
                if processed % 10 == 0:
                    print(f"  Processed {processed}/{len(email_ids)} emails...")
                
            except Exception as e:
                print(f"  ⚠️  Error processing email: {e}")
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
        
        print(f"\n✅ Success! Exported {processed} emails to {OUTPUT_FILE}")
        print(f"📊 File saved in the current directory")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
    
    finally:
        # Close connection
        try:
            mail.close()
            mail.logout()
            print("\n🔒 Disconnected from Outlook")
        except:
            pass


# ============================================
# MAIN - This runs when you execute the script
# ============================================

if __name__ == "__main__":
    print("=" * 50)
    print("  OUTLOOK TO EXCEL EXPORTER")
    print("=" * 50)
    
    # Check if .env file exists
    if not os.path.exists('.env'):
        print("\n⚠️  .env FILE NOT FOUND!")
        print("\nPlease create a .env file with your credentials.")
        print("Steps:")
        print("  1. Copy .env.example to .env")
        print("  2. Edit .env and add your email and password")
        print("  3. Run this script again")
        print("\nExample .env file content:")
        print("-" * 50)
        print("EMAIL=yourname@outlook.com")
        print("PASSWORD=your-app-password")
        print("FOLDER=Job Hunting")
        print("MAX_EMAILS=100")
        print("OUTPUT_FILE=outlook_emails.xlsx")
        print("-" * 50)
    elif EMAIL == "your-email@outlook.com" or PASSWORD == "your-app-password-here":
        print("\n⚠️  CONFIGURATION NEEDED!")
        print("Please edit your .env file and set your:")
        print("  1. EMAIL address")
        print("  2. PASSWORD (app password, not regular password)")
        print("  3. FOLDER name (if different from 'Job Hunting')")
        print("\nSee the setup guide for instructions!")
    else:
        # Run the export
        export_emails()
    
    print("\n" + "=" * 50)
    print("Script finished!")
    print("=" * 50)
