# Python Outlook to Excel - Complete Setup Guide

## What You Need
- ✅ A computer (Windows, Mac, or Linux)
- ✅ Python installed (we'll show you how)
- ✅ Your Outlook email address
- ✅ An app password (we'll show you how to get this)

---

## STEP 1: Install Python

### For Windows:
1. Go to https://www.python.org/downloads/
2. Click "Download Python" (get the latest version)
3. Run the installer
4. **IMPORTANT**: Check the box "Add Python to PATH"
5. Click "Install Now"
6. Wait for installation to complete

### For Mac:
1. Open Terminal (press Cmd+Space, type "terminal", press Enter)
2. Type this command and press Enter:
   ```bash
   brew install python3
   ```
3. If you don't have Homebrew, first install it from https://brew.sh

### For Linux:
Python is usually pre-installed. If not:
```bash
sudo apt-get update
sudo apt-get install python3 python3-pip
```

### Verify Python is Installed:
1. Open Command Prompt (Windows) or Terminal (Mac/Linux)
2. Type: `python --version` or `python3 --version`
3. You should see something like "Python 3.11.x"

✅ **Python installed!**

---

## STEP 2: Install Required Python Libraries

1. Open Command Prompt (Windows) or Terminal (Mac/Linux)
2. Type this command and press Enter:
   ```bash
   pip install openpyxl
   ```
3. Wait for it to finish (should take 10-30 seconds)

✅ **Libraries installed!**

---

## STEP 3: Get Your Outlook App Password

**CRITICAL**: You need an **App Password**, NOT your regular Outlook password!

### What is an App Password?
It's a special password that lets programs (like this script) access your email safely, without using your actual password.

### How to Get Your App Password:

1. **Go to Microsoft Security Settings:**
   - Visit: https://account.microsoft.com/security
   - Sign in with your Outlook account

2. **Enable Two-Step Verification (if not already on):**
   - Click "Advanced security options"
   - Find "Two-step verification"
   - Click "Turn on" if it's not already on
   - Follow the prompts to set it up (you'll need your phone)

3. **Create an App Password:**
   - After two-step verification is on, scroll down
   - Find "App passwords"
   - Click "Create a new app password"
   - A password will appear (like: `abcd-efgh-ijkl-mnop`)
   - **COPY THIS PASSWORD IMMEDIATELY** (you won't see it again!)
   - Save it somewhere safe (like a password manager)

✅ **App password created!**

---

## STEP 4: Enable IMAP in Outlook

IMAP is what lets programs connect to your email.

1. Go to Outlook on the web: https://outlook.live.com
2. Click the Settings gear icon (⚙️) in the top right
3. Click "View all Outlook settings" at the bottom
4. Go to "Mail" → "Sync email"
5. Under "POP and IMAP", make sure IMAP is **ON**
6. If it says "Let devices and apps use POP", turn it ON
7. Click "Save"

✅ **IMAP enabled!**

---

## STEP 5: Set Up the Python Script

1. **Download the script file:**
   - Save the `outlook_to_excel_python.py` file to your computer
   - Save it somewhere easy to find (like your Desktop or Documents folder)

2. **Open the script in a text editor:**
   - Right-click the file
   - Choose "Open with" → "Notepad" (Windows) or "TextEdit" (Mac)

3. **Edit the configuration at the top:**
   
   Find these lines near the top:
   ```python
   EMAIL = "your-email@outlook.com"
   PASSWORD = "your-app-password-here"
   FOLDER = "Job Hunting"
   MAX_EMAILS = 100
   OUTPUT_FILE = "outlook_emails.xlsx"
   ```

   Change them to your info:
   ```python
   EMAIL = "yourname@outlook.com"  # Your actual email
   PASSWORD = "abcd-efgh-ijkl-mnop"  # The app password you created
   FOLDER = "Job Hunting"  # Your folder name (or "INBOX" for inbox)
   MAX_EMAILS = 100  # How many emails to export
   OUTPUT_FILE = "outlook_emails.xlsx"  # Name of Excel file to create
   ```

4. **Save the file** (Ctrl+S or Cmd+S)

✅ **Script configured!**

---

## STEP 6: Run the Script

### On Windows:
1. Open Command Prompt
2. Navigate to where you saved the script:
   ```bash
   cd Desktop
   ```
   (or wherever you saved it)
3. Run the script:
   ```bash
   python outlook_to_excel_python.py
   ```

### On Mac/Linux:
1. Open Terminal
2. Navigate to where you saved the script:
   ```bash
   cd ~/Desktop
   ```
3. Run the script:
   ```bash
   python3 outlook_to_excel_python.py
   ```

### What You'll See:
```
==================================================
  OUTLOOK TO EXCEL EXPORTER
==================================================

📧 Starting email export from folder: Job Hunting
📊 Will create Excel file: outlook_emails.xlsx
--------------------------------------------------

Connecting to Outlook...
✅ Successfully connected to Outlook!

📁 Opening folder: Job Hunting
🔍 Searching for emails...
✅ Found 47 emails

📝 Creating Excel file...

📥 Processing emails...
  Processed 10/47 emails...
  Processed 20/47 emails...
  Processed 30/47 emails...
  Processed 40/47 emails...

✅ Success! Exported 47 emails to outlook_emails.xlsx
📊 File saved in the current directory

🔒 Disconnected from Outlook

==================================================
Script finished!
==================================================
```

4. **Check your folder** - you should see a file called `outlook_emails.xlsx`
5. **Open it in Excel** - your emails are there!

🎉 **Done!**

---

## Customization

### Change Which Folder to Export:
In the script, change:
```python
FOLDER = "Job Hunting"
```
to:
```python
FOLDER = "INBOX"  # For your main inbox
# or
FOLDER = "Sent Items"  # For sent emails
# or
FOLDER = "Any/Folder/Name"  # For any other folder
```

### Export More Emails:
Change:
```python
MAX_EMAILS = 100
```
to whatever number you want (like `MAX_EMAILS = 500`)

### Change Output File Name:
Change:
```python
OUTPUT_FILE = "outlook_emails.xlsx"
```
to whatever you want (like `OUTPUT_FILE = "job_applications.xlsx"`)

---

## Troubleshooting

### "Login failed" Error:
- ❌ Make sure you're using the **APP PASSWORD**, not your regular password
- ❌ Check that your email address is spelled correctly
- ❌ Make sure two-step verification is enabled
- ❌ Make sure IMAP is enabled in Outlook settings

### "Could not open folder" Error:
- The folder name might be wrong
- The script will show you available folders
- Folder names are case-sensitive!
- Use quotes if the folder has spaces: `FOLDER = "Job Hunting"`

### "Module not found" Error:
- You need to install openpyxl: `pip install openpyxl`

### "Python not found" Error:
- Python isn't installed correctly
- Try `python3` instead of `python`
- Make sure you checked "Add Python to PATH" during installation

### Folder Names Don't Match:
Different Outlook versions use different folder names:
- Try "INBOX" instead of "Inbox"
- Try "Sent Items" instead of "Sent"
- The script will show available folder names if it can't find yours

### Script is Slow:
- Processing lots of emails takes time (about 1-2 seconds per email)
- Reduce MAX_EMAILS to process fewer at a time
- Large emails with attachments take longer

---

## Running the Script Regularly

### Option 1: Manual (easiest)
Just run the script whenever you want updated emails.

### Option 2: Schedule It (Windows)
1. Open Task Scheduler
2. Create a new task
3. Set it to run: `python C:\path\to\outlook_to_excel_python.py`
4. Set schedule (daily, weekly, etc.)

### Option 3: Schedule It (Mac/Linux)
1. Open Terminal
2. Type: `crontab -e`
3. Add a line like: `0 9 * * * python3 /path/to/outlook_to_excel_python.py`
   (This runs it every day at 9 AM)

---

## Explain It Like You're 5 👶

**What is Python?**
Python is like a magic recipe book for your computer. You give it instructions (a recipe), and it follows them!

**What does this script do?**
Imagine you have a friend who:
1. Knocks on your email box's door (connects to Outlook)
2. Shows their special pass (app password) to get in
3. Walks to your "Job Hunting" drawer
4. Takes out all the letters (emails)
5. Reads each one and writes down:
   - When did it arrive?
   - Who sent it?
   - What does it say?
6. Writes everything in a neat table in Excel
7. Gives you the Excel file

**Why do I need an "app password"?**
It's like giving your friend a special key that only opens the mail drawer, not your whole house! It's safer than giving them your real password.

**What is IMAP?**
IMAP is like the mail slot in your door that lets people check your mail without having to break in!

---

## Advantages of This Method

**Compared to Power Automate:**
- ✅ Works with personal Outlook accounts
- ✅ No cloud service needed
- ✅ More control over what you export
- ✅ Can process unlimited emails

**Compared to Manual Export:**
- ✅ Much faster (automatic)
- ✅ Can be scheduled to run regularly
- ✅ Customizable

**Compared to Gmail Forwarding:**
- ✅ Keeps emails in Outlook
- ✅ Don't need to use Gmail
- ✅ More private

---

## Security Notes 🔒

- Your app password is stored in the script file - keep it safe!
- Don't share the script file with your password in it
- The script only reads your emails, it doesn't delete or modify anything
- Your emails never go to any cloud service or website
- Everything happens on your computer only

---

## Need Help?

If you run into issues:
1. Read the error message carefully
2. Check the Troubleshooting section above
3. Make sure Python and openpyxl are installed
4. Verify your app password is correct
5. Check that IMAP is enabled
6. Try with FOLDER = "INBOX" first to test

Good luck! 🚀
