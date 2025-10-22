# Email to Excel Exporters

A collection of scripts to export emails from Gmail and Outlook to Excel/Google Sheets for easy organization and analysis.

## 📋 Overview

This repository contains two main solutions:

1. **Gmail to Google Sheets** - Google Apps Script that transfers Gmail emails to a Google Sheet
2. **Outlook to Excel** - Python script that exports Outlook emails to an Excel file via IMAP

Perfect for organizing job applications, tracking correspondence, archiving emails, or any situation where you need your emails in a spreadsheet format!

## 🚀 Quick Start

### Gmail Solution
- ✅ Works with personal Gmail accounts
- ✅ Runs directly in Google Sheets
- ✅ No installation required
- ✅ Can run automatically or on-demand

[Jump to Gmail Setup →](#gmail-to-google-sheets)

### Outlook Solution
- ✅ Works with personal Outlook accounts
- ✅ Runs on your computer
- ✅ Exports to Excel (.xlsx)
- ✅ Connects via IMAP

[Jump to Outlook Setup →](#outlook-to-excel-python)

---

## Gmail to Google Sheets

### Features
- Export emails from any Gmail label/folder
- Captures: Date, From, Subject, Body Preview, Labels
- Customizable search queries
- Multiple export options (all emails, unread only, specific sender)

### Setup Instructions

1. **Open Google Sheets**
   - Go to [sheets.google.com](https://sheets.google.com)
   - Create a new blank spreadsheet

2. **Open Apps Script Editor**
   - Click `Extensions` → `Apps Script`

3. **Add the Script**
   - Delete any existing code
   - Copy the code from [`gmail_to_sheets.gs`](gmail_to_sheets.gs)
   - Paste it into the editor
   - Click Save (💾)

4. **Configure the Script**
   - Find line 23: `var searchQuery = "label:job-hunting";`
   - Change `"label:job-hunting"` to your label name
   - Change line 26 if you want more/fewer emails: `var maxEmails = 100;`

5. **Run It**
   - Select `transferEmailsToSheets` from the dropdown
   - Click Run (▶️)
   - Authorize the script when prompted
   - Check your spreadsheet - your emails are there!

### Customization

**Change the label:**
```javascript
var searchQuery = "label:your-label-name";
```

**Change number of emails:**
```javascript
var maxEmails = 200;  // Export 200 emails
```

**Search examples:**
```javascript
var searchQuery = "is:unread";  // Only unread emails
var searchQuery = "from:someone@example.com";  // From specific sender
var searchQuery = "after:2024/01/01";  // After a specific date
var searchQuery = "subject:interview";  // Emails with "interview" in subject
```

---

## Outlook to Excel (Python)

### Features
- Export emails from any Outlook folder
- Captures: Date, From, Subject, Body Preview, Folder
- Works with personal Outlook accounts
- Exports to Excel (.xlsx) format
- Can be scheduled to run automatically

### Requirements
- Python 3.6 or higher
- `openpyxl` library
- Outlook account with IMAP enabled
- Microsoft app password

### Installation

1. **Install Python**
   - Download from [python.org](https://www.python.org/downloads/)
   - **Important:** Check "Add Python to PATH" during installation

2. **Install Required Library**
   ```bash
   pip install openpyxl
   ```

3. **Get Your Outlook App Password**
   - Go to [Microsoft Security](https://account.microsoft.com/security)
   - Enable Two-Step Verification
   - Create an App Password
   - Save the password securely

4. **Enable IMAP in Outlook**
   - Go to [Outlook Settings](https://outlook.live.com)
   - Settings (⚙️) → View all Outlook settings
   - Mail → Sync email → Enable IMAP

### Setup Instructions

1. **Download the Script**
   - Download [`outlook_to_excel_python.py`](outlook_to_excel_python.py)

2. **Configure the Script**
   - Open the file in a text editor
   - Find the configuration section at the top:
   ```python
   EMAIL = "your-email@outlook.com"  # Your email
   PASSWORD = "your-app-password-here"  # Your app password
   FOLDER = "Job Hunting"  # Folder to export from
   MAX_EMAILS = 100  # Number of emails to export
   OUTPUT_FILE = "outlook_emails.xlsx"  # Output filename
   ```
   - Update with your information
   - Save the file

3. **Run the Script**
   ```bash
   python outlook_to_excel_python.py
   ```
   Or on Mac/Linux:
   ```bash
   python3 outlook_to_excel_python.py
   ```

4. **Check the Output**
   - Look for `outlook_emails.xlsx` in the same folder
   - Open it in Excel!

### Customization

**Change folder:**
```python
FOLDER = "INBOX"  # Main inbox
FOLDER = "Sent Items"  # Sent emails
```

**Export more emails:**
```python
MAX_EMAILS = 500  # Export 500 emails
```

**Change output file:**
```python
OUTPUT_FILE = "my_emails.xlsx"
```

---

## 📁 Repository Structure

```
.
├── README.md                          # This file
├── gmail_to_sheets.gs                 # Gmail Google Apps Script
├── SETUP_GUIDE.md                     # Gmail detailed setup guide
├── outlook_to_excel_python.py         # Outlook Python script
├── PYTHON_SETUP_GUIDE.md              # Outlook detailed setup guide
├── OFFICE_SCRIPT_COMPLETE_GUIDE.md    # Alternative Office Scripts method
├── outlook_office_script.ts           # Office Scripts version (requires Microsoft 365 Business)
└── POWER_AUTOMATE_SIMPLE.md           # Power Automate guide (requires Microsoft 365 Business)
```

---

## 🔧 Troubleshooting

### Gmail Issues

**"Script requires authorization"**
- This is normal on first run
- Click "Review Permissions" → Choose your account → Allow

**"No emails found"**
- Check your label name (use hyphens for spaces: `job-hunting` not `Job Hunting`)
- Try `"in:inbox"` to test with all inbox emails
- Make sure you have emails in that label

### Outlook Issues

**"Login failed"**
- Make sure you're using an **app password**, not your regular password
- Verify IMAP is enabled in Outlook settings
- Check your email address is correct

**"Could not open folder"**
- Folder names are case-sensitive
- Try `"INBOX"` for main inbox
- Script will list available folders if it can't find yours

**"Module not found: openpyxl"**
- Run: `pip install openpyxl`

---

## 🎯 Use Cases

- **Job Hunting**: Track applications, responses, and interviews
- **Email Archiving**: Keep a backup of important emails
- **Email Analysis**: Analyze email patterns, response times, senders
- **Data Migration**: Move emails between services
- **Compliance**: Keep records of business correspondence
- **Organization**: Better search and filter capabilities in Excel

---

## 🔒 Security & Privacy

### Gmail Script
- Runs entirely within your Google account
- Only you can run it
- No data leaves Google's servers
- You can revoke access anytime in Google Account settings

### Outlook Script
- Runs locally on your computer
- Uses app password (not your real password)
- No data sent to external services
- Read-only access (won't modify or delete emails)
- Store your app password securely (consider a password manager)

**Important:** 
- Never share your app password
- Never commit passwords to GitHub
- Use environment variables for sensitive data in production

---

## 📝 License

MIT License - Feel free to use, modify, and distribute these scripts!

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

Some ideas for contributions:
- Add support for other email providers
- Add email attachment handling
- Add filtering options
- Improve error handling
- Add progress indicators
- Create a GUI version

---

## ⭐ Show Your Support

If these scripts helped you, please consider:
- ⭐ Starring this repository
- 🐛 Reporting issues
- 💡 Suggesting new features
- 🔀 Submitting pull requests

---

## 📚 Additional Resources

- [Gmail Search Operators](https://support.google.com/mail/answer/7190)
- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Python IMAP Documentation](https://docs.python.org/3/library/imaplib.html)
- [Openpyxl Documentation](https://openpyxl.readthedocs.io/)

---

## 🙋 FAQ

**Q: Can I export emails with attachments?**
A: Currently, only email metadata and body text are exported. Attachment support could be added as a future feature.

**Q: How many emails can I export at once?**
A: Gmail script: Limited by execution time (typically 1000-2000 emails). Outlook script: No hard limit, but processing 1000+ emails may take several minutes.

**Q: Can I schedule these to run automatically?**
A: 
- Gmail: Use Google Apps Script triggers (time-driven)
- Outlook: Use Task Scheduler (Windows) or cron (Mac/Linux)

**Q: Does this work with other email providers?**
A: The Outlook/Python script works with any email provider that supports IMAP (Gmail, Yahoo, etc.). Just change the IMAP server address.

**Q: Will this delete my emails?**
A: No! Both scripts are read-only. They only read and export emails.

---

## 📧 Contact

Found a bug? Have a question? [Open an issue](../../issues)!

---

Made with ❤️ for anyone tired of manual email organization!
