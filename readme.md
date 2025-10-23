# Email to Excel Exporters

A collection of scripts to export emails from Gmail and Outlook to Excel/Google Sheets for easy organization and analysis.

## 📋 Overview

This repository contains two main solutions:

1. **Gmail to Google Sheets** - Google Apps Script that transfers Gmail emails to a Google Sheet
2. **Outlook to Excel** - Python script that exports Outlook emails to an Excel file using OAuth2

Perfect for organizing job applications, tracking correspondence, archiving emails, or any situation where you need your emails in a spreadsheet format!

> **⚠️ Important Update (2024):** Microsoft has disabled Basic Authentication (app passwords + IMAP) for most personal Outlook accounts. This repository uses OAuth2 Modern Authentication ([outlook_to_excel.py](outlook_to_excel.py)) that works with current Microsoft security requirements. See [OAUTH_SETUP_PERSONAL.md](OAUTH_SETUP_PERSONAL.md) for setup instructions.

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
- ✅ Uses Modern Authentication (OAuth2)
- ⚠️ **Note:** IMAP/app password method no longer works for most accounts

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
- Captures: Date, From, Subject, Body Preview, Status (Read/Unread)
- Works with personal Outlook accounts
- Exports to Excel (.xlsx) format
- Uses Modern Authentication (OAuth2)
- Can be scheduled to run automatically

### Requirements
- Python 3.6 or higher
- Required libraries: `openpyxl`, `msal`, `requests`, `python-dotenv`
- Azure app registration (free, one-time setup)
- Personal Microsoft account (@outlook.com, @hotmail.com, @live.com)

### Installation

1. **Install Python**
   - Download from [python.org](https://www.python.org/downloads/)
   - **Important:** Check "Add Python to PATH" during installation

2. **Install Required Libraries**
   ```bash
   pip install -r requirements.txt
   ```
   Or manually:
   ```bash
   pip install openpyxl msal requests python-dotenv
   ```

3. **Register Azure Application** (One-time, FREE setup)
   - Follow the detailed guide: [OAUTH_SETUP_PERSONAL.md](OAUTH_SETUP_PERSONAL.md)
   - Takes about 5 minutes
   - Completely free for personal use
   - You'll get a Client ID to use in your configuration

### Setup Instructions

1. **Download the Script**
   - Download [`outlook_to_excel.py`](outlook_to_excel.py)

2. **Configure Your Credentials**
   - Copy `.env.example` to `.env`
   - Add your Azure app Client ID (from registration step above)
   - Configure your preferences:
   ```bash
   CLIENT_ID=your-azure-app-client-id
   TENANT_ID=consumers
   FOLDER=inbox
   MAX_EMAILS=100
   OUTPUT_FILE=outlook_emails.xlsx
   ```

3. **Run the Script**
   ```bash
   python3 outlook_to_excel.py
   ```

4. **Authenticate**
   - The script will display a code and open your browser
   - Go to https://microsoft.com/devicelogin
   - Enter the code
   - Sign in with your personal Microsoft account
   - Grant permission to read your mail

5. **Check the Output**
   - Look for `outlook_emails.xlsx` in the same folder
   - Open it in Excel!


### Customization

Edit your `.env` file to customize settings:

**Change folder:**
```bash
FOLDER=inbox  # Main inbox (lowercase for OAuth2)
FOLDER=sentitems  # Sent emails
```

**Export more emails:**
```bash
MAX_EMAILS=500  # Export 500 emails
```

**Change output file:**
```bash
OUTPUT_FILE=my_emails.xlsx
```

---

## 📁 Repository Structure

```
.
├── README.md                          # This file
├── gmail_to_sheets.gs                 # Gmail Google Apps Script
├── SETUP_GUIDE.md                     # Gmail detailed setup guide
├── outlook_to_excel.py                # Outlook OAuth2 script (uses Modern Authentication)
├── OAUTH_SETUP_PERSONAL.md            # OAuth2 setup guide for personal accounts
├── requirements.txt                   # Python dependencies
├── .env.example                       # Configuration template
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

### Outlook OAuth2 Issues

**"Work or school account only" error**
- Make sure you selected "Personal Microsoft accounts only" when registering your Azure app
- Check that TENANT_ID is set to "consumers" in your .env file

**"Invalid client" error**
- Double-check your CLIENT_ID in the .env file
- Make sure you copied the entire Application (client) ID from Azure

**"Permission denied" or "Consent required"**
- Make sure you added Mail.Read permission in Azure Portal
- Try revoking and re-granting consent

**"Module not found" errors**
- Run: `pip install -r requirements.txt`
- Or manually: `pip install openpyxl msal requests python-dotenv`

**Authentication keeps asking for login**
- The token cache may be corrupted
- Look for hidden cache files and delete them


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

### Outlook OAuth2 Script
- Runs locally on your computer
- Uses OAuth2 Modern Authentication (no passwords stored)
- No data sent to external services (besides Microsoft for authentication)
- Read-only access (won't modify or delete emails)
- Your Azure app is private to you only
- Authentication tokens cached locally for convenience

**Important:**
- Never share your CLIENT_ID publicly (keep .env file private)
- Never commit .env files to GitHub (already in .gitignore)
- Your Azure app registration is free and only accessible by you
- You can revoke access anytime from your Microsoft account settings

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
- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/api/overview)
- [Azure App Registration Guide](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [MSAL Python Documentation](https://msal-python.readthedocs.io/)
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
A: The Gmail script works with Gmail. The Outlook OAuth2 script works with personal Microsoft accounts (@outlook.com, @hotmail.com, @live.com). The legacy IMAP script may work with other providers, but Microsoft has deprecated this method.

**Q: Will this delete my emails?**
A: No! Both scripts are read-only. They only read and export emails.

---

## 📧 Contact

Found a bug? Have a question? [Open an issue](../../issues)!

---

Made with ❤️ for anyone tired of manual email organization!
