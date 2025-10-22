# Gmail to Google Sheets - Setup Guide

## How to Set Up This Script

### Step 1: Open Google Sheets
1. Go to Google Sheets (sheets.google.com)
2. Create a new blank spreadsheet
3. Give it a name like "My Gmail Backup"

### Step 2: Open the Script Editor
1. In your spreadsheet, click on "Extensions" in the top menu
2. Click "Apps Script"
3. This opens the script editor in a new tab

### Step 3: Add the Script
1. Delete any code that's already there
2. Copy ALL the code from the "gmail_to_sheets.gs" file
3. Paste it into the script editor
4. Click the save icon (💾) or press Ctrl+S (Cmd+S on Mac)
5. Give your project a name like "Gmail Transfer"

### Step 4: Run the Script
1. In the script editor, select "transferEmailsToSheets" from the dropdown menu at the top
2. Click the "Run" button (▶️ play icon)
3. The first time you run it, Google will ask for permissions:
   - Click "Review Permissions"
   - Choose your Google account
   - Click "Advanced" then "Go to [your project name] (unsafe)"
   - Click "Allow"
4. Go back to your spreadsheet - your emails should now be there!

## What Each Function Does

- **transferEmailsToSheets()** - Gets all emails from your "Job Hunting" label
- **transferUnreadEmails()** - Gets only unread emails (from all folders)
- **transferEmailsFromSender()** - Gets emails from a specific person (you need to edit the email address in the code)

## How to Customize

### Change How Many Emails to Get
Find this line in the code:
```javascript
var maxEmails = 100;
```
Change 100 to any number you want (like 50 or 200)

### Change Which Emails to Get
Find this line in the code:
```javascript
var searchQuery = "label:Job Hunting";
```

You can change it to:
- `"in:inbox"` - all emails in your inbox
- `"is:unread"` - only unread emails
- `"from:someone@example.com"` - emails from a specific person
- `"subject:interview"` - emails with "interview" in the subject
- `"after:2024/01/01"` - emails after January 1, 2024
- `"label:Work"` - emails with a different label
- You can also combine searches like: `"label:Job Hunting is:unread"` for unread job hunting emails

## Tips

- The script gets email "threads" (conversations), so if there are multiple emails in one conversation, they'll all be added
- The body preview is limited to 500 characters to keep the sheet manageable
- You can run the script as many times as you want - it clears the sheet each time
- To update your sheet with new emails, just run the script again

## Troubleshooting

**"Script requires authorization"**
- This is normal the first time. Follow Step 4 above to grant permissions.

**"No emails showing up"**
- Check your search query - make sure it matches emails you actually have
- Try changing `"in:inbox"` to `"in:anywhere"` to search all folders

**"Script takes too long to run"**
- Reduce the number of emails (change maxEmails to a smaller number like 50)
- Gmail has limits on how much you can do at once
