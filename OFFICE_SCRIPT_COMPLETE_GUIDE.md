# Complete Guide: Outlook to Excel with Office Scripts + Power Automate

## What You Need
- ✅ Microsoft 365 subscription (Business, Education, or Family plan)
- ✅ Access to Excel for the web (online version)
- ✅ Access to Power Automate (comes with Microsoft 365)
- ✅ Work/School Outlook account (NOT personal Outlook.com)

---

## PART 1: Set Up the Office Script in Excel

### Step 1: Create Your Excel File
1. Go to **Excel for the web**: https://office.com/launch/excel
2. Click **"New blank workbook"**
3. Name it: **"Job Hunting Emails"** (or whatever you want)
4. The file will auto-save to your OneDrive

### Step 2: Open the Automate Tab
1. In your Excel file, click the **"Automate"** tab at the top of the screen
2. Click **"New Script"**
3. A code editor will open on the right side

### Step 3: Add the Script Code
1. **Delete** all the code that's already in the editor
2. Copy **ALL** the code from the `outlook_office_script.ts` file
3. Paste it into the script editor
4. Click the **"Save"** button (💾 disk icon)
5. The script will auto-name itself "main" - you can rename it to "TransferEmails" if you want
6. Close the script editor

**✅ Part 1 Complete!** Your Excel file is ready to receive emails.

---

## PART 2: Set Up Power Automate (Automatic - New Emails)

This flow will automatically add new emails to Excel as they arrive.

### Step 1: Go to Power Automate
1. Open: https://make.powerautomate.com
2. Sign in with your Microsoft 365 account

### Step 2: Create a New Automated Flow
1. Click **"+ Create"** in the left sidebar
2. Select **"Automated cloud flow"**
3. In the flow name box, type: **"Auto-add Outlook Emails"**
4. In the search box below, type: **"When a new email arrives"**
5. Select **"When a new email arrives (V3)"** from Office 365 Outlook
6. Click **"Create"**

### Step 3: Configure the Email Trigger
You'll see a box titled "When a new email arrives (V3)". Configure it:

1. **Folder**: 
   - Click the folder icon (📁)
   - Navigate to and select your **"Job Hunting"** folder
   
2. **Include Attachments**: No

3. Leave everything else as default

### Step 4: Add the Office Script Action
1. Click **"+ New step"**
2. In the search box, type: **"Run script"**
3. Select **"Run script"** from Excel Online (Business)

### Step 5: Configure the Run Script Action
You'll see several fields to fill in:

1. **Location**: Select **"OneDrive for Business"**

2. **Document Library**: Select **"OneDrive"**

3. **File**: 
   - Click the folder icon
   - Browse to find your **"Job Hunting Emails.xlsx"** file
   - Select it

4. **Script**: Select your script (it will be called "main" or "TransferEmails")

Now you'll see parameters for the script:

5. **emailDate**: 
   - Click in the field
   - From the dynamic content panel, select **"Received Time"**

6. **emailFrom**: 
   - Click in the field
   - Select **"From"**

7. **emailSubject**: 
   - Click in the field
   - Select **"Subject"**

8. **emailBody**: 
   - Click in the field
   - Select **"Body Preview"**

9. **emailFolder**: 
   - Click in the field
   - Just type: **"Job Hunting"** (or your folder name)

### Step 6: Save and Test
1. Click **"Save"** at the top right
2. Click **"Test"** → **"Manually"** → **"Test"**
3. Send yourself a test email to your Job Hunting folder
4. Wait 1-2 minutes
5. Go check your Excel file - the email should appear!

**🎉 Automatic flow complete!** New emails will now be added automatically.

---

## PART 3: Get All Existing Emails (Manual Flow)

This flow lets you grab all your existing emails with one click.

### Step 1: Create a New Manual Flow
1. In Power Automate, click **"+ Create"**
2. Select **"Instant cloud flow"**
3. Name it: **"Get All Job Hunting Emails"**
4. Select **"Manually trigger a flow"**
5. Click **"Create"**

### Step 2: List Emails from Folder
1. Click **"+ New step"**
2. Search for: **"List emails"**
3. Select **"List emails (V3)"** from Office 365 Outlook
4. Configure:
   - **Folder**: Select your **"Job Hunting"** folder
   - **Top**: **100** (number of emails to get - change if needed)
   - **Include Attachments**: No

### Step 3: Add a Loop
1. Click **"+ New step"**
2. Search for: **"Apply to each"**
3. Select **"Apply to each"** (Control)
4. In "Select an output from previous steps":
   - Click in the field
   - Select **"value"** from the List emails dynamic content

### Step 4: Inside the Loop - Add Script Action
Inside the "Apply to each" box:

1. Click **"Add an action"**
2. Search for: **"Run script"**
3. Select **"Run script"** from Excel Online (Business)
4. Configure:
   - **Location**: OneDrive for Business
   - **Document Library**: OneDrive
   - **File**: Your Excel file
   - **Script**: Your script

5. Fill in the parameters (make sure to select from "Apply to each" dynamic content):
   - **emailDate**: **"Received Time"**
   - **emailFrom**: **"From"**
   - **emailSubject**: **"Subject"**
   - **emailBody**: **"Body Preview"**
   - **emailFolder**: Type **"Job Hunting"**

### Step 5: Run the Flow
1. Click **"Save"**
2. Click **"Test"** → **"Manually"** → **"Run flow"**
3. Wait for it to complete (may take a minute or two)
4. Check your Excel file - all emails should be there!

**🎉 Manual flow complete!** You can run this anytime to grab all emails.

---

## Customization Tips

### Change the Folder Name
In the Power Automate trigger, just select a different folder when configuring.

### Get More Than 100 Emails
In the "List emails" action, change the **Top** value from 100 to a higher number (max is usually 250).

### Get Only Unread Emails
In the "List emails" action:
1. Click **"Show advanced options"**
2. In **Filter Query**, type: `isRead eq false`

### Get Emails from a Specific Date Range
In the "List emails" action's **Filter Query**, use:
- `receivedDateTime ge 2024-01-01` (emails after Jan 1, 2024)
- `receivedDateTime ge 2024-01-01 and receivedDateTime le 2024-12-31` (emails in 2024)

### Add More Columns to Excel
1. Edit the Office Script to add more parameters
2. Update the `rowData` array in the script
3. Add the corresponding parameter in Power Automate

---

## Troubleshooting

### "I don't see the Automate tab in Excel"
- Make sure you're using Excel **for the web** (not desktop Excel)
- Check that you have a Microsoft 365 subscription
- Try a different browser (Chrome or Edge work best)

### "I can't find my script in Power Automate"
- Make sure you saved the script in Excel
- Try refreshing the Power Automate page
- Make sure you're selecting the correct Excel file

### "The flow ran but no emails appeared"
- Check the flow's run history for errors
- Make sure you're looking at the correct Excel file
- Try running the flow again
- Check that the folder name matches exactly

### "Error: Script timed out"
- This can happen with very large emails
- Try limiting the body preview to fewer characters in the script
- Reduce the number of emails you're processing at once

### "I keep getting duplicate emails"
- The automatic flow adds emails as they arrive
- Don't also run the manual flow on the same emails
- Consider clearing your Excel file before running the manual flow

---

## Explain It Like You're 5 👶

Remember the two-robot system? Here's how it works with Office Scripts:

**Robot #1 - The Mail Carrier (Power Automate):**
- Watches your Outlook mailbox 👀
- When a new email arrives in "Job Hunting", it picks it up 📬
- Reads all the important info (who sent it, what it says, when it arrived)
- Carries that info to Robot #2

**Robot #2 - The Filing Robot (Office Script in Excel):**
- Lives inside your Excel file 📊
- Waits for Robot #1 to bring email info
- When info arrives, it writes it down in a new row
- Keeps everything neat and organized

**The Cool Part:**
Robot #1 knows how to talk to Robot #2! When you set up Power Automate, you're basically teaching Robot #1 to say:

"Hey Robot #2, I have an email!"
- "The date is: [date]"
- "It's from: [person]"
- "The subject is: [subject]"
- "Here's what it says: [body]"
- "It was in the: [folder] folder"

And Robot #2 says "Got it! Writing it down!" and adds a new row to Excel.

**Two Modes:**
1. **Guard Mode** - Robot #1 stands guard and catches EVERY new email as it arrives
2. **Collection Mode** - You press a button and Robot #1 goes and grabs ALL the old emails at once

You get to choose which mode(s) you want! 🎮

---

## Why This is Better Than the "Power Automate Only" Version

**Office Scripts version:**
- ✅ Cleaner code that's easier to understand
- ✅ More professional and maintainable
- ✅ Can add custom logic easily
- ✅ Better error handling
- ✅ More control over formatting

**Power Automate only version:**
- ✅ Simpler setup (no coding)
- ✅ Good for beginners
- ❌ Less flexible
- ❌ Harder to customize

Since you said you like Office Scripts better, you made the right choice! It's the more powerful option. 💪
