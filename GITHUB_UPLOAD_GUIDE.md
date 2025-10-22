# How to Upload to GitHub

This guide will walk you through uploading your Email to Excel project to GitHub.

## Prerequisites
- A GitHub account (create one at [github.com](https://github.com) if needed)
- Git installed on your computer

---

## Method 1: Using GitHub Web Interface (Easiest!)

This method is perfect if you're not comfortable with command line.

### Step 1: Create a New Repository

1. Go to [github.com](https://github.com)
2. Click the **"+"** icon in the top right
3. Select **"New repository"**
4. Fill in the details:
   - **Repository name**: `email-to-excel-exporters` (or whatever you like)
   - **Description**: "Scripts to export Gmail and Outlook emails to Excel/Google Sheets"
   - **Public** or **Private**: Choose based on your preference
   - **DON'T** check "Initialize with README" (we already have one)
5. Click **"Create repository"**

### Step 2: Upload Files

1. On the repository page, click **"uploading an existing file"**
2. Drag and drop ALL these files:
   - `README.md`
   - `LICENSE`
   - `.gitignore`
   - `CONTRIBUTING.md`
   - `gmail_to_sheets.gs`
   - `SETUP_GUIDE.md`
   - `outlook_to_excel_python.py`
   - `PYTHON_SETUP_GUIDE.md`
   - `outlook_office_script.ts`
   - `OFFICE_SCRIPT_COMPLETE_GUIDE.md`
   - `POWER_AUTOMATE_SIMPLE.md`

3. **IMPORTANT**: Make sure you **remove your personal information** from `outlook_to_excel_python.py`:
   - Change `EMAIL = "your-email@outlook.com"` to `EMAIL = "your-email@outlook.com"`
   - Change `PASSWORD = "your-app-password"` to `PASSWORD = "your-app-password-here"`
   
4. Add a commit message: "Initial commit - Email to Excel exporters"
5. Click **"Commit changes"**

✅ **Done!** Your project is now on GitHub!

---

## Method 2: Using Git Command Line

This is for those comfortable with the terminal/command prompt.

### Step 1: Install Git

**Windows:**
- Download from [git-scm.com](https://git-scm.com/download/win)
- Run the installer

**Mac:**
- Open Terminal
- Type: `git --version`
- If not installed, it will prompt you to install

**Linux:**
```bash
sudo apt-get install git
```

### Step 2: Create Repository on GitHub

1. Go to [github.com](https://github.com)
2. Click **"+"** → **"New repository"**
3. Name it: `email-to-excel-exporters`
4. Make it Public or Private
5. **DON'T** initialize with README
6. Click **"Create repository"**
7. **Copy the repository URL** (looks like: `https://github.com/yourusername/email-to-excel-exporters.git`)

### Step 3: Prepare Your Local Files

1. Create a new folder for your project
2. Copy all the files into this folder:
   - `README.md`
   - `LICENSE`
   - `.gitignore`
   - `CONTRIBUTING.md`
   - `gmail_to_sheets.gs`
   - `SETUP_GUIDE.md`
   - `outlook_to_excel_python.py`
   - `PYTHON_SETUP_GUIDE.md`
   - `outlook_office_script.ts`
   - `OFFICE_SCRIPT_COMPLETE_GUIDE.md`
   - `POWER_AUTOMATE_SIMPLE.md`

3. **CRITICAL**: Edit `outlook_to_excel_python.py` and remove your personal credentials:
   ```python
   EMAIL = "your-email@outlook.com"  # Change back to placeholder
   PASSWORD = "your-app-password-here"  # Change back to placeholder
   ```

### Step 4: Initialize Git and Push

Open Terminal (Mac/Linux) or Command Prompt (Windows) in your project folder and run:

```bash
# Initialize git repository
git init

# Add all files
git add .

# Commit the files
git commit -m "Initial commit - Email to Excel exporters"

# Add your GitHub repository as remote
git remote add origin https://github.com/yourusername/email-to-excel-exporters.git

# Push to GitHub
git branch -M main
git push -u origin main
```

When prompted, enter your GitHub username and password (or personal access token).

✅ **Done!** Your project is now on GitHub!

---

## Step 5: Customize Your Repository

### Add Topics (Tags)
1. On your repository page, click the gear icon ⚙️ next to "About"
2. Add topics like:
   - `email`
   - `gmail`
   - `outlook`
   - `excel`
   - `google-sheets`
   - `python`
   - `google-apps-script`
   - `automation`
   - `imap`

### Add a Description
In the same "About" section, add:
"Scripts to export Gmail and Outlook emails to Excel/Google Sheets for easy organization and analysis"

### Add a Website
If you have a personal website, add it here

---

## Important: Protect Your Credentials! 🔒

### Before Uploading

**Double-check these files DON'T contain your personal info:**

❌ **Never commit:**
- Your email address (real one)
- Your app password
- Any actual credentials

✅ **Always use placeholders:**
```python
EMAIL = "your-email@outlook.com"
PASSWORD = "your-app-password-here"
```

### If You Accidentally Committed Credentials

**If you uploaded your password by mistake:**

1. **Immediately** change your password/app password in your Outlook account
2. Delete the repository:
   - Go to repository Settings
   - Scroll to bottom → "Danger Zone"
   - Click "Delete this repository"
3. Create a new repository with cleaned files

**Never try to just edit the file** - the password will still be in Git history!

---

## Maintaining Your Repository

### Updating Files

**Web Interface:**
1. Go to the file on GitHub
2. Click the pencil icon ✏️
3. Make changes
4. Commit changes

**Command Line:**
```bash
# Make changes to files locally
# Then:
git add .
git commit -m "Description of changes"
git push
```

### Responding to Issues

When people open issues:
1. Be friendly and helpful
2. Ask for clarification if needed
3. Label issues appropriately
4. Close resolved issues

### Accepting Pull Requests

When someone submits improvements:
1. Review the code
2. Test if possible
3. Request changes if needed
4. Merge when satisfied

---

## Promoting Your Repository

### Share it on:
- Reddit (r/Python, r/productivity, r/gmail)
- Twitter/X with hashtags (#Python #Automation #Productivity)
- LinkedIn
- Dev.to or Medium (write a blog post about it)
- Hacker News

### Write a Blog Post
Share your experience creating this and how it helps with job hunting!

---

## Making It Popular ⭐

**Good README = More Stars**

Your README already has:
- ✅ Clear description
- ✅ Quick start guide
- ✅ Examples
- ✅ Troubleshooting
- ✅ Use cases
- ✅ Contributing guide

**To get more attention:**
- Add screenshots or GIFs showing it in action
- Create a demo video
- Add badges (build status, license, etc.)
- Keep it updated
- Respond to issues quickly
- Be welcoming to contributors

---

## Adding Badges to README

Add these to the top of your README.md:

```markdown
![Python Version](https://img.shields.io/badge/python-3.6+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Contributions Welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg)
```

They'll look like: 
![Python Version](https://img.shields.io/badge/python-3.6+-blue.svg) ![License](https://img.shields.io/badge/license-MIT-green.svg)

---

## GitHub Features to Enable

### Issues
- Already enabled by default
- Great for bug reports and feature requests

### Discussions
1. Go to Settings
2. Scroll to "Features"
3. Enable "Discussions"
4. Good for Q&A and community

### Wiki
- Great for extended documentation
- Enable in Settings → Features

### Projects
- Use GitHub Projects for tracking TODOs
- Plan new features

---

## Example Repository Structure

```
email-to-excel-exporters/
├── README.md                           # Main documentation (landing page)
├── LICENSE                             # MIT License
├── .gitignore                          # Ignored files
├── CONTRIBUTING.md                     # How to contribute
│
├── gmail/
│   ├── gmail_to_sheets.gs             # Gmail script
│   └── SETUP_GUIDE.md                 # Gmail setup guide
│
├── outlook/
│   ├── python/
│   │   ├── outlook_to_excel_python.py # Python IMAP script
│   │   └── PYTHON_SETUP_GUIDE.md      # Python setup guide
│   │
│   └── office-scripts/
│       ├── outlook_office_script.ts    # Office Scripts version
│       ├── OFFICE_SCRIPT_GUIDE.md      # Office Scripts guide
│       └── POWER_AUTOMATE_SIMPLE.md    # Power Automate guide
│
└── examples/
    ├── sample_output.xlsx              # Example output file
    └── screenshots/                    # Screenshots folder
```

**Optional**: Reorganize your files into folders like above for better organization!

---

## Checklist Before Publishing

- [ ] Removed all personal credentials
- [ ] README.md is clear and complete
- [ ] LICENSE file included
- [ ] .gitignore includes sensitive files
- [ ] Code has helpful comments
- [ ] All files uploaded
- [ ] Repository description added
- [ ] Topics/tags added
- [ ] Repository is Public (if you want others to use it)

---

## After Publishing

1. **Share the link** with friends and colleagues
2. **Star your own repo** (to show it's active)
3. **Tweet about it** with relevant hashtags
4. **Post on LinkedIn** (especially good for job hunting tools!)
5. **Submit to awesome lists** (search GitHub for "awesome email" or "awesome automation")

---

## Need Help?

- [GitHub Documentation](https://docs.github.com)
- [Git Tutorial](https://git-scm.com/docs/gittutorial)
- [Markdown Guide](https://www.markdownguide.org/)

---

Congratulations on publishing your first (or next) GitHub project! 🎉

Your repository link will be: `https://github.com/YOUR-USERNAME/email-to-excel-exporters`
