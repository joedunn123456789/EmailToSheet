# OAuth2 Setup for Personal Microsoft Accounts

Since Microsoft has disabled Basic Authentication (app passwords) for your account, you need to set up OAuth2. This requires registering an application with Microsoft.

## Step-by-Step Guide

### 1. Register Your Application

1. Go to: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
2. Click **"New registration"**
3. Fill in the form:
   - **Name**: `Email to Excel Exporter` (or any name you like)
   - **Supported account types**: Select **"Accounts in any organizational directory and personal Microsoft accounts"**
   - **Redirect URI**: Leave blank for now
4. Click **Register**

### 2. Get Your Application (Client) ID

After registration, you'll see the Overview page:
- Copy the **Application (client) ID** (looks like: `12345678-1234-1234-1234-123456789abc`)
- Save this - you'll need it for your .env file

### 3. Configure Authentication

1. In the left menu, click **"Authentication"**
2. Click **"Add a platform"**
3. Select **"Mobile and desktop applications"**
4. Check the box for: `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Also check: `http://localhost`
6. Click **Configure**
7. Scroll down to **"Allow public client flows"**
8. Set it to **"Yes"**
9. Click **Save**

### 4. Configure API Permissions

1. In the left menu, click **"API permissions"**
2. Click **"Add a permission"**
3. Select **"Microsoft Graph"**
4. Select **"Delegated permissions"**
5. Search for and check: **"Mail.Read"**
6. Click **"Add permissions"**
7. (Optional) Click **"Grant admin consent"** if available

### 5. Update Your .env File

Open your `.env` file and update the CLIENT_ID:

```
CLIENT_ID=YOUR-APPLICATION-CLIENT-ID-HERE
TENANT_ID=common
```

Replace `YOUR-APPLICATION-CLIENT-ID-HERE` with the Application (client) ID from step 2.

### 6. Run the Script

```bash
python3 outlook_to_excel.py
```

The script will:
1. Give you a code and URL
2. Open your browser to https://microsoft.com/devicelogin
3. You'll paste the code
4. Sign in with your **personal** Microsoft account (jdunn0423@live.com)
5. Grant permission to read your mail
6. The script will download your emails!

## Troubleshooting

**"Work or school account only" error**
- Make sure you selected "Accounts in any organizational directory **and personal Microsoft accounts**" in step 1

**"Admin approval required" error**
- This shouldn't happen for personal accounts, but if it does, make sure "Allow public client flows" is set to "Yes"

**"Invalid client" error**
- Double-check your CLIENT_ID in the .env file
- Make sure you copied the entire Application (client) ID

## Security Notes

- This app registration is private to you
- Only you can use it
- You can delete it anytime from the Azure portal
- The app only has permission to read emails (Mail.Read)
- No one else can access your emails through this app
