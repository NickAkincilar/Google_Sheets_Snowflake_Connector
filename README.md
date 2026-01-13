# Google Sheets Snowflake Connector
Snowflake Connector for Google Sheets to fetch data from Snowflake tables or using SQL Queries via Snowflake SQL API endpoints. THis plug-in is provided as-is and is a personal project not directly associated with Snoflake

## Disclaimer

This project (“Snowflake Connector for Google Sheets”) is an independent, personal project and is **not affiliated with, endorsed by, or supported by Snowflake Inc.**

The plugin is provided **“AS IS”**, without warranty of any kind. Use at your own risk.

“Snowflake” is a trademark of Snowflake Inc. All other trademarks belong to their respective owners.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

# Quick Installation Guide

Install it by making a local copy of the Sheet with the App Script and allowing the script to access your Google Sheets account.

### Steps:
1. **Click this link:** [Snowflake Connector Template](https://docs.google.com/spreadsheets/d/1TUO_-nBQYKkQTxJkwbZW-VZkf1LkkITHjpfjl7B3RmY/edit?usp=sharing) 
2. **File > Make a copy**  (Creates an editable copy in your Google drive)
3. **Refresh the page**
4. **Snowflake menu appears!**






### 1. Configure Snowflake Connection

1. Click **Snowflake** > **Configure Connection** in the menu
2. Fill in your Snowflake credentials:
   - **Account Identifier**: Your Snowflake account (e.g., `xy12345.us-east-1` or `xy12345.snowflakecomputing.com`)
   - **Username**: Your Snowflake username
   - **Password / PAT Token**: Your Snowflake password OR Personal Access Token (PAT)
     - **Password**: Use for traditional username/password authentication
     - **PAT Token**: Paste your PAT token for token-based authentication (automatically detected)
   - **Warehouse** (Optional): Warehouse to use (leave empty for default)
   - **Role** (Optional): Role to use (leave empty for default)
3. Click **Test Connection** to verify credentials
4. Click **Save Configuration**

**Authentication Methods Supported:**
- **Basic Auth** (Username + Password) - Traditional authentication
- **Bearer Token** (PAT) - Personal Access Token authentication (auto-detected)
- **OAuth 2.0** (Azure AD / SSO) - Enterprise SSO with automatic token refresh

## Using OAUTH Requires additional setup
Use [OAUTH SETUP GUIDE](https://github.com/NickAkincilar/Google_Sheets_Snowflake_Connector/blob/main/OAUTH_SETUP_GUIDE.md) document to configure.


<br><br>

## USING CONNECTOR

### 2. Browse and Fetch Data

1. Click **Snowflake** > **Query Data** in the menu
2. A sidebar will appear on the right
3. Use the **Table Browser** section:
   - Select a **Database** from the dropdown
   - Select a **Schema** from the dropdown
   - Select a **Table** from the dropdown
   - Set a row limit (optional)
   - Click **📥 Fetch Data to Sheet**
4. Data will be loaded into your active sheet with formatted headers

### 3. Execute Custom Queries

1. In the sidebar, scroll to the **Custom Query** section
2. Enter your SQL query in the text area
3. Check/uncheck "Clear sheet before inserting data" as needed
4. Click **▶️ Execute Query**
---
###  Managing Credentials & Tokens

The Snowflake menu provides several options for managing authentication:

#### Clear OAuth Token
- **Menu**: Snowflake > OAuth: Clear Token
- **What it does**: Removes OAuth authorization token only (keeps configuration)
- **When to use**: 
  - OAuth token expired or invalid
  - Need to re-authorize with different account
  - Troubleshooting OAuth issues
- **Note**: Configuration (Client ID, Tenant ID) remains saved

#### Refresh All Credentials
- **Menu**: Snowflake > Refresh All Credentials
- **What it does**: 
  - Clears all cached tokens
  - Keeps your configuration settings
  - For OAuth: requires re-authorization
  - For PAT: verifies current token status
  - For Basic: confirms credentials are active
- **When to use**:
  - Switching between accounts
  - Token seems stale or expired
  - Getting authentication errors

#### Clear Configuration
- **Menu**: Snowflake > Clear Configuration
- **What it does**: Removes ALL stored credentials and settings
- **When to use**:
  - Starting fresh with new account
  - Switching authentication methods completely
  - Uninstalling the connector
- **Warning**: You'll need to reconfigure everything

#### OAuth: Check Status
- **Menu**: Snowflake > OAuth: Check Status
- **What it does**: Shows current OAuth authorization status
- **Shows**:
  - Authorization state (authorized/not authorized)
  - Token preview
  - Tenant and Client ID
  - Quick refresh option
- **When to use**: 
  - Verify OAuth is working
  - Check which account is authorized
  - Quick access to refresh token

## Available Menu Options

Complete Snowflake menu structure:

```
Snowflake
├── Configure Connection          - Set up credentials and connection
├── Query Data                     - Browse & fetch data (sidebar)
├── ─────────────────────────────
├── OAuth: Authorize               - Complete OAuth authorization flow
├── OAuth: Check Status            - View OAuth status & refresh
├── OAuth: Clear Token            - Remove OAuth token only
├── ─────────────────────────────
├── Clear Configuration            - Remove all settings
└── Refresh All Credentials        - Clear tokens, keep config
```


<br/><br/>

## Snowflake Configuration Instructions

This guide helps you prepare your Snowflake account for use with the Google Sheets Connector.



## Prerequisites

You need a Snowflake account with:
- User configured to access Snowflake via PAT tokens
- Login credentials (username and pat token)
- Appropriate permissions to access databases
- Access to an existing warehouse 



### Key Requirements for using PAT tokens from external apps for connections
While you don't need a special role to create a token for yourself, there are a few prerequisites your account administrator must have configured first:

- **User Type:** Your user must be set to `TYPE = PERSON`(standard for human users) or `TYPE = SERVICE`.

- **Authentication Policy:**
    - Your user must be governed by an authentication policy that explicitly includes `orgname-PROGRAMMATIC_ACCESS_TOKEN` in its allowed methods.
    - Added OAuth2 for Snowflake External Aouth provider. use [OAUTH SETUP GUIDE](https://github.com/NickAkincilar/Google_Sheets_Snowflake_Connector/blob/main/OAUTH_SETUP_GUIDE.md) document to configure.

- **Network Policy:** For the token to actually work, you must be subject to a Network Policy. This is a security requirement to ensure PATs are only used from trusted locations.

---

### Finding Your Snowflake Account Identifier



1. Log into Snowflake Snowsight UI
2. Click on your UserID in **LOWER-LEFT** corner
3. Click on **Connect a tool to Snowflake**  
4. Copy **ACCOUNT IDENTIFIER** value which has the  `orgname-accountid` format


### Generating a PAT Token for password


Use Snowsight (Web UI)
Log in to Snowsight.

1. In the navigation menu, go to Governance & Security » Users & roles.

2. Select your own user profile.

3. Scroll down to the Programmatic access tokens section.

4. Choose the **ROLE** that will be used for access & the expiration **TIME LIMIT**  

5. Click + Generate new token.

6. Copy the token immediately. For security, Snowflake will never show the token secret again after you close the dialog.


<br/><br/><br/>


## SECURITY INFO FOR THIS PLUGIN 
### 🔐 Credentials Storage
**Where They're Stored:**

Google Apps Script's Script Properties Service - Line 9 in **Code.gs**:

## 🔒 Security Details:

### Location:
- NOT in the spreadsheet
- NOT in the script code
- NOT visible to other users
- Stored in Google's secure property store for the Apps Script project

### Access Control:

**✅ Only accessible by:**
- The script itself
- The owner of the Google Sheet/Apps Script project
- Google's secure backend

**❌ NOT accessible by:**
- Other users with view/edit access to the sheet
- People viewing the script code
- External applications

**Encryption:**
- Google encrypts Script Properties at rest
- Transmitted over HTTPS only
- Never exposed in client-side JavaScript

---




