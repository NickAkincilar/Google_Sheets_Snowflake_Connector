# Google Sheets Snowflake Connector
Snowflake Connector for Google Sheets to fetch data from Snowflake tables or using SQL Queries via Snowflake SQL API endpoints

# Quick Installation Guide

Install it by making a local copy of the Sheet with the App Script and allowing the script to access your Google Sheets account.

### Steps:
1. **Click this link:** [Snowflake Connector Template](https://docs.google.com/spreadsheets/d/1TUO_-nBQYKkQTxJkwbZW-VZkf1LkkITHjpfjl7B3RmY/edit?usp=sharing) 
2. **File > Make a copy**  (Creates an editable copy in your Google drive)
3. **Refresh the page**
4. **Snowflake menu appears!**
5. **Configure your credentials**
    -  **It will ask you for authorization the first time**
    -  **Allow your Google Account to use it**
    -  **When you see a warning about UNVERIFIED APP**
        -   **Click "ADVANCED"**
        -   **Click "Go to Snowflake_GS_Sheets_Connector (unsafe)"**
        - **Pick "Select All" the click CONTINUE**

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

- **Authentication Policy:** Your user must be governed by an authentication policy that explicitly includes `orgname-PROGRAMMATIC_ACCESS_TOKEN` in its allowed methods.

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




