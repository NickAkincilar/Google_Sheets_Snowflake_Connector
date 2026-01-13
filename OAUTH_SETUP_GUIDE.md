# OAuth 2.0 Setup Guide - Azure AD + Snowflake External OAuth

Complete step-by-step instructions for setting up OAuth 2.0 authentication between Google Sheets Connector, Azure AD, and Snowflake.

---

## 📋 Overview

This guide covers:
1. **Azure AD App Registration** - Register and configure application
2. **Snowflake External OAuth Integration** - Configure Snowflake to trust Azure AD tokens
3. **Google Sheets Connector Configuration** - Set up the connector
4. **Testing & Verification** - Ensure everything works
5. **Troubleshooting** - Common issues and solutions

**Time Required**: 20-30 minutes

---

## 🔐 Prerequisites

### Required Access:
- ✅ Azure AD **Application Administrator** or **Global Administrator** role
- ✅ Snowflake **ACCOUNTADMIN** role
- ✅ Google Sheets access
- ✅ Google Apps Script access

### Required Information:
- Your Azure AD Tenant ID
- Your Snowflake account identifier
- User's email address (for user mapping)

---

## Part 1: Azure AD Configuration

### Step 1.1: Register New Application

1. **Open Azure Portal**
   - Go to: https://portal.azure.com
   - Sign in with your Azure AD admin account

2. **Navigate to App Registrations**
   ```
   Azure Active Directory → App registrations → + New registration
   ```

3. **Fill in Registration Details**
   - **Name**: `Snowflake Google Sheets Connector`
   - **Supported account types**: 
     - Select: **Accounts in this organizational directory only (Single tenant)**
   - **Redirect URI**: 
     - Leave blank for now (we'll add this later)
   - Click **Register**

4. **Save Your Application Details**
   After registration, you'll see the overview page. **Copy and save these values**:
   - ✅ **Application (client) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - ✅ **Directory (tenant) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

---

### Step 1.2: Create Client Secret

1. **Navigate to Certificates & secrets**
   ```
   Your App → Certificates & secrets → + New client secret
   ```

2. **Create the Secret**
   - **Description**: `Google Sheets Connector Secret`
   - **Expires**: 
     - Recommended: **180 days (6 months)** or **365 days (12 months)**
     - ⚠️ Set a calendar reminder to rotate before expiration!
   - Click **Add**

3. **Copy the Secret Value**
   - ⚠️ **CRITICAL**: Copy the **Value** immediately - you cannot see it again!
   - ✅ **Save**: Client Secret value somewhere secure
   - **Format**: Long alphanumeric string (e.g., `abc123~...`)

---

### Step 1.3: Expose an API (Create Resource ID)

This step creates the Application ID URI that Snowflake will validate.

1. **Navigate to Expose an API**
   ```
   Your App → Expose an API
   ```

2. **Set Application ID URI**
   - Click **+ Set** next to "Application ID URI"
   - **Default format**: `api://<your-client-id>`
   - **Example**: `api://2a70a01e-7d51-4849-9228-928a7c05b5cf`
   - ✅ Accept the default (recommended)
   - Click **Save**
   - ✅ **Copy and save** this Application ID URI

3. **Add a Scope** (Optional but Recommended)
   - Click **+ Add a scope**
   - **Scope name**: `session:role-any`
   - **Who can consent**: **Admins and users**
   - **Admin consent display name**: `Access Snowflake with any role`
   - **Admin consent description**: `Allows the application to access Snowflake on behalf of the signed-in user`
   - **User consent display name**: `Access your Snowflake account`
   - **User consent description**: `Allows the application to access Snowflake on your behalf`
   - **State**: **Enabled**
   - Click **Add scope**

---

### Step 1.4: Configure Token Claims (Get preferred_username)

1. **Navigate to Token Configuration**
   ```
   Your App → Token configuration
   ```

2. **Add Optional Claim for UPN** (Optional)
   - Click **+ Add optional claim**
   - **Token type**: Select **Access**
   - Scroll down and check: ☑️ **upn**
   - Click **Add**
   - If prompted about Microsoft Graph permissions, click **Add** to grant

   **Note**: Even with this configured, UPN may not appear in v2 tokens. The token will have `preferred_username` instead, which works perfectly with Snowflake.

---

### Step 1.5: Grant API Permissions

1. **Navigate to API Permissions**
   ```
   Your App → API permissions
   ```

2. **Review Default Permissions**
   You should see:
   - Microsoft Graph → `User.Read` (Delegated)

3. **Add Additional Permissions** (if needed)
   - Click **+ Add a permission**
   - Select **Microsoft Graph**
   - Select **Delegated permissions**
   - Search and add:
     - ☑️ `openid`
     - ☑️ `profile`
     - ☑️ `offline_access`
   - Click **Add permissions**

4. **Grant Admin Consent**
   - Click **✓ Grant admin consent for [Your Organization]**
   - Confirm by clicking **Yes**
   - All permissions should show green checkmarks

---

### Step 1.6: Configure Redirect URI (Done Later)

We'll add the redirect URI after configuring the Google Sheets connector, as it will provide the exact URL needed.

**Skip this for now - we'll come back to it in Part 3.**

---

### Step 1.7: Summary - Azure AD Information to Save

At this point, you should have saved:

| Item | Example | Where to Find |
|------|---------|---------------|
| **Tenant ID** | `3976235e-9bd9-4eed-aee8-1de491c63438` | App → Overview |
| **Client ID** | `01cce8e2-5ffb-4e3f-b6ac-b6ee2a5cfd6b` | App → Overview |
| **Client Secret** | `abc123~xyz...` | App → Certificates & secrets |
| **Application ID URI** | `api://2a70a01e-7d51-4849-9228-928a7c05b5cf` | App → Expose an API |

**Resource App ID**: This is the GUID portion of the Application ID URI:
- Full URI: `api://2a70a01e-7d51-4849-9228-928a7c05b5cf`
- **Resource App ID**: `2a70a01e-7d51-4849-9228-928a7c05b5cf` (just the GUID)

---

## Part 2: Snowflake Configuration

### Step 2.1: Verify Prerequisites


1. **Verify ACCOUNTADMIN Access**
   ```sql
   SELECT CURRENT_ROLE();
   ```
   - Should show: `ACCOUNTADMIN`
   - If not: `USE ROLE ACCOUNTADMIN;`

---

### Step 2.2: Create External OAuth Security Integration

1. **Open Snowflake Web UI**
   - Navigate to: `https://app.snowflake.com`
   - Sign in with ACCOUNTADMIN privileges

2. **Open a Worksheet**
   - Click **Worksheets** (left sidebar)
   - Click **+ Worksheet** (top right)

3. **Run the CREATE SECURITY INTEGRATION Command**

```sql
-- Create External OAuth Security Integration for Azure AD
CREATE OR REPLACE SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS
  TYPE = EXTERNAL_OAUTH
  ENABLED = TRUE
  EXTERNAL_OAUTH_TYPE = AZURE
  EXTERNAL_OAUTH_ISSUER = 'https://login.microsoftonline.com/<YOUR_TENANT_ID>/v2.0'
  EXTERNAL_OAUTH_JWS_KEYS_URL = 'https://login.microsoftonline.com/<YOUR_TENANT_ID>/discovery/v2.0/keys'
  EXTERNAL_OAUTH_AUDIENCE_LIST = ('<YOUR_RESOURCE_APP_ID>')
  EXTERNAL_OAUTH_TOKEN_USER_MAPPING_CLAIM = 'preferred_username'
  EXTERNAL_OAUTH_SNOWFLAKE_USER_MAPPING_ATTRIBUTE = 'LOGIN_NAME'
  EXTERNAL_OAUTH_ANY_ROLE_MODE = 'ENABLE';
```

4. **Replace Placeholders with Your Values**

| Placeholder | Replace With | Example |
|-------------|--------------|---------|
| `<YOUR_TENANT_ID>` | Your Azure AD Tenant ID | `3976235e-9bd9-4eed-aee8-1de491c63438` |
| `<YOUR_RESOURCE_APP_ID>` | GUID from Application ID URI | `2a70a01e-7d51-4849-9228-928a7c05b5cf` |

**Example with real values:**

```sql
CREATE OR REPLACE SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS
  TYPE = EXTERNAL_OAUTH
  ENABLED = TRUE
  EXTERNAL_OAUTH_TYPE = AZURE
  EXTERNAL_OAUTH_ISSUER = 'https://login.microsoftonline.com/3976235e-9bd9-4eed-aee8-1de491c63438/v2.0'
  EXTERNAL_OAUTH_JWS_KEYS_URL = 'https://login.microsoftonline.com/3976235e-9bd9-4eed-aee8-1de491c63438/discovery/v2.0/keys'
  EXTERNAL_OAUTH_AUDIENCE_LIST = ('2a70a01e-7d51-4849-9228-928a7c05b5cf')
  EXTERNAL_OAUTH_TOKEN_USER_MAPPING_CLAIM = 'preferred_username'
  EXTERNAL_OAUTH_SNOWFLAKE_USER_MAPPING_ATTRIBUTE = 'LOGIN_NAME'
  EXTERNAL_OAUTH_ANY_ROLE_MODE = 'ENABLE';
```

5. **Execute the Command**
   - Click **Run** (▶️ button)
   - Should see: `Security Integration AZURE_AD_OAUTH_GOOGLE_SHEETS successfully created.`

---

### Step 2.3: Verify the Integration

```sql
-- Show the integration details
SHOW INTEGRATIONS LIKE 'AZURE_AD_OAUTH_GOOGLE_SHEETS';

-- Describe the integration
DESC SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS;
```

**Expected Output:**
- `enabled: true`
- `external_oauth_type: AZURE`
- `external_oauth_issuer: https://login.microsoftonline.com/.../v2.0`
- `external_oauth_token_user_mapping_claim: ['preferred_username']`

---

### Step 2.4: Configure User Mapping

The OAuth token contains `preferred_username` claim (e.g., `john.smith@snowflake.com`). Your Snowflake user's `LOGIN_NAME` must match this value.

1. **Check Your Current User**

```sql
-- Find your Snowflake username
SELECT CURRENT_USER() AS my_username;

-- Check your user's LOGIN_NAME
SELECT 
  USER_NAME,
  LOGIN_NAME,
  EMAIL,
  DISABLED
FROM SNOWFLAKE.ACCOUNT_USAGE.USERS 
WHERE USER_NAME = CURRENT_USER();
```

2. **Update LOGIN_NAME to Match Your Email**

```sql
-- Replace <YOUR_USERNAME> with your actual Snowflake username
-- Replace <YOUR_EMAIL> with your Azure AD email (from preferred_username claim)

ALTER USER <YOUR_USERNAME> SET LOGIN_NAME = '<YOUR_EMAIL>';

-- Example:
ALTER USER Nsmith SET LOGIN_NAME = 'john.smith@snowflake.com';
```

3. **Verify the Update**

```sql
-- Confirm LOGIN_NAME is now set correctly
SELECT USER_NAME, LOGIN_NAME, EMAIL 
FROM SNOWFLAKE.ACCOUNT_USAGE.USERS 
WHERE USER_NAME = CURRENT_USER();
```

**Expected Output:**
- `LOGIN_NAME` should match your Azure AD email

---

### Step 2.5: Configure Network Policy (Handle IP Restrictions)

Google Apps Script runs from Google Cloud IPs. If you have network policies enabled, you need to whitelist these IPs.

**Option A: Check if You Have Network Policies**

```sql
-- Check account-level network policy
SHOW PARAMETERS LIKE 'NETWORK_POLICY' IN ACCOUNT;

-- If a policy is set, view its details
SHOW NETWORK POLICIES;
DESC NETWORK POLICY <YOUR_POLICY_NAME>;
```

**Option B: Temporarily Disable for Testing - NOT RECOMMENDED**



```sql
-- Remove network policy restriction on account(TESTING ONLY)
ALTER USER JOHN SET NETWORK_POLICY = NULL;
```

**Option C: Whitelist Google Cloud IPs (Production)**

```sql
-- Update your network policy to include Google Cloud IP ranges
ALTER NETWORK POLICY <YOUR_POLICY_NAME> SET 
  ALLOWED_IP_LIST = (
    -- Keep your existing IPs
    '203.0.113.0/24',      -- Example: Your office network
    '198.51.100.0/24',     -- Example: Your VPN
    
    -- Add Google Cloud IP ranges
    '34.116.0.0/16',       -- Google Cloud range
    '35.191.0.0/16',       -- Google Cloud range
    '130.211.0.0/16'       -- Google Cloud range
  );
```

**Note**: You may encounter specific Google IPs like `34.116.22.37`. Add the /16 range to cover all IPs in that subnet.

---

### Step 2.6: Summary - Snowflake Configuration Complete

At this point, Snowflake is configured to:
- ✅ Accept Azure AD OAuth tokens
- ✅ Validate tokens against your Azure AD tenant
- ✅ Map `preferred_username` claim to user's `LOGIN_NAME`
- ✅ Allow users to assume any role (controlled by Snowflake RBAC)
- ✅ (Optional) Allow Google Cloud IPs

---

## Part 3: Google Sheets Connector Configuration



### Step 3.1: Get the OAuth Redirect URI

1. **Open Configuration Dialog**
   ```
   Snowflake menu → Configure Connection
   ```

2. **Select OAuth Authentication**
   - **Authentication Method**: Select `OAuth 2.0 (Azure AD / External OAuth)`

3. **Copy the Redirect URI**
   - Scroll down to the blue box labeled **"📋 Azure AD Redirect URI"**
   - Copy the entire URL shown
   - **Format**: `https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback`
   - **Example**: `https://script.google.com/macros/d/AKfycbzXXXXX.../usercallback`

---

### Step 3.2: Add Redirect URI to Azure AD

1. **Return to Azure Portal**
   - Go to your app registration
   ```
   Azure AD → App registrations → Your App → Authentication
   ```

2. **Add Web Redirect URI**
   - Under **Platform configurations**, click **+ Add a platform**
   - Select **Web**
   - **Redirect URIs**: Paste the URL from Step 3.2
   - **Example**: `https://script.google.com/macros/d/AKfycbzXXXXX.../usercallback`
   - ✅ **Check both boxes**:
     - ☑️ Access tokens (used for implicit flows)
     - ☑️ ID tokens (used for implicit and hybrid flows)
   - Click **Configure**

3. **Verify the Redirect URI**
   - Should now see the redirect URI listed under **Web** platform
   - Status should be **Configured**

---

### Step 3.3: Configure the Connector

1. **Return to Google Sheets Configuration Dialog**
   ```
   Snowflake menu → Configure Connection
   ```

2. **Fill in Connection Details**

| Field | Value | Example |
|-------|-------|---------|
| **Account Identifier** | Your Snowflake account | `xy12345.us-east-1` |
| **Authentication Method** | OAuth 2.0 (Azure AD / External OAuth) | — |
| **Azure AD Tenant ID** | From Azure (Step 1.7) | `3976235e-9bd9-4eed-aee8-1de491c63438` |
| **Client (Application) ID** | From Azure (Step 1.7) | `01cce8e2-5ffb-4e3f-b6ac-b6ee2a5cfd6b` |
| **Client Secret** | From Azure (Step 1.2) | `abc123~xyz...` |
| **Snowflake Resource App ID** | From Azure (Step 1.7) | `2a70a01e-7d51-4849-9228-928a7c05b5cf` |
| **OAuth Scope** | Default (leave as-is) | `profile offline_access` |
| **Warehouse** | Your warehouse (optional) | `COMPUTE_WH` |
| **Role** | Your role (optional) | `ACCOUNTADMIN` |

4. **Save Configuration**
   - Click **Save Configuration**
   - Should see: "Configuration saved successfully!"
   - ⚠️ Do NOT click "Test Connection" yet

---

### Step 3.4: Authorize OAuth

1. **Start OAuth Authorization**
   ```
   Snowflake menu → OAuth: Authorize
   ```

2. **Authorization Dialog**
   - A dialog will appear with an authorization link
   - Click the **"Authorize Access"** link
   - A new browser tab will open

3. **Sign In to Azure AD**
   - Enter your Azure AD credentials
   - Username: Your work email
   - Password: Your work password
   - Complete MFA if required

4. **Grant Permissions**
   - You'll see a permissions consent screen:
     ```
     Snowflake Google Sheets Connector wants to:
     - View your basic profile info
     - Maintain access to data you have given it access to
     ```
   - Click **Accept** or **Yes**

5. **Success Page**
   - You should see: "Success! You can close this tab."
   - Close the authorization tab
   - Close the dialog in Google Sheets

---

### Step 3.5: Verify OAuth Status

1. **Check OAuth Status**
   ```
   Snowflake menu → OAuth: Check Status
   ```

2. **Expected Result**
   ```
   ✅ AUTHORIZED
   
   Token Preview: eyJhbGciOiJSUzI1N...
   Tenant ID: 3976235e-9bd9-4eed-aee8-1de491c63438
   Client ID: 01cce8e2-5ffb-4e3f-b6ac-b6ee2a5cfd6b
   
   Token is active and ready to use.
   ```

3. **Debug Token Details** (Optional but Recommended)
   ```
   Snowflake menu → OAuth: Debug Token
   ```

4. **Check Execution Logs**
   ```
   Extensions → Apps Script → Executions (left sidebar)
   ```

   **Look for**:
   - ✅ `Audience (aud): 2a70a01e-7d51-4849-9228-928a7c05b5cf`
   - ✅ `preferred_username: your.email@company.com`
   - ✅ `Token is still valid`
   - ✅ `Audience matches expected value` (or close enough)

---

## Part 4: Testing & Verification

### Step 4.1: Test the Connection

1. **Open Configuration Dialog**
   ```
   Snowflake menu → Configure Connection
   ```

2. **Click Test Connection**
   - Should see: "Testing OAuth connection..."
   - Wait 5-10 seconds

3. **Expected Success**
   ```
   ✅ Connection successful!
   
   Snowflake Version: 8.x.x
   User: YOUR_USERNAME
   Role: YOUR_ROLE
   Warehouse: YOUR_WAREHOUSE
   ```

4. **If Test Fails** - See Troubleshooting section below

---

### Step 4.2: Query Data

1. **Open Query Sidebar**
   ```
   Snowflake menu → Query Data
   ```

2. **Select Database, Schema, Table**
   - **Database**: Select from dropdown
   - **Schema**: Select from dropdown
   - **Table**: Select from dropdown
   - **Row Limit**: 100 (for testing)

3. **Fetch Data**
   - Click **📥 Fetch Data to Sheet**
   - Data should appear in your spreadsheet
   - Headers should be formatted (bold, background color)

4. **Try Custom Query** (Optional)
   ```sql
   SELECT CURRENT_USER(), CURRENT_ROLE(), CURRENT_WAREHOUSE();
   ```
   - Scroll down in sidebar to **Custom Query** section
   - Paste the query
   - Click **▶️ Execute Query**
   - Should show your user, role, and warehouse

---

### Step 4.3: Verify User Identity

Run this query to confirm the OAuth user mapping is working:

```sql
-- In Google Sheets sidebar, run this custom query:
SELECT 
  CURRENT_USER() AS snowflake_username,
  CURRENT_ROLE() AS current_role,
  CURRENT_WAREHOUSE() AS current_warehouse,
  SESSION_ID() AS session_id
```

**Expected Result**:
- `snowflake_username`: Should match your Snowflake user
- `current_role`: Should match the role you're using
- Shows that OAuth authentication is working correctly

---

## Part 5: Troubleshooting

### Issue 1: "Invalid OAuth access token"

**Symptoms:**
```
Error: Snowflake API Error (401): Invalid OAuth access token
```

**Possible Causes & Solutions:**

#### A. Audience Mismatch

**Debug:**
```
Snowflake menu → OAuth: Debug Token
Check logs: Audience (aud) value
```

**Fix:**
```sql
-- Update Snowflake to accept correct audience
-- Use the GUID only, not the full api:// URI
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET
  EXTERNAL_OAUTH_AUDIENCE_LIST = ('<RESOURCE_APP_ID_GUID_ONLY>');

-- Example:
-- If token shows: aud: "2a70a01e-7d51-4849-9228-928a7c05b5cf"
-- Then use:
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET
  EXTERNAL_OAUTH_AUDIENCE_LIST = ('2a70a01e-7d51-4849-9228-928a7c05b5cf');
```

#### B. Token Expired

**Debug:**
```
Snowflake menu → OAuth: Debug Token
Check logs: "Expires at" timestamp
```

**Fix:**
```
Snowflake menu → OAuth: Clear Token
Snowflake menu → OAuth: Authorize
(Re-authorize)
```

#### C. Integration Not Enabled

**Debug:**
```sql
SHOW INTEGRATIONS LIKE 'AZURE_AD_OAUTH_GOOGLE_SHEETS';
-- Check: enabled = true
```

**Fix:**
```sql
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET ENABLED = TRUE;
```

---

### Issue 2: "IP not allowed to access Snowflake"

**Symptoms:**
```
Error: Incoming request with IP 34.116.22.37 is not allowed to access Snowflake
```

**This means OAuth is working! Just IP restrictions blocking it.**

**Solution A: Temporarily Disable (Testing)**
```sql
ALTER ACCOUNT SET NETWORK_POLICY = NULL;
```

**Solution B: Whitelist Google Cloud IPs (Production)**
```sql
-- Get your network policy name
SHOW PARAMETERS LIKE 'NETWORK_POLICY' IN ACCOUNT;

-- Update the policy
ALTER NETWORK POLICY YOUR_POLICY_NAME SET 
  ALLOWED_IP_LIST = (
    -- Your existing IPs
    '203.0.113.0/24',
    -- Add Google Cloud IPs
    '34.116.0.0/16',
    '35.191.0.0/16',
    '130.211.0.0/16'
  );
```

---

### Issue 3: "OAuth not authorized"

**Symptoms:**
```
Error: OAuth not authorized. Please use Snowflake > OAuth: Authorize menu.
```

**Solution:**
```
1. Snowflake menu → OAuth: Clear Token
2. Snowflake menu → OAuth: Authorize
3. Complete sign-in and grant permissions
4. Try again
```

---

### Issue 4: User Mapping Fails

**Symptoms:**
```
Token is valid but Snowflake rejects it
"User not found" or similar error
```

**Debug:**
```
Snowflake menu → OAuth: Debug Token
Check: preferred_username value
```

**Fix:**
```sql
-- Update your Snowflake user's LOGIN_NAME to match preferred_username
ALTER USER <YOUR_USERNAME> SET LOGIN_NAME = '<EMAIL_FROM_TOKEN>';

-- Example:
ALTER USER Nsmith SET LOGIN_NAME = 'john.smith@snowflake.com';
```

---

### Issue 5: Redirect URI Mismatch

**Symptoms:**
```
Azure AD error: "The redirect URI does not match"
```

**Solution:**
1. Get correct redirect URI from Google Sheets config dialog
2. Azure AD → Your App → Authentication
3. Verify redirect URI matches exactly (including /usercallback)
4. Must be added under **Web** platform, not Single-page application

---

### Issue 6: Token Missing Claims

**Symptoms:**
```
Debug shows: UPN (upn): not set
preferred_username: not set
```

**Solution:**
1. Azure AD → Your App → Token configuration
2. Add optional claim: `upn` (Access token)
3. Azure AD → Your App → API permissions
4. Add: Microsoft Graph → `profile`, `openid`
5. Grant admin consent
6. Re-authorize in Google Sheets

---

## Part 6: Security Best Practices

### 6.1: Rotate Client Secrets Regularly

- Set expiration: 180 days or 365 days maximum
- Create calendar reminder 2 weeks before expiration
- Generate new secret before old one expires
- Update Google Sheets configuration with new secret
- Revoke old secret after verification

### 6.2: Limit OAuth Integration Scope

```sql
-- Instead of ENABLE_ANY_ROLE, specify allowed roles
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET
  EXTERNAL_OAUTH_ANY_ROLE_MODE = 'DISABLE'
  EXTERNAL_OAUTH_ALLOWED_ROLES_LIST = ('DATA_ANALYST', 'REPORTING_USER');
```

### 6.3: Monitor OAuth Usage

```sql
-- Check OAuth token usage
SELECT 
  USER_NAME,
  CLIENT_ENVIRONMENT,
  REPORTED_CLIENT_TYPE,
  AUTHENTICATION_METHOD,
  SESSION_ID,
  LOGIN_TIME
FROM SNOWFLAKE.ACCOUNT_USAGE.SESSIONS
WHERE AUTHENTICATION_METHOD = 'OAUTH_ACCESS_TOKEN'
ORDER BY LOGIN_TIME DESC
LIMIT 100;
```

### 6.4: Audit OAuth Configuration Changes

```sql
-- Monitor changes to security integrations
SELECT 
  QUERY_TEXT,
  USER_NAME,
  ROLE_NAME,
  EXECUTION_STATUS,
  START_TIME
FROM SNOWFLAKE.ACCOUNT_USAGE.QUERY_HISTORY
WHERE QUERY_TEXT ILIKE '%SECURITY INTEGRATION%'
  AND QUERY_TEXT ILIKE '%AZURE_AD_OAUTH%'
ORDER BY START_TIME DESC;
```

---

## Part 7: Maintenance & Operations

### 7.1: Updating Client Secret

When your Azure AD client secret expires:

1. **Generate New Secret in Azure AD**
   ```
   Azure AD → Your App → Certificates & secrets → + New client secret
   ```

2. **Update Google Sheets Connector**
   ```
   Snowflake menu → Configure Connection
   Update "Client Secret" field
   Save Configuration
   ```

3. **Test Connection**
   ```
   Snowflake menu → Configure Connection → Test Connection
   ```

4. **Revoke Old Secret**
   ```
   Azure AD → Your App → Certificates & secrets → Delete old secret
   ```

---

### 7.2: Adding New Users

For each new user who needs access:

1. **Ensure User Exists in Azure AD**
   - User must be in the same tenant

2. **Create/Update Snowflake User**
   ```sql
   -- Create user if doesn't exist
   CREATE USER IF NOT EXISTS NEW_USER
     LOGIN_NAME = 'new.user@company.com'
     DISPLAY_NAME = 'New User'
     EMAIL = 'new.user@company.com';
   
   -- Grant appropriate roles
   GRANT ROLE DATA_ANALYST TO USER NEW_USER;
   ```

3. **User Follows Part 3 Steps**
   - Install connector
   - Configure with same Azure AD app credentials
   - Authorize OAuth
   - Start querying

---

### 7.3: Disabling OAuth Access

To temporarily or permanently disable OAuth:

```sql
-- Disable the integration
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET ENABLED = FALSE;

-- Or drop it entirely
DROP SECURITY INTEGRATION IF EXISTS AZURE_AD_OAUTH_GOOGLE_SHEETS;
```

---

## Part 8: Quick Reference

### Azure AD Information

| Item | Where to Find |
|------|---------------|
| Tenant ID | App → Overview |
| Client ID | App → Overview |
| Client Secret | App → Certificates & secrets → Generate new |
| Application ID URI | App → Expose an API |
| Resource App ID | GUID portion of Application ID URI |
| Redirect URI | Get from Google Sheets connector config |

### Snowflake Commands

```sql
-- View integration
SHOW INTEGRATIONS LIKE 'AZURE_AD_OAUTH_GOOGLE_SHEETS';
DESC SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS;

-- Update audience
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET
  EXTERNAL_OAUTH_AUDIENCE_LIST = ('your-resource-app-id');

-- Update user mapping
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET
  EXTERNAL_OAUTH_TOKEN_USER_MAPPING_CLAIM = 'preferred_username';

-- Set user LOGIN_NAME
ALTER USER YOUR_USERNAME SET LOGIN_NAME = 'your.email@company.com';

-- Disable/Enable
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET ENABLED = FALSE;
ALTER SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS SET ENABLED = TRUE;
```

### Google Sheets Menu

```
Snowflake → Configure Connection  (Set up credentials)
Snowflake → Query Data             (Fetch data)
Snowflake → OAuth: Authorize       (Complete OAuth flow)
Snowflake → OAuth: Check Status    (Verify authorization)
Snowflake → OAuth: Debug Token     (Inspect token claims)
Snowflake → OAuth: Clear Token     (Force re-authorization)
Snowflake → Refresh All Credentials (Clear cached tokens)
Snowflake → Clear Configuration    (Remove all settings)
```

---

## Part 9: Testing Checklist

Use this checklist to verify your setup:

### Azure AD Configuration
- [ ] App registration created
- [ ] Client secret generated and saved
- [ ] Application ID URI configured (`api://...`)
- [ ] Optional claims added (upn for access tokens)
- [ ] API permissions granted (profile, openid, offline_access)
- [ ] Admin consent granted for permissions
- [ ] Redirect URI added (from Google Sheets)

### Snowflake Configuration
- [ ] External OAuth integration created
- [ ] Integration enabled (`ENABLED = TRUE`)
- [ ] Audience matches resource app ID (GUID only, no `api://`)
- [ ] Issuer URL correct (includes `/v2.0`)
- [ ] JWS keys URL correct
- [ ] User mapping claim set (`preferred_username`)
- [ ] User LOGIN_NAME matches Azure AD email
- [ ] Network policy allows Google Cloud IPs (if applicable)

### Google Sheets Connector
- [ ] Code deployed to Apps Script
- [ ] Snowflake menu appears in spreadsheet
- [ ] OAuth configuration saved (all fields filled)
- [ ] OAuth authorization completed (signed in to Azure AD)
- [ ] OAuth status shows "AUTHORIZED"
- [ ] Debug token shows correct audience
- [ ] Debug token shows preferred_username claim
- [ ] Test connection succeeds
- [ ] Can load database/schema/table lists
- [ ] Can fetch data to spreadsheet

---

## Part 10: Support & Resources

### Official Documentation

- **Snowflake External OAuth**: https://docs.snowflake.com/en/user-guide/oauth-azure
- **Azure AD OAuth 2.0**: https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow
- **Google Apps Script OAuth2**: https://github.com/googleworkspace/apps-script-oauth2

### Common Links

- Azure Portal: https://portal.azure.com
- Snowflake Web UI: https://app.snowflake.com
- Google Sheets: https://sheets.google.com

### Getting Help

1. **Check Logs**
   ```
   Extensions → Apps Script → Executions
   Look for error messages and stack traces
   ```

2. **Use Debug Tools**
   ```
   Snowflake menu → OAuth: Debug Token
   Review token claims and validation
   ```

3. **Verify Snowflake Integration**
   ```sql
   DESC SECURITY INTEGRATION AZURE_AD_OAUTH_GOOGLE_SHEETS;
   ```

4. **Check Azure AD Logs**
   ```
   Azure Portal → Azure AD → Sign-in logs
   Filter by application name
   ```

---

## Summary

You've now configured:
1. ✅ Azure AD application with OAuth2 credentials
2. ✅ Snowflake External OAuth Security Integration
3. ✅ Google Sheets connector with OAuth authentication
4. ✅ End-to-end OAuth flow from Google Sheets → Azure AD → Snowflake

**Total Configuration Time**: ~20-30 minutes

**Result**: Secure, SSO-enabled access to Snowflake from Google Sheets using Azure AD credentials!

---

**Last Updated**: January 2026  
**Tested With**:
- Azure AD (Microsoft Entra ID)
- Google Apps Script OAuth2 Library v43

---

For additional help or to report issues, check the other documentation files:
- `README.md` - Overview and features
- `TROUBLESHOOTING.md` - Common issues
- `TOKEN_MANAGEMENT.md` - Managing credentials
