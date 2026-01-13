/**
 * Snowflake Google Sheets Connector
 * 
 * This Google Apps Script connects to Snowflake using SQL API
 * and allows users to fetch data into Google Sheets
 */

// Configuration - Store these in Script Properties for security
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const USER_PROPERTIES = PropertiesService.getUserProperties();

/**
 * Creates a custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Snowflake')
    .addItem('Configure Connection', 'showConfigDialog')
    .addItem('Query Data', 'showQuerySidebar')
    .addSeparator()
    .addItem('OAuth: Authorize', 'showOAuthAuthorization')
    .addItem('OAuth: Check Status', 'checkOAuthStatus')
    .addItem('OAuth: Debug Token', 'debugOAuthToken')
    .addItem('OAuth: Clear Token', 'clearOAuthToken')
    .addSeparator()
    .addItem('Clear Configuration', 'clearConfiguration')
    .addItem('Refresh All Credentials', 'refreshCredentials')
    .addToUi();
}

/**
 * Shows the configuration dialog for Snowflake credentials
 */
function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigDialog')
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Snowflake Connection Configuration');
}

/**
 * Shows the query sidebar for selecting databases, schemas, and tables
 */
function showQuerySidebar() {
  const config = getConfiguration();
  if (!config.account) {
    SpreadsheetApp.getUi().alert('Please configure Snowflake connection first.');
    return;
  }
  
  // Check if OAuth is configured but not authorized
  if (config.authMethod === 'oauth') {
    const oauthService = getSnowflakeOAuthService();
    if (!oauthService || !oauthService.hasAccess()) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'OAuth Not Authorized',
        'OAuth authentication is configured but not authorized yet.\n\n' +
        'Please authorize first by going to:\n' +
        'Snowflake menu > OAuth: Authorize\n\n' +
        'Would you like to authorize now?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        showOAuthAuthorization();
      }
      return;
    }
  }
  
  const html = HtmlService.createHtmlOutputFromFile('QuerySidebar')
    .setTitle('Snowflake Query Builder')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Saves Snowflake configuration to Script Properties
 */
function saveConfiguration(config) {
  try {
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_ACCOUNT', config.account);
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_AUTH_METHOD', config.authMethod || 'basic');
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_USERNAME', config.username || '');
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_PASSWORD', config.password || '');
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_WAREHOUSE', config.warehouse || '');
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_ROLE', config.role || '');
    
    // OAuth2 settings
    if (config.authMethod === 'oauth') {
      SCRIPT_PROPERTIES.setProperty('OAUTH_TENANT_ID', config.oauthTenantId || '');
      SCRIPT_PROPERTIES.setProperty('OAUTH_CLIENT_ID', config.oauthClientId || '');
      SCRIPT_PROPERTIES.setProperty('OAUTH_CLIENT_SECRET', config.oauthClientSecret || '');
      SCRIPT_PROPERTIES.setProperty('OAUTH_RESOURCE_APP_ID', config.oauthResourceAppId || '');
      SCRIPT_PROPERTIES.setProperty('OAUTH_SCOPE', config.oauthScope || 'profile offline_access');
    }
    
    return { success: true, message: 'Configuration saved successfully!' };
  } catch (error) {
    return { success: false, message: 'Error saving configuration: ' + error.toString() };
  }
}

/**
 * Gets Snowflake configuration from Script Properties
 */
function getConfiguration() {
  return {
    account: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_ACCOUNT') || '',
    authMethod: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_AUTH_METHOD') || 'basic',
    username: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_USERNAME') || '',
    password: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_PASSWORD') || '',
    warehouse: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_WAREHOUSE') || '',
    role: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_ROLE') || '',
    oauthTenantId: SCRIPT_PROPERTIES.getProperty('OAUTH_TENANT_ID') || '',
    oauthClientId: SCRIPT_PROPERTIES.getProperty('OAUTH_CLIENT_ID') || '',
    oauthClientSecret: SCRIPT_PROPERTIES.getProperty('OAUTH_CLIENT_SECRET') || '',
    oauthResourceAppId: SCRIPT_PROPERTIES.getProperty('OAUTH_RESOURCE_APP_ID') || '',
    oauthScope: SCRIPT_PROPERTIES.getProperty('OAUTH_SCOPE') || 'session:role-any offline_access'
  };
}

/**
 * Clears stored configuration
 */
function clearConfiguration() {
  // Clear basic/PAT settings
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_ACCOUNT');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_AUTH_METHOD');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_USERNAME');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_PASSWORD');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_WAREHOUSE');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_ROLE');
  
  // Clear OAuth settings
  SCRIPT_PROPERTIES.deleteProperty('OAUTH_TENANT_ID');
  SCRIPT_PROPERTIES.deleteProperty('OAUTH_CLIENT_ID');
  SCRIPT_PROPERTIES.deleteProperty('OAUTH_CLIENT_SECRET');
  SCRIPT_PROPERTIES.deleteProperty('OAUTH_RESOURCE_APP_ID');
  SCRIPT_PROPERTIES.deleteProperty('OAUTH_SCOPE');
  
  // Clear OAuth tokens
  const service = getSnowflakeOAuthService();
  if (service) {
    service.reset();
  }
  
  SpreadsheetApp.getUi().alert('Configuration cleared successfully!');
}

/**
 * Clears only OAuth authorization token (keeps configuration)
 */
function clearOAuthToken() {
  const config = getConfiguration();
  
  if (!config.authMethod || config.authMethod !== 'oauth') {
    SpreadsheetApp.getUi().alert(
      'OAuth Not Configured',
      'Your current authentication method is not OAuth.\n\n' +
      'Current method: ' + (config.authMethod || 'basic') + '\n\n' +
      'This function only clears OAuth tokens.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const service = getSnowflakeOAuthService();
  if (service) {
    service.reset();
    Logger.log('OAuth token cleared');
  }
  
  // Also clear from USER_PROPERTIES as a backup
  USER_PROPERTIES.deleteProperty('oauth2.SnowflakeOAuth');
  
  SpreadsheetApp.getUi().alert(
    'OAuth Token Cleared',
    'OAuth authorization token has been removed.\n\n' +
    'Your OAuth configuration (Client ID, Tenant ID, etc.) is still saved.\n\n' +
    'To re-authorize:\n' +
    'Snowflake menu > OAuth: Authorize',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Refreshes all credentials and tokens
 */
function refreshCredentials() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfiguration();
  
  if (!config.authMethod) {
    ui.alert('No configuration found. Please configure connection first.');
    return;
  }
  
  const response = ui.alert(
    'Refresh Credentials',
    'This will:\n' +
    '• Clear all cached tokens\n' +
    '• Keep your configuration settings\n' +
    '• Require re-authorization (for OAuth)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  let message = '';
  
  if (config.authMethod === 'oauth') {
    // Clear OAuth tokens
    const service = getSnowflakeOAuthService();
    if (service) {
      service.reset();
      Logger.log('OAuth tokens cleared');
    }
    USER_PROPERTIES.deleteProperty('oauth2.SnowflakeOAuth');
    message = 'OAuth tokens cleared.\n\n' +
              'Next steps:\n' +
              '1. Snowflake menu > OAuth: Authorize\n' +
              '2. Complete authorization\n' +
              '3. Use Query Data normally';
  } else if (config.authMethod === 'pat') {
    // For PAT tokens, there's no cached token - just verify current one works
    message = 'PAT token configuration refreshed.\n\n' +
              'Your Personal Access Token is still active.\n\n' +
              'If you need a new token:\n' +
              '1. Generate new PAT in Snowflake\n' +
              '2. Snowflake menu > Configure Connection\n' +
              '3. Update the PAT token field';
  } else {
    // Basic auth
    message = 'Basic authentication refreshed.\n\n' +
              'Your username/password configuration is still active.\n\n' +
              'If credentials changed:\n' +
              'Snowflake menu > Configure Connection';
  }
  
  ui.alert('Credentials Refreshed', message, ui.ButtonSet.OK);
}

/**
 * Executes a SQL query against Snowflake SQL API
 */
function executeSnowflakeQuery(sqlText) {
  const config = getConfiguration();
  
  // Validate configuration based on auth method
  if (!config.account) {
    throw new Error('Snowflake connection not configured. Please configure connection first.');
  }
  
  if (config.authMethod === 'oauth') {
    // For OAuth, we need OAuth config but not username/password
    if (!config.oauthTenantId || !config.oauthClientId) {
      throw new Error('OAuth not configured. Please configure OAuth settings first.');
    }
  } else {
    // For Basic Auth and PAT, we need username and password/token
    if (!config.username || !config.password) {
      throw new Error('Snowflake connection not configured. Please configure connection first.');
    }
  }
  
  // Snowflake SQL API endpoint
  const accountUrl = config.account.includes('.') ? config.account : `${config.account}.snowflakecomputing.com`;
  const url = `https://${accountUrl}/api/v2/statements`;
  
  // Prepare request payload
  const payload = {
    statement: sqlText,
    timeout: 60,
    database: null,
    schema: null,
    warehouse: config.warehouse || null,
    role: config.role || null
  };
  
  // Set up authentication based on configured method
  let authHeader;
  let tokenType;
  
  if (config.authMethod === 'oauth') {
    // OAuth2 authentication (External OAuth with Azure AD)
    try {
      const oauthService = getSnowflakeOAuthService();
      if (!oauthService) {
        throw new Error('OAuth not configured. Please configure OAuth settings first.');
      }
      if (!oauthService.hasAccess()) {
        throw new Error('OAuth not authorized. Please use Snowflake > OAuth: Authorize menu.');
      }
      const accessToken = oauthService.getAccessToken();
      Logger.log('OAuth token obtained, length: ' + accessToken.length);
      
      // Decode and log token claims for debugging
      const tokenClaims = decodeJWT(accessToken);
      if (tokenClaims) {
        Logger.log('Token audience (aud): ' + tokenClaims.aud);
        Logger.log('Token issuer (iss): ' + tokenClaims.iss);
        Logger.log('Token subject (sub): ' + tokenClaims.sub);
        Logger.log('Token UPN: ' + (tokenClaims.upn || 'not set'));
        Logger.log('Token expiry: ' + new Date(tokenClaims.exp * 1000));
      }
      
      authHeader = 'Bearer ' + accessToken;
      // For External OAuth (Azure AD), don't set X-Snowflake-Authorization-Token-Type
      // Snowflake will validate the token directly against the External OAuth integration
      tokenType = undefined;
    } catch (oauthError) {
      Logger.log('OAuth error details: ' + oauthError.toString());
      throw new Error('OAuth error: ' + oauthError.toString());
    }
  } else {
    // Detect if using PAT token (Bearer) or username/password (Basic Auth)
    // PAT tokens are typically long alphanumeric strings without special characters
    const isPAT = config.password && config.password.length > 50 && !/[\s:@]/.test(config.password);
    
    if (isPAT) {
      // Use Bearer token authentication for PAT
      authHeader = 'Bearer ' + config.password;
      tokenType = 'PROGRAMMATIC_ACCESS_TOKEN';
    } else {
      // Use Basic authentication for username/password
      if (!config.username || !config.password) {
        throw new Error('Username and password required. Please configure connection.');
      }
      const auth = Utilities.base64Encode(config.username + ':' + config.password);
      authHeader = 'Basic ' + auth;
      tokenType = undefined;
    }
  }
  
  // Build headers - only include X-Snowflake-Authorization-Token-Type if tokenType is defined
  const headers = {
    'Authorization': authHeader,
    'Accept': 'application/json',
    'User-Agent': 'GoogleSheetsConnector/1.0'
  };
  
  // Only add token type header for PAT tokens
  // For Basic auth and External OAuth, this header should NOT be set
  if (tokenType) {
    headers['X-Snowflake-Authorization-Token-Type'] = tokenType;
    Logger.log('Setting token type header: ' + tokenType);
  }
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      const errorData = JSON.parse(responseText);
      throw new Error(`Snowflake API Error (${responseCode}): ${errorData.message || responseText}`);
    }
    
    const result = JSON.parse(responseText);
    
    // Check if there are additional partitions to retrieve
    if (result.resultSetMetaData && result.resultSetMetaData.partitionInfo && 
        result.resultSetMetaData.partitionInfo.length > 1) {
      // Retrieve all additional partitions
      const statementHandle = result.statementHandle;
      const allData = result.data || [];
      const totalPartitions = result.resultSetMetaData.partitionInfo.length;
      
      Logger.log(`Query returned ${totalPartitions} partitions. Fetching additional partitions...`);
      
      // Fetch partitions 1, 2, 3, etc. (partition 0 is already in result)
      for (let i = 1; i < totalPartitions; i++) {
        Logger.log(`Fetching partition ${i}/${totalPartitions - 1}...`);
        const partitionData = fetchPartition(statementHandle, i, config, authHeader, tokenType);
        if (partitionData && partitionData.length > 0) {
          allData.push(...partitionData);
          Logger.log(`Partition ${i} added ${partitionData.length} rows`);
        }
      }
      
      Logger.log(`Total rows retrieved: ${allData.length}`);
      
      // Update result with all combined data
      result.data = allData;
    }
    
    return result;
  } catch (error) {
    throw new Error('Failed to execute query: ' + error.toString());
  }
}

/**
 * Fetches a specific partition of query results
 */
function fetchPartition(statementHandle, partitionNumber, config, authHeader, tokenType) {
  const accountUrl = config.account.includes('.') ? config.account : `${config.account}.snowflakecomputing.com`;
  const url = `https://${accountUrl}/api/v2/statements/${statementHandle}?partition=${partitionNumber}`;
  
  // Build headers - only include X-Snowflake-Authorization-Token-Type if tokenType is defined
  const headers = {
    'Authorization': authHeader,
    'Accept': 'application/json',
    'Accept-Encoding': 'gzip',
    'User-Agent': 'GoogleSheetsConnector/1.0'
  };
  
  // Only add token type header for PAT tokens
  if (tokenType) {
    headers['X-Snowflake-Authorization-Token-Type'] = tokenType;
  }
  
  const options = {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      const errorText = response.getContentText();
      Logger.log(`Warning: Failed to fetch partition ${partitionNumber}: HTTP ${responseCode}`);
      Logger.log(`Error response: ${errorText}`);
      return [];
    }
    
    // Parse response - Snowflake may claim gzip but send plain JSON
    const contentEncoding = response.getHeaders()['Content-Encoding'] || response.getHeaders()['content-encoding'];
    let responseData;
    
    try {
      // First, try parsing as plain JSON (Snowflake often sends plain JSON despite gzip header)
      const textData = response.getContentText();
      responseData = JSON.parse(textData);
      Logger.log(`Partition ${partitionNumber}: Parsed as plain JSON (${textData.length} bytes)`);
    } catch (jsonError) {
      // If plain JSON fails, try gzip decompression
      try {
        Logger.log(`Partition ${partitionNumber}: Plain JSON failed, trying gzip decompression...`);
        const compressed = response.getContent();
        const decompressed = Utilities.ungzip(compressed);
        const textData = decompressed.getDataAsString();
        responseData = JSON.parse(textData);
        Logger.log(`Partition ${partitionNumber}: Successfully decompressed (${textData.length} bytes)`);
      } catch (gzipError) {
        Logger.log(`Partition ${partitionNumber}: Both JSON and gzip parsing failed`);
        Logger.log(`JSON error: ${jsonError.toString()}`);
        Logger.log(`Gzip error: ${gzipError.toString()}`);
        throw new Error(`Failed to parse partition ${partitionNumber} response`);
      }
    }
    
    // Partition responses can be either:
    // 1. Just the data array directly: [[row1], [row2], ...]
    // 2. Wrapped in an object: { data: [[row1], [row2], ...] }
    if (Array.isArray(responseData)) {
      return responseData;
    } else if (responseData.data && Array.isArray(responseData.data)) {
      return responseData.data;
    } else {
      Logger.log(`Warning: Unexpected partition ${partitionNumber} response format`);
      return [];
    }
  } catch (error) {
    Logger.log(`Error fetching partition ${partitionNumber}: ${error.toString()}`);
    return [];
  }
}

/**
 * Fetches list of databases from Snowflake
 */
function getDatabases() {
  try {
    const result = executeSnowflakeQuery('SHOW DATABASES');
    
    // Parse the results
    if (result.data && result.data.length > 0) {
      // The result format for SHOW DATABASES includes database name typically in first column
      const databases = result.data.map(row => row[1]); // Database name is usually in column 1
      return { success: true, data: databases };
    }
    
    return { success: true, data: [] };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Fetches list of schemas for a given database
 */
function getSchemas(database) {
  try {
    const result = executeSnowflakeQuery(`SHOW SCHEMAS IN DATABASE "${database}"`);
    
    if (result.data && result.data.length > 0) {
      // Schema name is typically in column 1
      const schemas = result.data.map(row => row[1]);
      return { success: true, data: schemas };
    }
    
    return { success: true, data: [] };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Fetches list of tables for a given database and schema
 */
function getTables(database, schema) {
  try {
    const result = executeSnowflakeQuery(`SHOW TABLES IN "${database}"."${schema}"`);
    
    if (result.data && result.data.length > 0) {
      // Table name is typically in column 1
      const tables = result.data.map(row => row[1]);
      return { success: true, data: tables };
    }
    
    return { success: true, data: [] };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Fetches data from a specific table and writes it to the active sheet
 */
function fetchTableData(database, schema, table, limit) {
  try {
    const limitClause = limit && limit > 0 ? ` LIMIT ${limit}` : '';
    const query = `SELECT * FROM "${database}"."${schema}"."${table}"${limitClause}`;
    
    const result = executeSnowflakeQuery(query);
    
    if (!result.data) {
      return { success: false, error: 'No data returned from query' };
    }
    
    // Get the active sheet
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Clear existing content
    sheet.clear();
    
    // Write headers
    if (result.resultSetMetaData && result.resultSetMetaData.rowType) {
      const headers = result.resultSetMetaData.rowType.map(col => col.name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
    }
    
    // Write data
    if (result.data.length > 0) {
      const startRow = 2;
      sheet.getRange(startRow, 1, result.data.length, result.data[0].length).setValues(result.data);
      
      // Auto-resize columns
      for (let i = 1; i <= result.data[0].length; i++) {
        sheet.autoResizeColumn(i);
      }
    }
    
    return { 
      success: true, 
      message: `Successfully loaded ${result.data.length} rows from ${database}.${schema}.${table}`,
      rowCount: result.data.length
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Executes a custom SQL query and writes results to the active sheet
 */
function executeCustomQuery(query, clearSheet) {
  try {
    const result = executeSnowflakeQuery(query);
    
    if (!result.data) {
      return { success: false, error: 'No data returned from query' };
    }
    
    // Get the active sheet
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Clear existing content if requested
    if (clearSheet) {
      sheet.clear();
    }
    
    // Write headers
    let startRow = 1;
    if (result.resultSetMetaData && result.resultSetMetaData.rowType) {
      const headers = result.resultSetMetaData.rowType.map(col => col.name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
      startRow = 2;
    }
    
    // Write data
    if (result.data.length > 0) {
      sheet.getRange(startRow, 1, result.data.length, result.data[0].length).setValues(result.data);
      
      // Auto-resize columns
      for (let i = 1; i <= result.data[0].length; i++) {
        sheet.autoResizeColumn(i);
      }
    }
    
    return { 
      success: true, 
      message: `Successfully executed query. Returned ${result.data.length} rows.`,
      rowCount: result.data.length
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Creates and configures the OAuth2 service for Snowflake
 */
function getSnowflakeOAuthService() {
  try {
    const config = getConfiguration();
    
    if (!config.oauthTenantId || !config.oauthClientId) {
      Logger.log('OAuth not configured: missing tenant ID or client ID');
      return null;
    }
    
    // Construct OAuth scope properly for Azure AD + Snowflake External OAuth
    let scope;
    if (config.oauthResourceAppId) {
      // For External OAuth with Azure AD, we need:
      // 1. api://resource-app-id/.default - for API access
      // 2. profile - to get UPN and other profile claims
      // 3. offline_access - for refresh tokens
      // 4. openid - for ID token claims
      scope = `api://${config.oauthResourceAppId}/.default profile openid offline_access`;
      Logger.log('Using Snowflake External OAuth scope: ' + scope);
    } else {
      // Fallback to user-provided scope (for standard Snowflake OAuth - not External OAuth)
      scope = config.oauthScope || 'profile offline_access';
      Logger.log('Using standard scope: ' + scope);
    }
    
    Logger.log('Creating OAuth service with tenant: ' + config.oauthTenantId);
    Logger.log('Client ID: ' + config.oauthClientId);
    Logger.log('Resource App ID: ' + config.oauthResourceAppId);
    
    return OAuth2.createService('SnowflakeOAuth')
        .setAuthorizationBaseUrl(`https://login.microsoftonline.com/${config.oauthTenantId}/oauth2/v2.0/authorize`)
        .setTokenUrl(`https://login.microsoftonline.com/${config.oauthTenantId}/oauth2/v2.0/token`)
        .setClientId(config.oauthClientId)
        .setClientSecret(config.oauthClientSecret)
        .setCallbackFunction('authCallback')
        .setPropertyStore(USER_PROPERTIES)
        .setScope(scope)
        .setParam('prompt', 'consent');
  } catch (error) {
    Logger.log('Error creating OAuth service: ' + error.toString());
    return null;
  }
}

/**
 * OAuth callback function
 */
function authCallback(request) {
  const service = getSnowflakeOAuthService();
  const authorized = service.handleCallback(request);
  
  if (authorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab and return to your spreadsheet.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab.');
  }
}

/**
 * Show OAuth authorization URL
 */
function showOAuthAuthorization() {
  const service = getSnowflakeOAuthService();
  
  if (!service) {
    SpreadsheetApp.getUi().alert('OAuth not configured. Please configure OAuth settings in Connection Configuration first.');
    return;
  }
  
  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('Already authorized! Use "OAuth: Check Status" to see details.');
    return;
  }
  
  const authorizationUrl = service.getAuthorizationUrl();
  const template = HtmlService.createHtmlOutput(`
    <html>
      <body>
        <h3>Snowflake OAuth Authorization</h3>
        <p>Click the link below to authorize this application:</p>
        <p><a href="${authorizationUrl}" target="_blank">Authorize Access</a></p>
        <p><small>After authorizing, you can close this dialog.</small></p>
      </body>
    </html>
  `).setWidth(400).setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(template, 'OAuth Authorization');
}

/**
 * Check OAuth status
 */
function checkOAuthStatus() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfiguration();
  const service = getSnowflakeOAuthService();
  
  if (!config.authMethod) {
    ui.alert(
      'No Configuration',
      'No authentication configured.\n\n' +
      'Please use: Snowflake menu > Configure Connection',
      ui.ButtonSet.OK
    );
    return;
  }
  
  if (config.authMethod !== 'oauth') {
    ui.alert(
      'Not Using OAuth',
      'Current authentication method: ' + config.authMethod.toUpperCase() + '\n\n' +
      'This status check is only for OAuth authentication.\n\n' +
      'Your current credentials are stored and active.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  if (!service) {
    ui.alert(
      'OAuth Not Configured',
      'OAuth authentication method is selected but not properly configured.\n\n' +
      'Please reconfigure: Snowflake menu > Configure Connection',
      ui.ButtonSet.OK
    );
    return;
  }
  
  if (service.hasAccess()) {
    const token = service.getAccessToken();
    const response = ui.alert(
      'OAuth Status',
      '✅ AUTHORIZED\n\n' +
      'Token Preview: ' + token.substring(0, 20) + '...\n' +
      'Tenant ID: ' + config.oauthTenantId + '\n' +
      'Client ID: ' + config.oauthClientId + '\n\n' +
      'Token is active and ready to use.\n\n' +
      'Need to refresh token?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      clearOAuthToken();
    }
  } else {
    const response = ui.alert(
      'OAuth Status',
      '❌ NOT AUTHORIZED\n\n' +
      'OAuth is configured but not authorized.\n\n' +
      'Would you like to authorize now?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      showOAuthAuthorization();
    }
  }
}

/**
 * Gets the OAuth redirect URI for Azure AD configuration
 * @returns {string} The redirect URI to use in Azure AD App Registration
 */
function getOAuthRedirectUri() {
  return ScriptApp.getService().getUrl();
}

/**
 * Decodes a JWT token to inspect its claims (for debugging)
 * @param {string} token - The JWT token
 * @returns {object} The decoded token payload
 */
function decodeJWT(token) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) {
      return null;
    }
    
    // Decode the payload (second part)
    const payload = parts[1];
    // Add padding if needed
    const paddedPayload = payload + '='.repeat((4 - payload.length % 4) % 4);
    const decoded = Utilities.base64Decode(paddedPayload, Utilities.Charset.UTF_8);
    const jsonString = Utilities.newBlob(decoded).getDataAsString();
    return JSON.parse(jsonString);
  } catch (error) {
    Logger.log('Error decoding JWT: ' + error.toString());
    return null;
  }
}

/**
 * Checks if OAuth is authorized (without showing UI)
 * Used by ConfigDialog to determine if testing is allowed
 * @returns {boolean} True if OAuth is authorized, false otherwise
 */
function isOAuthAuthorized() {
  try {
    const config = getConfiguration();
    
    // If not using OAuth, return false
    if (!config.authMethod || config.authMethod !== 'oauth') {
      return false;
    }
    
    // Check if OAuth service is configured
    const service = getSnowflakeOAuthService();
    if (!service) {
      return false;
    }
    
    // Check if service has access
    return service.hasAccess();
  } catch (error) {
    Logger.log('Error checking OAuth authorization: ' + error.toString());
    return false;
  }
}

/**
 * Debug OAuth token - run this from Apps Script to inspect the OAuth token claims
 * This helps diagnose OAuth issues with External OAuth integration
 */
function debugOAuthToken() {
  try {
    const config = getConfiguration();
    
    if (!config.authMethod || config.authMethod !== 'oauth') {
      Logger.log('Not using OAuth authentication');
      return;
    }
    
    const service = getSnowflakeOAuthService();
    if (!service) {
      Logger.log('ERROR: OAuth service not configured');
      return;
    }
    
    if (!service.hasAccess()) {
      Logger.log('ERROR: OAuth not authorized. Use Snowflake > OAuth: Authorize');
      return;
    }
    
    const token = service.getAccessToken();
    Logger.log('=== OAuth Token Information ===');
    Logger.log('Token length: ' + token.length);
    Logger.log('Token preview: ' + token.substring(0, 50) + '...');
    
    const claims = decodeJWT(token);
    if (claims) {
      Logger.log('\n=== Token Claims ===');
      Logger.log('Audience (aud): ' + claims.aud);
      Logger.log('Issuer (iss): ' + claims.iss);
      Logger.log('Subject (sub): ' + claims.sub);
      Logger.log('UPN (upn): ' + (claims.upn || 'not set'));
      Logger.log('Email: ' + (claims.email || 'not set'));
      Logger.log('Issued at: ' + new Date(claims.iat * 1000));
      Logger.log('Expires at: ' + new Date(claims.exp * 1000));
      Logger.log('App ID (appid): ' + (claims.appid || 'not set'));
      Logger.log('\n=== Full Token Claims ===');
      Logger.log(JSON.stringify(claims, null, 2));
      
      // Check if audience matches expected value
      Logger.log('\n=== Validation ===');
      const expectedAudience = `api://${config.oauthResourceAppId}`;
      if (claims.aud === expectedAudience) {
        Logger.log('✅ Audience matches expected value: ' + expectedAudience);
      } else {
        Logger.log('❌ Audience MISMATCH!');
        Logger.log('   Expected: ' + expectedAudience);
        Logger.log('   Actual: ' + claims.aud);
        Logger.log('   This may cause Snowflake External OAuth to reject the token.');
      }
      
      // Check expiration
      const now = new Date();
      const expiry = new Date(claims.exp * 1000);
      if (now < expiry) {
        Logger.log('✅ Token is still valid (expires in ' + Math.round((expiry - now) / 1000 / 60) + ' minutes)');
      } else {
        Logger.log('❌ Token has EXPIRED!');
      }
    } else {
      Logger.log('ERROR: Failed to decode JWT token');
    }
    
    SpreadsheetApp.getUi().alert('OAuth token debug complete. Check Execution logs:\nExtensions > Apps Script > Executions');
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Debug test function - run this from Apps Script to test with your credentials
 */
function debugTest() {
  // Test query on CUSTOMER_CHURN_DATA
  try {
    Logger.log('=== Starting debug test ===');
    const query = 'SELECT * FROM DELETETHIS.PUBLIC.CUSTOMER_CHURN_DATA LIMIT 100';
    Logger.log('Query: ' + query);
    
    const result = executeSnowflakeQuery(query);
    
    Logger.log('=== Query Results ===');
    Logger.log('Statement Handle: ' + result.statementHandle);
    Logger.log('Rows in data array: ' + (result.data ? result.data.length : 0));
    
    if (result.resultSetMetaData) {
      Logger.log('Total rows (metadata): ' + result.resultSetMetaData.numRows);
      Logger.log('Partition info: ' + JSON.stringify(result.resultSetMetaData.partitionInfo));
    }
    
    SpreadsheetApp.getUi().alert('Debug test complete. Check logs in Apps Script > Executions');
    return result;
  } catch (error) {
    Logger.log('=== ERROR ===');
    Logger.log(error.toString());
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Tests the Snowflake connection
 */
function testConnection() {
  try {
    const result = executeSnowflakeQuery('SELECT CURRENT_VERSION() AS VERSION, CURRENT_USER() AS USER, CURRENT_ROLE() AS ROLE, CURRENT_WAREHOUSE() AS WAREHOUSE');
    
    if (result.data && result.data.length > 0) {
      return { 
        success: true, 
        message: 'Connection successful!',
        details: {
          version: result.data[0][0],
          user: result.data[0][1],
          role: result.data[0][2],
          warehouse: result.data[0][3]
        }
      };
    }
    
    return { success: false, error: 'Unexpected response from Snowflake' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

