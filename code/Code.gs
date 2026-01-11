/**
 * Snowflake Google Sheets Connector
 * 
 * This Google Apps Script connects to Snowflake using SQL API
 * and allows users to fetch data into Google Sheets
 */

// Configuration - Store these in Script Properties for security
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

/**
 * Creates a custom menu in Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Snowflake')
    .addItem('Configure Connection', 'showConfigDialog')
    .addItem('Query Data', 'showQuerySidebar')
    .addItem('Clear Configuration', 'clearConfiguration')
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
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_USERNAME', config.username);
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_PASSWORD', config.password);
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_WAREHOUSE', config.warehouse || '');
    SCRIPT_PROPERTIES.setProperty('SNOWFLAKE_ROLE', config.role || '');
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
    username: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_USERNAME') || '',
    password: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_PASSWORD') || '',
    warehouse: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_WAREHOUSE') || '',
    role: SCRIPT_PROPERTIES.getProperty('SNOWFLAKE_ROLE') || ''
  };
}

/**
 * Clears stored configuration
 */
function clearConfiguration() {
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_ACCOUNT');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_USERNAME');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_PASSWORD');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_WAREHOUSE');
  SCRIPT_PROPERTIES.deleteProperty('SNOWFLAKE_ROLE');
  SpreadsheetApp.getUi().alert('Configuration cleared successfully!');
}

/**
 * Executes a SQL query against Snowflake SQL API
 */
function executeSnowflakeQuery(sqlText) {
  const config = getConfiguration();
  
  if (!config.account || !config.username || !config.password) {
    throw new Error('Snowflake connection not configured. Please configure connection first.');
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
  
  // Set up authentication
  // Detect if using PAT token (Bearer) or username/password (Basic Auth)
  // PAT tokens are typically long alphanumeric strings without special characters
  const isPAT = config.password.length > 50 && !/[\s:@]/.test(config.password);
  
  let authHeader;
  if (isPAT) {
    // Use Bearer token authentication for PAT
    authHeader = 'Bearer ' + config.password;
  } else {
    // Use Basic authentication for username/password
    const auth = Utilities.base64Encode(config.username + ':' + config.password);
    authHeader = 'Basic ' + auth;
  }
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': authHeader,
      'Accept': 'application/json',
      'User-Agent': 'GoogleSheetsConnector/1.0',
      'X-Snowflake-Authorization-Token-Type': isPAT ? 'PROGRAMMATIC_ACCESS_TOKEN' : undefined
    },
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
        const partitionData = fetchPartition(statementHandle, i, config, isPAT);
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
function fetchPartition(statementHandle, partitionNumber, config, isPAT) {
  const accountUrl = config.account.includes('.') ? config.account : `${config.account}.snowflakecomputing.com`;
  const url = `https://${accountUrl}/api/v2/statements/${statementHandle}?partition=${partitionNumber}`;
  
  let authHeader;
  if (isPAT) {
    authHeader = 'Bearer ' + config.password;
  } else {
    const auth = Utilities.base64Encode(config.username + ':' + config.password);
    authHeader = 'Basic ' + auth;
  }
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': authHeader,
      'Accept': 'application/json',
      'Accept-Encoding': 'gzip',
      'User-Agent': 'GoogleSheetsConnector/1.0',
      'X-Snowflake-Authorization-Token-Type': isPAT ? 'PROGRAMMATIC_ACCESS_TOKEN' : undefined
    },
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

