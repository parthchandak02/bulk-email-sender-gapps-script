// email.gs
// Google Apps Script for Email Management and Response Tracking

/**
 * @OnlyCurrentDoc
 */

// Global variables and configuration
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE', // Replace with your spreadsheet ID
  SHEET_NAME: 'Sheet1',
  EMAIL_SUBJECT: 'Follow-up Request',
  SENDER_NAME: 'Your Name'
};

// Column indices (0-based)
const COLS = {
  FIRST_NAME: 0,    // Column A
  EMAIL: 1,         // Column B
  ACTION: 2,        // Column C
  SENT_DATE: 3,     // Column D - Changed from STATUS
  EMAIL_LINK: 4,    // Column E - New column
  RESPONSES_START: 5 // Column F onwards - Moved one column right
};

// Add this at the top with other constants
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

// Logging utility
const Logger = {
  log: function(component, message, data = null) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] [${component}] ${message}`;
    console.log(logMessage);
    if (data) {
      console.log('Data:', JSON.stringify(data, null, 2));
    }
    return logMessage;
  },
  error: function(component, error, context = null) {
    const timestamp = new Date().toISOString();
    const errorMessage = `[${timestamp}] [${component}] ERROR: ${error.message || error}`;
    console.error(errorMessage);
    if (context) {
      console.error('Context:', JSON.stringify(context, null, 2));
    }
    if (error.stack) {
      console.error('Stack:', error.stack);
    }
    return errorMessage;
  }
};

// UI Constants
const UI = {
  ICONS: {
    EMAIL: '📧',
    REFRESH: '🔄',
    WARNING: '⚠️',
    SUCCESS: '✅',
    ERROR: '❌'
  },
  MESSAGES: {
    PROCESSING: 'Processing emails...',
    SUCCESS: 'Operation completed successfully',
    ERROR: 'An error occurred'
  }
};

// Add these constants at the top of your file
const ONE_SECOND = 1000;
const ONE_MINUTE = ONE_SECOND * 60;
const MAX_EXECUTION_TIME = ONE_MINUTE * 5; // Leave 1 minute buffer from 6-minute limit

/**
 * Enhanced menu creation with icons
 */
function onOpen() {
  try {
    Logger.log('Menu', 'Adding custom menu to spreadsheet');
    SpreadsheetApp.getUi()
      .createMenu(`${UI.ICONS.EMAIL} Email Manager`)
      .addItem(`${UI.ICONS.EMAIL} Send Emails...`, 'showEmailDialog')
      .addItem(`${UI.ICONS.REFRESH} Check Responses...`, 'showResponseDialog')
      .addToUi();
    Logger.log('Menu', 'Custom menu added successfully');
  } catch (error) {
    Logger.error('Menu', error);
  }
}

/**
 * Shows dialog for email processing options
 */
function showEmailDialog() {
  const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  // Analyze rows - only include rows with "Send Email" action
  const pendingRows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLS.ACTION] === 'Send Email') {
      pendingRows.push({
        row: i + 1,
        name: data[i][COLS.FIRST_NAME],
        email: data[i][COLS.EMAIL]
      });
    }
  }

  if (pendingRows.length === 0) {
    ui.alert('No Pending Emails',
      'There are no new emails to send. Rows marked as "Awaiting Response" are skipped.',
      ui.ButtonSet.OK);
    return;
  }

  // Create HTML for dialog
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <style>
        body {
          font-family: 'Google Sans', Roboto, Arial, sans-serif;
          padding: 24px;
          margin: 0;
          color: #202124;
          background: #fff;
          line-height: 1.5;
        }

        h2 {
          margin: 0 0 16px;
          font-size: 20px;
          font-weight: 500;
          color: #1a73e8;
        }

        .pending-row {
          margin: 8px 0;
          padding: 12px 16px;
          background: #f8f9fa;
          border-radius: 8px;
          border: 1px solid #dadce0;
          transition: all 0.2s ease;
        }

        .pending-row:hover {
          background: #f1f3f4;
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        .button-container {
          margin-top: 24px;
          display: flex;
          gap: 8px;
          flex-wrap: wrap;
        }

        button {
          padding: 8px 24px;
          border-radius: 24px;
          border: none;
          font-family: inherit;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s ease;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        button.primary {
          background: #1a73e8;
          color: white;
        }

        button.primary:hover {
          background: #1557b0;
          box-shadow: 0 1px 3px rgba(0,0,0,0.2);
        }

        button.secondary {
          background: #fff;
          color: #1a73e8;
          border: 1px solid #dadce0;
        }

        button.secondary:hover {
          background: #f8f9fa;
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        .warning {
          margin-top: 16px;
          padding: 12px 16px;
          background: #fce8e6;
          color: #c5221f;
          border-radius: 8px;
          font-size: 14px;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .email-count {
          display: inline-flex;
          align-items: center;
          padding: 4px 12px;
          background: #e8f0fe;
          color: #1a73e8;
          border-radius: 16px;
          font-size: 14px;
          margin-bottom: 16px;
        }
      </style>

      <h2>${UI.ICONS.EMAIL} Email Processing Options</h2>

      <div class="email-count">
        ${UI.ICONS.EMAIL} ${pendingRows.length} pending email${pendingRows.length > 1 ? 's' : ''} to process
      </div>

      <div id="pendingEmails">
        ${pendingRows.map(row => `
          <div class="pending-row">
            <strong>Row ${row.row}:</strong> ${row.name}
            <br>
            <span style="color: #5f6368; font-size: 14px;">${row.email}</span>
          </div>
        `).join('')}
      </div>

      <div class="button-container">
        <button class="primary" onclick="processSelected('single')">
          ${UI.ICONS.EMAIL} Process First Row
        </button>
        <button class="primary" onclick="processSelected('all')">
          ${UI.ICONS.EMAIL} Process All Rows
        </button>
        <button class="secondary" onclick="google.script.host.close()">
          Cancel
        </button>
      </div>

      <div class="warning">
        ${UI.ICONS.WARNING} Processing all rows will send ${pendingRows.length} email${pendingRows.length > 1 ? 's' : ''}.
      </div>

      <script>
        function processSelected(mode) {
          // Disable buttons during processing
          document.querySelectorAll('button').forEach(btn => {
            btn.disabled = true;
            btn.style.opacity = '0.7';
            btn.style.cursor = 'wait';
          });

          google.script.run
            .withSuccessHandler(onSuccess)
            .withFailureHandler(onError)
            .processEmailsWithMode(mode);
        }

        function onSuccess(result) {
          const dialog = document.createElement('div');
          dialog.style = \`
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            padding: 24px;
            border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            text-align: center;
          \`;
          dialog.innerHTML = \`
            <div style="font-size: 48px; margin-bottom: 16px;">${UI.ICONS.SUCCESS}</div>
            <div style="margin-bottom: 16px;">\${result}</div>
            <button class="primary" onclick="google.script.host.close()">Close</button>
          \`;
          document.body.appendChild(dialog);
        }

        function onError(error) {
          const dialog = document.createElement('div');
          dialog.style = \`
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: white;
            padding: 24px;
            border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            text-align: center;
          \`;
          dialog.innerHTML = \`
            <div style="font-size: 48px; margin-bottom: 16px;">${UI.ICONS.ERROR}</div>
            <div style="color: #c5221f; margin-bottom: 16px;">Error: \${error.message}</div>
            <button class="primary" onclick="google.script.host.close()">Close</button>
          \`;
          document.body.appendChild(dialog);
        }
      </script>
    `)
    .setWidth(450)
    .setHeight(500)
    .setTitle('Email Processing Options');

  ui.showModalDialog(htmlOutput, 'Email Processing Options');
}

/**
 * Process emails based on selected mode
 */
function processEmailsWithMode(mode) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

  try {
    let processed = 0;

    // Start from row 1 (skipping header)
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS.ACTION] === 'Send Email') {
      processRow(sheet, data[i], i);
        processed++;

        if (mode === 'single') break; // Stop after first row if single mode
      }
    }

    return `Successfully processed ${processed} email(s)`;
  } catch (error) {
    Logger.error('ProcessEmails', error);
    throw error;
  }
}

/**
 * Shows dialog for checking email responses with preview
 */
function showResponseDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 16px;
          font-size: 14px;
        }
        .log-container {
          height: 500px;  /* Increased height */
          overflow-y: auto;
          padding: 16px;
          border: 1px solid #ccc;
          border-radius: 8px;
          font-family: 'Courier New', monospace;
          background: #f8f9fa;
          margin: 16px 0;
          font-size: 13px;
          line-height: 1.5;
        }
        .log-entry {
          padding: 8px 12px;
          margin: 4px 0;
          border-radius: 6px;
          white-space: pre-wrap;
          word-break: break-word;
        }
        .log-entry.info { background: #e8f0fe; color: #1a73e8; }
        .log-entry.success { background: #e6f4ea; color: #188038; }
        .log-entry.error { background: #fce8e6; color: #ea4335; }
        .log-entry.warning { background: #fef7e0; color: #ea8600; }
        #status {
          font-size: 16px;
          font-weight: 500;
          margin-bottom: 16px;
          padding: 12px;
          border-radius: 6px;
          background: #f1f3f4;
        }
        .button-container {
          display: flex;
          gap: 12px;
          margin-top: 16px;
        }
        .button {
          padding: 10px 20px;
          border-radius: 6px;
          border: none;
          cursor: pointer;
          font-size: 14px;
          display: flex;
          align-items: center;
          gap: 8px;
          font-weight: 500;
          transition: all 0.2s ease;
        }
        .button:hover {
          opacity: 0.9;
          transform: translateY(-1px);
        }
        .button:disabled {
          opacity: 0.6;
          cursor: not-allowed;
          transform: none;
        }
        .primary {
          background: #1a73e8;
          color: white;
        }
        .secondary {
          background: #f1f3f4;
          color: #1a73e8;
          border: 1px solid #dadce0;
        }
        .spinner {
          display: inline-block;
          width: 18px;
          height: 18px;
          border: 2px solid #ffffff;
          border-radius: 50%;
          border-top-color: transparent;
          animation: spin 1s linear infinite;
        }
        @keyframes spin {
          to {transform: rotate(360deg);}
        }
      </style>

      <h3>Check Email Responses</h3>
      <div>Status: <span id="status">Initializing...</span></div>
      <div id="log-container" class="log-container"></div>

      <div class="button-container">
        <button id="startBtn" class="button primary" onclick="startProcessing()">
          🔄 Check Responses
        </button>
        <button id="cancelBtn" class="button secondary" onclick="cancelProcessing()" disabled>
          🛑 Cancel
        </button>
        <button id="refreshBtn" class="button secondary" onclick="refreshUI()">
          🔄 Refresh UI
        </button>
      </div>

      <script>
        let isProcessing = false;
        let logCheckInterval;
        let processStarted = false;

        // Create async wrapper for google.script.run
        const createAsyncFunction = (name) => {
          return (...args) => {
            return new Promise((resolve, reject) => {
              google.script.run
                .withSuccessHandler(resolve)
                .withFailureHandler(reject)
                [name](...args);
            });
          };
        };

        const serverFunctions = {
          processEmailResponses: createAsyncFunction('processEmailResponses'),
          getProcessLogs: createAsyncFunction('getProcessLogs'),
          clearProcessLogs: createAsyncFunction('clearProcessLogs'),
          cancelResponseScan: createAsyncFunction('cancelResponseScan'),
          forceStopProcess: createAsyncFunction('forceStopProcess')
        };

        function addLog(message, type = 'info') {
          const container = document.getElementById('log-container');
          const entry = document.createElement('div');
          entry.className = 'log-entry ' + type;
          entry.textContent = new Date().toLocaleTimeString() + ': ' + message;
          container.appendChild(entry);
          container.scrollTop = container.scrollHeight;
        }

        async function checkForNewLogs() {
          if (!isProcessing) return;

          try {
            const logs = await serverFunctions.getProcessLogs();
            if (logs && logs.length > 0) {
              logs.forEach(log => {
                addLog(log.message, log.type);
              });
              await serverFunctions.clearProcessLogs();
            }
          } catch (error) {
            console.error('Error checking logs:', error);
          }
        }

        function startLogChecking() {
          stopLogChecking();
          logCheckInterval = setInterval(checkForNewLogs, 500);
        }

        function stopLogChecking() {
          if (logCheckInterval) {
            clearInterval(logCheckInterval);
            logCheckInterval = null;
          }
        }

        async function startProcessing() {
          if (processStarted) return;
          processStarted = true;

          try {
            setProcessing(true);
            document.getElementById('status').textContent = 'Processing...';

            await serverFunctions.forceStopProcess();

            await new Promise(resolve => setTimeout(resolve, 500));

            startLogChecking();

            const result = await serverFunctions.processEmailResponses();

            if (result.error) {
              addLog(result.error, 'error');
              document.getElementById('status').textContent = 'Error';
            } else {
              const message = result.message ||
                'Processed ' + result.processedCount + ' of ' + result.totalCount + ' rows';
              addLog(message, result.partial ? 'warning' : 'success');
              document.getElementById('status').textContent =
                result.partial ? 'Partially Complete' : 'Complete';
            }
          } catch (error) {
            addLog('Error: ' + error.message, 'error');
            document.getElementById('status').textContent = 'Error';
          } finally {
            setProcessing(false);
            processStarted = false;
            setTimeout(stopLogChecking, 1000);
          }
        }

        function setProcessing(processing) {
          isProcessing = processing;
          document.getElementById('startBtn').disabled = processing;
          document.getElementById('refreshBtn').disabled = processing;
          document.getElementById('cancelBtn').disabled = !processing;
          document.getElementById('startBtn').innerHTML =
            processing ? '<div class="spinner"></div> Processing...' : '🔄 Check Responses';
        }

        async function cancelProcessing() {
          try {
            addLog('Cancelling process...', 'warning');
            await serverFunctions.cancelResponseScan();
            addLog('Process cancelled', 'info');
            document.getElementById('status').textContent = 'Cancelled';
          } catch (error) {
            addLog('Error cancelling: ' + error.message, 'error');
          } finally {
            setProcessing(false);
            processStarted = false;
            stopLogChecking();
          }
        }

        function refreshUI() {
          document.getElementById('log-container').innerHTML = '';
          document.getElementById('status').textContent = 'Ready';
          addLog('UI refreshed', 'info');
        }
      </script>
    `)
    .setWidth(800)
    .setHeight(700)
    .setTitle('Check Email Responses');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Check Email Responses');
}

// Add these helper functions
function getProcessLogs() {
  const logsJson = PropertiesService.getScriptProperties().getProperty('PROCESS_LOGS');
  return logsJson ? JSON.parse(logsJson) : [];
}

function clearProcessLogs() {
  PropertiesService.getScriptProperties().deleteProperty('PROCESS_LOGS');
}

// Modify processEmailResponses to use the logging callback
function processEmailResponses() {
  // First check if already processing
  if (PropertiesService.getScriptProperties().getProperty('IS_PROCESSING') === 'true') {
    return {
      success: false,
      error: 'Another process is already running'
    };
  }

  const startTime = Date.now();
  const isTimeLeft = () => MAX_EXECUTION_TIME > Date.now() - startTime;
  let processedEmails = new Set(); // Track processed emails to avoid duplicates

  try {
    // Set processing flag
    PropertiesService.getScriptProperties().setProperty('IS_PROCESSING', 'true');

    // Reset cancel flag and logs
    PropertiesService.getScriptProperties().deleteProperty('CANCEL_SCAN');
    PropertiesService.getScriptProperties().deleteProperty('PROCESS_LOGS');

    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    const log = (message, type = 'info') => {
      console.log(message);
      const logs = JSON.parse(PropertiesService.getScriptProperties().getProperty('PROCESS_LOGS') || '[]');
      logs.push({ message, type, timestamp: new Date().toISOString() });
      PropertiesService.getScriptProperties().setProperty('PROCESS_LOGS', JSON.stringify(logs));
    };

    log('Starting email response check...');

    // Get unique rows awaiting response
    const awaitingRows = data.reduce((acc, row, index) => {
      if (index > 0 &&
          row[COLS.ACTION] === 'Awaiting Response' &&
          !processedEmails.has(row[COLS.EMAIL])) {
        processedEmails.add(row[COLS.EMAIL]);
        acc.push({
          rowIndex: index,
          name: row[COLS.FIRST_NAME],
          email: row[COLS.EMAIL],
          sentDate: row[COLS.SENT_DATE]
        });
      }
      return acc;
    }, []);

    log(`Found ${awaitingRows.length} unique rows awaiting response`);

    if (awaitingRows.length === 0) {
      PropertiesService.getScriptProperties().deleteProperty('IS_PROCESSING');
      return {
        success: true,
        processedCount: 0,
        totalCount: 0,
        message: 'No rows found with "Awaiting Response" status'
      };
    }

    let processedCount = 0;

    // Process each row sequentially
    for (const row of awaitingRows) {
      // Check for cancellation
      if (PropertiesService.getScriptProperties().getProperty('CANCEL_SCAN') === 'true') {
        log('Operation cancelled by user', 'warning');
        PropertiesService.getScriptProperties().deleteProperty('IS_PROCESSING');
        return {
          success: false,
          cancelled: true,
          processedCount: processedCount,
          totalCount: awaitingRows.length,
          message: `Operation cancelled. Processed ${processedCount} of ${awaitingRows.length} rows.`
        };
      }

      try {
        log(`Processing ${row.name} (${row.email})`);

        // Process emails for this row
        const [datePart, timePart] = row.sentDate.split(' at ');
        const sentDate = new Date(`${datePart} ${timePart}`);
        const searchQuery = `(from:${row.email} OR to:${row.email}) after:${sentDate.toISOString().split('T')[0]}`;

        log(`Searching for emails: ${row.email}`);
        const threads = GmailApp.search(searchQuery);

        // Collect responses
        const responses = [];
        for (const thread of threads) {
          const messages = thread.getMessages();
          for (const message of messages) {
            if (message.getDate() > sentDate) {
              responses.push({
                date: Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a"),
                subject: message.getSubject(),
                from: message.getFrom(),
                to: message.getTo(),
                body: message.getPlainBody().substring(0, 200) + '...',
                link: message.getThread().getPermalink()
              });
            }
          }
        }

        // Sort and write responses
        if (responses.length > 0) {
          responses.sort((a, b) => new Date(a.date) - new Date(b.date));
          log(`Found ${responses.length} responses for ${row.name}`);

          // Clear existing responses
          const lastColumn = sheet.getLastColumn();
          if (lastColumn > COLS.RESPONSES_START) {
            sheet.getRange(row.rowIndex + 1, COLS.RESPONSES_START + 1, 1, lastColumn - COLS.RESPONSES_START)
              .clearContent();
          }

          // Write responses
          responses.forEach((response, index) => {
            const formattedResponse = `
📅 ${response.date}
👤 From: ${response.from}
👥 To: ${response.to}
📧 ${response.subject}
💬 ${response.body}
🔗 ${response.link}`.trim();

            sheet.getRange(row.rowIndex + 1, COLS.RESPONSES_START + index + 1)
              .setValue(formattedResponse);
          });
        }

        processedCount++;
        log(`Completed processing ${row.name}`, 'success');

        // Small delay between rows
        Utilities.sleep(100);

      } catch (error) {
        log(`Error processing ${row.name}: ${error.message}`, 'error');
        console.error(`Error processing row ${row.rowIndex}:`, error);
      }
    }

    const finalMessage = `Successfully processed ${processedCount} rows`;
    log(finalMessage, 'success');

    PropertiesService.getScriptProperties().deleteProperty('IS_PROCESSING');
    return {
      success: true,
      processedCount: processedCount,
      totalCount: awaitingRows.length,
      complete: true,
      message: finalMessage
    };

  } catch (error) {
    const errorMessage = `Process error: ${error.message}`;
    console.error(errorMessage);
    log(errorMessage, 'error');

    PropertiesService.getScriptProperties().deleteProperty('IS_PROCESSING');
    return {
      success: false,
      error: error.message,
      processedCount: processedCount || 0,
      totalCount: awaitingRows?.length || 0
    };
  }
}

// Helper function to log to the client
function logToClient(text, type = 'info') {
  try {
    return google.script.run
      .withSuccessHandler(() => {})
      .withFailureHandler(() => {})
      .logProgress({ text, type });
  } catch (error) {
    console.error('Logging error:', error);
  }
}

// Function to receive logs from server
function logProgress(message) {
  // This function exists to receive logs from the server
  // The client-side code will handle displaying the message
  return message;
}

// Simple test function to verify callback works
function testSimpleCallback(callback) {
  try {
    callback('Checking permissions...');

    // Verify permissions
    const token = ScriptApp.getOAuthToken();
    callback('Permissions verified');
    Utilities.sleep(1000);

    callback('Reading spreadsheet...');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    callback('Found sheet: ' + sheet.getName());
    Utilities.sleep(1000);

    callback('Testing Gmail access...');
    const unread = GmailApp.getInboxUnreadCount();
    callback('Gmail access confirmed');
    Utilities.sleep(1000);

    callback('Test complete!');
    return { success: true };

  } catch (error) {
    console.error('Test error:', error);
    callback('Error: ' + error.message);
    return { error: error.message };
  }
}

/**
 * Cancel the response scan operation
 */
function cancelResponseScan() {
  try {
    PropertiesService.getScriptProperties().setProperty('CANCEL_SCAN', 'true');
    return { success: true, message: 'Operation cancelled' };
  } catch (error) {
    console.error('Cancel error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Preview responses with progress updates
 */
function previewResponsesWithProgress(callback) {
  Logger.log('ResponseCheck', 'Starting response check process');

  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    Logger.log('ResponseCheck', `Loaded ${data.length - 1} rows from spreadsheet`);
    callback({ log: 'Scanning spreadsheet for awaiting responses...' });

    const awaitingRows = [];

    // First identify all rows awaiting response
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS.ACTION] === 'Awaiting Response') {
        const name = data[i][COLS.FIRST_NAME];
        const email = data[i][COLS.EMAIL];
        const sentDateStr = data[i][COLS.SENT_DATE];

        Logger.log('ResponseCheck', `Found row ${i + 1}: ${name} (${email}) sent on ${sentDateStr}`);
        awaitingRows.push({ rowIndex: i, name, email, sentDateStr });
      }
    }

    Logger.log('ResponseCheck', `Found ${awaitingRows.length} rows awaiting response`);

    if (awaitingRows.length === 0) {
      callback({ log: 'No rows found with "Awaiting Response" status' });
      return { error: 'No rows found with "Awaiting Response" status' };
    }

    callback({ log: `Found ${awaitingRows.length} emails to check` });

    const results = [];

    // Process each awaiting row
    for (const row of awaitingRows) {
      try {
        callback({ log: `Checking responses from ${row.email}...` });

        // Parse the sent date
        const [datePart, timePart] = row.sentDateStr.split(' at ');
        const sentDate = new Date(`${datePart} ${timePart}`);

        // Search for emails received after sent date
        const searchQuery = `from:${row.email} after:${sentDate.toISOString().split('T')[0]}`;
        Logger.log('ResponseCheck', `Searching with query: ${searchQuery}`);

        const threads = GmailApp.search(searchQuery);
        let responseCount = 0;
        let responseDetails = [];

        // Process each thread
        for (const thread of threads) {
          const messages = thread.getMessages();
          for (const message of messages) {
            if (message.getDate() > sentDate) {
              responseCount++;
              responseDetails.push({
                date: Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a"),
                subject: message.getSubject(),
                link: message.getThread().getPermalink()
              });
            }
          }
        }

        Logger.log('ResponseCheck', `Found ${responseCount} responses from ${row.name}`);
        callback({
          log: `Found ${responseCount} response(s) from ${row.name}`,
          type: responseCount > 0 ? 'success' : 'info'
        });

        results.push({
          rowNum: row.rowIndex + 1,
          name: row.name,
          email: row.email,
          count: responseCount,
          responses: responseDetails,
          sentDate: row.sentDateStr
        });

      } catch (error) {
        Logger.error('ResponseCheck', error, { row });
        callback({
          log: `Error checking ${row.email}: ${error.message}`,
          type: 'error'
        });
      }
    }

    Logger.log('ResponseCheck', 'Completed checking all rows');
    return { rows: results };

  } catch (error) {
    Logger.error('ResponseCheck', error);
    callback({
      log: `Error: ${error.message}`,
      type: 'error'
    });
    return { error: error.message };
  }
}

/**
 * Format email information in a more structured way
 */
function formatEmailInfo(message) {
  try {
    const messageDate = Utilities.formatDate(
      message.getDate(),
      Session.getScriptTimeZone(),
      "MMMM d, yyyy 'at' h:mm a"
    );
    const subject = message.getSubject();
    const body = message.getPlainBody();
    const snippet = body.length > 200 ? body.substring(0, 200) + '...' : body;
    const emailLink = message.getThread().getPermalink();

    return {
      date: messageDate,
      subject: subject,
      body: snippet,
      link: emailLink
    };
  } catch (error) {
    Logger.error('FormatEmail', error, { messageId: message.getId() });
    return null;
  }
}

/**
 * Format responses for spreadsheet columns
 */
function formatResponsesForSheet(responses) {
  if (typeof responses === 'string' && responses.startsWith('Error:')) {
    return [responses];
  }

  const formattedResponses = [];
  const responseArray = responses.split('\n\n-------------------\n\n');

  responseArray.forEach(response => {
    if (response.trim()) {
      const formattedResponse = `
📅 ${response.match(/Date: (.*)/)[1]}

📧 ${response.match(/Subject: (.*)/)[1]}

💬 ${response.match(/Body: (.*)/)[1]}

🔗 ${response.match(/Link: (.*)/)[1]}
`.trim();

      formattedResponses.push(formattedResponse);
    }
  });

  return formattedResponses;
}

/**
 * Custom function to check received emails after a given date
 */
function RECEIVEDEMAILSAFTERDATE(email, dateTime) {
  if (SCRIPT_PROPERTIES.getProperty('GMAIL_AUTHORIZED') !== 'true') {
    return 'Please run the authorizeGmail function first from the Apps Script editor';
  }

  Logger.log('CheckEmails', `Checking emails from ${email} after ${dateTime}`);

  try {
    if (!email || !dateTime) {
      throw new Error('Email and date are required');
    }

    const filterDate = new Date(dateTime);
    if (isNaN(filterDate.getTime())) {
      throw new Error('Invalid date format');
    }

    // Search for emails
    const searchQuery = `from:${email} after:${filterDate.toISOString().split('T')[0]}`;
    const threads = GmailApp.search(searchQuery);

    Logger.log('CheckEmails', `Found ${threads.length} matching threads`);

    // Process threads
    const emailsInfo = threads.flatMap(thread => {
      return thread.getMessages()
        .filter(message => message.getDate() > filterDate)
        .map(message => formatEmailInfo(message))
        .filter(info => info !== null);
    });

    // Format the email information
    return emailsInfo.map(info => [
      `Date: ${info.date}`,
      `Subject: ${info.subject}`,
      `Body: ${info.body}`,
      `Link: ${info.link}`
    ].join('\n')).join('\n\n-------------------\n\n');

  } catch (error) {
    Logger.error('CheckEmails', error, { email, dateTime });
    return `Error: ${error.message}`;
  }
}

// Process individual row
function processRow(sheet, row, rowIndex) {
  const action = row[COLS.ACTION];

  if (action !== 'Send Email') {
    return; // Skip if no action needed or already sent
  }

  const firstName = row[COLS.FIRST_NAME];
  const email = row[COLS.EMAIL];

  Logger.log('ProcessRow', `Processing row ${rowIndex + 1}`, { firstName, email });

  try {
    // Validate inputs
    if (!email || !email.includes('@')) {
      throw new Error('Invalid email address');
    }
    if (!firstName) {
      throw new Error('First name is required');
    }

    // Send email and get message ID
    const messageId = sendEmailAndGetId(email, firstName);

    // Update spreadsheet with success status
    updateRowStatus(sheet, rowIndex, messageId);

    // Update action column to "Awaiting Response"
    sheet.getRange(rowIndex + 1, COLS.ACTION + 1).setValue('Awaiting Response');

    Logger.log('ProcessRow', `Successfully processed row ${rowIndex + 1}`);
  } catch (error) {
    Logger.error('ProcessRow', error, { rowIndex, firstName, email });
    const errorCell = sheet.getRange(rowIndex + 1, COLS.SENT_DATE + 1);
    errorCell.setValue(`Error: ${error.message}`);
  }
}

// Send email and return message ID
function sendEmailAndGetId(email, firstName) {
  Logger.log('SendEmail', `Sending email to ${email}`);

  try {
    const emailBody = getEmailBody(firstName);
    const options = {
      name: CONFIG.SENDER_NAME,
      from: Session.getActiveUser().getEmail(),
      htmlBody: emailBody // Use HTML formatting
    };

    // Send email
    GmailApp.sendEmail(email, CONFIG.EMAIL_SUBJECT, '', options); // Empty string for plain text version

    // Search for the just-sent email
    Utilities.sleep(2000); // Wait for email to be processed
    const searchQuery = `to:${email} subject:"${CONFIG.EMAIL_SUBJECT}" newer_than:1d`;
    const threads = GmailApp.search(searchQuery, 0, 1);

    if (threads.length === 0) {
      throw new Error('Sent email thread not found');
    }

    const messageId = threads[0].getMessages()[0].getId();
    Logger.log('SendEmail', `Email sent successfully`, { messageId });
    return messageId;

  } catch (error) {
    Logger.error('SendEmail', error, { email, firstName });
    throw error;
  }
}

// Update row status after successful send
function updateRowStatus(sheet, rowIndex, messageId) {
  try {
    // Format date in a human-readable way
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a");

    // Get direct link to the email message
    const email = GmailApp.getMessageById(messageId);
    const emailLink = email.getThread().getPermalink();

    // Update status in separate columns
    sheet.getRange(rowIndex + 1, COLS.SENT_DATE + 1).setValue(formattedDate);
    sheet.getRange(rowIndex + 1, COLS.EMAIL_LINK + 1).setValue(emailLink);

    Logger.log('UpdateStatus', `Updated status for row ${rowIndex + 1}`);
  } catch (error) {
    Logger.error('UpdateStatus', error, { rowIndex, messageId });
    throw error;
  }
}

/**
 * Function to authorize Gmail access
 */
function authorizeGmail() {
  try {
    // Get Gmail service
    const gmailService = GmailApp.getInboxUnreadCount(); // This will trigger auth

    // Store authorization status
    SCRIPT_PROPERTIES.setProperty('GMAIL_AUTHORIZED', 'true');

    return 'Gmail authorization successful';
  } catch (error) {
    Logger.error('Auth', error);
    return 'Failed to authorize Gmail: ' + error.message;
  }
}

// Email template function with HTML formatting
function getEmailBody(recipientName) {
  Logger.log('Template', `Generating email for ${recipientName}`);

  return `
    <div style="font-family: Arial, sans-serif; line-height: 1.6;">
      Hello ${recipientName},<br><br>

      I hope this message finds you well. My name is <b>[Your Name]</b>, and I am reaching out regarding [Your Purpose].
      I would greatly appreciate your time in [Your Request].<br><br>

      I would be happy to discuss how we can collaborate on [Topic]. Could you share your thoughts on this?<br><br>

      For your reference, here's an example document:
      <a href="#">[Reference Document]</a><br><br>

      Here are my key areas of focus:<br><br>

      <ul style="list-style-type: none; padding-left: 0;">
        <li style="margin-bottom: 12px;">• <b>Area 1:</b> Description of first area of expertise</li>
        <li style="margin-bottom: 12px;">• <b>Area 2:</b> Description of second area of expertise</li>
        <li style="margin-bottom: 12px;">• <b>Area 3:</b> Description of third area of expertise</li>
        <li style="margin-bottom: 12px;">• <b>Area 4:</b> Description of fourth area of expertise</li>
      </ul><br>

      Thank you for your time and consideration.<br><br>

      Best regards,<br>
      [Your Name]
    </div>
  `;
}

// Test function with generic data
function testReceivedEmailsFunction() {
  Logger.log('Test', 'Starting test of RECEIVEDEMAILSAFTERDATE');

  // Test data with generic values
  const testEmail = 'test.user@example.com';
  const testDate = new Date('2024-02-22');

  try {
    // Run the function
    const result = RECEIVEDEMAILSAFTERDATE(testEmail, testDate);

    // Log results
    Logger.log('Test', 'Function returned:', result);

    // Basic validation
    if (result.includes('Error:')) {
      Logger.error('Test', 'Function returned error:', result);
    } else {
      Logger.log('Test', 'Test completed successfully');
    }

  } catch (error) {
    Logger.error('Test', error);
  }
}

// Optional: Add a function to reset status if needed
function resetEmailStatus(rowIndex) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);

  // Clear sent date and email link
  sheet.getRange(rowIndex + 1, COLS.SENT_DATE + 1).clearContent();
  sheet.getRange(rowIndex + 1, COLS.EMAIL_LINK + 1).clearContent();

  // Reset action to "Send Email"
  sheet.getRange(rowIndex + 1, COLS.ACTION + 1).setValue('Send Email');
}

function checkAllResponsesWithProgress(callback) {
  try {
    // Initial callback to confirm function is running
    callback({
      log: 'Starting response check...',
      type: 'info'
    });

    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    callback({
      log: `Loaded spreadsheet with ${data.length - 1} rows`,
      type: 'info'
    });

    // Count "Awaiting Response" rows first
    let awaitingCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS.ACTION] === 'Awaiting Response') {
        awaitingCount++;
      }
    }

    callback({
      log: `Found ${awaitingCount} rows with "Awaiting Response" status`,
      type: 'info',
      current: 0,
      total: awaitingCount
    });

    if (awaitingCount === 0) {
      return { error: 'No rows found with "Awaiting Response" status' };
    }

    let processedCount = 0;

    // Process each row
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS.ACTION] === 'Awaiting Response') {
        const name = data[i][COLS.FIRST_NAME];
        const email = data[i][COLS.EMAIL];

        callback({
          log: `Processing ${name} (${email})`,
          type: 'info',
          current: processedCount + 1,
          total: awaitingCount
        });

        // ... rest of your email processing code ...

        processedCount++;
      }
    }

    callback({
      log: 'Completed processing all rows',
      type: 'success',
      current: awaitingCount,
      total: awaitingCount
    });

    return { success: true, processedCount };

  } catch (error) {
    Logger.error('CheckResponses', error);
    callback({
      log: `Error: ${error.message}`,
      type: 'error'
    });
    return { error: error.message };
  }
}

/**
 * Test function to simulate email response checking
 */
function testEmailResponseCheck() {
  console.log('Starting email response check test');

  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    console.log(`Loaded spreadsheet with ${data.length - 1} rows`);

    // Count "Awaiting Response" rows
    let awaitingRows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS.ACTION] === 'Awaiting Response') {
        awaitingRows.push({
          row: i + 1,
          name: data[i][COLS.FIRST_NAME],
          email: data[i][COLS.EMAIL],
          sentDate: data[i][COLS.SENT_DATE]
        });
      }
    }

    console.log(`Found ${awaitingRows.length} rows with "Awaiting Response" status`);
    console.log('Awaiting rows details:');
    awaitingRows.forEach(row => {
      console.log(`Row ${row.row}: ${row.name} (${row.email}) - Sent on ${row.sentDate}`);
    });

    // Process each awaiting row
    for (const row of awaitingRows) {
      console.log(`\nProcessing row ${row.row} - ${row.name}`);

      try {
        // Parse the sent date
        const [datePart, timePart] = row.sentDate.split(' at ');
        const sentDate = new Date(`${datePart} ${timePart}`);
        console.log(`Parsed date: ${sentDate.toISOString()}`);

        // Search for emails
        const searchQuery = `from:${row.email} after:${sentDate.toISOString().split('T')[0]}`;
        console.log(`Search query: ${searchQuery}`);

        const threads = GmailApp.search(searchQuery);
        console.log(`Found ${threads.length} email threads`);

        // Process threads
        let responseCount = 0;
        let responses = [];

        for (const thread of threads) {
          const messages = thread.getMessages();
          for (const message of messages) {
            if (message.getDate() > sentDate) {
              responseCount++;
              responses.push({
                date: Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a"),
                subject: message.getSubject(),
                snippet: message.getPlainBody().substring(0, 100) + '...',
                link: message.getThread().getPermalink()
              });
            }
          }
        }

        console.log(`Found ${responseCount} responses for ${row.name}`);
        if (responseCount > 0) {
          console.log('Response details:');
          responses.forEach((response, index) => {
            console.log(`\nResponse ${index + 1}:`);
            console.log(`Date: ${response.date}`);
            console.log(`Subject: ${response.subject}`);
            console.log(`Preview: ${response.snippet}`);
            console.log(`Link: ${response.link}`);
          });
        }

      } catch (error) {
        console.error(`Error processing row ${row.row}:`, error.message);
      }
    }

    console.log('\nEmail response check completed');

  } catch (error) {
    console.error('Test error:', error.message);
    console.error('Stack:', error.stack);
  }
}

// Add these scopes at the top of your script
function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

// Create a function to set up triggers and permissions
function setup() {
  // This will prompt for permissions
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  GmailApp.getInboxUnreadCount(); // This ensures Gmail permission

  console.log('Setup complete - permissions granted');
}

function showTestDialog() {
  // First verify we have permissions
  try {
    const token = ScriptApp.getOAuthToken();
    console.log('Have OAuth token, proceeding...');
  } catch (e) {
    console.error('Missing permissions - please run setup() first');
    return;
  }

  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <style>
        body { font-family: Arial, sans-serif; padding: 16px; }
        .log-container {
          height: 300px;
          overflow-y: auto;
          padding: 12px;
          border: 1px solid #ccc;
          font-family: monospace;
          background: #f8f9fa;
          margin-top: 16px;
        }
        .log-entry {
          padding: 4px 8px;
          margin: 2px 0;
          border-radius: 4px;
        }
        .log-entry.info { background: #e8f0fe; color: #1a73e8; }
        .log-entry.success { background: #e6f4ea; color: #188038; }
        .log-entry.error { background: #fce8e6; color: #ea4335; }
      </style>

      <h3>Test Dialog</h3>
      <div>Status: <span id="status">Initializing...</span></div>
      <div id="log-container" class="log-container"></div>

      <script>
        let logCount = 0;

        function addLog(message, type = 'info') {
          logCount++;
          const container = document.getElementById('log-container');
          const entry = document.createElement('div');
          entry.className = 'log-entry ' + type;
          entry.textContent = logCount + '. ' + new Date().toLocaleTimeString() + ': ' + message;
          container.appendChild(entry);
          container.scrollTop = container.scrollHeight;

          // Also update status
          document.getElementById('status').textContent = message;
        }

        // Initial log
        addLog('Dialog opened, starting test...');

        // Run the test function
        google.script.run
          .withSuccessHandler(function(result) {
            addLog('Test completed successfully!', 'success');
            document.getElementById('status').textContent = 'Complete';
          })
          .withFailureHandler(function(error) {
            addLog('Error: ' + error.message, 'error');
            document.getElementById('status').textContent = 'Error';
          })
          .testSimpleCallback(function(message) {
            addLog(message);
          });
      </script>
    `)
    .setWidth(500)
    .setHeight(400)
    .setTitle('Test Dialog');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Test Dialog');
}

function continueProcessing(startFromRow) {
  // Similar to processEmailResponses but starts from a specific row
  // Can be implemented if needed
}

// Add this helper function to force stop any running process
function forceStopProcess() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('IS_PROCESSING');
    PropertiesService.getScriptProperties().deleteProperty('CANCEL_SCAN');
    PropertiesService.getScriptProperties().deleteProperty('PROCESS_LOGS');
    return { success: true };
  } catch (error) {
    console.error('Force stop error:', error);
    return { success: false, error: error.message };
  }
}


