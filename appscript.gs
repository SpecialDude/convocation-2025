/**
 * CONVOCATION RSVP EMAIL PROCESSOR
 * Automatically processes RSVP emails and adds them to Google Sheets
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any existing code and paste this entire script
 * 4. Update CONFIGURATION section below
 * 5. Save the project (Ctrl+S or Cmd+S)
 * 6. Run 'setupEmailProcessor' function once to authorize
 * 7. Set up time-based trigger for 'processRSVPEmails'
 */

// ========================================
// CONFIGURATION - UPDATE THESE VALUES
// ========================================

const CONFIG = {
  // Your email address where RSVPs are sent
  targetEmail: 'your-email@gmail.com',
  
  // Gmail label to organize processed emails (will be created automatically)
  labelName: 'Convocation-RSVP',
  
  // Subject keywords to identify RSVP emails (case-insensitive)
  subjectKeywords: ['RSVP', 'Convocation', 'Attendance', 'Graduation'],
  
  // Sheet name where data will be written
  sheetName: 'Sheet1',
  
  // Column mapping (A=1, B=2, C=3, etc.)
  columns: {
    name: 1,      // Column A
    email: 2,     // Column B
    celebrating: 3, // Column C
    notes: 4,     // Column D
    timestamp: 5   // Column E
  },
  
  // How far back to check for emails (in days)
  daysToCheck: 7,
  
  // Maximum emails to process per run (to avoid timeout)
  maxEmailsPerRun: 50
};

// ========================================
// MAIN PROCESSING FUNCTION
// ========================================

/**
 * Main function to process RSVP emails
 * This should be set up to run automatically via time-based trigger
 */
function processRSVPEmails() {
  try {
    Logger.log('=== Starting RSVP Email Processing ===');
    
    // Get or create the label
    const label = getOrCreateLabel(CONFIG.labelName);
    const processedLabel = getOrCreateLabel(CONFIG.labelName + '-Processed');
    
    // Search for unprocessed RSVP emails
    const searchQuery = buildSearchQuery();
    Logger.log('Search query: ' + searchQuery);
    
    const threads = GmailApp.search(searchQuery, 0, CONFIG.maxEmailsPerRun);
    Logger.log('Found ' + threads.length + ' email threads to process');
    
    if (threads.length === 0) {
      Logger.log('No new RSVP emails found');
      return;
    }
    
    // Get the spreadsheet
    const sheet = getActiveSheet();
    
    // Process each thread
    let processedCount = 0;
    let errorCount = 0;
    
    threads.forEach(function(thread) {
      try {
        const messages = thread.getMessages();
        
        messages.forEach(function(message) {
          try {
            // Extract data from email
            const rsvpData = extractRSVPData(message);
            
            if (rsvpData) {
              // Add to spreadsheet
              addRowToSheet(sheet, rsvpData);
              processedCount++;
              Logger.log('✓ Processed: ' + rsvpData.name);
            } else {
              Logger.log('⚠ Could not extract data from email: ' + message.getSubject());
              errorCount++;
            }
          } catch (e) {
            Logger.log('✗ Error processing message: ' + e.toString());
            errorCount++;
          }
        });
        
        // Mark thread as processed
        thread.addLabel(processedLabel);
        thread.removeLabel(label);
        thread.markRead();
        
      } catch (e) {
        Logger.log('✗ Error processing thread: ' + e.toString());
        errorCount++;
      }
    });
    
    Logger.log('=== Processing Complete ===');
    Logger.log('Successfully processed: ' + processedCount);
    Logger.log('Errors: ' + errorCount);
    
    // Send summary email if configured
    if (processedCount > 0) {
      sendSummaryEmail(processedCount, errorCount);
    }
    
  } catch (e) {
    Logger.log('CRITICAL ERROR: ' + e.toString());
    sendErrorEmail(e);
  }
}

// ========================================
// EMAIL PARSING FUNCTIONS
// ========================================

/**
 * Extract RSVP data from email message
 */
function extractRSVPData(message) {
  const body = message.getPlainBody();
  const subject = message.getSubject();
  const from = message.getFrom();
  const date = message.getDate();
  
  // Try different parsing strategies
  let rsvpData = null;
  
  // Strategy 1: Structured format (Form-style emails)
  rsvpData = parseStructuredEmail(body, from, date);
  
  // Strategy 2: Natural language format
  if (!rsvpData || !rsvpData.name) {
    rsvpData = parseNaturalLanguageEmail(body, subject, from, date);
  }
  
  // Strategy 3: Web3Forms format
  if (!rsvpData || !rsvpData.name) {
    rsvpData = parseWeb3FormsEmail(body, from, date);
  }
  
  return rsvpData;
}

/**
 * Parse structured email format (key: value pairs)
 * Example:
 * Name: John Doe
 * Email: john@example.com
 * Celebrating: Ada Okafor, Bola Ahmed
 * Notes: Vegetarian meal
 */
function parseStructuredEmail(body, from, date) {
  const data = {
    name: extractField(body, ['name', 'full name', 'guest name']),
    email: extractField(body, ['email', 'email address', 'e-mail']) || extractEmailFromString(from),
    celebrating: extractField(body, ['celebrating', 'attending for', 'graduate', 'graduates']),
    notes: extractField(body, ['notes', 'message', 'special requests', 'comments', 'dietary']),
    timestamp: date
  };
  
  return data.name ? data : null;
}

/**
 * Parse natural language email
 * Example: "Hi, I'm John Doe and I'll be attending to celebrate Ada Okafor..."
 */
function parseNaturalLanguageEmail(body, subject, from, date) {
  // Try to extract name from common patterns
  let name = null;
  
  // Pattern 1: "I'm [Name]" or "I am [Name]"
  let match = body.match(/I['']?m\s+([A-Z][a-z]+\s+[A-Z][a-z]+)/i);
  if (match) name = match[1];
  
  // Pattern 2: "My name is [Name]"
  if (!name) {
    match = body.match(/my\s+name\s+is\s+([A-Z][a-z]+\s+[A-Z][a-z]+)/i);
    if (match) name = match[1];
  }
  
  // Pattern 3: Email signature
  if (!name) {
    const lines = body.split('\n');
    // Check last few lines for name
    for (let i = Math.max(0, lines.length - 5); i < lines.length; i++) {
      const line = lines[i].trim();
      if (/^[A-Z][a-z]+\s+[A-Z][a-z]+$/.test(line)) {
        name = line;
        break;
      }
    }
  }
  
  // Extract celebrating
  let celebrating = null;
  const celebrantNames = ['Ada Okafor', 'Bola Ahmed', 'Chidi Nwankwo', 'Damilola Adeyemi', 'Emeka Obi'];
  const foundCelebrants = [];
  
  celebrantNames.forEach(function(celebrant) {
    if (body.toLowerCase().includes(celebrant.toLowerCase())) {
      foundCelebrants.push(celebrant);
    }
  });
  
  if (foundCelebrants.length > 0) {
    celebrating = foundCelebrants.join(', ');
  }
  
  return name ? {
    name: name,
    email: extractEmailFromString(from),
    celebrating: celebrating || 'Not specified',
    notes: 'Extracted from email body',
    timestamp: date
  } : null;
}

/**
 * Parse Web3Forms notification email format
 */
function parseWeb3FormsEmail(body, from, date) {
  // Web3Forms sends data in a specific format
  if (body.includes('New submission from') || body.includes('Web3Forms')) {
    return parseStructuredEmail(body, from, date);
  }
  return null;
}

/**
 * Extract field value using multiple possible field names
 */
function extractField(text, fieldNames) {
  for (let i = 0; i < fieldNames.length; i++) {
    const fieldName = fieldNames[i];
    
    // Try pattern: "Field Name: Value"
    let pattern = new RegExp(fieldName + '\\s*:?\\s*(.+?)(?:\\n|$)', 'i');
    let match = text.match(pattern);
    
    if (match && match[1]) {
      let value = match[1].trim();
      // Clean up common artifacts
      value = value.replace(/^[:\-\s]+/, '');
      value = value.replace(/[\r\n]+.*$/, ''); // Remove everything after newline
      if (value && value.length > 0 && value.length < 200) {
        return value;
      }
    }
  }
  return null;
}

/**
 * Extract email address from string
 */
function extractEmailFromString(str) {
  const match = str.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/);
  return match ? match[1] : str;
}

// ========================================
// SPREADSHEET FUNCTIONS
// ========================================

/**
 * Get the active sheet or create headers if needed
 */
function getActiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetName);
  
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }
  
  // Check if headers exist, if not create them
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.appendRow(['Name', 'Email', 'Celebrating', 'Notes', 'Timestamp']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Add RSVP data as a new row in the sheet
 */
function addRowToSheet(sheet, data) {
  // Check for duplicates
  if (isDuplicate(sheet, data.email, data.name)) {
    Logger.log('⚠ Skipping duplicate: ' + data.name + ' (' + data.email + ')');
    return;
  }
  
  const row = [
    data.name || 'Unknown',
    data.email || 'Not provided',
    data.celebrating || 'Not specified',
    data.notes || 'None',
    data.timestamp || new Date()
  ];
  
  sheet.appendRow(row);
}

/**
 * Check if this RSVP already exists in the sheet
 */
function isDuplicate(sheet, email, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false; // Only header row
  
  const emailColumn = CONFIG.columns.email;
  const nameColumn = CONFIG.columns.name;
  
  const emails = sheet.getRange(2, emailColumn, lastRow - 1, 1).getValues();
  const names = sheet.getRange(2, nameColumn, lastRow - 1, 1).getValues();
  
  for (let i = 0; i < emails.length; i++) {
    if (emails[i][0] === email || names[i][0] === name) {
      return true;
    }
  }
  
  return false;
}

// ========================================
// GMAIL LABEL FUNCTIONS
// ========================================

/**
 * Get existing label or create new one
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log('Created new label: ' + labelName);
  }
  
  return label;
}

/**
 * Build Gmail search query
 */
function buildSearchQuery() {
  const keywords = CONFIG.subjectKeywords.map(k => 'subject:' + k).join(' OR ');
  const dateLimit = new Date();
  dateLimit.setDate(dateLimit.getDate() - CONFIG.daysToCheck);
  const dateStr = Utilities.formatDate(dateLimit, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  return '(' + keywords + ') to:' + CONFIG.targetEmail + ' after:' + dateStr + ' -label:' + CONFIG.labelName + '-Processed';
}

// ========================================
// NOTIFICATION FUNCTIONS
// ========================================

/**
 * Send summary email after processing
 */
function sendSummaryEmail(successCount, errorCount) {
  const recipient = Session.getActiveUser().getEmail();
  const subject = '✓ RSVP Emails Processed: ' + successCount + ' new responses';
  const body = 'Convocation RSVP Processing Summary\n\n' +
               'Successfully processed: ' + successCount + '\n' +
               'Errors: ' + errorCount + '\n\n' +
               'View your spreadsheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  try {
    MailApp.sendEmail(recipient, subject, body);
  } catch (e) {
    Logger.log('Could not send summary email: ' + e.toString());
  }
}

/**
 * Send error notification
 */
function sendErrorEmail(error) {
  const recipient = Session.getActiveUser().getEmail();
  const subject = '✗ Error Processing RSVP Emails';
  const body = 'An error occurred while processing RSVP emails:\n\n' +
               error.toString() + '\n\n' +
               'Please check the Apps Script logs for details.';
  
  try {
    MailApp.sendEmail(recipient, subject, body);
  } catch (e) {
    Logger.log('Could not send error email: ' + e.toString());
  }
}

// ========================================
// SETUP FUNCTIONS
// ========================================

/**
 * One-time setup function - Run this first!
 * Creates necessary labels and authorizes the script
 */
function setupEmailProcessor() {
  Logger.log('=== Starting Setup ===');
  
  try {
    // Create labels
    getOrCreateLabel(CONFIG.labelName);
    getOrCreateLabel(CONFIG.labelName + '-Processed');
    Logger.log('✓ Labels created');
    
    // Verify sheet access
    const sheet = getActiveSheet();
    Logger.log('✓ Sheet access verified: ' + sheet.getName());
    
    // Test email search
    const query = buildSearchQuery();
    Logger.log('✓ Search query: ' + query);
    
    Logger.log('=== Setup Complete ===');
    Logger.log('Next steps:');
    Logger.log('1. Update CONFIG section with your email address');
    Logger.log('2. Set up time-based trigger for processRSVPEmails()');
    Logger.log('3. Send a test RSVP email to yourself');
    
  } catch (e) {
    Logger.log('✗ Setup error: ' + e.toString());
  }
}

/**
 * Create time-based trigger automatically
 * Run this after setupEmailProcessor() completes successfully
 */
function createTimeTrigger() {
  // Delete existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processRSVPEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger to run every 10 minutes
  ScriptApp.newTrigger('processRSVPEmails')
    .timeBased()
    .everyMinutes(10)
    .create();
  
  Logger.log('✓ Trigger created: processRSVPEmails will run every 10 minutes');
}

/**
 * Test function to verify email parsing
 * Processes the most recent RSVP email without adding to sheet
 */
function testEmailParsing() {
  Logger.log('=== Testing Email Parsing ===');
  
  const searchQuery = buildSearchQuery();
  const threads = GmailApp.search(searchQuery, 0, 1);
  
  if (threads.length === 0) {
    Logger.log('No RSVP emails found to test');
    return;
  }
  
  const message = threads[0].getMessages()[0];
  Logger.log('Testing email: ' + message.getSubject());
  Logger.log('From: ' + message.getFrom());
  Logger.log('Date: ' + message.getDate());
  Logger.log('---');
  
  const data = extractRSVPData(message);
  
  if (data) {
    Logger.log('✓ Successfully extracted:');
    Logger.log('  Name: ' + data.name);
    Logger.log('  Email: ' + data.email);
    Logger.log('  Celebrating: ' + data.celebrating);
    Logger.log('  Notes: ' + data.notes);
  } else {
    Logger.log('✗ Could not extract data');
    Logger.log('Email body:');
    Logger.log(message.getPlainBody());
  }
}

// ========================================
// MANUAL PROCESSING (For testing)
// ========================================

/**
 * Process a specific email by subject (for testing)
 */
function processSpecificEmail(subjectContains) {
  const threads = GmailApp.search('subject:' + subjectContains, 0, 1);
  
  if (threads.length === 0) {
    Logger.log('No email found with subject containing: ' + subjectContains);
    return;
  }
  
  const message = threads[0].getMessages()[0];
  const data = extractRSVPData(message);
  
  if (data) {
    const sheet = getActiveSheet();
    addRowToSheet(sheet, data);
    Logger.log('✓ Added to sheet: ' + data.name);
  } else {
    Logger.log('✗ Could not extract data from email');
  }
}