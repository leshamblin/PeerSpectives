
/**
 * --- CONFIGURATION ---
 */
const CONFIG = {
  // Set to true to send to ADMIN_EMAIL only. Set to false to send to actual team members.
  TEST_MODE: true,
  
  // The email address to receive the test emails
  ADMIN_EMAIL: 'INSERT_ADMIN_EMAIL_HERE',
  
  // The ID of the Google Drive folder containing the FINAL (processed) reviews to send
  REVIEWS_FOLDER_ID: 'INSERT_REVIEWS_FOLDER_ID_HERE',
  
  // The ID of the Google Sheet containing Team Member Roster
  // EXPECTED STRUCTURE: Must have columns for "Team" and "Email" (names configurable below)
  ROSTER_SHEET_ID: 'INSERT_ROSTER_SHEET_ID_HERE',
  ROSTER_TAB_NAME: 'Sheet1', // The tab name to read from
  TEAM_COL_HEADER: 'Team',   // Text in the header row for the Team column
  EMAIL_COL_HEADER: 'Email', // Text in the header row for the Email column
  
  // Optional: Filter to a specific team name (e.g., "Team A") or leave empty to process all or use specific functions
  TARGET_TEAM_NAME: '',

  // Name of the log files
  TEST_LOG_FILE_NAME: 'Email_Run_Log_TEST',
  LIVE_LOG_FILE_NAME: 'Email_Run_Log_LIVE'
};

/**
 * Main function to run. 
 * Can be run manually from the editor.
 */
function runReviewMailer() {
  const reviewsFolder = DriveApp.getFolderById(CONFIG.REVIEWS_FOLDER_ID);
  
  if (!reviewsFolder) {
    Logger.log("Error: Could not find Reviews folder. Please check ID.");
    return;
  }

  // 1. Get Roster Data
  const rosterData = getRosterData();
  if (!rosterData) return; // Error logged in helper

  // 2. Select file(s)
  const files = reviewsFolder.getFilesByType(MimeType.GOOGLE_DOCS);
  const fileList = [];
  while (files.hasNext()) {
    fileList.push(files.next());
  }
  fileList.sort((a, b) => a.getName().localeCompare(b.getName()));
  
  if (CONFIG.TARGET_TEAM_NAME === '') {
    Logger.log("--- No specific team selected in CONFIG.TARGET_TEAM_NAME ---");
    Logger.log("Available Teams in Reviews Folder:");
    fileList.forEach(f => Logger.log("- " + f.getName()));
    Logger.log("-------------------------------------------------------------");
    Logger.log("Please set CONFIG.TARGET_TEAM_NAME to one of the above names and run again.");
    return;
  }
  
  // 3. Process the targeted team
  const targetFile = fileList.find(f => f.getName().includes(CONFIG.TARGET_TEAM_NAME));
  
  if (!targetFile) {
    Logger.log("Error: Could not find file matching: " + CONFIG.TARGET_TEAM_NAME);
    return;
  }
  
  Logger.log("Processing: " + targetFile.getName());
  processTeamFile(targetFile, rosterData, reviewsFolder);
}

/**
 * Handles the logic for a single team file
 */
function processTeamFile(reviewFile, rosterData, reviewsFolder) {
  // 1. Extract Team Name from Filename
  // Expecting format like "Team A (1).docx" or "Team A.docx"
  // logic: grab "Team [Word]"
  // This might need adjustment depending on exact naming convention
  const fileName = reviewFile.getName();
  
  // Regex to match "Team A", "Team B", "Team Alpha", etc.
  // Assumes team name is at the start
  let teamName = "";
  const match = fileName.match(/^(Team\s+[A-Za-z0-9]+)/i);
  if (match) {
    teamName = match[1].trim(); 
  } else {
    Logger.log("Could not parse Team Name from filename: " + fileName);
    return;
  }
  
  // 2. Lookup Emails from Roster
  const emails = rosterData[teamName]; // Lookup in our pre-built map
  
  if (!emails || emails.length === 0) {
    Logger.log("No emails found in Roster Sheet for team: '" + teamName + "'");
    return;
  }
  Logger.log(`Found ${emails.length} emails for ${teamName}: ${emails.join(", ")}`);
  
  // 3. Prepare Email
  const subject = "Peer Review Results for " + teamName;
  let body = "Attached are your team's peer review results.\n\nBest regards,\nAdmin";
  let recipients = emails.join(",");
  
  // 4. Handle Test Mode & Logging
  const logFileName = CONFIG.TEST_MODE ? CONFIG.TEST_LOG_FILE_NAME : CONFIG.LIVE_LOG_FILE_NAME;
  
  if (CONFIG.TEST_MODE) {
    body = "This is a test email. If it were a real email, it would have sent the following message to [" + emails.join(", ") + "].\n\n" + body;
    recipients = CONFIG.ADMIN_EMAIL;
    Logger.log("TEST MODE ACTIVE. Sending to: " + recipients);
  } else {
    Logger.log("SENDING LIVE EMAIL to: " + recipients);
  }
  
  // Log the run
  logRunToDrive(reviewsFolder, logFileName, teamName, emails, CONFIG.TEST_MODE);

  // 5. Send
  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      body: body,
      attachments: [reviewFile.getAs(MimeType.PDF)]
    });
    Logger.log("Email sent successfully.");
  } catch (e) {
    Logger.log("Error sending email: " + e.toString());
  }
}

/**
 * Reads the Google Sheet and returns a map of Team Name -> [Array of Emails]
 */
function getRosterData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ROSTER_SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.ROSTER_TAB_NAME) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log("Roster sheet seems empty or missing header.");
      return null;
    }
    
    const headers = data[0];
    const teamColIdx = headers.indexOf(CONFIG.TEAM_COL_HEADER);
    const emailColIdx = headers.indexOf(CONFIG.EMAIL_COL_HEADER);
    
    if (teamColIdx === -1 || emailColIdx === -1) {
      Logger.log(`Error: Could not find columns '${CONFIG.TEAM_COL_HEADER}' or '${CONFIG.EMAIL_COL_HEADER}' in sheet headers.`);
      Logger.log("Headers found: " + headers.join(", "));
      return null;
    }
    
    const rosterMap = {};
    
    // Iterate rows (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const team = String(row[teamColIdx]).trim();
      const email = String(row[emailColIdx]).trim();
      
      if (team && email) {
        if (!rosterMap[team]) {
          rosterMap[team] = [];
        }
        rosterMap[team].push(email);
      }
    }
    
    return rosterMap;
  } catch (e) {
    Logger.log("Error reading Roster Sheet: " + e.toString());
    return null;
  }
}

/**
 * Logs the run details to a Google Doc
 */
function logRunToDrive(folder, logFileName, teamName, recipientsList, isTest) {
  const iterator = folder.getFilesByName(logFileName);
  let doc;
  if (iterator.hasNext()) {
    doc = DocumentApp.openById(iterator.next().getId());
  } else {
    // Create new Google Doc
    const newDocId = DocumentApp.create(logFileName).getId();
    const newFile = DriveApp.getFileById(newDocId);
    newFile.moveTo(folder); 
    doc = DocumentApp.openById(newDocId);
    doc.getBody().setText(`Peer Review Email Log (${isTest ? "TEST" : "LIVE"})\n\n`);
  }
  
  const body = doc.getBody();
  const timestamp = new Date().toString();
  
  const statusLine = isTest ? "STATUS: TEST (Sent to Admin)" : "STATUS: LIVE (Sent to Recipients)";
  
  body.appendParagraph("------------------------------------------------");
  body.appendParagraph("Date: " + timestamp);
  body.appendParagraph("Team: " + teamName);
  body.appendParagraph("Intended Recipients: " + recipientsList.join(", "));
  body.appendParagraph(statusLine);
  
  doc.saveAndClose();
  Logger.log("Logged run details to file: " + logFileName);
}
