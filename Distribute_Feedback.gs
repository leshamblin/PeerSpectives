
/**
 * --- CONFIGURATION ---
 */
const FEEDBACK_CONFIG = {
  // ID of the Source Google Doc (which has the instructor feedback)
  SOURCE_FILE_ID: 'REPLACE_WITH_SOURCE_DOC_ID',

  // ID of the Reviews Folder where the Team docs are located
  REVIEWS_FOLDER_ID: 'REPLACE_WITH_REVIEWS_FOLDER_ID'
};

/**
 * Main function to distribute feedback sections to Team Docs.
 */
function distributeFeedback() {
  const sourceDoc = getSourceDocById(FEEDBACK_CONFIG.SOURCE_FILE_ID);
  if (!sourceDoc) {
    Logger.log("Could not find the source Google Doc with ID: " + FEEDBACK_CONFIG.SOURCE_FILE_ID);
    return;
  }
  
  const body = sourceDoc.getBody();
  const numChildren = body.getNumChildren();
  
  let currentTeamName = null;
  let elementBuffer = [];
  
  Logger.log(`Scanning source doc: ${sourceDoc.getName()} (${numChildren} elements)`);
  
  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    const type = element.getType();
    let isTeamHeader = false;
    let text = "";
    
    // Check for explicit Page Break element
    if (type === DocumentApp.ElementType.PAGE_BREAK) {
      // A page break often signifies the END of the previous section
      // But usually it's attached to a paragraph or is its own child
      // We'll treat it as a break.
      // If we encounter a hard page break, we might want to reset for the next header scan
      // But typically the header comes right after.
    }
    
    // Check for Paragraphs to find Header
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      text = element.asParagraph().getText().trim();
      
      // Strategy 1: Check for Asterisks: *Team A*
      if (/^\*\s*Team\s+.+\*$/.test(text) ||  // Matches "*Team A*"
          /^Team\s+.+$/.test(text) && element.asParagraph().getHeading() !== DocumentApp.ParagraphHeading.NORMAL) { // Fallback: Heading style?
          
         // Removing asterisks for clean name
         const cleanText = text.replace(/\*/g, '').trim(); 
         
         // Verify it starts with Team after cleaning
         if (cleanText.toLowerCase().startsWith("team")) {
            isTeamHeader = true;
            text = cleanText; // Update text for extraction
         }
      }
    }
    
    if (isTeamHeader) {
      Logger.log(`Found header: "${text}"`);
      
      // Flush previous team
      if (currentTeamName) {
        appendBufferToTeamFile(currentTeamName, elementBuffer);
      }
      
      // Start new team
      currentTeamName = extractTeamName(text); 
      elementBuffer = []; // Reset buffer
      
      // We do NOT add the header line itself to the buffer to avoid duplication?
      // Or do we want the header in the feedback? Let's keep it to be safe.
      elementBuffer.push(element); 
      
    } else {
      // If we are currently tracking a team, add content to buffer
      if (currentTeamName) {
        // Option: Don't add if it's just an empty paragraph at start? 
        elementBuffer.push(element);
      }
    }
  }
  
  // Flush the last team
  if (currentTeamName && elementBuffer.length > 0) {
    appendBufferToTeamFile(currentTeamName, elementBuffer);
  }
  
  Logger.log("Distribution Complete.");
}

/**
 * Clean cleanup of team name from header text.
 * e.g. "*Team A*" -> "Team A"
 */
function extractTeamName(headerText) {
  // Remove asterisks first
  let clean = headerText.replace(/\*/g, '').trim();
  
  // Match "Team [Name]"
  // This allows "Team A - Comments" to become "Team A"
  const match = clean.match(/^(Team\s+[A-Za-z0-9]+)/i);
  if (match) {
    return match[1].trim(); // Returns "Team A"
  }
  return clean; // Fallback
}

/**
 * Appends the collected elements to the Team's review file.
 */
function appendBufferToTeamFile(teamName, elements) {
  const folder = getReviewsFolder();
  if (!folder) return;
  
  // Updated logic: Search for "Team A" exactly or "Team A.docx"
  // No longer looking for (1)
  const files = folder.getFiles();
  let targetFile = null;
  
  // We need to be careful not to match "Team Alpha" when looking for "Team A" if using startsWith
  // Better to check specific patterns
  
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName().replace(/\.[^/.]+$/, ""); // Remove extension
    
    // Strict match? or startsWith?
    // If your file is called "Team A", strict match is safer.
    if (name.toLowerCase() === teamName.toLowerCase()) {
       targetFile = f;
       break;
    }
    // Fallback: If file is "Team A Feedback.docx"
    if (name.toLowerCase().startsWith(teamName.toLowerCase() + " ")) {
       // Only match if followed by space to avoid Team A matching Team Alpha
       targetFile = f;
       break;
    }
  }
  
  if (!targetFile) {
    Logger.log(`Warning: matching file for "${teamName}" in Reviews folder. Skipping.`);
    return;
  }
  
  Logger.log(`Inserting ${elements.length} elements to TOP of: ${targetFile.getName()}`);
  
  const doc = DocumentApp.openById(targetFile.getId());
  const body = doc.getBody();
  
  // Logic: Insert at the beginning (Index 0)
  // We maintain an index that increments as we add items so they stay in order.
  let insertIndex = 0;
  
  // 1. Add Header
  body.insertParagraph(insertIndex++, "--- INSTRUCTOR FEEDBACK ---")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  // 2. Insert Elements
  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    const type = element.getType();
    
    // We clone the elements to move them
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      body.insertParagraph(insertIndex++, element.asParagraph().copy());
    } else if (type === DocumentApp.ElementType.TABLE) {
      body.insertTable(insertIndex++, element.asTable().copy());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      body.insertListItem(insertIndex++, element.asListItem().copy());
    }
    // Skip unknowns
  }
  
  // 3. Add Page Break to separate from the original content below
  body.insertPageBreak(insertIndex++);
  
  doc.saveAndClose();
}

/**
 * Helper to get the Reviews folder by ID
 */
function getReviewsFolder() {
  try {
    return DriveApp.getFolderById(FEEDBACK_CONFIG.REVIEWS_FOLDER_ID);
  } catch(e) {
    Logger.log("Error: Could not open Reviews Folder with ID provided. " + e.toString());
    return null;
  }
}

/**
 * Helper to get source doc by ID
 */
function getSourceDocById(id) {
  try {
    return DocumentApp.openById(id);
  } catch (e) {
    Logger.log("Error finding source doc: " + e.toString());
    return null;
  }
}
