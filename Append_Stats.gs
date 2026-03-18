/**
 * Appends team stats from a source Google Doc to individual team files in a target folder.
 * 
 * INSTRUCTIONS:
 * 1. Upload 'Team_Stats_Tables_With_Summary.docx' to Drive and open it as a Google Doc.
 * 2. Copy the ID of that Google Doc (from the URL) into SOURCE_DOC_ID below.
 * 3. Copy the ID of the Folder containing your team files into TARGET_FOLDER_ID below.
 * 4. Run the 'appendTeamStats' function.
 */

// *** CONFIGURATION ***
const SOURCE_DOC_ID = 'REPLACE_WITH_SOURCE_DOC_ID'; 
const TARGET_FOLDER_ID = 'REPLACE_WITH_TARGET_FOLDER_ID';

function appendTeamStats() {
  const sourceDoc = DocumentApp.openById(SOURCE_DOC_ID);
  const body = sourceDoc.getBody();
  const numChildren = body.getNumChildren();
  
  const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
  
  let currentTeamName = null;
  let currentTeamElements = [];
  
  // Iterate through the source document
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    const type = child.getType();
    
    // Check for Team Header
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asParagraph().getText();
      // Pattern: "Team A – Team Peer Review Stats"
      if (text.startsWith('Team ') && text.includes('Peer Review Stats')) {
        
        // If we were already collecting for a team, process the previous team now
        if (currentTeamName) {
          processTeam(currentTeamName, currentTeamElements, targetFolder);
        }
        
        // Start new team
        currentTeamName = text.split('–')[0].split('-')[0].trim(); // "Team A"
        currentTeamElements = []; // Reset elements
        
        // We generally don't want to append the header itself if the target file already has one,
        // but let's include it for clarity or make it optional. 
        // For now, let's include it so they know what the stats are.
        currentTeamElements.push(child.copy()); 
        continue;
      }
    }
    
    // If we are inside a team section, collect elements
    if (currentTeamName) {
      // Stop if we hit an empty paragraph that might be a separator? 
      // Actually, just collect everything until the next header.
      currentTeamElements.push(child.copy());
    }
  }
  
  // Process the last team
  if (currentTeamName) {
    processTeam(currentTeamName, currentTeamElements, targetFolder);
  }
}

function processTeam(teamName, elements, folder) {
  Logger.log('Processing: ' + teamName);
  
  // Find the target file
  const files = folder.getFiles();
  let targetFile = null;
  
  while (files.hasNext()) {
    const file = files.next();
    // Simple matching: check if filename contains "Team A" (case insensitive)
    // Be careful with "Team A" matching "Team AB" - adding space or boundary check is better
    // But for now, let's assume standard naming "Team A ..."
    if (file.getName().toLowerCase().includes(teamName.toLowerCase()) && 
        file.getMimeType() === MimeType.GOOGLE_DOCS) {
      targetFile = file;
      break;
    }
  }
  
  if (!targetFile) {
    Logger.log('  -> No matching file found for ' + teamName);
    return;
  }
  
  Logger.log('  -> Inserting at TOP of: ' + targetFile.getName());
  
  const doc = DocumentApp.openById(targetFile.getId());
  const body = doc.getBody();
  
  // We want to insert at the very beginning (index 0).
  // We must maintain the order of elements, so we insert the first element at 0,
  // the second at 1, etc.
  let insertIndex = 0;

  // Insert all collected elements
  elements.forEach(element => {
    const type = element.getType();
    let inserted = null;
    
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      inserted = body.insertParagraph(insertIndex, element);
    } else if (type === DocumentApp.ElementType.TABLE) {
      inserted = body.insertTable(insertIndex, element);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      inserted = body.insertListItem(insertIndex, element);
    }
    
    if (inserted) {
      insertIndex++;
    }
  });
  
  // Add a page break after the new stats to separate them from the previous content
  body.insertPageBreak(insertIndex);
  
  doc.saveAndClose();
  Logger.log('  -> Done.');
}
