function getSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('Set the SPREADSHEET_ID script property.');
  }
  return active;
}

function getSheet_(name) {
  var sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) {
    throw new Error('Missing sheet: ' + name);
  }
  return sheet;
}

function getConfig() {
  var sheet = getSheet_(SHEET_NAMES.config);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var config = {};

  // Detect header/value layout (row 1 headers, row 2 values)
  var headers = data[0] || [];
  var values = data[1] || [];
  var headerMode = headers.some(function (cell) {
    return (cell || '').toString().trim();
  });
  if (headerMode) {
    headers.forEach(function (header, idx) {
      var key = (header || '').toString().trim();
      if (!key) return;
      assignConfigKey_(config, key, values[idx]);
    });
    return config;
  }

  // Fallback: key/value pairs per row (original behavior)
  var rows = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 2)).getValues();
  rows.forEach(function (row) {
    var key = (row[0] || '').toString().trim();
    if (!key) return;
    assignConfigKey_(config, key, row[1]);
  });
  return config;
}

function assignConfigKey_(config, key, value) {
  switch (key.toLowerCase()) {
    case 'course number':
    case 'course':
      config.courseNumber = value || '';
      break;
    case 'techsupport email':
    case 'techsupport':
    case 'support':
      config.techsupportEmail = value || '';
      break;
    default:
      config[key] = value;
      break;
  }
}

function findRosterEntry(email) {
  if (!email) return null;
  var sheet = getSheet_(SHEET_NAMES.roster);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var lowered = email.toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    var rowEmail = (rows[i][1] || '').toString().trim().toLowerCase();
    if (rowEmail === lowered) {
      return {
        name: rows[i][0] || '',
        email: rows[i][1] || '',
        team: rows[i][2] || '',
      };
    }
  }
  return null;
}

function findTeamAssignments(teamName) {
  var result = [];
  if (!teamName) return result;
  var sheet = getSheet_(SHEET_NAMES.assignments);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;
  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var target = teamName.toString().trim().toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    var current = (rows[i][0] || '').toString().trim().toLowerCase();
    if (current === target) {
      for (var c = 1; c < rows[i].length; c++) {
        var name = (rows[i][c] || '').toString().trim();
        if (name) {
          result.push({ teamName: name, label: name });
        }
      }
      break;
    }
  }
  return result;
}

function loadQuestions() {
  var sheet = getSheet_(SHEET_NAMES.questions);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return rows
    .filter(function (row) {
      return (row[0] || '').toString().trim();
    })
    .map(function (row) {
      var rawType = (row[1] || 'scale').toString().trim().toLowerCase();
      var normalizedType;
      if (rawType === 'text') {
        normalizedType = 'text';
      } else if (/info|instruction|note/.test(rawType)) {
        normalizedType = 'info';
      } else {
        normalizedType = 'scale';
      }
      var scaleMax = Number(row[2]);
      if (!scaleMax || scaleMax < 1) scaleMax = 5;
      return {
        name: row[0].toString().trim(),
        type: normalizedType,
        scaleMax: scaleMax,
        label: row[3] || row[0],
        category: row[4] || '',
        hover: row[5] || '',
      };
    });
}

function ensureFolderExists(name) {
  var parentId = SCRIPT_PROPS.getProperty('REVIEW_PARENT_FOLDER_ID');
  var parentFolder;
  try {
    parentFolder = parentId ? DriveApp.getFolderById(parentId) : DriveApp.getRootFolder();
  } catch (err) {
    Logger.log('Invalid REVIEW_PARENT_FOLDER_ID; using root. ' + err);
    parentFolder = DriveApp.getRootFolder();
  }
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(name);
}

function generateDoc(payload) {
  var folder = ensureFolderExists(REVIEW_FOLDER_NAME);
  var docName = payload.revieweeTeam + ' (1)';
  var file = getOrCreateReviewDoc_(folder, docName);
  var doc = DocumentApp.openById(file.getId());
  var body = doc.getBody();

  body.appendParagraph('Review submitted ' + formatTimestamp_(payload.timestamp)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Reviewer: ' + payload.reviewer.name + ' (' + payload.reviewer.email + ')');
  body.appendParagraph('Reviewer Team: ' + payload.reviewer.team);
  body.appendParagraph('Reviewee Team: ' + payload.revieweeTeam);

  var feedbackAnswers = payload.answers.filter(function (answer) {
    return answer.type === 'text';
  });
  if (feedbackAnswers.length) {
    body.appendParagraph('Feedback').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    feedbackAnswers.forEach(function (answer) {
      body.appendParagraph(answer.display);
      body.appendParagraph('');
    });
  }

  body.appendParagraph('— — —');

  doc.saveAndClose();
  return doc.getUrl();
}

function getOrCreateReviewDoc_(folder, name) {
  var iterator = folder.getFilesByName(name);
  if (iterator.hasNext()) {
    return iterator.next();
  }
  var doc = DocumentApp.create(name);
  var file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (err) {
    Logger.log('Could not remove file from root: ' + err);
  }
  return file;
}

function writeLogRow(entry) {
  var sheet = getSheet_(SHEET_NAMES.logs);
  sheet.appendRow([
    entry.timestamp || new Date(),
    entry.reviewer.name,
    entry.reviewer.email,
    entry.reviewer.team,
    entry.revieweeTeam,
    entry.scoreSummary.total,
    entry.scoreSummary.max,
    entry.scoreSummary.percent,
    JSON.stringify(entry.answers),
    entry.docUrl,
  ]);
}

function formatTimestamp_(date) {
  return Utilities.formatDate(date || new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a z');
}

function updateDocHeaders(teamName) {
  if (!teamName) throw new Error('Provide a team name.');
  var docs = DriveApp.getFolderByName(REVIEW_FOLDER_NAME)
    .getFilesByName(teamName + ' (1)');
  if (!docs.hasNext()) throw new Error('Doc not found for team: ' + teamName);
  var file = docs.next();
  var doc = DocumentApp.openById(file.getId());
  var body = doc.getBody();

  var rosterSheet = getSheet_(SHEET_NAMES.roster);
  var rosterRows = rosterSheet.getRange(2, 1, Math.max(rosterSheet.getLastRow() - 1, 0), 3).getValues();
  var members = rosterRows
    .filter(function (row) {
      return (row[2] || '').toString().trim().toLowerCase() === teamName.toString().trim().toLowerCase();
    })
    .map(function (row) {
      return row[1];
    });

  var teamSheet = getSheet_(SHEET_NAMES.teams);
  var teamRows = teamSheet.getRange(2, 1, Math.max(teamSheet.getLastRow() - 1, 0), 5).getValues();
  var meta = teamRows.find(function (row) {
    return (row[0] || '').toString().trim().toLowerCase() === teamName.toString().trim().toLowerCase();
  }) || [];

  body.appendParagraph('Reviews for ' + teamName).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Team member emails: ' + (members.join(', ') || '—'));
  body.appendParagraph('Organization: ' + (meta[2] || '—'));
  body.appendParagraph('Topic: ' + (meta[3] || '—'));
  doc.saveAndClose();
}

function updateTeamStats() {
  var logsSheet = getSheet_(SHEET_NAMES.logs);
  var statsSheetName = 'Team Stats';
  var statsSheet = getSpreadsheet_().getSheetByName(statsSheetName) || getSpreadsheet_().insertSheet(statsSheetName);
  statsSheet.clearContents();
  var questions = loadQuestions().filter(function (q) {
    return q.type === 'scale';
  });
  var headers = ['Team', 'Review Count'].concat(
    questions.map(function (q) {
      return q.label || q.name;
    })
  );
  statsSheet.appendRow(headers);

  var lastRow = logsSheet.getLastRow();
  if (lastRow < 2) return;
  var data = logsSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var averages = {};

  data.forEach(function (row) {
    var revieweeTeam = (row[4] || '').toString().trim();
    if (!revieweeTeam) return;
    try {
      var answers = JSON.parse(row[8] || '[]');
      var byQuestion = {};
      answers.forEach(function (ans) {
        var key = ans.qName || ans.name;
        var val = Number(ans.value);
        if (key && !isNaN(val)) {
          byQuestion[key] = val;
        }
      });
      if (!averages[revieweeTeam]) {
        averages[revieweeTeam] = { count: 0, totals: {} };
      }
      averages[revieweeTeam].count += 1;
      questions.forEach(function (q) {
        var value = byQuestion[q.name];
        if (typeof value !== 'number') return;
        if (!averages[revieweeTeam].totals[q.name]) {
          averages[revieweeTeam].totals[q.name] = 0;
        }
        averages[revieweeTeam].totals[q.name] += value;
      });
    } catch (err) {
      Logger.log('Could not parse answers for team ' + revieweeTeam + ': ' + err);
    }
  });

  Object.keys(averages)
    .sort()
    .forEach(function (team) {
      var entry = averages[team];
      var row = [team, entry.count];
      questions.forEach(function (q) {
        var total = entry.totals[q.name] || 0;
        var avg = entry.count ? Math.round((total / entry.count) * 100) / 100 : 0;
        row.push(avg);
      });
      statsSheet.appendRow(row);
    });
}
function getCompletedTeams_(reviewerEmail) {
  var sheet = getSheet_(SHEET_NAMES.logs);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return rows
    .filter(function (row) {
      return (row[2] || '').toString().trim().toLowerCase() === reviewerEmail;
    })
    .map(function (row) {
      return (row[4] || '').toString().trim().toLowerCase();
    });
}
