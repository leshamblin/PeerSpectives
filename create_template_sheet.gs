/**
 * PeerSpectives — Template Sheet Generator
 *
 * Run createTemplateSheet() once to generate a fully structured Google Sheet
 * with all required tabs, column headers, and example data.
 *
 * After running:
 *   1. Replace example data with your actual course data
 *   2. Copy the Spreadsheet ID from the URL and paste it into Script Properties
 *      as SPREADSHEET_ID
 */

function createTemplateSheet() {
  const ss = SpreadsheetApp.create('PeerSpectives Template — YOUR COURSE NAME');
  const url = ss.getUrl();
  const id  = ss.getId();

  _buildRoster(ss);
  _buildTeamAssignments(ss);
  _buildTeams(ss);
  _buildQuestions(ss);
  _buildConfig(ss);
  _buildLogs(ss);

  // Remove the default blank sheet
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  Logger.log('✅ Template sheet created!');
  Logger.log('URL: ' + url);
  Logger.log('Spreadsheet ID: ' + id);
  Logger.log('');
  Logger.log('Next steps:');
  Logger.log('1. Replace example data with your course data');
  Logger.log('2. Set SPREADSHEET_ID = ' + id + ' in Apps Script → Project Settings → Script Properties');
}

// ─────────────────────────────────────────────────────────────────────────────

function _buildRoster(ss) {
  const sheet = ss.insertSheet('Roster');
  const headers = ['Name', 'Email', 'Team'];
  const examples = [
    ['Jane Doe',    'jadoe@ncsu.edu',   'Team Alpha'],
    ['John Smith',  'jsmith@ncsu.edu',  'Team Alpha'],
    ['Alice Chen',  'achen@ncsu.edu',   'Team Beta'],
    ['Bob Patel',   'bpatel@ncsu.edu',  'Team Beta'],
    ['Carol Jones', 'cjones@ncsu.edu',  'Team Gamma'],
    ['David Kim',   'dkim@ncsu.edu',    'Team Gamma'],
  ];
  _writeHeaderRow(sheet, headers);
  sheet.getRange(2, 1, examples.length, headers.length).setValues(examples);
  sheet.setFrozenRows(1);
  _autoResize(sheet, headers.length);
}

function _buildTeamAssignments(ss) {
  const sheet = ss.insertSheet('Team Assignments');
  const headers = ['Reviewer Team', 'Reviewee 1', 'Reviewee 2', 'Reviewee 3'];
  const examples = [
    ['Team Alpha', 'Team Beta',  'Team Gamma', ''],
    ['Team Beta',  'Team Alpha', 'Team Gamma', ''],
    ['Team Gamma', 'Team Alpha', 'Team Beta',  ''],
  ];
  _writeHeaderRow(sheet, headers);
  sheet.getRange(2, 1, examples.length, headers.length).setValues(examples);
  sheet.setFrozenRows(1);
  _autoResize(sheet, headers.length);

  const note = sheet.getRange('A1');
  note.setNote('Team names must match exactly (case-sensitive) across all tabs.\nAdd more "Reviewee N" columns if teams review more than 3 others.');
}

function _buildTeams(ss) {
  const sheet = ss.insertSheet('Teams');
  const headers = ['Team Name', 'Link', 'Organization', 'Topic'];
  const examples = [
    ['Team Alpha', 'https://docs.google.com/presentation/...', 'Acme Corp',    'Supply chain optimization'],
    ['Team Beta',  'https://docs.google.com/presentation/...', 'Blue Ridge Co','Customer retention strategy'],
    ['Team Gamma', 'https://docs.google.com/presentation/...', 'Cedar LLC',    'Market entry analysis'],
  ];
  _writeHeaderRow(sheet, headers);
  sheet.getRange(2, 1, examples.length, headers.length).setValues(examples);
  sheet.setFrozenRows(1);
  _autoResize(sheet, headers.length);

  sheet.getRange('B1').setNote('Link shown to reviewers before they fill out the rubric (e.g., slides, video, website).');
}

function _buildQuestions(ss) {
  const sheet = ss.insertSheet('Questions');
  const headers = ['Question Name', 'Type', 'Scale Max', 'Label', 'Category', 'Hover Text'];
  const examples = [
    // Category: Presentation
    ['Intro Section Instructions', 'info',  '',  'Review the team\'s presentation materials using the link above before filling out this rubric.', 'Presentation', ''],
    ['Overall Clarity',            'scale', '7', 'Clarity of communication',   'Presentation', 'Rate from 1 (very unclear) to 7 (exceptionally clear)'],
    ['Visual Design',              'scale', '7', 'Quality of slides/visuals',  'Presentation', 'Rate from 1 (needs work) to 7 (professional quality)'],
    ['Content Depth',              'scale', '7', 'Depth of analysis',          'Presentation', 'Rate from 1 (surface-level) to 7 (thorough and insightful)'],
    // Category: Teamwork
    ['Teamwork Instructions',      'info',  '',  'Answer the following based on your interactions with this team during the project.', 'Teamwork', ''],
    ['Collaboration',              'scale', '5', 'Collaboration with other teams', 'Teamwork', 'Rate from 1 (difficult to work with) to 5 (excellent collaborator)'],
    // Category: Written Feedback
    ['Strengths',                  'text',  '',  '',  'Written Feedback', 'What did this team do particularly well?'],
    ['Areas to Improve',           'text',  '',  '',  'Written Feedback', 'What could this team strengthen or do differently?'],
    ['Additional Comments',        'text',  '',  '',  'Written Feedback', 'Any other thoughts for this team? (optional)'],
  ];
  _writeHeaderRow(sheet, headers);
  sheet.getRange(2, 1, examples.length, headers.length).setValues(examples);
  sheet.setFrozenRows(1);
  _autoResize(sheet, headers.length);

  sheet.getRange('B1').setNote('Supported types:\n  scale — interactive slider (1 to Scale Max)\n  text  — free-form textarea\n  info  — read-only instruction block');
  sheet.getRange('E1').setNote('Questions with the same Category are grouped into a collapsible section in the UI.');
}

function _buildConfig(ss) {
  const sheet = ss.insertSheet('Config');
  const data = [
    ['course number',    'MBA 531'],
    ['techsupport email','helpdesk@ncsu.edu'],
  ];
  _writeHeaderRow(sheet, ['Setting', 'Value']);
  sheet.getRange(2, 1, data.length, 2).setValues(data);
  sheet.setFrozenRows(1);
  _autoResize(sheet, 2);
}

function _buildLogs(ss) {
  const sheet = ss.insertSheet('Logs');
  const headers = [
    'Timestamp', 'Reviewer Name', 'Reviewer Email', 'Reviewer Team',
    'Reviewee Team', 'Score', 'Max Score', 'Percent', 'Answers JSON', 'Doc URL'
  ];
  _writeHeaderRow(sheet, headers);
  sheet.setFrozenRows(1);
  _autoResize(sheet, headers.length);
  sheet.getRange('A1').setNote('This tab is written to automatically by the app. Do not edit manually.');
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function _writeHeaderRow(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  range.setBackground('#cc0000');
  range.setFontColor('#ffffff');
}

function _autoResize(sheet, numCols) {
  for (let i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
  }
}
