const REVIEW_FOLDER_NAME = 'Reviews';
const SHEET_NAMES = {
  roster: 'Roster',
  assignments: 'Team Assignments',
  teams: 'Teams',
  questions: 'Questions',
  logs: 'Logs',
  config: 'Config',
};
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = SCRIPT_PROPS.getProperty('SPREADSHEET_ID');

function doGet() {
  var email = '';
  try {
    email = (Session.getActiveUser().getEmail() || '').trim();
  } catch (err) {
    Logger.log(err);
  }
  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = email;
  template.cacheKey = Date.now();
  return template
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('NC State PeerSpectives')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns the email of the authenticated user running the script.
 * Used to verify that payload emails match the actual session user.
 * NOTE: Requires deployment as "Anyone at NC State University" — Session.getActiveUser()
 * only returns the viewer's email within the same Google Workspace domain.
 */
function getSessionEmail_() {
  try {
    return (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function getRevieweeChoices(payload) {
  try {
    var email = (payload && payload.email ? payload.email : '').trim().toLowerCase();
    if (!email) {
      throw new Error('Missing email address.');
    }
    var sessionEmail = getSessionEmail_();
    if (!sessionEmail || sessionEmail !== email) {
      return { ok: false, message: 'Authentication error. Please reload the page and try again.' };
    }
    var config = getConfig();
    var reviewer = findRosterEntry(email);
    if (!reviewer) {
      return {
        ok: false,
        message: 'We could not find your email in the roster. Please contact the course staff.',
        config: config,
      };
    }
    var completed = getCompletedTeams_(email);
    var assignments = findTeamAssignments(reviewer.team);
    var reviewees = assignments.filter(function (team) {
      return completed.indexOf(team.teamName.toLowerCase()) === -1;
    });
    return {
      ok: true,
      reviewer: reviewer,
      reviewees: reviewees,
      totalAssignments: assignments.length,
      config: config,
    };
  } catch (err) {
    Logger.log(err);
    return { ok: false, message: err.message };
  }
}

function getRubric() {
  try {
    var questions = loadQuestions();
    if (!questions.length) {
      throw new Error('No questions configured.');
    }
    var groups = {};
    var totalScale = 0;
    questions.forEach(function (q) {
      var cat = (q.category || '').toString().trim();
      if (!groups[cat]) {
        groups[cat] = [];
      }
      groups[cat].push(q);
      if (q.type === 'scale') {
        totalScale += q.scaleMax;
      }
    });
    var rubric = Object.keys(groups)
      .sort()
      .map(function (cat) {
        return { category: cat, questions: groups[cat] };
      });
    return { ok: true, rubric: rubric, totals: { maxScale: totalScale } };
  } catch (err) {
    Logger.log(err);
    return { ok: false, message: err.message };
  }
}

function getVideoLink(teamName) {
  try {
    var sheet = getSheet_(SHEET_NAMES.teams);
    var numRows = Math.max(sheet.getLastRow() - 1, 0);
    var valueRows = sheet.getRange(2, 1, numRows, 5).getValues();
    var richRows = sheet.getRange(2, 2, numRows, 1).getRichTextValues();
    var matchIndex = valueRows.findIndex(function (row) {
      return (row[0] || '').toString().trim().toLowerCase() === (teamName || '').toString().trim().toLowerCase();
    });
    if (matchIndex === -1) {
      return { ok: false, message: 'No project link found.' };
    }
    var match = valueRows[matchIndex];
    var richCell = (richRows[matchIndex] && richRows[matchIndex][0]) || null;
    var rawUrl = (match[1] || '').toString().trim();
    var linkUrl = richCell && richCell.getLinkUrl ? richCell.getLinkUrl() : '';
    var finalUrl = linkUrl || rawUrl;
    if (!finalUrl) {
      return { ok: false, message: 'No project link found.' };
    }
    return {
      ok: true,
      url: finalUrl,
      orgName: match[2] || '',
      topic: match[3] || '',
      slidesUrl: (match[4] || '').toString().trim(),
    };
  } catch (err) {
    Logger.log(err);
    return { ok: false, message: err.message };
  }
}

function submitReview(formData) {
  try {
    if (!formData) {
      throw new Error('Submission payload missing.');
    }
    var reviewer = formData.reviewer || {};
    var revieweeTeam = (formData.revieweeTeam || '').trim();
    if (!revieweeTeam) {
      throw new Error('Please choose a team to review.');
    }
    var email = (reviewer.email || '').trim().toLowerCase();
    if (!email) {
      throw new Error('Reviewer email missing.');
    }
    var sessionEmail = getSessionEmail_();
    if (!sessionEmail || sessionEmail !== email) {
      throw new Error('Authentication error. Please reload and try again.');
    }
    var rosterEntry = findRosterEntry(email);
    if (!rosterEntry) {
      throw new Error('Reviewer not found in roster.');
    }

    var questions = loadQuestions();
    var answersInput = formData.answers || [];
    var answerMap = {};
    answersInput.forEach(function (ans) {
      if (ans && ans.qName) {
        answerMap[ans.qName] = ans.value;
      }
    });

    var missing = [];
    var recordedAnswers = [];
    var scoreTotal = 0;
    var scoreMax = 0;

    questions.forEach(function (q) {
      var response = answerMap[q.name];
      if (response === undefined || response === null || response === '') {
        missing.push(q.label);
        return;
      }
      if (q.type === 'scale') {
        var numeric = Number(response);
        if (isNaN(numeric) || numeric < 1 || numeric > q.scaleMax) {
          missing.push(q.label);
          return;
        }
        scoreTotal += numeric;
        scoreMax += q.scaleMax;
        recordedAnswers.push({
          name: q.name,
          label: q.label,
          type: q.type,
          value: numeric,
          display: numeric + ' / ' + q.scaleMax,
        });
      } else {
        var text = response.toString().trim();
        if (!text) {
          missing.push(q.label);
          return;
        }
        recordedAnswers.push({
          name: q.name,
          label: q.label,
          type: q.type,
          value: text,
          display: text,
        });
      }
    });

    if (missing.length) {
      return { ok: false, message: 'Please complete: ' + missing.join(', ') };
    }

    var percent = scoreMax ? Math.round((scoreTotal / scoreMax) * 1000) / 10 : 0;
    var timestamp = new Date();
    var scoreSummary = { total: scoreTotal, max: scoreMax, percent: percent };

    // Acquire a script-level lock to prevent race conditions when multiple
    // students submit simultaneously (e.g., two reviewers hitting the same
    // team's doc at the same moment).
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    var docUrl;
    try {
      docUrl = generateDoc({
        reviewer: rosterEntry,
        revieweeTeam: revieweeTeam,
        timestamp: timestamp,
        answers: recordedAnswers,
        scoreSummary: scoreSummary,
      });

      writeLogRow({
        timestamp: timestamp,
        reviewer: rosterEntry,
        revieweeTeam: revieweeTeam,
        answers: recordedAnswers,
        scoreSummary: scoreSummary,
        docUrl: docUrl,
      });
    } finally {
      lock.releaseLock();
    }

    return { ok: true, fileUrl: docUrl, scoreSummary: scoreSummary };
  } catch (err) {
    Logger.log(err);
    return { ok: false, message: err.message };
  }
}

function getReviewCounts(opts) {
  try {
    var email = (opts && opts.email ? opts.email : '').trim().toLowerCase();
    if (!email) {
      throw new Error('Missing email address.');
    }
    var sessionEmail = getSessionEmail_();
    if (!sessionEmail || sessionEmail !== email) {
      throw new Error('Authentication error.');
    }
    var sheet = getSheet_(SHEET_NAMES.logs);
    var lastRow = sheet.getLastRow();
    var completed = 0;
    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      completed = rows.filter(function (row) {
        return (row[2] || '').toString().trim().toLowerCase() === email;
      }).length;
    }
    var choices = getRevieweeChoices({ email: email });
    var total = choices && choices.ok && choices.reviewees ? choices.reviewees.length : 0;
    return { completed: completed, total: total };
  } catch (err) {
    Logger.log(err);
    return { completed: 0, total: 0 };
  }
}
