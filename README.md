# PeerSpectives

A Google Apps Script web app for peer evaluation in team-based university courses.
Students submit rubric-based reviews of other teams through a simple web interface.
Review documents are auto-generated in Google Drive and can be emailed to teams
when the review period closes.

Built and maintained by the Instructional Design Group, Poole College of Management,
NC State University.

---

## What it does

- Presents each student with a rubric for every team they are assigned to review
- Accepts slider-based scores and written feedback
- Generates a Google Doc per team with all reviews received
- Logs every submission to a Google Sheet for tracking
- Includes optional scripts to email review PDFs to teams and distribute instructor feedback

---

## Prerequisites

- An NC State Google account (`@ncsu.edu`)
- Access to Google Drive, Sheets, and Apps Script
- [Node.js](https://nodejs.org) and `clasp` installed (for command-line setup)
  — OR — willingness to copy files manually in the Apps Script editor

---

## Setup: Step by Step

### Step 1 — Copy the template Google Sheet

1. Open the [PeerSpectives Template Sheet](REPLACE_WITH_TEMPLATE_SHEET_URL)
2. Click **File → Make a copy** and save it to your Google Drive
3. Rename it something like `MBA 531 PeerSpectives Spring 2026`
4. Copy the **Spreadsheet ID** from the URL:
   `https://docs.google.com/spreadsheets/d/`**`THIS-PART`**`/edit`

> The template sheet has all required tabs with example data and column headers pre-filled.
> Replace the example data with your course data before deploying.

---

### Step 2 — Fill in the sheet tabs

#### Roster tab
One row per student.

| Name | Email | Team |
|------|-------|------|
| Jane Doe | jadoe@ncsu.edu | Team Alpha |
| John Smith | jsmith@ncsu.edu | Team Alpha |

#### Team Assignments tab
One row per reviewer team. List the teams they should review in columns C onward.
Leave cells blank if a team reviews fewer than the maximum number of teams.

| Reviewer Team | Reviewee 1 | Reviewee 2 | Reviewee 3 |
|---------------|------------|------------|------------|
| Team Alpha | Team Beta | Team Gamma | |
| Team Beta | Team Alpha | Team Delta | |

> Team names must match exactly between all tabs (case-sensitive).

#### Teams tab
One row per team. The Link column is shown to reviewers before they fill out the rubric
so they can reference the team's work.

| Team Name | Link | Organization | Topic |
|-----------|------|--------------|-------|
| Team Alpha | https://... | Acme Corp | Supply chain optimization |

#### Questions tab
Defines the rubric. Three supported types:

| Question Name | Type | Scale Max | Label | Category | Hover Text |
|---------------|------|-----------|-------|----------|------------|
| Overall Quality | scale | 7 | Quality of deliverable | Presentation | Rate 1 (needs work) to 7 (excellent) |
| Strengths | text | | | Written Feedback | What did this team do well? |
| Areas to Improve | text | | | Written Feedback | What could be stronger? |

- `scale` — renders an interactive slider from 1 to Scale Max
- `text` — renders a textarea for written feedback
- `info` — renders a read-only text block (use for instructions or section dividers)

Questions with the same Category are grouped into a collapsible section in the UI.

#### Config tab
Two-column layout: setting name in column A, value in column B.

| Setting | Value |
|---------|-------|
| course number | MBA 531 |
| techsupport email | helpdesk@ncsu.edu |

#### Logs tab
Leave this blank — the app writes a row here for every submission automatically.

---

### Step 3 — Copy the code to Apps Script

#### Option A: Using clasp (recommended)

```bash
npm install -g @google/clasp
clasp login
git clone https://github.com/YOUR-ORG/peerspectives.git
cd peerspectives
clasp create --type webapp --title "PeerSpectives - YOUR COURSE"
clasp push
```

#### Option B: Manual copy in the browser

1. Go to [script.google.com](https://script.google.com) and click **New project**
2. Rename the project (e.g., `PeerSpectives MBA 531`)
3. For each file below, create a matching file in the editor and paste the contents:

| File | Type in editor |
|------|---------------|
| `Code.gs` | Script (.gs) |
| `Utilities.gs` | Script (.gs) |
| `Email_Reviews.gs` | Script (.gs) |
| `Distribute_Feedback.gs` | Script (.gs) |
| `Append_Stats.gs` | Script (.gs) |
| `Index.html` | HTML |
| `Scripts.html` | HTML |
| `Styles.html` | HTML |

---

### Step 4 — Set Script Properties

1. In the Apps Script editor, go to **Project Settings** (gear icon ⚙️)
2. Scroll to **Script Properties** and click **Add script property**
3. Add the following property:

| Property | Value |
|----------|-------|
| `SPREADSHEET_ID` | The spreadsheet ID you copied in Step 1 |

Optionally add:

| Property | Value |
|----------|-------|
| `REVIEW_PARENT_FOLDER_ID` | Drive folder ID where the Reviews folder will be created. If omitted, it will be created in your root Drive. |

---

### Step 5 — Deploy as a web app

1. In the Apps Script editor, click **Deploy → New deployment**
2. Click the gear icon next to "Type" and select **Web app**
3. Configure:
   - **Description**: e.g., `Spring 2026 initial deploy`
   - **Execute as**: Me
   - **Who has access**: Anyone at NC State University

   > ⚠️ **"Anyone at NC State University" is required** — do not change this to "Anyone."
   > The app verifies each submission by checking the authenticated Google session against
   > the submitted email. This check only works within the NC State Google Workspace domain.
   > If you set access to "Anyone" (all Google accounts), session emails will not be returned
   > and all submissions will fail with an authentication error.

4. Click **Deploy** and authorize any permission prompts
5. Copy the **Web app URL** — this is what you share with students

> After pushing any code changes, go to **Deploy → Manage deployments**,
> click the pencil icon on your deployment, and select **New version** to
> make changes live.

---

### Step 6 — Test before sharing

1. Open the web app URL in your browser while signed into your `@ncsu.edu` account
2. Verify your name appears and you can see your assigned teams
3. Submit a test review and confirm:
   - A new row appears in the **Logs** tab of your sheet
   - A Google Doc was created in the Reviews folder in Drive

---

### Step 7 — Share with students

Send students the web app URL. They must be signed into their `@ncsu.edu` Google account.
The app authenticates them automatically against the Roster tab — no login screen required.

---

## Post-review workflows (optional)

These scripts are run manually from the Apps Script editor after the review period closes.

### Send reviews to teams — `Email_Reviews.gs`
Fill in the configuration at the top of the file, then run `runReviewMailer()`.
Each team receives a PDF of all reviews submitted about them.

### Distribute instructor feedback — `Distribute_Feedback.gs`
1. Prepare a Google Doc with feedback sections separated by team name headers
2. Fill in `SOURCE_FILE_ID` and `REVIEWS_FOLDER_ID` at the top of the file
3. Run `distributeFeedback()`

### Append team statistics — `Append_Stats.gs`
Fill in `SOURCE_DOC_ID` and `TARGET_FOLDER_ID` at the top of the file,
then run `appendTeamStats()` to add score summary tables to each team's review doc.

---

## Customizing the rubric

All rubric questions live in the **Questions tab** of your Google Sheet — no code changes needed.
Add, remove, or reorder rows to change what reviewers see. Changes take effect immediately
without redeploying.

---

## Troubleshooting

**"You are not on the roster"**
Check that the student's email in the Roster tab exactly matches their NC State login email.

**Reviews folder not appearing in Drive**
Verify `SPREADSHEET_ID` is set correctly in Script Properties (Step 4).

**Code changes not visible after editing**
You must create a new version in Deploy → Manage deployments after pushing changes.

**Submission spinner that never completes**
This can happen during PDF generation under load. The app will time out after 50 seconds
and prompt the student to try again — the submission was not saved and they should resubmit.

**Resubmissions**
Students can resubmit. The latest submission overwrites the previous review doc for that team.

---

## Notes

- Only `@ncsu.edu` accounts can authenticate. This is enforced in `Scripts.html`.
- The Logs tab keeps a permanent record of every submission including all answers as JSON.
- To reset a student's submission (allow a fresh start), delete their row(s) from the Logs tab
  and delete their review doc from the Reviews folder in Drive.

---

## Questions / support

Built by the Instructional Design Group, Poole College of Management, NC State University.
Contact [elizabeth@ncsu.edu](mailto:elizabeth@ncsu.edu) for help with setup.
