# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## THIS REPOSITORY IS PUBLIC

Everything committed here is world-readable. Before adding any file, assume it
will be read by strangers.

Never commit:

- **Live Google resource IDs** — Sheet, Drive folder, Apps Script project or
  deployment IDs. `.clasp.json` carries the script ID and is gitignored.
- **Instructor guides or any `.docx`** — they embed live Sheet, Drive and
  web-app hyperlinks inside the Word XML, where no text search will find them.
  Gitignored for that reason.
- **Rosters, exports, real email addresses, or student work** of any kind.
- Local scratch: one-off debugging scripts, editor workspace files.

The `.gitignore` encodes all of this. If something belongs to a running course
rather than to the template, it does not belong in this repo.

Before pushing, scan **all history**, not the working tree — a commit keeps its
own snapshot, so fixing a file today leaves the old value readable in the commit
that introduced it:

```bash
git grep -hoE '[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}' $(git rev-list --all) -- | sort -u
git grep -hoE 'https://(docs|script|drive)\.google\.com/[^ ")`]+' $(git rev-list --all) -- | sort -u
git ls-files | grep -iE '\.(docx|xlsx|pptx|pdf|zip)$'
```

Expect only placeholder addresses (`jsmith@`, `you@`, `no-reply@`) and no
Google URLs or binaries at all.

## What this is

A distributable Google Apps Script template for team peer review. Students
submit rubric-based reviews through a web app; review documents are generated in
Drive and can be emailed to teams when the review period closes.

It is a **starting point others copy**, not a fan-out template — there is no
`sync.sh` and no registered downstream courses. `README.md` is the setup guide
written for that audience, and is the reference for how the app is configured.

## Shape of the code

- Apps Script server files use `.gs`: `Code.gs`, `Utilities.gs`,
  `Email_Reviews.gs`, `Distribute_Feedback.gs`, `Append_Stats.gs`,
  `create_template_sheet.gs`
- Client: `Index.html`, `Scripts.html`, `Styles.html`
- Python helpers for the document workflow: `append_docs.py`, `split_teams.py`
- Configuration lives in **Script Properties**, not a config file (README
  step 4). Nothing here should hardcode a Sheet or folder ID.

## Note on the instructor guide

An earlier instructor guide in this folder documents a deployment whose
spreadsheet has since been deleted; the web app errors because it cannot open
it. Treat those IDs as stale. Regenerating that guide for a live course will
embed live links — which is precisely why `*.docx` is gitignored here.
