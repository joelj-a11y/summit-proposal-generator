# Summit Quote & Proposal Generator

A web app hosted on Render that does three things:

1. **Quote calculator.** Staff enter project details and get per-inspection pricing. Each team has its own page:
   - Summit Services Group: `https://summit-proposal-generator.onrender.com/`
   - SSG Western Colorado: `https://summit-proposal-generator.onrender.com/western-colorado`
2. **Proposal generator** (`/generate`). Upload a quote file (`.json`) and download a finished Word proposal.
3. **Quote log.** Every downloaded quote adds one row to the Google Sheet, on that team's own tab.

## Files in this repository

| File | What it is | Who edits it |
|---|---|---|
| `teams/summit.js` | Summit's pricing assumptions and inspector roster | Managers, when prices change |
| `teams/western-colorado.js` | SSG Western Colorado's pricing assumptions and roster | Managers, when prices change |
| `Summit_Proposal_Template.docx` | The Word proposal template, with `{{PLACEHOLDERS}}` | Anyone, in Word |
| `summit-quote-generator.html` | The calculator page, shared by all teams | Developer only |
| `app.py` | The web server: pages, Word proposals, quote log | Developer only |
| `requirements.txt` | Python packages Render installs | Developer only |

## Everyday workflow

1. Open your team's calculator page and fill in the project details.
2. Click **Generate Quote** and review the numbers. Optionally enter a manual routine rate.
3. Click **Download Quote Data**. This saves a `.json` file *and* adds the quote to your team's tab in the quote log. A green message confirms the save; a red message means the file downloaded but the log did not update.
4. Click the **Proposal Generator** link, upload the `.json` file, and download the Word proposal.

## Changing prices or the inspector roster

Every assumption lives in the team's file in the `teams/` folder: hourly rate, admin fee, drive speed, round-trip distance, road-distance factor, time on site by acreage, storm-event rate, post-construction percentage, flat-rate multipliers and premium, rate sheet, project types, consulting rate, and inspectors.

1. In GitHub, open `teams/` and click the team's file.
2. Click the pencil icon (**Edit this file**).
3. Change only the numbers or text to the right of the colons. Keep commas and quote marks as they are.
4. Click **Commit changes**. Render redeploys automatically in a few minutes.
5. Run one test quote to confirm the page loads and the numbers look right.

If a calculator page shows "Team settings could not be loaded," a comma or quote mark was probably removed by mistake. Undo the last edit in GitHub's history.

**Inspector locations** are approximate town-center coordinates. To add an inspector, copy an existing line and change the name, town, and coordinates. To find coordinates, right-click the town in Google Maps. To stop assigning someone quotes without deleting them, set `active: false`.

## Adding another team

1. Copy `teams/summit.js` to a new file such as `teams/new-team.js`. Change `teamKey`, `teamName`, `pageTitle`, and the pricing.
2. In `app.py`, add an entry to the `TEAMS` list with the same key, a web address, the settings file name, and a Google Sheet tab name.
3. Commit both files. The team's Sheet tab is created automatically with its first logged quote.

## Editing the Word template

Open `Summit_Proposal_Template.docx` in Word and edit freely: wording, fonts, layout, logo. Text in double curly braces is filled in automatically. You can move a placeholder, copy it to a new spot, or delete it.

Type placeholders exactly as listed below, in capitals. A misspelled placeholder appears unchanged in the finished proposal, so a typo is easy to spot.

| Placeholder | Filled with |
|---|---|
| `{{PROJECT_NAME}}`, `{{CLIENT_NAME}}`, `{{PROJECT_TYPE}}`, `{{DISTURBED_ACRES}}` | Project details |
| `{{PROJECT_CITY}}`, `{{PROJECT_COUNTY}}`, `{{PROJECT_STATE}}`, `{{PROJECT_ADDRESS}}` | Location (state is always Colorado) |
| `{{SITE_CONTACT_NAME}}`, `{{SITE_CONTACT_PHONE}}`, `{{SITE_CONTACT_EMAIL}}` | Client's site contact |
| `{{PREPARED_BY}}`, `{{PREPARED_BY_CONTACT}}` | Summit contact name, and phone / email |
| `{{TEAM_NAME}}`, `{{DATE}}` | Team name; today's date (not currently used in the template) |
| `{{ITEM_1A_QTY}}`, `{{ITEM_1A_RATE}}`, `{{ITEM_1A_TOTAL}}` (also `1B`, `1C`) | SWMP items; blank if not selected |
| `{{SWMP_SUBTOTAL}}` | Items 1A + 1B + 1C |
| `{{ROUTINE_QTY}}`, `{{ROUTINE_RATE}}`, `{{ROUTINE_TOTAL}}` | Item 2A, routine inspections |
| `{{STORM_QTY}}`, `{{STORM_RATE}}`, `{{STORM_TOTAL}}` | Item 2B, post-storm inspections |
| `{{STORM_EVENTS_PER_MONTH}}` | Storm-event assumption from the team settings |
| `{{INSPECTION_SUBTOTAL}}` | Items 2A + 2B |
| `{{FLAT_MONTHS}}`, `{{FLAT_MONTHLY_RATE}}`, `{{FLAT_TOTAL}}` | Item 3, flat monthly rate |
| `{{POST_QTY}}`, `{{POST_RATE}}`, `{{POST_TOTAL}}`, `{{POST_SUBTOTAL}}` | Item 4, post-construction inspections |
| `{{ITEM_5_DESC}}`, `{{ITEM_5_QTY}}`, `{{ITEM_5_RATE}}`, `{{ITEM_5_TOTAL}}` (also `6`) | Consulting items; blank if not selected |
| `{{OTHER_SUBTOTAL}}` | Items 5 + 6 |
| `{{OPTION_1_TOTAL}}` | Project total, per-inspection billing |
| `{{OPTION_2_TOTAL}}` | Project total, flat monthly billing |
| `{{CONSULTING_RATE}}` | Consulting hourly rate from the team settings |
| `{{SCOPE_OF_WORK}}` | Must be alone on its own line. Becomes one paragraph per selected item that has a scope description. |

## Quote log (Google Sheet)

The sheet is the one whose ID is saved in Render as `GOOGLE_SHEET_ID`. Open it at `https://docs.google.com/spreadsheets/d/<that ID>`. Each team writes to its own tab, which is created automatically with a header row. The old `Sheet1` tab is left untouched as an archive.

## Render settings

| Setting | Value |
|---|---|
| Build command | `pip install -r requirements.txt` |
| Start command | `gunicorn app:app` |
| Environment variables | `GOOGLE_SHEET_ID`, `GOOGLE_SHEETS_CREDENTIALS` |

To check the setup, visit `/health`. It reports whether the template, calculator, team files, and Google settings are all in place.

## References

- Render, deploying a Flask app: <https://render.com/docs/deploy-flask>
- python-docx, the library that fills in the Word template: <https://python-docx.readthedocs.io/>
- Google Sheets API, appending rows: <https://developers.google.com/sheets/api/guides/values#append_values>
