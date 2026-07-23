---
name: qlik-refresh
description: Pull fresh YTD specialty-team exports from Qlik Cloud (via the signed-in Chrome) and refresh the Specialty Teams dashboard. Use when Dean asks to refresh the dashboard / grab Qlik exports.
---

Refresh the Specialty Teams Performance Dashboard end-to-end: pull fresh exports from
Qlik Cloud, then rebuild the data embedded in `specialty_teams_dashboard.html`.

## Ground rules
- Use the connected Chrome (Claude in Chrome tools) for all Qlik work. NEVER enter
  Dean's credentials or complete Andersen SSO. If Qlik isn't signed in / the SSO
  session is expired, STOP immediately and report that sign-in is needed.
- A full **YTD** export (Jan 1 of the current year → today) is sufficient each time.
  The refresh script carries the dashboard's existing data forward (prior years and
  anything the new export doesn't cover persist automatically).

## 1) Pull the export from Qlik
Source sheet (ISC Agent Performance app):
https://rbi.us.qlikcloud.com/sense/app/cf823d56-5a15-4cd1-8306-5f14a707efe3/sheet/dceae61a-91a1-4f05-976f-26ccaddeeefc/state/analysis

1. Clear any existing selections.
2. Filter **ApptSetBy** to the specialty-team agents: search each prefix —
   `@OUTB`, `@CXSV`, `@RSLE`, `@SWPS`, `@RVTS` — and select all matching values
   (apply one prefix at a time, confirm the selection chip, then add the next).
3. IMPORTANT — former/inactive reps: the Qlik data model is dynamic and keeps no
   history. When a rep leaves a specialty team their prefix changes (usually back to
   @OUTB or off the specialty prefixes entirely). Read `roster.csv` in this repo and,
   for any agent whose name did NOT get captured by the prefix selection, search
   ApptSetBy for their name and add whatever value they now appear under. (Their
   already-captured history is safe via carry-forward; this step only catches records
   that changed or arrived since the last refresh.)
4. Date filter: full YTD in ONE range makes the table error with "The request
   exceeds the memory limit" — pull it in **quarter-sized chunks** instead
   (verified working with 42 agents selected): Jan 1–Mar 31, Apr 1–Jun 30, then
   Jul 1–today, etc. Set each range by opening the date picker and typing both
   dates (From field, Enter, then To field, Enter, then click outside to apply).
   Note the picker RESETS both fields each time it opens — always set From and To
   together; editing just one field creates a single-day selection.
5. For each chunk, download the detail table (the "Agent" tab's table — columns
   start CalendarDate, ApptSetBy, ... and end AppointmentSetOn, # of attempts,
   Source; 23 columns): right-click the table → Table → Download → Data →
   Download. Wait for the xlsx to land in Downloads (takes 1–3 minutes to
   generate; ~13MB per quarter) before starting the next chunk. Leave filenames
   alone ("Untitled - <date> (N).xlsx").
6. **Cost-center store export** (feeds the "YTD vs Entire Cost Center" card): go to
   the Agent Overview (Store) sheet
   (https://rbi.us.qlikcloud.com/sense/app/cf823d56-5a15-4cd1-8306-5f14a707efe3/sheet/4c72cd7f-6e24-4b6e-bdea-893259fe9131/state/analysis),
   **Store** tab. NO agent filter — this is the entire cost center. Set the date
   range to the same YTD window (Jan 1 → today; one range is fine, this table is
   only 12 store rows) and download the Store table the same way. The refresh
   script recognizes it by schema (StoreName + Net Order $, no CalendarDate) and
   updates the card; if you skip this export, the card keeps its previous data.
7. When done, clear the CalendarDate and ApptSetBy selection chips (toolbar X on
   each) to leave the sheet clean.

## 2) Rebuild the dashboard
Run from the repo root:

    python refresh_dashboard.py

It auto-discovers the last 24h of matching exports in Downloads, merges them over the
existing data (newest wins per calendar day + rep + team), rewrites the embedded data
in `specialty_teams_dashboard.html`, and stamps "Data refreshed / Data through" in the
header. Review the printed before/after totals — big unexplained drops mean a bad
export (nothing is lost; fix the Qlik selection and rerun, or `git checkout` the HTML).

## 3) Verify and hand off
- Open the dashboard (or reload it in the browser) and confirm the header shows
  today's refresh time and the expected "Data through" date, and KPIs look sane.
- Roster changes (someone joined/left a team) are edited by hand in `roster.csv`
  (Team, Agent, StartDate, EndDate; blank EndDate = active), then rerun the refresh.
- Remind Dean to commit & push the updated HTML (GitHub Desktop) to publish.
