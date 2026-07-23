# Specialty Teams Performance Dashboard

A self-contained HTML dashboard for the specialty teams (@OUTB, @CXSV, @RSLE, @SWPS, @RVTS).
All data is embedded inside [specialty_teams_dashboard.html](specialty_teams_dashboard.html) —
open it in any browser, no server needed.

## Refreshing the data

Tip: in a Claude Code session in this repo, `/qlik-refresh` runs this whole flow
(Claude pulls the export through your signed-in Chrome, then rebuilds).

1. **Export from Qlik.** Source: the [ISC Agent Performance sheet](https://rbi.us.qlikcloud.com/sense/app/cf823d56-5a15-4cd1-8306-5f14a707efe3/sheet/dceae61a-91a1-4f05-976f-26ccaddeeefc/state/analysis)
   — the appointment-detail table whose export has these 23 columns:

   > CalendarDate, ApptSetBy, Ownership Group, StoreName, # of Contacts, Appts Set,
   > Set Rate, Appts Set Same Day, Appts Set Next Day, Appts Issued, Demos, No Demo,
   > Not Home, Demo Rate, Net Sales, Net Sale $, Avg Job Size, RPA,
   > Net Close to Issued %, DaysFromAppointmentCreatedToAppointment,
   > AppointmentSetOn, # of attempts, Source

   Filter ApptSetBy to the specialty-team prefixes (@OUTB, @CXSV, @RSLE, @SWPS,
   @RVTS) and export **year-to-date** (Jan 1 → today). A full-YTD export every
   refresh is the recipe: the Qlik data model is dynamic and keeps no history, so
   this re-pulls the whole year fresh, while the dashboard carries forward anything
   the export doesn't cover (prior years; reps who went inactive or back to @OUTB
   and fell out of the prefix filter — see carry-forward below). If Qlik's export
   limit forces it, split into multiple exports — by month, by team, any chunking
   is fine. No renaming needed; "Untitled - … .xlsx" in Downloads is expected.

   Also grab the **cost-center store export** for the "YTD vs Entire Cost Center"
   card: the [Agent Overview (Store) sheet](https://rbi.us.qlikcloud.com/sense/app/cf823d56-5a15-4cd1-8306-5f14a707efe3/sheet/4c72cd7f-6e24-4b6e-bdea-893259fe9131/state/analysis),
   Store tab, **no agent filter**, same YTD range, download the Store table. The
   script recognizes it automatically; skipping it just leaves that card's data as-is.

2. **Run the refresh.** Double-click **`Refresh Dashboard.bat`** (or run
   `python refresh_dashboard.py`). It scans Downloads for detail exports from the
   last 24 hours, rebuilds the data blob inside the HTML, stamps the refresh time,
   and prints a before/after comparison of the totals. Sanity-check that comparison —
   if something looks way off, the export was probably filtered oddly; nothing is
   lost, just re-export and run it again (or `git checkout` the HTML to revert).

3. **Publish.** Commit and push the updated HTML (e.g. in GitHub Desktop).

Useful flags: `--check` builds and prints the comparison without touching the HTML;
`--files <path> …` uses specific exports instead of scanning Downloads;
`--lookback N` widens the Downloads scan to the last N hours.

## How the merge works (and carry-forward)

Exports often overlap (re-downloads, per-team chunks, partial date windows). For each
**(calendar day, rep + team)**, the **newest** export covering it wins; older files only
fill gaps. So a fresh YTD export supersedes everything it covers, and
per-team/per-period chunks merge without double counting.

The dashboard's **current embedded data participates as the oldest source** in every
refresh. That's what preserves history the dynamic Qlik model can't give back: 2025
records, and records of reps who since left a team (their prefix changes, so they drop
out of prefix-filtered exports — but their captured history stays in the dashboard).
`--no-carry` disables this and rebuilds strictly from the exports you pass.

## The roster

[roster.csv](roster.csv) is the source of truth for who is on which team and when
(StartDate–EndDate; blank end = still active). Every export row is attributed to a
team by **rep name + tenure window** — the `@XXXX` prefix in ApptSetBy is ignored, so
a rep's history follows them across prefix changes. Rows outside every tenure window
are dropped.

To add, move, or retire an agent: edit roster.csv, then run the refresh again
(the roster is embedded into the HTML along with the data).

## Files

| File | Purpose |
|---|---|
| `specialty_teams_dashboard.html` | The dashboard (data embedded; this is the deliverable) |
| `refresh_dashboard.py` | Rebuilds the embedded data from Qlik exports |
| `Refresh Dashboard.bat` | One-click runner for the above |
| `roster.csv` | Team roster + tenure windows (edit by hand) |

Requires Python 3 with `openpyxl` (`pip install openpyxl`).
