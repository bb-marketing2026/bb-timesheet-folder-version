> **Strictly private & confidential.** Internal Brainbox use only. Do not forward, publish or distribute outside the group.

# Brainbox Timesheet Tracker (folder edition)

A start/end timer that writes client hours into Brainbox timesheets (.xlsx) kept in a folder on your computer or in a OneDrive or SharePoint folder synced to it. Several people can track time for the same client, and a Team log shows everyone's work so it can be added to the timesheet. Built with React 18, Vite 5, ExcelJS (reading) and JSZip (saving). No server or database.

There is a separate Microsoft 365 edition, which saves directly through SharePoint and handles people saving at the same moment more strongly. See its own README.

## How it works

| Panel or field | What it does |
|---|---|
| Open timesheet | Choose the folder that holds the client timesheet. If it has several .xlsx files, pick one. Recent clients are remembered. |
| Your name (under the title) | Type your name and click **Save**. There is no default. Colleagues see it next to your entries in the Team log after they press **Refresh**. It is never written into the timesheet. Click **Change** to update it. Your existing log keeps working under the new name. |
| Activity tag / Activity | Written to **Project Description** and **Activities**. Tags are suggested from the file. |
| Start / End | The timer. It survives a refresh or closed tab. |
| Hours / Hours in total | Decimal hours to two places, matching the sheet. |
| Add hours to the file | Saves your sessions to your team log first, then adds them to the timesheet in date order. |
| Save to team log only | Saves to your team log without touching the timesheet, so someone else can add it later. |
| Add another session | Clears the timer for the next task. |
| Team log | Everyone's entries for this timesheet, by person and status. Select entries and click **Add selected to timesheet**. |
| Refresh | Re-reads the timesheet and everyone's logs. The button flashes "new updates" when the app sees changes (checked every minute). |
| Insert rows | Inserts 1 to 500 blank, formatted rows above Total Hours and extends the total formula. |
| Save a copy | Saves an exact backup named `[file] - Backup - vYYYYMMDD-rHHMM.xlsx` (Perth time). |
| Delete (bin icon) | Removes a row from the file after you confirm. Rows below move up. |
| Undo | Restores the file exactly as it was before your last change (last 10 changes, kept in this browser). |

Team log statuses:

| Status | Meaning |
|---|---|
| Not added yet | In someone's log, not in the timesheet. |
| In timesheet | Added, and a matching row is in the file. |
| Check | Marked as added but no matching row was found. It may have been edited or deleted in Excel, or lost in a OneDrive conflict. Add it again only if it really is missing. |

## Files the app creates

```
Clients/RAWCS/
  Brainbox - Timesheet - RAWCS.xlsx            the timesheet (template unchanged, no staff names)
  Timesheet Logs/
    Brainbox - Timesheet - RAWCS - Shan.json   one person's log (only their app writes it)
    Brainbox - Timesheet - RAWCS - Jo.json     a colleague's log
```

Each person writes only their own log, so logs never clash. A log is named after the name first saved, and each log carries a fixed ID, so changing your name later updates the same log. If a colleague already uses the name you type, the app asks whether that log is yours (for example from another computer) before linking it, otherwise you choose a different name.

## Recommended way of working when several people share a client

1. Everyone keeps the client folder synced with OneDrive and opens it in the app.
2. Each person tracks their sessions and clicks **Add hours to the file** when done.
3. Before adding, and whenever "Refresh: new updates" flashes, click **Refresh**.
4. If a save fails, the work is already in that person's log as "Not added yet". Add it from the Team log later.

When two people are very likely to save at the same time, use **Save to team log only** and let one person (for example the account lead) add everything from the Team log at the end of the day.

## How the app protects the timesheet

- It changes only the cells it writes. Other sheets, sensitivity labels, SharePoint metadata, custom properties and printer settings are kept byte for byte.
- Before saving, it checks the file has not changed since it was read. If it has, it applies the change again to the newer file.
- Every add, delete and insert is checked before saving (existing rows unchanged, totals add up, dates in order). Nothing is saved if a check fails.
- It warns when it finds a OneDrive conflict copy (for example `... RAWCS-DESKTOP-ABC.xlsx`) or the file is open in Excel on this computer.

## Run locally

Requires Node.js 18 or later.

```bash
npm install
npm run dev
```

Open http://localhost:5173 in Chrome or Edge.

## Deploy: GitHub, then Vercel

1. Create a private GitHub repository named `brainbox-timesheet-tracker` (no README, no .gitignore).
2. Upload everything in this folder, so `package.json` sits at the top of the repository.
3. At vercel.com/new, import the repository with these settings:

| Setting | Value |
|---|---|
| Framework Preset | Vite |
| Root Directory | `./` |
| Build Command | `npm run build` |
| Output Directory | `dist` |
| Install Command | `npm install` |
| Environment Variables | None |

4. Click **Deploy**. Share the link with the team. Each person uses it with their own folder access.

## Known issues and limits

- **Chrome or Edge only for the full app.** Firefox and Safari can open one file and download updated copies, with no team log.
- **OneDrive sync delay.** Each computer saves its own copy and OneDrive takes seconds to minutes to sync. Two saves to the same timesheet inside that gap can still produce a conflict copy. The app warns when it sees one, and the Team log "Check" status shows rows that went missing. The Microsoft 365 edition removes this risk.
- **Close the timesheet in Excel before saving from the app.** Windows locks open files.
- **Permission prompt.** Chrome asks to allow editing of the folder once per browser session.
- **Inserting rows is refused** on sheets with tables, charts, images, comments or pivot tables. Do that in Excel.
- **Timers, pending sessions, recent clients and undo copies** stay in each person's browser.

## Test the Excel logic

```bash
npm test -- "path/to/Brainbox - Timesheet - RAWCS.xlsx"
```

## Files

| File | Purpose |
|---|---|
| `src/App.jsx` | Interface, timer, Team log |
| `src/components.jsx` | Date, time and new timesheet forms |
| `src/core/timesheet.js` | Reads the timesheet and creates new client timesheets |
| `src/core/xlsxPatch.js` | Cell-level add (in date order), delete and insert rows, with checks |
| `src/core/team.js` | Team log format and statuses |
| `src/core/util.js` | Australian dates (Perth time), file names, downloads |
| `src/store/local.js` | Folder access, logs, change checks, conflict copy warnings |
| `src/store/idb.js` | Recent clients and undo history in the browser |
