// User manual. One source for the in-app manual and MANUAL.md (scripts/build-manual.mjs).
// Each block: { h: heading, p: paragraph } or { steps: [...] } or { table: [[head...], [row...]] } or { list: [...] }.

export function manual(edition) {
  const m365 = edition === 'm365';
  return {
    title: m365 ? 'Brainbox Timesheet Tracker user manual (Microsoft 365 edition)' : 'Brainbox Timesheet Tracker user manual (local edition)',
    notice: 'Strictly private & confidential. Internal Brainbox use only. Do not forward, publish or distribute outside the group.',
    sections: [
      {
        h: '1. What the tracker does',
        blocks: [
          { p: 'It times your work for a client with Start and End, then writes the hours into that client\'s Brainbox timesheet (.xlsx) in the standard template: Month, Date Log, Project Description, Activities and Total Hours. Colleagues working on the same client see each other\'s work in the Team log and can add it to the timesheet. Staff names are never written into the timesheet.' },
          { p: m365
            ? 'This edition opens timesheets stored in SharePoint or OneDrive for Business. Every save is checked by SharePoint, so two people saving at the same moment cannot overwrite each other.'
            : 'This edition opens timesheets saved on your computer or in a OneDrive folder synced to it. Use Chrome or Edge.' },
        ],
      },
      {
        h: '2. First-time setup (about 2 minutes)',
        blocks: [
          { steps: m365 ? [
            'Open the tracker link in Chrome or Edge.',
            'Click Sign in with Microsoft and use your Brainbox account.',
            'Under the title, type Your name and click Save. Colleagues see this name in the Team log.',
            'In SharePoint or OneDrive, click Copy link on the client timesheet. Paste it into Timesheet link and click Open.',
            'Check the Client tag beside your name (for example RAWCS). Click Change if it is wrong.',
          ] : [
            'Open the tracker link in Chrome or Edge.',
            'Under the title, type Your name and click Save. Colleagues see this name in the Team log.',
            'Click Open timesheet and choose the client\'s .xlsx file. Allow editing when Chrome asks.',
            'Check the Client tag beside your name (for example RAWCS). Click Change if it is wrong.',
            'In the Team log panel, click Connect team log. The folder picker opens where the timesheet is saved. Click Select folder, then Edit files. You do this once per timesheet.',
          ] },
        ],
      },
      {
        h: '3. The screen',
        blocks: [
          { table: [
            ['Area', 'What it is for'],
            ['How to use / User manual', 'Buttons at the top right. How to use opens short steps under the title. User manual opens this manual.'],
            ['Your name / Client tag', 'Your name for the Team log. The client\'s short name that starts every activity tag, for example RAWCS.'],
            ['Timesheet file', m365 ? 'Sign in, open by link, file details, Insert rows, Save a copy, Undo, Recent clients.' : 'Open timesheet, file details, warnings, Insert rows, Save a copy, Undo, Recent clients.'],
            ['Session', 'Activity tag, Activity, the timer, Start and End.'],
            ['Sessions to add', 'Finished sessions waiting to be saved. Edit dates and times here.'],
            ['Team log', 'Everyone\'s entries for this timesheet, with Refresh and Add selected to timesheet.'],
            ['Latest rows in the file', 'The last rows in the timesheet, with the bin icon to delete one.'],
          ] },
        ],
      },
      {
        h: '4. Tracking time',
        blocks: [
          { steps: [
            'Pick an Activity tag, for example Content Planning. The app shows how it is written, for example RAWCS Content Planning.',
            'Describe the Activity, for example Drafted November posts.',
            'Click Start. You can close the tab and the timer keeps counting.',
            'Click End when you finish. The session moves to Sessions to add.',
            'Click Add another session for the next task.',
          ] },
          { p: 'Forgot to start or stop? Edit the date (DD/MM/YYYY) and times (24 hour, Perth time) in Sessions to add, or use + Manual entry.' },
        ],
      },
      {
        h: '5. Saving your hours',
        blocks: [
          { table: [
            ['Button', 'What happens'],
            ['Add hours to the file', 'Saves to your team log first, then adds rows to the timesheet in date order. Month and date appear on the first row of each day only.'],
            ['Save to team log only', 'Saves to your team log without touching the timesheet, so a colleague can add it later.'],
            ['Combine sessions (tick box)', 'Same tag on the same day becomes one row, activities joined with "; ".'],
          ] },
          { p: 'If the timesheet cannot be saved (for example it is open in Excel), your work is already in your team log as Not added yet. Add it from the Team log later.' },
        ],
      },
      {
        h: '6. Working with colleagues',
        blocks: [
          { steps: [
            'Click Refresh to load the latest timesheet and everyone\'s entries. The button flashes "Refresh: new updates" when the app sees changes (it checks every minute).',
            'Filter by person or status.',
            'Tick entries and choose an action. Add selected to timesheet puts them in the timesheet in date order. Archive selected hides them from the list and leaves the timesheet as it is (see them under Archived, and use Restore selected to bring them back). Delete selected removes them from the team log for everyone, and deletes their rows from the timesheet if they are in it, after you confirm.',
          ] },
          { table: [
            ['Status', 'Meaning'],
            ['Not added yet', 'In someone\'s log, not in the timesheet.'],
            ['In timesheet', 'Added, and the matching row is in the file.'],
            ['Removed from timesheet', 'Was added, but the row is no longer in the timesheet. "by [name]" means it was deleted in the tracker. "not found in file" means it was deleted or edited in Excel, or lost in a sync conflict. Add it again only if it should be there, otherwise archive or delete it.'],
            ['Archived', 'Hidden from Needs action. Shown under Archived in the status filter.'],
          ] },
          { p: m365
            ? 'Two people can save at the same time. If a colleague saved first, the app re-applies your change to their version automatically.'
            : 'When two people are likely to save at the same time, use Save to team log only and let one person add everything from the Team log, for example at the end of the day.' },
        ],
      },
      {
        h: '7. Managing the timesheet',
        blocks: [
          { table: [
            ['Tool', 'How to use it'],
            ['Insert rows', 'Enter 1 to 500 and click Insert rows when Blank rows left is low (it turns red below 10). Formatting is copied and the Total Hours formula is extended.'],
            ['Save a copy', 'Saves an exact backup named "[timesheet] - Backup - vYYYYMMDD-rHHMM.xlsx". You keep working in the original.'],
            ['Delete (bin icon)', 'Removes a row after you confirm. Rows below move up. Its team log entry shows Removed from timesheet for everyone after they press Refresh.'],
            ['Undo', 'Puts the timesheet back exactly as it was before your last change (last 10 changes). Warns you if a colleague saved since.'],
            ['New timesheet', 'Creates a blank Brainbox timesheet for a new client.'],
          ] },
        ],
      },
      {
        h: '8. Troubleshooting',
        blocks: [
          { table: [
            ['Problem', 'What to do'],
            ['"Could not save" or "SharePoint is holding it"', 'Close the timesheet in Excel and click the button again. Nothing was changed.'],
            ['Not enough blank rows', 'Use Insert rows. The number needed is filled in for you.'],
            [m365 ? 'Cannot open the link' : '"Not the folder that holds" message', m365 ? 'Use Copy link in SharePoint or OneDrive, not the address bar of Excel for the web.' : 'In the folder picker, choose the folder the timesheet is saved in.'],
            ...(m365 ? [] : [['Possible OneDrive conflict copy', 'OneDrive saved two versions. Open the copy in Excel, check for rows that are missing from the main file, then add them from the Team log.']]),
            ['Name already used', 'Choose Yes only if that log is yours from another computer. Otherwise add a surname or initial.'],
            ['A colleague\'s entries are not showing', 'Click Refresh.' + (m365 ? '' : ' Their OneDrive may still be syncing.')],
          ] },
        ],
      },
      {
        h: '9. Good practice',
        blocks: [
          { list: [
            'Save your name once per browser and check the client tag when you open a new timesheet.',
            'Click Refresh before adding colleagues\' entries.',
            'Save a copy at month end.',
            'Close the timesheet in Excel before saving from the tracker.',
            'Do not rename the timesheet' + (m365 ? '.' : ' or the Timesheet Logs folder.'),
          ] },
        ],
      },
    ],
  };
}

/** The manual as Markdown, for MANUAL.md. */
export function manualMarkdown(edition) {
  const m = manual(edition);
  const out = [`> **${m.notice.split('. ')[0]}.** ${m.notice.split('. ').slice(1).join('. ')}`, '', `# ${m.title}`, ''];
  for (const s of m.sections) {
    out.push(`## ${s.h}`, '');
    for (const b of s.blocks) {
      if (b.p) out.push(b.p, '');
      if (b.steps) { b.steps.forEach((t, i) => out.push(`${i + 1}. ${t}`)); out.push(''); }
      if (b.list) { b.list.forEach((t) => out.push(`- ${t}`)); out.push(''); }
      if (b.table) {
        const [head, ...rows] = b.table;
        out.push(`| ${head.join(' | ')} |`, `|${head.map(() => '---').join('|')}|`);
        rows.forEach((r) => out.push(`| ${r.join(' | ')} |`));
        out.push('');
      }
    }
  }
  return out.join('\n');
}
