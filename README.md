<!--
================================================================
README.md
----------------------------------------------------------------
PURPOSE : Front page of the GitHub repo - what this is and how to run it.
CHANGE LOG
- 2026-10-14 : Initial version. No email addresses, IDs or secrets on
  purpose: the repo is PUBLIC.
================================================================
-->

# Mirage Tracker Reminder

Every evening at **8:00 PM IST**, emails each person the drawings that are **due today** in the
project tracker Excel files on SharePoint, plus a **one-time** reminder for drawings that became
overdue. Each tracker's lead gets a nightly summary (sent, failed, names not matched, bad cells).

The tracker-reading logic is a 1:1 JavaScript port of the **Excel Notifier** add-in (C#), so both
read a tracker the same way (checked with `tools/parityCheck.js`).

## Folder map

| Folder | What is in it |
|---|---|
| `core/` | Pure logic, no network: read the workbook, detect columns, decide who gets what, build the emails. Runs in Node and (later) in the browser. |
| `job/` | The nightly run: sign in as the job app, read files from SharePoint, send mail via Microsoft Graph. |
| `tools/` | Offline helpers: `parityCheck.js` (compare with Excel Notifier), `previewEmails.js` (write tonight's emails as .html, sends nothing). |
| `assets/` | Logo used at the end of every email. |
| `.github/workflows/` | The 8 PM schedule (GitHub Actions). |

## Run on your PC

1. Node.js 20+ installed, then `npm install`
2. Copy `.env.example` to `.env` and fill it in (never commit `.env`)
3. `npm run job` - with `DRY_RUN=true` it only reads and prints

Offline checks (no Microsoft 365 needed):

```
npm run parity -- "path\to\tracker.xlsx"
node tools/previewEmails.js "path\to\tracker.xlsx" --today 2026-10-14 --demo
```

## GitHub settings (Settings -> Secrets and variables -> Actions)

- **Secret:** `CLIENT_SECRET`
- **Variables:** `TENANT_ID`, `CLIENT_ID`, `SENDER_MAILBOX`, `CONFIG_SITE_HOST`, `CONFIG_SITE_PATH`,
  `CONFIG_FOLDER`, `ALERT_EMAIL`, `TEST_RECIPIENT`, `GLOBAL_TEST_MODE`

Meaning of each: `job/config.js`. Manual run: Actions -> *Nightly tracker reminders* -> *Run workflow*.

## Safety rules built in

- Mail only to `@<company domain>` addresses; a name with no sure match is **never** emailed (it goes to the lead summary).
- The job app can only send as the one shared mailbox and can only read the SharePoint sites IT granted.
- `GLOBAL_TEST_MODE=true` (default) sends every email to the test address only.
- A sent log stops duplicate reminders, including on re-runs.
