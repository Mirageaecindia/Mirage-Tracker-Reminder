# Tracker Reminder - Design

Oct 14, 2026 · @mirage

## 1. Overview

- A **scheduled Node job** (GitHub Actions) wakes every 30 minutes at :07 and :37.
- It reads the **config files in SharePoint**, runs only the trackers whose IST send time has arrived, and reads each tracker Excel through **Microsoft Graph**.
- Names are resolved to emails, reminders are built with the **Excel\_Notify logic** (ported to JavaScript), and mail goes out from the **shared mailbox**.
- A **React admin page** (GitHub Pages) edits the same config files, using the admin's own Microsoft sign-in.
- **SharePoint is the only database.** Nothing is stored on GitHub except code.

## 2. Architecture

&#91;embedded content: system architecture · job, page, SharePoint, Graph, Exchange\]

The job and the page never talk to each other. Both read and write the same SharePoint files through Graph, so there is no server to run.

## 3. Nightly run flow

&#91;embedded content: one run · 9 steps, 1 decision\]

In test mode every email in the Send and Lead summary steps goes to the test address with a `[TEST -> x]` subject prefix. Test sends use separate sent-log keys, so they never block real sends.

## 4. Components and modules

**Rule:** the core has no I/O and no browser or Node specifics. The job and the page plug into it. SheetJS is passed in, not imported (dependency inversion).

| Layer | Modules | Responsibility |
| --- | --- | --- |
| **core/** | `roles`, `roleMatcher`, `cellAddress`, `nameCellParser`, `dateCellParser` | Excel\_Notify parsing rules, ported 1:1 |
|  | `workbookSnapshot`, `sheetLayoutDetector`, `assignmentItem`, `trackerReader` | Read a workbook, find layout, produce assignment items |
|  | `recipientResolver`, `peopleSheetReader` | Name to email (map, People sheet, directory) |
|  | `recipientBuilder`, `sentLog` | Which reminders to send; duplicate protection |
|  | `emailComposer`, `leadSummaryComposer`, `htmlBuilder` | Email content |
|  | `trackerConfig`, `schedule` | Config validation; IST send-time and run-state rules |
| **job/** | `config`, `graphClient`, `sharePointFiles`, `directoryService`, `mailSender`, `runNightly` | Environment, Graph calls with retry, config files, mail, orchestration |
| **web/** | `graph`, `configStore`, `trackerLinks`, `trackerReading`, pages and `App` | Sign-in, load and save with eTag, link resolution, the four screens |
| **tools/** | `parityCheck`, `previewEmails` | Prove parity with Excel\_Notify; preview emails offline |

## 5. Data and config files

All live in SharePoint: Automation, Documents, Application\_Files, TrackerReminder.

| File | Written by | Holds |
| --- | --- | --- |
| `trackers.json` | Admin page | Registered trackers: file link, lead, send time, sheets and roles, backlog policy, test mode, enabled |
| `name_map.json` | Admin page | Shared name to email map (first lookup) |
| `admins.json` | Admin page | Who may use the page |
| `sentlog.json` | Job | What was already sent; due entries pruned after 7 days, overdue after 365 |
| `runstate.json` | Job | Last run date per tracker; alert throttle (one per day per problem) |

**Conflicts:** every save uses the file's eTag (`If-Match`). A 412 becomes `ConflictError` and the page asks the admin to reload.

## 6. Security and permissions

- **Job identity:** an Entra app using client credentials. Secret stored in GitHub Secrets only.
- **Mail:** `Application Mail.Send`, restricted by Exchange role-based access to the `TrackerReminderSender` scope. The app **cannot** send as anyone except `project-reminders@`.
- **SharePoint:** `Sites.Selected`. Write on `/sites/Automation`, read on `/sites/MirageAEC_Engg_BIM`.
- **Page identity:** the admin's own Microsoft sign-in (MSAL redirect flow). The page holds no secret.
- **Admin gate:** the page loads `admins.json` and refuses anyone not listed.
- **Domain lock:** recipients must be `@mirageaec.com`.
- **Test mode:** on by default; both the global and the per-tracker switch must be off for real delivery.
- **Secret expiry:** the client secret was created for 2 years. Renew it before it expires.

## 7. Design decisions and known risks

**Decisions**

- **SharePoint as the only database:** no server or DB to host; admins already have access control there.
- **Cron wakes every 30 min, config decides who runs:** per-tracker send times are editable without touching GitHub.
- **Cron at :07/:37, not :00/:30:** GitHub drops or delays runs at the top of the hour under load.
- **Shared core for Node and browser:** the page Preview and the job use the same code, so they cannot disagree.
- **Dates as ISO strings:** avoids time-zone shifts; "today" in IST is UTC + 5:30.

**Risks**

- **GitHub scheduled runs are best-effort.** They can be late or skipped. Not yet confirmed on this repo; this is the open test. Fallback: move the scheduler to netcup, core unchanged.
- **GitHub 2FA is required from 2026-10-18.** Without it the account is restricted, which can stop Actions and Pages.
- **Name matching misses.** An unresolved name gets no email. Mitigation: Preview screen and the lead summary list unresolved names.
- **Large workbooks in the browser** may be slow in Preview.
- **Sharing-link resolution is untested** on the page.
