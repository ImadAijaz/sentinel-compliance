# Sentinel — project handoff brief

Paste this whole file as your first message to a new assistant, then say what you want changed.

---

## 1. What this is

**Sentinel** is a compliance dashboard for **WCGTX** (Wellness & Care Group of Texas). It tracks
credential expiry for three populations:

- **Providers** (physicians) — ~182 tracked
- **Staff** (nursing, front office, midshift) — per site
- **Facilities** — Frisco ER, Castle Hills ER (plus Urgent Care Ennis / Plano, newly added)

Each tracked credential is a card showing a status derived from its expiry date:
`expired` (past) / `critical` (≤30d) / `due soon` (≤90d) / `valid` / `on file` (no expiry).

I am **Imad Aijaz**, the compliance manager and owner. I am **not a developer** — explain things
plainly, make the routine technical decisions yourself, and tell me only what I actually need to
decide. Do not hand me long essays. Keep answers short.

---

## 2. Where everything lives

### Code
```
C:\sfmig\s
```
Never run git from inside a OneDrive folder.

- Frontend: `app.js`, `index.html`, `styles.css`, `data.json`, `data.js`
- Backend: `api/*.js` (Vercel serverless), `lib/*.js` (shared)
- Local tools: `tools/*.js`, `tools/install-schedule.ps1`

### Deployment
- **Live:** https://sentinel-compliance-kappa.vercel.app/  ← the one that matters
- Second project `sentinel-compliance-delta.vercel.app` is **frozen at cloud96** (a Vercel-side
  Git integration problem, not a code bug — needs reconnecting in the Vercel dashboard)
- Two git remotes, **push to BOTH**:
  - `kappa` → https://github.com/ImadAijaz/sentinel-compliance.git
  - `origin` → https://github.com/iticcotx/sentinel-compliance.git

**Deploy procedure (follow exactly):**
1. Bump the cache-buster in `index.html`: `?v=cloudNN` → `?v=cloudNN+1` (4 occurrences)
2. `git -c user.email=payroll@wcgtx.com -c user.name="Imad Aijaz" commit`
3. `git push kappa main:main` **and** `git push origin main:main` (slow, 3–5 min — run in background)
4. Vercel auto-builds kappa in 1–2 min. Verify: `curl` the site and grep for the new `cloudNN`.

**Current version: cloud116.**

### Microsoft 365 / OneDrive — there are TWO drives, do not confuse them
| Env var | Points at | Holds |
|---|---|---|
| `MS_DRIVE_ID` | Sama Farooqui's **personal OneDrive** | master roster Excel + app state (`_Sentinel/`) |
| `MS_DOCS_DRIVE_ID` | **CorporateArchivesDirectory** SharePoint team site | all documents the dashboard reads |

Graph credentials (`MS_CLIENT_ID` / `MS_CLIENT_SECRET` / `MS_TENANT`) exist **only in Vercel env
vars**. They are NOT on the local machine, so nothing local can call Microsoft Graph.

### Local synced folders (this is the important part)
Both of these are OneDrive-synced onto Imad's PC, so local file I/O works with no Graph API:

```
SOURCE (master, source of truth, ~28,000 files / 30 GB):
C:\Users\imada\Wellness & Care Group of Texas Inc\Sama Farooqui - WCGTX Phyicians_04.08.2020\

DESTINATION (what the dashboard reads):
C:\Users\imada\Wellness & Care Group of Texas Inc\Corporate Archives Directory - Documents\Sama Farooqui\Sentinel\
```

The master folder contains:
- `..WCGTX Master Physician File\` — 98 provider folders named `Last, First MD`, 16,920 files
- `..Frisco ER\` and `..Castle Hills ER\` — each has `..State Requirements_*` (facility compliance)
  plus `Nursing Staff` / `Front Office` / `Midshift` (staff credential folders)
- `..Urgent Care Ennis\`, `..Urgent Care Plano\` — staff folders
- `..WCGTX Master Rosters\WCGTX Physician Roster.xlsx` — **the live roster**, readable locally.
  Sheets: `WCGTX Credentials`, `WCGTX COI Roster`, `Inactive Providers`
- `_Sentinel\` — the app's own state files (roster_delta.json, scan_state.json, doc_dates.json, …)
- Many untracked sites (Dallas Urgent Cares, Hill Regional, Van Zandt, etc.)

---

## 3. How data flows

### Documents (automatic, every 15 minutes)
```
master folder
   ↓  Windows Scheduled Task "Sentinel Master Sync", every 15 min, 2 steps
   ↓  step 1: tools/sync-master.js --apply --quiet
   ↓          files each document into Corporate Archives\Sentinel\<organized>\
   ↓  step 2: tools/refresh-data.js --live
   ↓          reads the expiry out of each filename, publishes _Sentinel/doc_dates.json
   ↓  OneDrive uploads both to SharePoint
   ↓  /api/data applies doc_dates.json at request time (forward-only)
   ↓  dashboard polls every 90 s
dashboard card updates
```

**This runs on Imad's PC, not in the cloud.** Reason: Vercel's Hobby plan allows one cron run per
DAY and the project is already at its 12-serverless-function ceiling. Both folders are synced
locally, so the sync is plain file I/O and OneDrive does the upload.

Task settings: wake-to-run enabled, survives sleep and battery, skips if the previous run is still
going. It does **not** run while the PC is fully powered off.

### Destination folder structure the sync produces
```
Sentinel\
  Provider\<First Last>\<N. Phase>\<Category>\file
  State Readiness\<Facility>\<NN. Section>\file
  Staff\<Site>\<Last, First_ROLE>\<Category>\file
```

The 6 provider phases:
`1. Application & Document Collection`, `2. Primary Source Verification`,
`3. Background & Compliance Review`, `4. Medical Staff Review`,
`5. Payer Enrollment & Facility Setup`, `6. Approval & Ongoing Monitoring`

The 14 facility sections (`lib/facility.js` → `STATE_SECTIONS`):
`01. Licensing & Regulatory Compliance` … `14. Daily Readiness Walkthrough`

### Roster (Excel)
`data.json` is a baked snapshot (~10 MB). `lib/delta.js applyRosterDelta()` merges live roster
changes from `_Sentinel/roster_delta.json`, regenerated by `lib/regen.js`. Roster reads/writes use
**exceljs** (download → edit → upload), never the Graph `/workbook` API.

**Expiry dates come from FILENAMES, never from inside the PDF.** `lib/graph.js dateFromName()`.
The filing convention puts the **expiry** in the name: `COLA_02_26_2028_CHER.pdf` expires
26 Feb 2028. A certificate that prints its *issue* date must still be named with its expiry.

---

## 4. Commands

```bash
# Dry run the document sync — prints the plan, writes nothing
node tools/sync-master.js

# Do it
node tools/sync-master.js --apply

# Trace one document end to end
node tools/sync-master.js --grep=cola

# Include HR-restricted folders (see §6)
node tools/sync-master.js --apply --include-hr

# Read dates off the documents; publish to the app (no deploy needed)
node tools/refresh-data.js --live

# Bake dates into data.json instead (needs a deploy to take effect)
node tools/refresh-data.js --apply

# Install / remove the 15-minute schedule
powershell -ExecutionPolicy Bypass -File tools\install-schedule.ps1
powershell -ExecutionPolicy Bypass -File tools\install-schedule.ps1 -Remove
```

Sync log: `...Corporate Archives Directory - Documents\Sama Farooqui\Sentinel\_sync\sync-log.txt`

---

## 5. Rules that were learned the hard way — do not regress these

1. **"Absent from the master" is NOT supersession.** The master has 98 provider folders; the
   dashboard tracks 182. A draft that treated absence as obsolescence proposed archiving **6,919
   live credentials** (ACLS, ATLS, BLS). Supersession needs evidence: the master moved that exact
   filename into an archive folder, or a dated newer filing of the same document arrived.
2. **Archive-name sets must be per person.** Generic names (`Photo.pdf`, `CV.pdf`) recur, so a
   global set let one provider's archiving retire everyone else's copy.
3. **A filename in both a live and an archive folder** caused copy → retire → copy, churning
   forever. Skip anything the current run is also filing in.
4. **Never flatten sub-folders into the category.** Two same-named documents then fight over one
   destination and 60 files rewrite every run. Only genuinely colliding files keep their sub-folder.
5. **Modified-time is not supersession evidence.** A timestamp says when someone touched a file,
   not which version of a credential it is.
6. **Dates only ever move FORWARD.** A stale copy must never drag a current credential backwards
   into looking expired.
7. **Never rewrite a TRACKED credential's expiry to silence it.** 72 of 83 expired items whose
   proof file sits in an archive folder are real expired certificates — an old cert filed under
   `z.Expired Docs` with no replacement *is* the finding.
8. **Event documents do not expire.** Meeting minutes, agendas, inspection reports, service calls,
   "as of" snapshots, emails, incident reports carry the date the thing *happened*. Treating that
   as an expiry put 54 false "expired" cards on the board. See `lib/facility.js isNonExpiring()`.
9. **Archive folder names vary wildly**: `z.Expired Docs`, `Z_Expired Documents_FriscoER`,
   `.Expired Health Docs`, `Licenses Expired`, `zArchive`, `ZZZ Archive`, `z.Superseded`.
   One shared test: `lib/facility.js isArchivedSegment()`. Do NOT match `Z_Firetrol` (a live vendor).
10. **Provider folder names do not match the roster.** `Couch, Chris` → roster says *Christopher*;
    `Nguyen DO, Lien` → credential glued to the surname; `Mohiuddin, Mohammed Amer` → first and
    last reversed *in the roster*. Match on surname + first name, then surname + first initial.
    Exact-name matching alone silently dropped 12 active providers.
11. **One implementation only.** `api/digest.js` once kept a private copy of the roster regen,
    drifted from the shared one, and corrupted the board nightly. Shared logic lives in `lib/`.
12. **`_` is a word character.** A date regex anchored with `\b` will not match
    `COLA_02_26_2028_CHER.pdf`. Use `lib/graph.js dateFromName()`; don't write another one.

---

## 6. Open items / known gaps

**A. Document sync is ONE-WAY (master → Sentinel).**
A document uploaded through the app lands in Corporate Archives\Sentinel but never reaches the
master folder, where the credentialing team actually works. Over time the two drift apart. The
roster *does* write back to the master Excel correctly — this gap is documents only.
**This is the most important unfinished piece.**

**B. `HR ONLY` / `HR File` folders are excluded from the sync by default.**
1,007 files: WCGTX initial applications, verifications of employment, and physician agreements
containing compensation terms. The destination is a shared library feeding a dashboard several
people can open. `--include-hr` overrides. Separately: **19 HR-type files already reached the
shared library** before any of this — worth reviewing regardless.

**C. Governance risk — the master roster sits in a departed employee's personal OneDrive.**
`lib/graph.js` notes Sama Farooqui has left the company, yet `MS_DRIVE_ID` still points at her
personal OneDrive. If IT deletes that account the master roster goes with it. It should be moved
into the CorporateArchivesDirectory team library and `MS_DRIVE_ID` / `MS_ROSTER_PATH` repointed.

**D. 24/7 without the PC.** The 15-minute sync only runs while Imad's machine is on (it wakes from
sleep, but not from powered-off). True cloud-side 24/7 would mean the Vercel app reading the master
folder through Graph. A partial facility-only version exists at
`/api/data?facsync=preview|run|auto|save|sources|reorg|reorg-run` (see `lib/facsync.js`), wired
into the nightly digest cron and a silent no-op until a source folder link is saved. A GitHub
Actions schedule could drive it — but that needs `CRON_SECRET` set first (item E).

**E. `CRON_SECRET` is not set in Vercel.** The daily cron returns 401 without it.

**F. Rotate exposed secrets.** `MS_*` and `GMAIL_*` credentials were committed at some point.

**G. Two provider folders in the master are unmatched and deliberately not guessed:**
- `Okorie Ikechukwu MD` — no comma, so which token is the surname is genuinely ambiguous
- `Zz_Facility Documents_Odessa` — not a person, correctly ignored

**H. Automatic digest emails are DISABLED.** The daily cron only regenerates data.

---

## 7. Hard constraints

- **Vercel Hobby: 12 serverless functions MAX** — the project is AT the cap. Extend an existing
  `api/*.js` file; never add a new one.
- **Vercel Hobby cron: once per day**, and cron paths cannot carry a query string.
- **PowerShell 5.1** on this machine: no `&&`, no ternary, no `??`. Em-dashes and other non-ASCII
  in `.ps1` files cause parser errors — keep PowerShell scripts ASCII-only.
- A locked Excel file returns **HTTP 423** — close the workbook before add/delete/sync.
- `PowerShell -replace` mangles `index.html`; use a proper editor/tool for it.
- The Sentinel app has no test framework. Verification has been ad-hoc Node scripts run against the
  real synced library and the real `data.json`. **Always dry-run before writing anything** — that
  practice is what caught the 6,919-credential mistake.

---

## 8. Recent history

| Version | What landed |
|---|---|
| cloud111 | Facility document sync via Graph; event documents stop being scored as credentials (facility "expired" 65 → 9); shared archive-folder test |
| cloud112 | `tools/sync-master.js` — local 15-min master-folder sync; filed 4,092 documents; scanner learns `Staff/` paths and non-hardcoded facility sites |
| cloud113 | `z.Superseded` recognised as an archive folder |
| cloud114 | Removed the manual "Sync documents" button — syncing is automatic |
| cloud115 | `tools/refresh-data.js` — read real dates off the documents; 380 credentials moved forward (COLA 2026-01-17 → 2028-02-26); expired 1,592 → 1,512 |
| cloud116 | Dates published to `_Sentinel/doc_dates.json` and applied by `/api/data` at request time — no deploy needed for a renewal to show |
