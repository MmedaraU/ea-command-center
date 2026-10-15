# EA Command Center

A local-first command center for a single Executive Assistant supporting
multiple executives across multiple companies. Runs on Streamlit + SQLite.
No cloud. No AI. No email sending. All data stays on your machine.

The app is the coordination and memory layer. Google Workspace and
Microsoft 365 continue to handle email, calendar, documents, video, and
forms. The app links out to those tools and tracks what they don't:
multi-executive context, pending responses, event budgets, speaking
pipelines, and relationship memory.

---

## Contents

- [EA Command Center](#ea-command-center)
  - [Contents](#contents)
  - [What It Does](#what-it-does)
  - [Requirements](#requirements)
    - [macOS](#macos)
    - [Ubuntu / Debian](#ubuntu--debian)
    - [Windows](#windows)
  - [Install](#install)
  - [Run](#run)
  - [First-Time Setup](#first-time-setup)
  - [Daily Workflow](#daily-workflow)
    - [Morning (15–20 minutes)](#morning-1520-minutes)
    - [Throughout the day](#throughout-the-day)
    - [End of day (5–10 minutes)](#end-of-day-510-minutes)
    - [Weekly (30–45 minutes)](#weekly-3045-minutes)
    - [Monthly (60 minutes)](#monthly-60-minutes)
  - [Modules](#modules)
    - [Dashboard](#dashboard)
    - [Meetings](#meetings)
    - [Events](#events)
    - [Speaking Engagements](#speaking-engagements)
    - [Trips](#trips)
    - [Contacts](#contacts)
    - [Directory](#directory)
    - [Library](#library)
  - [Prep Briefs](#prep-briefs)
  - [Mobile Export](#mobile-export)
  - [TickTick Integration](#ticktick-integration)
  - [Backup \& Restore](#backup--restore)
    - [Auto-backups](#auto-backups)
    - [Monthly reminder](#monthly-reminder)
    - [Offsite backups](#offsite-backups)
    - [Restore](#restore)
  - [File Structure](#file-structure)
  - [Troubleshooting](#troubleshooting)
  - [What This App Deliberately Does Not Do](#what-this-app-deliberately-does-not-do)
  - [License](#license)

---

## What It Does

**Tracks, links, and remembers across eight areas of EA work:**

- **Meetings** — records, attendees, action items, decisions, prep briefs
- **Events** — guest lists, RSVP imports from Google Forms, vendors, budget
- **Speaking engagements** — pipeline from invite to post-event, CFP
  tracking, deliverables, logistics
- **Trips** — multi-city itineraries, delegation, per diem, packing,
  weather, travel packs
- **Contacts** — directory with relationship memory
- **Directory** — companies, venues, vendors
- **Library** — destination guides, visa rules, templates, emergency
  information
- **Dashboard** — today at a glance, needs attention, what's on, search

**Cross-cutting features:**

- **Communications memory** — log every meaningful interaction; see last
  contacted, channel, next action
- **Cross-linking** — link meetings, events, and engagements to trips
- **Global search** — search across all modules from the Dashboard
- **Prep briefs** — HTML, PDF, and Word documents for meetings, events,
  and speaking engagements
- **Mobile export** — one-click encrypted HTML file for phone access
- **TickTick integration** — push and pull tasks, tag-to-executive mapping
- **Backup & restore** — hourly auto-backups, monthly reminder, restore log
- **CSV import/export** — every entity

---

## Requirements

- **Python 3.10+**
- **SQLite** (bundled with Python)
- **WeasyPrint system libraries** (for PDF briefs and travel packs)
- **cryptography** (optional — for encrypted mobile exports)
- **requests** (for TickTick and weather)

### macOS

```bash
brew install pango libffi
```

### Ubuntu / Debian

```bash
sudo apt install libpango-1.0-0 libpangoft2-1.0-0
```

### Windows

Install MSYS2 from <https://www.msys2.org/> (default location
`C:\msys64`), then open the **UCRT64** shell and run:

```bash
pacman -Syu
pacman -S mingw-w64-ucrt-x86_64-pango \
          mingw-w64-ucrt-x86_64-gdk-pixbuf2 \
          mingw-w64-ucrt-x86_64-cairo
```

The bundled `run.bat` adds `C:\msys64\ucrt64\bin` to `PATH` and sets
`WEASYPRINT_DLL_DIRECTORIES` so Python can find these DLLs. If your
MSYS2 lives elsewhere, edit `run.bat` accordingly.

Without this step, PDF exports fail with an error like
`OSError: cannot load library 'gobject-2.0-0'`.

---

## Install

Clone the repo, then:

```bash
python -m venv venv

# Windows:
venv\Scripts\activate
# macOS / Linux:
source venv/bin/activate

pip install -r requirements.txt
```

---

## Run

**Windows:** double-click `run.bat`

**macOS / Linux:**

```bash
venv/bin/python -m streamlit run app.py
```

Opens at `http://localhost:8501`. On first launch, `ea_command_center.db`
is created along with all tables.

---

## First-Time Setup

1. Open the app.
2. In the sidebar, click **➕ Add New Executive**.
3. Add a company (if none exist).
4. Add an executive: name, email, timezone, travel preferences.
5. Optionally add passports and memberships from the Full Profile view.

You're ready to use any module. The Directory and Library tabs work
without an executive.

---

## Daily Workflow

### Morning (15–20 minutes)

1. Open the app before email.
2. Scan the Dashboard's top six metrics.
3. Work through **Needs Attention**: overdue follow-ups, awaiting
   responses, payments due, unsigned contracts, over-budget events.
4. Check today's schedule.
5. Note any prep needed for today's events or engagements.

**Rule:** Don't open email before the Dashboard scan.

### Throughout the day

Log immediately after the interaction. Never batch.

- Send a meeting invite → log it on the attendee
- Get an RSVP → update status
- Call a contact → log communication
- Confirm a vendor → update engagement status
- Pay a vendor → record payment
- Submit a CFP → update CFP status

Target: 5–15 seconds per log.

### End of day (5–10 minutes)

1. Open Dashboard.
2. Clear remaining alerts.
3. Update incomplete action items.
4. Log anything missed.
5. Preview tomorrow's schedule.

### Weekly (30–45 minutes)

Friday afternoon or Monday morning.

1. Review Dashboard alerts.
2. Sync TickTick (sidebar → TickTick → Sync Now).
3. Review overdue deliverables (speaking).
4. Review next week's events and engagements.
5. Check payments due next week.
6. Review "gone quiet" contacts in the Contacts tab.
7. Plan next week's travel if any.

### Monthly (60 minutes)

Last Friday of the month.

1. Download the `.db` backup.
2. Store it offsite (external drive or cloud folder).
3. Click "Mark Done" in the reminder.
4. Review completed events — debrief notes done?
5. Review delivered engagements — thank-yous sent?
6. Review over-budget events — why?
7. Archive meetings older than 30 days.
8. Plan next month's events and engagements.

---

## Modules

### Dashboard

Single-screen overview:

- **Today** — six top metrics and today's schedule
- **Needs Attention** — grouped alerts (overdue follow-ups, awaiting
  responses, RSVPs, payments, contracts, over-budget events,
  deliverables, CFPs, thank-yous, meetings without notes)
- **What's On** — next 14 days, collapsed
- **Global search** — across meetings, events, speaking engagements,
  trips, contacts, and vendors
- **TickTick** — open tasks, due this week, overdue (collapsed)
- **Travel Overview** — executives, companies, contacts, spending,
  library counts (collapsed)

### Meetings

- Records with type, status, date, time, duration, location, video link
- Attendees with per-person status (pending, accepted, declined,
  tentative, no response, waitlist)
- VIP and internal/external flags
- Native quick notes plus links to Google Docs or Notion
- Action items with owners and due dates
- Decisions log
- Follow-up date
- **Prep brief** — HTML, PDF, Word with attendees, agenda, recent
  context, prior open items, decisions, action items
- Statuses: draft → request sent → pending → confirmed → completed,
  plus declined, tentative, no response, rescheduled, cancelled
- Auto-log attendance when a meeting is marked completed (creates a
  communication entry per attendee)

### Events

- Records with type, status, date, start/end time, venue, headcount,
  budget, currency
- Guest list with invite status (not invited, invited, pending,
  accepted, declined, waitlist, cancelled, checked in)
- Plus-ones, dietary, allergies, accessibility, badge names, seating
  preferences
- **Google Forms RSVP import** — wizard with column mapping, status
  mapping, guest matching (email → name), and idempotent re-import
- Vendors engaged per event with quote, agreed amount, contract,
  payment schedule
- Budget line items with estimated, actual, and paid amounts
- Auto-linking: confirming a vendor creates or updates a budget line
  item; recording a payment updates the paid amount
- **Prep brief** — guest summary, confirmed guests, awaiting RSVPs,
  vendors, budget with alerts
- Over-budget alerts at 90% and 100%

### Speaking Engagements

- Full pipeline: invited → interested → CFP submitted → under review
  → accepted → negotiating → confirmed → prep → ready → delivered
  → follow-up → closed
- CFP tracking with deadlines, proposal docs, and decision status
- Negotiation fields: fee offered, fee agreed, currency, travel and
  accommodation coverage, recording rights, exclusivity
- Five core deliverables: bio, headshot, slides, A/V requirements,
  intro script
- Logistics saved to Quick Notes (travel, accommodation, venue,
  green room, tech check, slides due)
- Post-engagement: thank-you, feedback, rating, testimonial,
  recording, would-repeat
- **Prep brief** — fee terms, topic, deliverables, CFP status
- Pipeline view (grouped by status) and list view
- Dashboard alerts for deliverables due, CFPs closing, thank-yous
  overdue

### Trips

Two sub-tabs:

**All Trips**

- Filter by executive, date range, search
- Trip cards with purpose, destination, executive, status, and
  financials
- Mass select and delete
- Export to CSV, Word, Excel, or `.ics` calendar

**New Trip**

- Multi-city stop builder with ↑ ↓ reorder
- Itinerary items with type, cost, currency, cost date, timezone,
  venue, delegation members, local support contacts
- Bulk paste from spreadsheet (auto header detection)
- Inline venue add
- Per-item duplicate, edit, and delete
- Travelers, local support, receipt folder URL
- Timezone display toggle (Home vs Destination)
- Weather preview for each stop

### Contacts

- Grouped by type: Emergency, Local Support, Staff, Partner, Other
- Tags with color-coded pills
- Custom fields: role, phone, email, country, city, notes, active
  status
- **Communications memory** — timeline per contact, last contacted,
  channel, next action, next action date
- **Relationship Health** panel: gone quiet (30/60/90 days), next
  actions due
- CSV import and export
- Default contacts per company (auto-included in new trips)

### Directory

Three sub-tabs:

**Companies** — name, cost center, policy notes, default contacts,
active status. Export profile to HTML, Word, or Excel.

**Venues** — name, address, city, country, WiFi credentials, badge
info, dress code, notes. Attachable to session-type itinerary items.

**Vendors** — name, category, contact, email, phone, website,
address, city, country, rating, preferred flag. Reusable across
events. Engagement history per vendor.

### Library

Six sub-tabs:

- **Trip Templates** — save any trip as a template, create new trips
  from templates
- **Destination Guides** — language, currency, emergency numbers,
  etiquette, phrases, packing tips, connectivity, recommended apps
- **Visa Rules** — passport nationality → destination pairs with
  visa requirement, max stay, processing time, fee
- **Checklist Templates** — reusable checklists, applied per trip
- **Packing Templates** — reusable packing lists, applied per traveler
- **Emergency Directory** — hospitals, emergency numbers per country,
  embassies and consulates

---

## Prep Briefs

Every meeting, event, and speaking engagement can generate a brief in
three formats:

| Format   | Use                                                    |
| -------- | ------------------------------------------------------ |
| **HTML** | Self-contained, printable from any browser, links work |
| **PDF**  | Print, send to executives, archive                     |
| **Word** | Edit, annotate, drop into corporate templates          |

Generate from the record's dialog → **📋 Prep Brief** → **Generate
Brief**. Preview inline or download any of the three formats.

Briefs include all relevant context: attendees, agenda, recent
communications, prior open action items, decisions, budget status,
guest summaries, vendor status, fee terms, deliverables, and CFP
status.

---

## Mobile Export

Sidebar → **📱 Mobile Export**.

Four scopes:

- **Full** — everything
- **Today only** — today's schedule plus a 14-day trip window
- **This event** — one event with guests and budget
- **This trip** — one trip with stops, itinerary, and linked items

Optionally set a password to encrypt the export with AES-256-GCM.
Download the HTML, email it to yourself, open on your phone. Works
offline. Can be added to home screen as a PWA.

Features:

- Sticky tabs: Today, Meetings, Events, Speaking, Trips, Contacts
- Search box that filters across the current tab
- Tap-to-call, tap-to-email, tap-to-open-links
- Dark mode follows OS preference
- Print-friendly stylesheet

---

## TickTick Integration

Register a TickTick developer app:

1. Go to <https://developer.ticktick.com/manage>
2. Create a new app
3. Redirect URI: `http://localhost:8502/callback`
4. Copy the Client ID and Secret

In the sidebar → **✅ TickTick**:

1. Paste Client ID and Secret
2. Click **🔗 Connect TickTick**
3. Click the authorize link → complete authorization on TickTick
4. Back in the app, click **🔄 Check Connection**

Once connected:

- **Push action items** — from any meeting, click ↗️ on an action item
- **Push next actions** — from any contact's communication dialog
- **Sync Now** — pulls all tasks from TickTick into the local mirror
- **Tag mapping** — map TickTick tags to executives so tasks display
  correctly in the Dashboard
- **Complete in either app** — completing a task in TickTick reflects
  on the linked local action item after the next sync

---

## Backup & Restore

Three layers:

### Auto-backups

Every hour, on app start, if the newest backup is older than 60
minutes. Stored in `.backups/`. Kept 60 deep. **Crash recovery only**
— will not survive disk failure.

Manual trigger: sidebar → Export Database → **🔄 Run Auto-Backup Now**.

### Monthly reminder

Fires in the last 5 days of each month if the last verified backup was
more than 25 days ago. Persistent until you click **✅ Mark as Backed
Up**.

### Offsite backups

The only thing that survives disk failure. Download the `.db` file
from the reminder or from sidebar → Export Database → **⬇️ Backup
.db**, then store on Dropbox, iCloud, an external drive, or a NAS.

### Restore

Sidebar → **💾 Import / Restore Database**.

- **Merge** — adds new data to existing
- **Replace** — full restore from a `.db` file

Before any Replace-mode restore, the app automatically creates a
pre-restore backup. Every restore is logged (sidebar → Export Database
→ 📜 Restore Log).

---

## File Structure

```
ea-command-center/
├── app.py                      # Main Streamlit app
├── database.py                 # Schema, migrations, all DB functions
├── brief_generator.py          # Meeting/event/speaking briefs (HTML, PDF, Word)
├── mobile_export.py            # Mobile HTML export with optional encryption
├── ticktick_sync.py            # TickTick API client
├── doc_generator.py            # Travel pack, profiles
├── excel_export.py             # Excel exports
├── currency.py                 # Exchange rates
├── duplicate_detection.py      # Fuzzy matching
├── utils.py                    # Timezone, ICS, bulk paste parsers
├── weather.py                  # Weather integration
├── templates/                  # Jinja2 templates
├── docs/                       # Documentation, screenshots
│   ├── DAILY_SOP.md
│   └── screenshots/
├── run.bat                     # Windows launcher
├── requirements.txt
├── ea_command_center.db        # Created on first run
├── .backups/                   # Auto-backups
├── ea_app_state.json           # Runtime state (backup reminder)
├── ea_dismissed_warnings.json  # Acknowledged warnings
└── ea_receipts/                # Receipt uploads
```

**Database highlights:**

- Automatic migrations — schema upgrades run on every startup
- WAL mode — a single writer, many readers, no blocking
- Foreign key cascades — deleting a trip removes its stops, items,
  delegation, expenses, packing lists, and checklists
- Indexes on all foreign keys and frequently filtered columns

---

## Troubleshooting

| Issue                                                                    | Fix                                                                                                                  |
| ------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------- |
| App won't start                                                          | Confirm `python` and `venv` are set up. Check the terminal for a traceback.                                          |
| `No module named 'streamlit'`                                            | Activate the venv, then `pip install -r requirements.txt`.                                                           |
| Port 8501 is busy                                                        | `streamlit run app.py --server.port 8502`.                                                                           |
| Port 8502 is busy (TickTick callback)                                    | Close whatever else is using it, then reconnect.                                                                     |
| Database is locked                                                       | Only one instance should be running. WAL mode reduces this to near-zero.                                             |
| PDF export fails                                                         | Check WeasyPrint system libraries. On Windows, use `run.bat`.                                                        |
| `OSError: cannot load library 'gobject-2.0-0'`                           | MSYS2's UCRT64 `bin` folder isn't on `PATH`. Use `run.bat`, or set `PATH` and `WEASYPRINT_DLL_DIRECTORIES` manually. |
| `TypeError: PDF.__init__() takes 1 positional argument but 2 were given` | WeasyPrint and pydyf are a mismatched pair. Fix: `pip install --upgrade --force-reinstall weasyprint pydyf`.         |
| TickTick 401                                                             | Disconnect, reconnect, re-authorize.                                                                                 |
| Mobile export says "Plaintext"                                           | Install `cryptography`: `pip install cryptography`.                                                                  |
| Backup reminder won't dismiss                                            | Click "Mark as Backed Up". If it doesn't clear, check write permissions on `ea_app_state.json`.                      |
| `StreamlitDuplicateElementKey`                                           | Clear Streamlit cache (press **C** in the browser) or restart the app.                                               |
| Import fails                                                             | Check the JSON is from a matching version. `.db` files must be raw SQLite from this app.                             |

---

## What This App Deliberately Does Not Do

- Send email
- Manage a calendar UI
- Host video calls
- Edit documents (links out to Google Docs / Notion)
- Build forms (uses Google Forms)
- Store files (only links)
- Use AI
- Sync to the cloud

Those live in Google Workspace or Microsoft 365. The app coordinates
and remembers.

---

## License

Internal use. Modify as needed.