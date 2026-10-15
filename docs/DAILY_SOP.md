# Daily SOP — EA Command Center

The app only works if it becomes a habit. This SOP is built around
natural transition points in an EA's day — not around the app's tabs.

**The core principle:** touch the app at transitions, not constantly.

---

## Morning Routine (15–20 minutes)

**Trigger:** You sit down with coffee. Before you open email.

### Step 1 — Open the app (1 min)

- Launch from `run.bat` (Windows) or `venv/bin/python -m streamlit run app.py`
- Land on Dashboard
- Scan the top six metrics:
  - 📅 Meetings today
  - 🎉 Events today
  - 🎤 Speaking today
  - ✅ Tasks (TickTick)
  - ⚠️ Overdue (follow-ups + tasks combined)
  - 📞 Follow-ups due today

### Step 2 — Triage Needs Attention (5–10 min)

Open **⚠️ Needs Attention**. Work through each category that has content:

- **🔴 Overdue Follow-ups** — open the contact, log the call or email
  you owed them, set a new next action
- **⏳ Meetings Awaiting Response** — send the nudge from your email
  client, update the attendee status in the app
- **✅ Action Items Due (7d)** — anything you can close today?
- **⏳ Overdue RSVPs** — check email, update guest statuses
- **💳 Payments Due (7d)** — note them, handle after email triage
- **✍️ Unsigned Contracts (7d)** — ping the vendor, log the contact
- **💰 Events Over 90% Budget** — flag for the executive
- **📦 Deliverables Due (7d)** — speaking prep: bio, slides, A/V
- **📤 CFPs Closing (14d)** — submit or defer
- **📮 Thank-Yous Overdue** — send and log
- **📝 Meetings Without Notes** — link notes docs

**Rule:** Handle overdue items before anything new.

### Step 3 — Check today's schedule (2 min)

Open **📅 Today's Schedule**. For each item:

- Do I have the agenda?
- Is the notes doc linked?
- Are attendees confirmed?
- Do I need to prep anything?

Flag anything missing in your head or on paper.

### Step 4 — Open email (10 min)

Process your inbox as normal. But this rule applies:

**Every email related to a meeting, event, vendor, or speaking
engagement gets logged in the app immediately after sending or
replying.**

Don't batch. Batching creates a backlog you'll never clear.

### Step 5 — Confirm the day (1 min)

Close the app. You know exactly what today looks like.

---

## Throughout the Day

**Rule:** Log immediately after the interaction. Never batch.

| Event                 | Action                                       | Time    |
| --------------------- | -------------------------------------------- | ------- |
| Send a meeting invite | Meeting → attendee → mark "sent"             | 10 sec  |
| Get an RSVP reply     | Meeting/event → update attendee/guest status | 10 sec  |
| Call a contact        | Contact → 📞 → log communication              | 15 sec  |
| Email a vendor        | Event → engagement → log it                  | 15 sec  |
| Confirm a vendor      | Update engagement status to "confirmed"      | 30 sec  |
| Pay a vendor          | Add payment (paid date)                      | 30 sec  |
| Submit a CFP          | Speaking → CFP → update status               | 30 sec  |
| Complete a task       | TickTick                                     | instant |
| Send a thank-you      | Speaking → post-engagement → mark sent       | 20 sec  |

**Friction reduction:**

- Keep the app open in a tab or second monitor
- Use the 📞 button on contact rows for one-click logging
- Use "Open" buttons from Dashboard alerts to jump to records

---

## End of Day (5–10 minutes)

**Trigger:** You're wrapping up. Before you close your laptop.

### Step 1 — Clear the Dashboard (3–5 min)

Open Dashboard. Work through:

- Any remaining **Needs Attention** items? Handle or push to tomorrow.
- Any action items still due today? Mark done or reschedule.

### Step 2 — Log missed items (2 min)

Scan your sent email folder from today. Anything related to a meeting,
event, vendor, or engagement that isn't logged? Log it now.

This is your safety net. If you logged throughout the day, this takes
30 seconds.

### Step 3 — Preview tomorrow (2 min)

- Check tomorrow's meetings
- Check tomorrow's follow-ups
- Check TickTick for tomorrow's tasks
- Note anything that needs prep

**Rule:** Don't close the app until the Dashboard's "Today" section is
empty or everything has been rescheduled.

---

## Weekly Review (30–45 minutes)

**Trigger:** Friday afternoon, or Monday morning. Pick one and stick
to it.

### Step 1 — Review Dashboard (10 min)

Walk through every alert category. Fix what's fixable. Flag what needs
follow-up next week.

### Step 2 — Sync TickTick (2 min)

Sidebar → **✅ TickTick** → **🔄 Sync Now**. This pulls any tasks
created outside the app.

### Step 3 — Meeting hygiene (5 min)

Meetings tab → filter by "Without notes". Any completed meetings
missing notes? Link them now or archive.

### Step 4 — Event and speaking check (10 min)

- **Events tab** → upcoming view. Any overdue RSVPs? Any unsigned
  contracts? Any payments due?
- **Speaking tab** → pipeline view. Any deliverables overdue? Any
  CFPs closing next week? Any thank-yous owed?

### Step 5 — Contact hygiene (5 min)

**Contacts tab** → expand **💬 Relationship Health**. Review:

- Gone quiet 60+ days — anyone you should reach out to?
- Next actions due this week — anyone you'd forgotten?

### Step 6 — Plan next week (5 min)

- Check next week's meetings
- Check next week's events
- Check next week's speaking engagements
- Check next week's deadlines (CFPs, deliverables, payments)
- Add prep tasks to TickTick

**End state:** You start Monday with a clear picture.

---

## Monthly Review (60 minutes)

**Trigger:** Last Friday of the month, or first business day.

### Step 1 — Backup (10 min)

- App shows the monthly reminder in the last 5 days of the month
- Download the `.db` backup
- Save to an external drive or cloud folder
- Click **✅ Mark as Backed Up** in the reminder

**This is non-negotiable.** Auto-backups protect against crashes; this
protects against disk failure.

### Step 2 — Completed work review (20 min)

**Events:**

- Any completed this month?
- Debrief notes done?
- Budget reconciled?
- Final thank-yous sent?

**Speaking:**

- Any delivered this month?
- Thank-yous sent?
- Feedback collected?
- Testimonials saved?
- Recordings received?

**Meetings:**

- Any recurring meetings that should be retired?
- Any meetings consistently without agendas?

**Vendors:**

- Any unpaid invoices?
- Any vendors you'd rate differently?

**Budget:**

- Any events over budget?
- Any patterns?

### Step 3 — Cleanup (15 min)

- Archive completed meetings older than 30 days (delete if no longer
  needed)
- Archive completed events
- Close stale communications ("no reply" after 30 days)
- Update vendor directory (add new vendors, mark preferred)
- Review "gone quiet 90+ days" contacts

### Step 4 — Plan next month (15 min)

- Check next month's events
- Check next month's speaking engagements
- Check CFP deadlines
- Check budget cycles
- Add major tasks to TickTick

---

## When You Fall Behind

You will. Here's the recovery SOP.

**If you miss a day:**
- Don't backfill everything
- Log only what's still relevant
- Skip yesterday's minor interactions

**If you miss a week:**
- Do a weekly review
- Log only active items
- Archive or close stale items
- Don't try to reconstruct history

**If you miss a month:**
- Do a monthly review
- Run backup
- Reset: log only current and future items
- Accept that history is incomplete

**The rule:** The app is for the present and future, not the past. Don't
drown in backlog.

---

## Minimum Viable SOP

If the full SOP is too much, do only this:

- **Every morning:** open app, scan Dashboard, handle overdue items
- **Every time you send an email:** log it
- **Every evening:** clear Dashboard
- **Every Friday:** weekly review
- **Every month:** backup

Everything else is optimization.

---

## Habit Stacking

Attach new habits to existing ones:

| Existing habit       | New habit                  |
| -------------------- | -------------------------- |
| Sit down with coffee | Open app, scan Dashboard   |
| Send an email        | Log it in the app          |
| Get a reply          | Update status              |
| Finish a meeting     | Log notes and action items |
| End of day           | Clear Dashboard            |
| Friday afternoon     | Weekly review              |
| Last Friday of month | Backup + monthly review    |

**The rule:** If you do X, you also do Y. No exceptions for the first
30 days. After that, it's automatic.

---

## 30-Day Adoption Plan

**Week 1 — Meetings only**

- Log every meeting
- Add attendees with statuses
- Log action items
- Do prep briefs for key meetings
- Don't touch events, speaking, or vendors yet

**Week 2 — Add events**

- Log every event
- Build guest lists
- Import RSVPs from Google Forms
- Keep meetings running

**Week 3 — Add speaking + contacts**

- Log speaking engagements into the pipeline
- Log every meaningful contact interaction
- Set next actions
- Keep meetings and events running

**Week 4 — Add TickTick, mobile export, backups**

- Connect TickTick
- Map tags to executives
- Try a mobile export
- Run a backup

**End of Week 4:**

- Run monthly backup
- Review what's working
- Drop what isn't
- Double down on what is

---

## Signs the SOP Is Working

- You open the app before email
- You log interactions without thinking
- You check the Dashboard before meetings
- You know what's pending without searching
- The backup reminder doesn't surprise you
- TickTick and the app feel like one system
- You stop using spreadsheets
- You stop searching your inbox for "when did I last contact..."

## Signs the SOP Is Failing

- You skip logging for 3+ days
- You avoid opening the app
- You're back to spreadsheets
- You can't remember the last backup
- TickTick and the app disagree
- You're logging retroactively more than in real time
- The Dashboard is full of red

**If failing:** Cut scope. Do only the Minimum Viable SOP for two weeks.
Add back slowly.

---

## One-Line Summary

Touch the app at transitions: **morning triage, log-as-you-send,
end-of-day clear, weekly review, monthly backup**. Log immediately, not
in batches. Start with meetings only, add modules over 30 days. When
you fall behind, log only what's current and let the past go.