"""
Prep brief generators.

Produces self-contained HTML documents for meetings, events, and
speaking engagements. Each brief is a single file: inline CSS, no
external assets, printable to PDF from any browser.
"""

from datetime import datetime
import database as db
import io

try:
    from weasyprint import HTML as WeasyHTML
    WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


# =========================================================
# SHARED HELPERS
# =========================================================


BRIEF_CSS = """
<style>
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
                 Roboto, Helvetica, Arial, sans-serif;
    color: #1a1a1a;
    background: #f7f7f7;
    margin: 0;
    padding: 24px;
    line-height: 1.5;
    -webkit-font-smoothing: antialiased;
  }
  .brief {
    max-width: 900px;
    margin: 0 auto;
    background: #ffffff;
    padding: 40px;
    border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
  }
  h1 { font-size: 26px; margin: 0 0 6px 0; line-height: 1.2; }
  h2 {
    font-size: 16px;
    margin: 28px 0 10px 0;
    padding-bottom: 6px;
    border-bottom: 2px solid #e5e5e5;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: #374151;
  }
  h3 { font-size: 15px; margin: 16px 0 6px 0; color: #111; }
  .subtitle { color: #6b7280; font-size: 14px; margin-bottom: 24px; }
  .badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 500;
    background: #e5e7eb;
    color: #374151;
    margin-right: 4px;
    margin-bottom: 4px;
  }
  .badge-green { background: #d1fae5; color: #065f46; }
  .badge-yellow { background: #fef3c7; color: #92400e; }
  .badge-red { background: #fee2e2; color: #991b1b; }
  .badge-blue { background: #dbeafe; color: #1e40af; }
  .badge-purple { background: #ede9fe; color: #5b21b6; }
  .meta {
    display: flex;
    flex-wrap: wrap;
    gap: 24px;
    margin: 16px 0 8px 0;
    padding: 16px;
    background: #f9fafb;
    border-radius: 6px;
  }
  .meta-item { font-size: 13px; }
  .meta-item strong {
    display: block;
    color: #6b7280;
    font-weight: 500;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 2px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 8px 0;
    font-size: 13px;
  }
  th, td {
    text-align: left;
    padding: 8px 10px;
    border-bottom: 1px solid #f0f0f0;
    vertical-align: top;
  }
  th {
    background: #f9fafb;
    font-weight: 600;
    color: #374151;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  tr:last-child td { border-bottom: none; }
  .note {
    background: #f9fafb;
    padding: 12px 16px;
    border-left: 3px solid #3b82f6;
    font-size: 13px;
    margin: 8px 0;
    border-radius: 0 4px 4px 0;
  }
  .empty {
    color: #9ca3af;
    font-style: italic;
    font-size: 13px;
    padding: 8px 0;
  }
  ul { margin: 8px 0; padding-left: 20px; }
  li { margin-bottom: 4px; font-size: 13px; }
  a { color: #2563eb; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .footer {
    margin-top: 40px;
    padding-top: 16px;
    border-top: 1px solid #f0f0f0;
    font-size: 11px;
    color: #9ca3af;
    text-align: center;
  }
  .pill-list { display: flex; flex-wrap: wrap; gap: 4px; margin: 4px 0; }
  @media print {
    body { background: #ffffff; padding: 0; }
    .brief { box-shadow: none; padding: 20px; border-radius: 0; }
    h2 { page-break-after: avoid; }
    table { page-break-inside: auto; }
    tr { page-break-inside: avoid; }
  }
</style>
"""


def _esc(s):
    if s is None:
        return ""
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _fmt_date(d):
    if not d:
        return "—"
    try:
        return datetime.fromisoformat(str(d)).strftime("%a, %d %b %Y")
    except Exception:
        return str(d)


def _badge(text, kind=None):
    kind_class = f" badge-{kind}" if kind else ""
    return f'<span class="badge{kind_class}">{_esc(text)}</span>'


def _meta_row(items):
    """items: list of (label, value) tuples."""
    cells = "".join(
        f'<div class="meta-item"><strong>{_esc(label)}</strong>'
        f'{_esc(value) or "—"}</div>'
        for label, value in items
    )
    return f'<div class="meta">{cells}</div>'


def _link_or_text(url, label=None):
    if not url:
        return ""
    display = label or url
    return f'<a href="{_esc(url)}">{_esc(display)}</a>'


def _status_badge_kind(status):
    """Map status string to a badge color hint."""
    s = (status or "").lower()
    if s in ("confirmed", "accepted", "signed", "completed", "delivered",
             "done", "paid", "ready", "closed", "checked_in"):
        return "green"
    if s in ("pending", "in_progress", "negotiating", "prep_in_progress",
             "tentative", "under_review", "quoted", "committed", "waitlist"):
        return "yellow"
    if s in ("declined", "cancelled", "rejected", "no_response", "overdue",
             "blocked"):
        return "red"
    if s in ("invited", "request_sent", "draft", "interested", "estimated",
             "planning", "shortlisted", "not_started"):
        return "blue"
    return None


def _footer():
    return (
        f'<div class="footer">'
        f'Generated {datetime.now().strftime("%d %b %Y · %H:%M")} · '
        f'EA Command Center'
        f'</div>'
    )


def _wrap_html(title, body):
    return (
        f'<!DOCTYPE html>\n<html lang="en">\n<head>\n'
        f'<meta charset="UTF-8">\n'
        f'<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f'<title>{_esc(title)}</title>\n'
        f'{BRIEF_CSS}\n'
        f'</head>\n<body>\n'
        f'<div class="brief">\n{body}\n</div>\n'
        f'</body>\n</html>'
    )


# =========================================================
# MEETING BRIEF
# =========================================================


def generate_meeting_brief_html(meeting_id):
    meeting = db.get_meeting(meeting_id)
    if not meeting:
        return None

    exec_id = meeting["exec_id"]
    attendees = db.get_meeting_attendees(meeting_id)
    action_items = db.get_meeting_action_items(meeting_id)
    decisions = db.get_meeting_decisions(meeting_id)

    # Split actions
    open_actions = [
        a for a in action_items
        if (a.get("status") or "not_started") != "done"
    ]

    # ---- Header ----
    status_kind = _status_badge_kind(meeting.get("status"))
    header = (
        f'<h1>{_esc(meeting.get("title"))}</h1>'
        f'<div class="subtitle">'
        f'{_badge(meeting.get("meeting_type") or "Meeting")} '
        f'{_badge(meeting.get("status") or "draft", status_kind)}'
        f'</div>'
    )

    # ---- Meta row ----
    time_str = meeting.get("meeting_time") or ""
    dur = (
        f'{meeting["duration_minutes"]} min'
        if meeting.get("duration_minutes") else ""
    )
    time_full = " · ".join(x for x in [time_str, dur] if x) or "—"
    meta = _meta_row([
        ("Executive", meeting.get("executive_name") or "—"),
        ("Date", _fmt_date(meeting.get("meeting_date"))),
        ("Time", time_full),
        ("Location", meeting.get("location") or "—"),
    ])

    # ---- Links ----
    links = []
    if meeting.get("calendar_event_url"):
        links.append(
            f'📅 {_link_or_text(meeting["calendar_event_url"], "Calendar event")}'
        )
    if meeting.get("video_link"):
        links.append(f'🎥 {_link_or_text(meeting["video_link"], "Video link")}')
    if meeting.get("notes_url"):
        label = meeting.get("notes_label") or "Notes doc"
        links.append(f'📝 {_link_or_text(meeting["notes_url"], label)}')
    links_html = (
        f'<h2>Links</h2><div class="pill-list">'
        + "".join(f'<span class="badge">{x}</span>' for x in links)
        + '</div>'
    ) if links else ""

    # ---- Agenda ----
    agenda_html = ""
    if meeting.get("agenda"):
        agenda_html = (
            f'<h2>Agenda</h2>'
            f'<div class="note">{_esc(meeting["agenda"]).replace(chr(10), "<br>")}</div>'
        )

    # ---- Attendees ----
    attendees_html = '<h2>Attendees</h2>'
    if not attendees:
        attendees_html += '<div class="empty">No attendees recorded.</div>'
    else:
        rows = []
        for a in attendees:
            name = a.get("contact_name") or "—"
            vip = " ⭐" if a.get("is_vip") else ""
            role = a.get("contact_role") or a.get("role") or ""
            kind = "Internal" if a.get("is_internal") else "External"
            stat = a.get("status") or "pending"
            stat_kind = _status_badge_kind(stat)
            rows.append(
                f'<tr>'
                f'<td><strong>{_esc(name)}</strong>{vip}'
                f'{f"<br><span style=\'color:#6b7280;font-size:12px;\'>{_esc(role)}</span>" if role else ""}'
                f'</td>'
                f'<td>{_esc(kind)}</td>'
                f'<td>{_badge(stat, stat_kind)}</td>'
                f'<td>{_esc(a.get("response_notes") or "")}</td>'
                f'</tr>'
            )
        attendees_html += (
            '<table><thead><tr>'
            '<th>Name</th><th>Type</th><th>Status</th><th>Notes</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- Context: recent touchpoints with attendees ----
    context_html = ""
    if attendees:
        context_rows = []
        for a in attendees[:5]:
            cid = a.get("contact_id")
            if not cid:
                continue
            comms = db.get_communications_for_contact(cid, limit=2)
            for cm in comms:
                context_rows.append(
                    f'<tr>'
                    f'<td>{_esc(a.get("contact_name") or "—")}</td>'
                    f'<td>{_fmt_date(cm.get("comm_date"))}</td>'
                    f'<td>{_esc(cm.get("channel") or "—")}</td>'
                    f'<td>{_esc(cm.get("summary") or "—")}</td>'
                    f'</tr>'
                )
        if context_rows:
            context_html = (
                '<h2>Recent Context</h2>'
                '<table><thead><tr>'
                '<th>Contact</th><th>Date</th><th>Channel</th><th>Summary</th>'
                '</tr></thead><tbody>'
                + "".join(context_rows)
                + '</tbody></table>'
            )

    # ---- Open action items from previous meetings with same attendees ----
    prior_actions_html = ""
    if attendees:
        contact_ids = {a["contact_id"] for a in attendees if a.get("contact_id")}
        prior = []
        for cid in contact_ids:
            # Query: any open action item owned by this contact
            conn_items = db.get_action_items_due()  # all open
            for ai in conn_items:
                if ai.get("owner_contact_id") == cid:
                    prior.append(ai)
        # Deduplicate
        seen = set()
        unique_prior = []
        for ai in prior:
            if ai["id"] not in seen:
                seen.add(ai["id"])
                unique_prior.append(ai)
        if unique_prior:
            rows = []
            for ai in unique_prior[:10]:
                rows.append(
                    f'<tr>'
                    f'<td>{_esc(ai.get("description") or "")}</td>'
                    f'<td>{_esc(ai.get("owner_contact_name") or ai.get("owner_exec_name") or "—")}</td>'
                    f'<td>{_fmt_date(ai.get("due_date"))}</td>'
                    f'<td>{_badge(ai.get("status") or "not_started", _status_badge_kind(ai.get("status")))}</td>'
                    f'</tr>'
                )
            prior_actions_html = (
                '<h2>Open Items From Prior Meetings</h2>'
                '<table><thead><tr>'
                '<th>Item</th><th>Owner</th><th>Due</th><th>Status</th>'
                '</tr></thead><tbody>'
                + "".join(rows)
                + '</tbody></table>'
            )

    # ---- Decisions ----
    decisions_html = ""
    if decisions:
        items = "".join(
            f'<li>{_esc(d.get("decision") or "")}'
            f'{f" — <em>{_esc(d.get('made_by'))}</em>" if d.get("made_by") else ""}'
            f'</li>'
            for d in decisions
        )
        decisions_html = f'<h2>Decisions</h2><ul>{items}</ul>'

    # ---- Action items (this meeting) ----
    actions_html = ""
    if open_actions:
        rows = []
        for a in open_actions:
            owner = a.get("owner_contact_name") or a.get("owner_exec_name") or "—"
            rows.append(
                f'<tr>'
                f'<td>{_esc(a.get("description") or "")}</td>'
                f'<td>{_esc(owner)}</td>'
                f'<td>{_fmt_date(a.get("due_date"))}</td>'
                f'<td>{_badge(a.get("status") or "not_started", _status_badge_kind(a.get("status")))}</td>'
                f'</tr>'
            )
        actions_html = (
            '<h2>Action Items</h2>'
            '<table><thead><tr>'
            '<th>Item</th><th>Owner</th><th>Due</th><th>Status</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- Quick notes ----
    quick_notes_html = ""
    if meeting.get("notes_native"):
        quick_notes_html = (
            f'<h2>Quick Notes</h2>'
            f'<div class="note">{_esc(meeting["notes_native"]).replace(chr(10), "<br>")}</div>'
        )

    body = (
        header + meta + links_html + agenda_html + attendees_html
        + context_html + prior_actions_html + actions_html
        + decisions_html + quick_notes_html + _footer()
    )
    return _wrap_html(meeting.get("title") or "Meeting Brief", body)


# =========================================================
# EVENT BRIEF
# =========================================================


def generate_event_brief_html(event_id):
    event = db.get_event(event_id)
    if not event:
        return None

    guests = db.get_event_guests(event_id)
    counts = db.get_event_guest_status_counts(event_id)
    engagements = db.get_vendor_engagements(event_id)
    budget = db.get_budget_summary(event_id)
    budget_items = db.get_budget_items(event_id)
    currency = event.get("currency") or "USD"

    # ---- Header ----
    status_kind = _status_badge_kind(event.get("status"))
    exported_date = datetime.now().strftime("%d %b %Y, %H:%M")
    header = (
        f'<h1>{_esc(event.get("name"))}</h1>'
        f'<div class="subtitle">'
        f'{_badge(event.get("event_type") or "Event")} '
        f'{_badge(event.get("status") or "planning", status_kind)}'
        f'<span style="color:#6b7280;margin-left:12px;font-size:13px;">'
        f"Exported {exported_date}</span>"
        f"</div>"
    )

    # ---- Meta row ----
    time_str = " · ".join(
        x for x in [event.get("start_time"), event.get("end_time")] if x
    ) or "—"
    venue = event.get("venue_name") or event.get("location_override") or "—"
    if event.get("venue_city"):
        venue += f" ({event['venue_city']})"
    meta = _meta_row([
        ("Executive", event.get("executive_name") or "—"),
        ("Date", _fmt_date(event.get("event_date"))),
        ("Time", time_str),
        ("Venue", venue),
        ("Headcount", event.get("headcount") or "—"),
    ])

    # ---- Objective ----
    objective_html = ""
    if event.get("objective"):
        objective_html = (
            f'<h2>Objective</h2>'
            f'<div class="note">{_esc(event["objective"]).replace(chr(10), "<br>")}</div>'
        )

    # ---- Guest summary ----
    accepted = counts.get("accepted", 0)
    pending = counts.get("pending", 0) + counts.get("invited", 0)
    declined = counts.get("declined", 0)
    waitlist = counts.get("waitlist", 0)
    total = sum(counts.values())

    summary_html = (
        f'<h2>Guest Summary</h2>'
        f'<div class="meta">'
        f'<div class="meta-item"><strong>Total</strong>{total}</div>'
        f'<div class="meta-item"><strong>Accepted</strong>'
        f'{_badge(str(accepted), "green")}</div>'
        f'<div class="meta-item"><strong>Pending</strong>'
        f'{_badge(str(pending), "yellow")}</div>'
        f'<div class="meta-item"><strong>Declined</strong>'
        f'{_badge(str(declined), "red")}</div>'
        f'<div class="meta-item"><strong>Waitlist</strong>'
        f'{_badge(str(waitlist), None)}</div>'
        f'</div>'
    )

    # ---- Accepted guest list ----
    guest_html = '<h2>Confirmed Guests</h2>'
    accepted_guests = [g for g in guests if g.get("invite_status") == "accepted"]
    if not accepted_guests:
        guest_html += '<div class="empty">No confirmed guests yet.</div>'
    else:
        rows = []
        for g in accepted_guests:
            name = g.get("guest_name") or g.get("contact_name") or "—"
            email = g.get("guest_email") or g.get("contact_email") or "—"
            diet = g.get("dietary") or ""
            access = g.get("accessibility") or ""
            badge_name = g.get("badge_name") or ""
            plus = g.get("plus_ones") or 0
            extras = []
            if diet:
                extras.append(f"🍽 {diet}")
            if access:
                extras.append(f"♿ {access}")
            if badge_name:
                extras.append(f"🏷 {badge_name}")
            extras_html = "<br>".join(_esc(x) for x in extras)
            rows.append(
                f'<tr>'
                f'<td><strong>{_esc(name)}</strong></td>'
                f'<td>{_esc(email)}</td>'
                f'<td>{plus if plus else ""}</td>'
                f'<td>{extras_html}</td>'
                f'</tr>'
            )
        guest_html += (
            '<table><thead><tr>'
            '<th>Name</th><th>Email</th><th>+</th><th>Notes</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- Pending guests ----
    pending_html = ""
    pending_guests = [
        g for g in guests
        if (g.get("invite_status") or "") in ("invited", "pending")
    ]
    if pending_guests:
        rows = []
        for g in pending_guests:
            name = g.get("guest_name") or g.get("contact_name") or "—"
            email = g.get("guest_email") or g.get("contact_email") or "—"
            stat = g.get("invite_status") or "pending"
            rows.append(
                f'<tr>'
                f'<td><strong>{_esc(name)}</strong></td>'
                f'<td>{_esc(email)}</td>'
                f'<td>{_badge(stat, _status_badge_kind(stat))}</td>'
                f'</tr>'
            )
        pending_html = (
            f'<h2>Awaiting RSVP ({len(pending_guests)})</h2>'
            '<table><thead><tr>'
            '<th>Name</th><th>Email</th><th>Status</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- Vendors ----
    vendor_html = ""
    if engagements:
        rows = []
        for e in engagements:
            vendor = e.get("vendor_name") or "—"
            cat = e.get("category") or e.get("vendor_category") or ""
            stat = e.get("status") or "shortlisted"
            agreed = e.get("agreed_amount")
            agreed_str = (
                f'{agreed:,.2f} {e.get("currency") or "USD"}'
                if agreed else "—"
            )
            contract = e.get("contract_status") or "not_sent"
            rows.append(
                f'<tr>'
                f'<td><strong>{_esc(vendor)}</strong>'
                f'{f"<br><span style=\'color:#6b7280;font-size:12px;\'>{_esc(cat)}</span>" if cat else ""}'
                f'</td>'
                f'<td>{_badge(stat, _status_badge_kind(stat))}</td>'
                f'<td>{agreed_str}</td>'
                f'<td>{_esc(contract.replace("_", " ").title())}</td>'
                f'</tr>'
            )
        vendor_html = (
            f'<h2>Vendors ({len(engagements)})</h2>'
            '<table><thead><tr>'
            '<th>Vendor</th><th>Status</th><th>Agreed</th><th>Contract</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- Budget ----
    budget_html = ""
    if budget_items:
        budget_total = event.get("budget_total") or 0
        est = budget["estimated"]
        act = budget["actual"]
        paid = budget["paid"]

        # Summary row
        summary = (
            f'<div class="meta">'
            f'<div class="meta-item"><strong>Budget</strong>'
            f'{budget_total:,.2f} {currency}</div>'
            f'<div class="meta-item"><strong>Estimated</strong>'
            f'{est:,.2f} {currency}</div>'
            f'<div class="meta-item"><strong>Actual</strong>'
            f'{act:,.2f} {currency}</div>'
            f'<div class="meta-item"><strong>Paid</strong>'
            f'{paid:,.2f} {currency}</div>'
            f'</div>'
        )
        if budget_total > 0:
            ratio = act / budget_total
            if ratio >= 1:
                summary += (
                    f'<div class="note" style="border-color:#dc2626;'
                    f'background:#fef2f2;color:#991b1b;">'
                    f'🔴 Over budget by {act - budget_total:,.2f} {currency}'
                    f'</div>'
                )
            elif ratio >= 0.9:
                summary += (
                    f'<div class="note" style="border-color:#d97706;'
                    f'background:#fffbeb;color:#92400e;">'
                    f'⚠️ At {ratio*100:.0f}% of budget'
                    f'</div>'
                )

        rows = []
        for b in budget_items:
            desc = b.get("description") or "—"
            cat = b.get("category") or ""
            est_amt = b.get("estimated_amount") or 0
            act_amt = b.get("actual_amount") or 0
            paid_amt = b.get("paid_amount") or 0
            vendor = b.get("vendor_name") or ""
            rows.append(
                f'<tr>'
                f'<td><strong>{_esc(desc)}</strong>'
                f'{f"<br><span style=\'color:#6b7280;font-size:12px;\'>{_esc(vendor or cat)}</span>" if (vendor or cat) else ""}'
                f'</td>'
                f'<td>{est_amt:,.2f}</td>'
                f'<td>{act_amt:,.2f}</td>'
                f'<td>{paid_amt:,.2f}</td>'
                f'<td>{_badge(b.get("status") or "estimated", _status_badge_kind(b.get("status")))}</td>'
                f'</tr>'
            )
        budget_html = (
            f'<h2>Budget</h2>{summary}'
            '<table><thead><tr>'
            f'<th>Item</th><th>Est ({currency})</th>'
            f'<th>Actual ({currency})</th><th>Paid ({currency})</th>'
            '<th>Status</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- Links ----
    links = []
    if event.get("calendar_event_url"):
        links.append(f'📅 {_link_or_text(event["calendar_event_url"], "Calendar")}')
    if event.get("rsvp_form_url"):
        links.append(f'📋 {_link_or_text(event["rsvp_form_url"], "RSVP form")}')
    if event.get("run_of_show_url"):
        links.append(f'🎬 {_link_or_text(event["run_of_show_url"], "Run of show")}')
    if event.get("logistics_url"):
        links.append(f'📦 {_link_or_text(event["logistics_url"], "Logistics doc")}')
    links_html = (
        f'<h2>Links</h2><div class="pill-list">'
        + "".join(f'<span class="badge">{x}</span>' for x in links)
        + '</div>'
    ) if links else ""

    # ---- Quick notes ----
    quick_notes_html = ""
    if event.get("notes_native"):
        quick_notes_html = (
            f'<h2>Notes</h2>'
            f'<div class="note">{_esc(event["notes_native"]).replace(chr(10), "<br>")}</div>'
        )

    body = (
        header + meta + objective_html + summary_html + guest_html
        + pending_html + vendor_html + budget_html + links_html
        + quick_notes_html + _footer()
    )
    return _wrap_html(event.get("name") or "Event Brief", body)


# =========================================================
# SPEAKING ENGAGEMENT BRIEF
# =========================================================


def generate_speaking_brief_html(engagement_id):
    eng = db.get_speaking_engagement(engagement_id)
    if not eng:
        return None

    deliverables = db.get_speaking_deliverables(engagement_id)
    cfps = db.get_speaking_cfps(engagement_id)

    # ---- Header ----
    status_kind = _status_badge_kind(eng.get("status"))
    header = (
        f'<h1>{_esc(eng.get("title"))}</h1>'
        f'<div class="subtitle">'
        f'{_badge(eng.get("engagement_type") or "Speaking")} '
        f'{_badge(eng.get("format") or "—")} '
        f'{_badge(eng.get("status") or "invited", status_kind)}'
        f'</div>'
    )

    # ---- Meta row ----
    meta = _meta_row([
        ("Executive", eng.get("executive_name") or "—"),
        ("Organizer", eng.get("organizer_name") or "—"),
        ("Date", _fmt_date(eng.get("engagement_date"))),
        ("Time", eng.get("engagement_time") or "—"),
        ("Duration", eng.get("duration") or "—"),
        ("Location", eng.get("location") or "—"),
    ])

    # ---- Fee & terms ----
    fee_bits = []
    if eng.get("fee_offered"):
        fee_bits.append(
            f'Offered: {eng["fee_offered"]:,.2f} {eng.get("currency") or "USD"}'
        )
    if eng.get("fee_agreed"):
        fee_bits.append(
            f'Agreed: {eng["fee_agreed"]:,.2f} {eng.get("currency") or "USD"}'
        )
    if eng.get("travel_covered"):
        fee_bits.append("Travel covered")
    if eng.get("accommodation_covered"):
        fee_bits.append("Accommodation covered")
    if eng.get("recording_rights"):
        fee_bits.append(f"Recording: {eng['recording_rights']}")

    terms_html = ""
    if fee_bits:
        terms_html = (
            f'<h2>Fee & Terms</h2>'
            f'<div class="pill-list">'
            + "".join(f'<span class="badge">{_esc(x)}</span>' for x in fee_bits)
            + '</div>'
        )

    # ---- Topic & audience ----
    topic_html = ""
    if eng.get("topic") or eng.get("audience"):
        topic_html = '<h2>Topic & Audience</h2><div class="note">'
        if eng.get("topic"):
            topic_html += f'<strong>Topic:</strong> {_esc(eng["topic"])}<br>'
        if eng.get("audience"):
            topic_html += f'<strong>Audience:</strong> {_esc(eng["audience"])}'
        topic_html += '</div>'

    # ---- Deliverables ----
    deliv_html = '<h2>Deliverables</h2>'
    if not deliverables:
        deliv_html += (
            '<div class="empty">No deliverables yet. Add them in the '
            'engagement record.</div>'
        )
    else:
        rows = []
        for d in deliverables:
            name = d.get("name") or "—"
            owner = d.get("owner") or "—"
            due = _fmt_date(d.get("due_date"))
            stat = d.get("status") or "not_started"
            submitted = "📤" if d.get("submitted_to_organizer") else ""
            file_link = ""
            if d.get("file_url"):
                file_link = _link_or_text(d["file_url"], "Open")
            rows.append(
                f'<tr>'
                f'<td><strong>{_esc(name)}</strong></td>'
                f'<td>{_esc(owner)}</td>'
                f'<td>{due}</td>'
                f'<td>{_badge(stat, _status_badge_kind(stat))}</td>'
                f'<td>{submitted} {file_link}</td>'
                f'</tr>'
            )
        deliv_html += (
            '<table><thead><tr>'
            '<th>Deliverable</th><th>Owner</th><th>Due</th>'
            '<th>Status</th><th>File</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- CFP ----
    cfp_html = ""
    if cfps:
        rows = []
        for cf in cfps:
            title = cf.get("proposal_title") or "Untitled"
            stat = cf.get("status") or "draft"
            deadline = _fmt_date(cf.get("submission_deadline"))
            rows.append(
                f'<tr>'
                f'<td><strong>{_esc(title)}</strong></td>'
                f'<td>{deadline}</td>'
                f'<td>{_badge(stat, _status_badge_kind(stat))}</td>'
                f'<td>{_link_or_text(cf.get("proposal_doc_url"), "Doc") if cf.get("proposal_doc_url") else ""}</td>'
                f'</tr>'
            )
        cfp_html = (
            f'<h2>CFP Submissions</h2>'
            '<table><thead><tr>'
            '<th>Proposal</th><th>Deadline</th><th>Status</th><th>Doc</th>'
            '</tr></thead><tbody>'
            + "".join(rows)
            + '</tbody></table>'
        )

    # ---- Links ----
    links = []
    if eng.get("calendar_event_url"):
        links.append(f'📅 {_link_or_text(eng["calendar_event_url"], "Calendar")}')
    if eng.get("notes_url"):
        links.append(f'📝 {_link_or_text(eng["notes_url"], "Notes doc")}')
    links_html = (
        f'<h2>Links</h2><div class="pill-list">'
        + "".join(f'<span class="badge">{x}</span>' for x in links)
        + '</div>'
    ) if links else ""

    # ---- Quick notes ----
    quick_notes_html = ""
    if eng.get("notes_native"):
        quick_notes_html = (
            f'<h2>Quick Notes</h2>'
            f'<div class="note">{_esc(eng["notes_native"]).replace(chr(10), "<br>")}</div>'
        )

    body = (
        header + meta + terms_html + topic_html + deliv_html + cfp_html
        + links_html + quick_notes_html + _footer()
    )
    return _wrap_html(eng.get("title") or "Speaking Brief", body)

# =========================================================
# PDF GENERATORS (HTML → PDF via WeasyPrint)
# =========================================================


def _html_to_pdf(html_string):
    if not WEASYPRINT_AVAILABLE:
        return None
    try:
        return WeasyHTML(string=html_string).write_pdf()
    except Exception:
        return None


def generate_meeting_brief_pdf(meeting_id):
    html = generate_meeting_brief_html(meeting_id)
    if not html:
        return None
    return _html_to_pdf(html)


def generate_event_brief_pdf(event_id):
    html = generate_event_brief_html(event_id)
    if not html:
        return None
    return _html_to_pdf(html)


def generate_speaking_brief_pdf(engagement_id):
    html = generate_speaking_brief_html(engagement_id)
    if not html:
        return None
    return _html_to_pdf(html)


# =========================================================
# WORD DOCX HELPERS
# =========================================================


def _docx_new():
    if not DOCX_AVAILABLE:
        return None
    doc = Document()
    # Tight margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
    # Base font
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    return doc


def _docx_title(doc, text):
    p = doc.add_heading(text or "—", level=0)
    return p


def _docx_subtitle(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text or "")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    run.italic = True
    return p


def _docx_heading(doc, text):
    return doc.add_heading(text or "", level=1)


def _docx_paragraph(doc, text, italic=False, size=None):
    p = doc.add_paragraph()
    run = p.add_run(text or "")
    if italic:
        run.italic = True
    if size:
        run.font.size = Pt(size)
    return p


def _docx_note(doc, text):
    """A quiet, indented note block."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text or "")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)
    return p


def _docx_meta_table(doc, items):
    """items: list of (label, value) tuples. Renders as 2-col table."""
    items = [(l, v) for l, v in items if v not in (None, "", "—")]
    if not items:
        return
    table = doc.add_table(rows=len(items), cols=2)
    table.autofit = False
    table.columns[0].width = Inches(1.6)
    table.columns[1].width = Inches(5.0)
    for i, (label, value) in enumerate(items):
        row = table.rows[i]
        c0 = row.cells[0]
        c1 = row.cells[1]
        c0.width = Inches(1.6)
        c1.width = Inches(5.0)
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(str(label).upper())
        r0.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(str(value))
        r1.font.size = Pt(10)


def _docx_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(str(item))
    return


def _docx_table(doc, headers, rows):
    """headers: list of strings. rows: list of lists of strings."""
    if not headers:
        return
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.style = "Light Grid Accent 1"

    # Header row
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(str(h))
        run.bold = True
        run.font.size = Pt(9)

    # Data rows
    for r_i, row in enumerate(rows):
        cells = table.rows[r_i + 1].cells
        for c_i, val in enumerate(row):
            if c_i >= n_cols:
                continue
            p = cells[c_i].paragraphs[0]
            run = p.add_run(str(val) if val is not None else "")
            run.font.size = Pt(9)

    return


def _docx_footer(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(
        f"Generated {datetime.now().strftime('%d %b %Y · %H:%M')} · "
        f"EA Command Center"
    )
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x9C, 0xA3, 0xAF)


def _docx_badge_line(doc, items):
    """Render a set of short badges as a single comma-separated line."""
    items = [str(x) for x in items if x]
    if not items:
        return
    p = doc.add_paragraph()
    run = p.add_run(" · ".join(items))
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x37, 0x41, 0x51)


# =========================================================
# MEETING BRIEF (Word)
# =========================================================


def generate_meeting_brief_docx(meeting_id):
    if not DOCX_AVAILABLE:
        return None
    meeting = db.get_meeting(meeting_id)
    if not meeting:
        return None

    exec_id = meeting["exec_id"]
    attendees = db.get_meeting_attendees(meeting_id)
    action_items = db.get_meeting_action_items(meeting_id)
    decisions = db.get_meeting_decisions(meeting_id)
    open_actions = [
        a for a in action_items
        if (a.get("status") or "not_started") != "done"
    ]

    doc = _docx_new()

    # Header
    _docx_title(doc, meeting.get("title"))
    subtitle_parts = [
        meeting.get("meeting_type") or "Meeting",
        (meeting.get("status") or "draft").replace("_", " ").title(),
    ]
    _docx_subtitle(doc, " · ".join(subtitle_parts))

    # Meta
    time_str = meeting.get("meeting_time") or ""
    dur = (
        f'{meeting["duration_minutes"]} min'
        if meeting.get("duration_minutes") else ""
    )
    time_full = " · ".join(x for x in [time_str, dur] if x) or "—"
    _docx_meta_table(doc, [
        ("Executive", meeting.get("executive_name")),
        ("Date", _fmt_date(meeting.get("meeting_date"))),
        ("Time", time_full),
        ("Location", meeting.get("location")),
    ])

    # Links
    links = []
    if meeting.get("calendar_event_url"):
        links.append(f"Calendar: {meeting['calendar_event_url']}")
    if meeting.get("video_link"):
        links.append(f"Video: {meeting['video_link']}")
    if meeting.get("notes_url"):
        links.append(
            f"{meeting.get('notes_label') or 'Notes'}: "
            f"{meeting['notes_url']}"
        )
    if links:
        _docx_heading(doc, "Links")
        _docx_bullets(doc, links)

    # Agenda
    if meeting.get("agenda"):
        _docx_heading(doc, "Agenda")
        _docx_note(doc, meeting["agenda"])

    # Attendees
    _docx_heading(doc, "Attendees")
    if not attendees:
        _docx_paragraph(doc, "No attendees recorded.", italic=True)
    else:
        rows = []
        for a in attendees:
            name = a.get("contact_name") or "—"
            vip = " ⭐" if a.get("is_vip") else ""
            role = a.get("contact_role") or a.get("role") or ""
            kind = "Internal" if a.get("is_internal") else "External"
            stat = a.get("status") or "pending"
            rows.append([
                f"{name}{vip}",
                role,
                kind,
                stat.title(),
                a.get("response_notes") or "",
            ])
        _docx_table(
            doc,
            ["Name", "Role", "Type", "Status", "Notes"],
            rows,
        )

    # Recent context
    if attendees:
        context_rows = []
        for a in attendees[:5]:
            cid = a.get("contact_id")
            if not cid:
                continue
            comms = db.get_communications_for_contact(cid, limit=2)
            for cm in comms:
                context_rows.append([
                    a.get("contact_name") or "—",
                    _fmt_date(cm.get("comm_date")),
                    cm.get("channel") or "—",
                    cm.get("summary") or "—",
                ])
        if context_rows:
            _docx_heading(doc, "Recent Context")
            _docx_table(
                doc,
                ["Contact", "Date", "Channel", "Summary"],
                context_rows,
            )

    # Prior open items
    if attendees:
        contact_ids = {a["contact_id"] for a in attendees if a.get("contact_id")}
        prior = []
        all_open = db.get_action_items_due()
        for ai in all_open:
            if ai.get("owner_contact_id") in contact_ids:
                prior.append(ai)
        seen = set()
        unique_prior = []
        for ai in prior:
            if ai["id"] not in seen:
                seen.add(ai["id"])
                unique_prior.append(ai)
        if unique_prior:
            _docx_heading(doc, "Open Items From Prior Meetings")
            rows = []
            for ai in unique_prior[:10]:
                rows.append([
                    ai.get("description") or "",
                    ai.get("owner_contact_name")
                    or ai.get("owner_exec_name") or "—",
                    _fmt_date(ai.get("due_date")),
                    (ai.get("status") or "not_started").replace("_", " ").title(),
                ])
            _docx_table(
                doc,
                ["Item", "Owner", "Due", "Status"],
                rows,
            )

    # Action items
    if open_actions:
        _docx_heading(doc, "Action Items")
        rows = []
        for a in open_actions:
            owner = (
                a.get("owner_contact_name")
                or a.get("owner_exec_name") or "—"
            )
            rows.append([
                a.get("description") or "",
                owner,
                _fmt_date(a.get("due_date")),
                (a.get("status") or "not_started").replace("_", " ").title(),
            ])
        _docx_table(doc, ["Item", "Owner", "Due", "Status"], rows)

    # Decisions
    if decisions:
        _docx_heading(doc, "Decisions")
        items = []
        for d in decisions:
            line = d.get("decision") or ""
            if d.get("made_by"):
                line += f" — {d['made_by']}"
            items.append(line)
        _docx_bullets(doc, items)

    # Quick notes
    if meeting.get("notes_native"):
        _docx_heading(doc, "Quick Notes")
        _docx_note(doc, meeting["notes_native"])

    _docx_footer(doc)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# =========================================================
# EVENT BRIEF (Word)
# =========================================================


def generate_event_brief_docx(event_id):
    if not DOCX_AVAILABLE:
        return None
    event = db.get_event(event_id)
    if not event:
        return None

    guests = db.get_event_guests(event_id)
    counts = db.get_event_guest_status_counts(event_id)
    engagements = db.get_vendor_engagements(event_id)
    budget = db.get_budget_summary(event_id)
    budget_items = db.get_budget_items(event_id)
    currency = event.get("currency") or "USD"

    doc = _docx_new()

    _docx_title(doc, event.get("name"))
    exported_date = datetime.now().strftime("%d %b %Y, %H:%M")
    subtitle_parts = [
        event.get("event_type") or "Event",
        (event.get("status") or "planning").replace("_", " ").title(),
        f"Exported {exported_date}",
    ]
    _docx_subtitle(doc, " · ".join(subtitle_parts))

    time_str = " · ".join(
        x for x in [event.get("start_time"), event.get("end_time")] if x
    ) or "—"
    venue = event.get("venue_name") or event.get("location_override") or "—"
    if event.get("venue_city"):
        venue += f" ({event['venue_city']})"

    _docx_meta_table(doc, [
        ("Executive", event.get("executive_name")),
        ("Date", _fmt_date(event.get("event_date"))),
        ("Time", time_str),
        ("Venue", venue),
        ("Headcount", event.get("headcount")),
    ])

    if event.get("objective"):
        _docx_heading(doc, "Objective")
        _docx_note(doc, event["objective"])

    # Guest summary
    accepted = counts.get("accepted", 0)
    pending = counts.get("pending", 0) + counts.get("invited", 0)
    declined = counts.get("declined", 0)
    waitlist = counts.get("waitlist", 0)
    total = sum(counts.values())

    _docx_heading(doc, "Guest Summary")
    summary_p = doc.add_paragraph()
    summary_run = summary_p.add_run(
        f"Total: {total}  ·  Accepted: {accepted}  ·  "
        f"Pending: {pending}  ·  Declined: {declined}  ·  "
        f"Waitlist: {waitlist}"
    )
    summary_run.font.size = Pt(11)
    summary_run.bold = True

    # Accepted guests
    _docx_heading(doc, "Confirmed Guests")
    accepted_guests = [
        g for g in guests if g.get("invite_status") == "accepted"
    ]
    if not accepted_guests:
        _docx_paragraph(doc, "No confirmed guests yet.", italic=True)
    else:
        rows = []
        for g in accepted_guests:
            name = g.get("guest_name") or g.get("contact_name") or "—"
            email = g.get("guest_email") or g.get("contact_email") or "—"
            diet = g.get("dietary") or ""
            access = g.get("accessibility") or ""
            badge_name = g.get("badge_name") or ""
            plus = g.get("plus_ones") or 0
            extras = []
            if diet:
                extras.append(f"Diet: {diet}")
            if access:
                extras.append(f"Access: {access}")
            if badge_name:
                extras.append(f"Badge: {badge_name}")
            rows.append([
                name,
                email,
                str(plus) if plus else "",
                "; ".join(extras),
            ])
        _docx_table(doc, ["Name", "Email", "+", "Notes"], rows)

    # Pending
    pending_guests = [
        g for g in guests
        if (g.get("invite_status") or "") in ("invited", "pending")
    ]
    if pending_guests:
        _docx_heading(doc, f"Awaiting RSVP ({len(pending_guests)})")
        rows = []
        for g in pending_guests:
            name = g.get("guest_name") or g.get("contact_name") or "—"
            email = g.get("guest_email") or g.get("contact_email") or "—"
            stat = g.get("invite_status") or "pending"
            rows.append([name, email, stat.title()])
        _docx_table(doc, ["Name", "Email", "Status"], rows)

    # Vendors
    if engagements:
        _docx_heading(doc, f"Vendors ({len(engagements)})")
        rows = []
        for e in engagements:
            vendor = e.get("vendor_name") or "—"
            cat = e.get("category") or e.get("vendor_category") or ""
            stat = e.get("status") or "shortlisted"
            agreed = e.get("agreed_amount")
            agreed_str = (
                f'{agreed:,.2f} {e.get("currency") or "USD"}'
                if agreed else "—"
            )
            contract = e.get("contract_status") or "not_sent"
            rows.append([
                f"{vendor}" + (f" ({cat})" if cat else ""),
                stat.replace("_", " ").title(),
                agreed_str,
                contract.replace("_", " ").title(),
            ])
        _docx_table(
            doc, ["Vendor", "Status", "Agreed", "Contract"], rows,
        )

    # Budget
    if budget_items:
        _docx_heading(doc, "Budget")
        budget_total = event.get("budget_total") or 0
        est = budget["estimated"]
        act = budget["actual"]
        paid = budget["paid"]

        summary_p = doc.add_paragraph()
        summary_p.add_run(
            f"Budget: {budget_total:,.2f} {currency}  ·  "
            f"Estimated: {est:,.2f}  ·  Actual: {act:,.2f}  ·  "
            f"Paid: {paid:,.2f}"
        ).bold = True

        if budget_total > 0:
            ratio = act / budget_total
            if ratio >= 1:
                warn_p = doc.add_paragraph()
                warn_run = warn_p.add_run(
                    f"🔴 Over budget by {act - budget_total:,.2f} {currency}"
                )
                warn_run.bold = True
                warn_run.font.color.rgb = RGBColor(0x99, 0x1B, 0x1B)
            elif ratio >= 0.9:
                warn_p = doc.add_paragraph()
                warn_run = warn_p.add_run(
                    f"⚠️ At {ratio*100:.0f}% of budget"
                )
                warn_run.bold = True
                warn_run.font.color.rgb = RGBColor(0x92, 0x40, 0x0E)

        rows = []
        for b in budget_items:
            desc = b.get("description") or "—"
            vendor = b.get("vendor_name") or b.get("category") or ""
            rows.append([
                f"{desc}" + (f" ({vendor})" if vendor else ""),
                f'{b.get("estimated_amount") or 0:,.2f}',
                f'{b.get("actual_amount") or 0:,.2f}',
                f'{b.get("paid_amount") or 0:,.2f}',
                (b.get("status") or "estimated").replace("_", " ").title(),
            ])
        _docx_table(
            doc,
            [
                "Item",
                f"Est ({currency})",
                f"Actual ({currency})",
                f"Paid ({currency})",
                "Status",
            ],
            rows,
        )

    # Links
    links = []
    if event.get("calendar_event_url"):
        links.append(f"Calendar: {event['calendar_event_url']}")
    if event.get("rsvp_form_url"):
        links.append(f"RSVP form: {event['rsvp_form_url']}")
    if event.get("run_of_show_url"):
        links.append(f"Run of show: {event['run_of_show_url']}")
    if event.get("logistics_url"):
        links.append(f"Logistics: {event['logistics_url']}")
    if links:
        _docx_heading(doc, "Links")
        _docx_bullets(doc, links)

    if event.get("notes_native"):
        _docx_heading(doc, "Notes")
        _docx_note(doc, event["notes_native"])

    _docx_footer(doc)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# =========================================================
# SPEAKING BRIEF (Word)
# =========================================================


def generate_speaking_brief_docx(engagement_id):
    if not DOCX_AVAILABLE:
        return None
    eng = db.get_speaking_engagement(engagement_id)
    if not eng:
        return None

    deliverables = db.get_speaking_deliverables(engagement_id)
    cfps = db.get_speaking_cfps(engagement_id)

    doc = _docx_new()

    _docx_title(doc, eng.get("title"))
    subtitle_parts = [
        eng.get("engagement_type") or "Speaking",
        eng.get("format") or "",
        (eng.get("status") or "invited").replace("_", " ").title(),
    ]
    _docx_subtitle(doc, " · ".join([x for x in subtitle_parts if x]))

    _docx_meta_table(doc, [
        ("Executive", eng.get("executive_name")),
        ("Organizer", eng.get("organizer_name")),
        ("Date", _fmt_date(eng.get("engagement_date"))),
        ("Time", eng.get("engagement_time")),
        ("Duration", eng.get("duration")),
        ("Location", eng.get("location")),
    ])

    # Fee & terms
    fee_bits = []
    currency = eng.get("currency") or "USD"
    if eng.get("fee_offered"):
        fee_bits.append(f"Offered: {eng['fee_offered']:,.2f} {currency}")
    if eng.get("fee_agreed"):
        fee_bits.append(f"Agreed: {eng['fee_agreed']:,.2f} {currency}")
    if eng.get("travel_covered"):
        fee_bits.append("Travel covered")
    if eng.get("accommodation_covered"):
        fee_bits.append("Accommodation covered")
    if eng.get("recording_rights"):
        fee_bits.append(f"Recording: {eng['recording_rights']}")
    if fee_bits:
        _docx_heading(doc, "Fee & Terms")
        _docx_badge_line(doc, fee_bits)

    # Topic & audience
    if eng.get("topic") or eng.get("audience"):
        _docx_heading(doc, "Topic & Audience")
        if eng.get("topic"):
            _docx_paragraph(doc, f"Topic: {eng['topic']}")
        if eng.get("audience"):
            _docx_paragraph(doc, f"Audience: {eng['audience']}")

    # Deliverables
    _docx_heading(doc, "Deliverables")
    if not deliverables:
        _docx_paragraph(doc, "No deliverables yet.", italic=True)
    else:
        rows = []
        for d in deliverables:
            rows.append([
                d.get("name") or "—",
                d.get("owner") or "—",
                _fmt_date(d.get("due_date")),
                (d.get("status") or "not_started").replace("_", " ").title(),
                "Yes" if d.get("submitted_to_organizer") else "",
                d.get("file_url") or "",
            ])
        _docx_table(
            doc,
            ["Deliverable", "Owner", "Due", "Status", "Submitted", "File"],
            rows,
        )

    # CFP
    if cfps:
        _docx_heading(doc, "CFP Submissions")
        rows = []
        for cf in cfps:
            rows.append([
                cf.get("proposal_title") or "Untitled",
                _fmt_date(cf.get("submission_deadline")),
                (cf.get("status") or "draft").replace("_", " ").title(),
                cf.get("proposal_doc_url") or "",
            ])
        _docx_table(
            doc, ["Proposal", "Deadline", "Status", "Doc"], rows,
        )

    # Links
    links = []
    if eng.get("calendar_event_url"):
        links.append(f"Calendar: {eng['calendar_event_url']}")
    if eng.get("notes_url"):
        links.append(f"Notes doc: {eng['notes_url']}")
    if links:
        _docx_heading(doc, "Links")
        _docx_bullets(doc, links)

    if eng.get("notes_native"):
        _docx_heading(doc, "Quick Notes")
        _docx_note(doc, eng["notes_native"])

    _docx_footer(doc)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
