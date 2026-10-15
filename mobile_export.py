"""
Mobile HTML export.

Generates a self-contained HTML file with all data embedded. Opens on
any phone browser with no server. Optional AES-256-GCM password
encryption using the `cryptography` library.

When encryption is enabled:
  - Python derives a key from the password (PBKDF2, 100k iterations)
  - AES-256-GCM encrypts the JSON payload
  - The HTML embeds ciphertext + salt + IV
  - Browser prompts for password, uses SubtleCrypto to decrypt

When encryption is disabled:
  - The JSON is embedded as plaintext in a <script> tag
"""

import base64
import json
import os
from datetime import datetime, timedelta

import database as db

try:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    CRYPTO_AVAILABLE = True
except Exception:
    CRYPTO_AVAILABLE = False


# =========================================================
# DATA COLLECTION
# =========================================================


def _serialize_row(row):
    """Convert a sqlite3.Row or dict to a plain dict with safe values."""
    if row is None:
        return None
    if hasattr(row, "keys"):
        return {k: row[k] for k in row.keys()}
    return dict(row)


def build_export_data(scope="full", scope_id=None):
    """
    Build a dict with all data needed for the mobile export.

    scope: "full" | "today" | "event" | "trip"
    scope_id: for event/trip, the record ID
    """
    data = {
        "generated_at": datetime.now().isoformat(),
        "scope": scope,
        "scope_id": scope_id,
    }

    if scope == "event":
        ev = db.get_event(scope_id)
        if not ev:
            return None
        data["executives"] = [
            _serialize_row(db.get_executive_profile(ev["exec_id"]))
        ] if ev.get("exec_id") else []
        data["events"] = [ev]
        data["event_guests"] = db.get_event_guests(scope_id)
        data["vendors"] = []
        data["budget_items"] = db.get_budget_items(scope_id)
        data["meetings"] = []
        data["speaking_engagements"] = []
        data["trips"] = []
        data["contacts"] = []
        return data

    if scope == "trip":
        trip = db.get_trip(scope_id)
        if not trip:
            return None
        data["executives"] = [
            _serialize_row(db.get_executive_profile(trip["exec_id"]))
        ] if trip.get("exec_id") else []
        data["trips"] = [trip]
        data["trip_stops"] = db.get_trip_stops(scope_id)
        data["itinerary_items"] = db.get_items_for_trip(scope_id)
        data["trip_delegation"] = db.get_trip_delegation_members(scope_id)
        # Linked items
        data["meetings"] = db.get_meetings_for_trip(scope_id)
        data["events"] = db.get_events_for_trip(scope_id)
        data["speaking_engagements"] = db.get_speaking_for_trip(scope_id)
        data["vendors"] = []
        data["budget_items"] = []
        data["contacts"] = []
        return data

    # full or today
    today = datetime.now().date()

    if scope == "today":
        today_iso = today.isoformat()
        data["meetings"] = db.get_meetings(
            start_date=today_iso, end_date=today_iso
        )
        data["events"] = db.get_events(
            start_date=today_iso, end_date=today_iso
        )
        data["speaking_engagements"] = db.get_speaking_engagements(
            start_date=today_iso, end_date=today_iso
        )
        window_end = (today + timedelta(days=14)).isoformat()
        data["trips"] = [
            t for t in db.get_trips_overlapping_dates(
                today_iso, window_end
            )
        ]
        data["vendors"] = []
        data["contacts"] = []
        data["event_guests"] = []
        data["budget_items"] = []
    else:
        # Full
        data["meetings"] = db.get_meetings()
        data["events"] = db.get_events()
        data["speaking_engagements"] = db.get_speaking_engagements()
        data["trips"] = db.get_spending_summary()
        # For full export, include only active contacts to keep size down
        data["contacts"] = [
            c for c in db.get_contacts(active_only=True)
        ]
        data["vendors"] = db.get_vendors(active_only=True)
        data["event_guests"] = []
        data["budget_items"] = []

    # Executive list for the filter
    if "executives" not in data:
        data["executives"] = db.get_all_executives(active_only=True)

    return data


# =========================================================
# ENCRYPTION
# =========================================================


def _derive_key(password, salt):
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100_000,
    )
    return kdf.derive(password.encode("utf-8"))


def encrypt_payload(payload_json, password):
    """Encrypt with AES-256-GCM. Returns (ciphertext_b64, salt_b64, iv_b64)."""
    if not CRYPTO_AVAILABLE:
        raise RuntimeError("cryptography library not available")
    salt = os.urandom(16)
    iv = os.urandom(12)
    key = _derive_key(password, salt)
    aesgcm = AESGCM(key)
    ciphertext = aesgcm.encrypt(iv, payload_json.encode("utf-8"), None)
    return (
        base64.b64encode(ciphertext).decode("ascii"),
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(iv).decode("ascii"),
    )


# =========================================================
# HTML TEMPLATE
# =========================================================


MOBILE_CSS = """
* { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
:root {
  --bg: #ffffff;
  --surface: #f9fafb;
  --text: #111827;
  --text-muted: #6b7280;
  --border: #e5e7eb;
  --accent: #2563eb;
  --accent-soft: #dbeafe;
  --green: #059669;
  --green-soft: #d1fae5;
  --yellow: #d97706;
  --yellow-soft: #fef3c7;
  --red: #dc2626;
  --red-soft: #fee2e2;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #0f172a;
    --surface: #1e293b;
    --text: #f1f5f9;
    --text-muted: #94a3b8;
    --border: #334155;
    --accent: #60a5fa;
    --accent-soft: #1e3a8a;
    --green: #34d399;
    --green-soft: #064e3b;
    --yellow: #fbbf24;
    --yellow-soft: #78350f;
    --red: #f87171;
    --red-soft: #7f1d1d;
  }
}
html, body {
  margin: 0;
  padding: 0;
  background: var(--bg);
  color: var(--text);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  font-size: 16px;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}
header {
  padding: 16px;
  border-bottom: 1px solid var(--border);
  position: sticky;
  top: 0;
  background: var(--bg);
  z-index: 10;
}
h1 {
  margin: 0 0 4px 0;
  font-size: 20px;
  font-weight: 700;
}
.meta {
  margin: 0;
  font-size: 12px;
  color: var(--text-muted);
}
.tabs {
  display: flex;
  overflow-x: auto;
  gap: 6px;
  padding: 12px 16px;
  border-bottom: 1px solid var(--border);
  background: var(--surface);
  position: sticky;
  top: 73px;
  z-index: 9;
  -webkit-overflow-scrolling: touch;
  scrollbar-width: none;
}
.tabs::-webkit-scrollbar { display: none; }
.tabs button {
  flex: 0 0 auto;
  padding: 8px 14px;
  font-size: 14px;
  font-weight: 500;
  background: var(--bg);
  color: var(--text);
  border: 1px solid var(--border);
  border-radius: 999px;
  cursor: pointer;
  white-space: nowrap;
}
.tabs button.active {
  background: var(--accent);
  color: #ffffff;
  border-color: var(--accent);
}
main {
  padding: 16px;
  max-width: 720px;
  margin: 0 auto;
}
.search {
  padding: 10px 14px;
  width: 100%;
  font-size: 16px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  color: var(--text);
  margin-bottom: 16px;
}
.search:focus {
  outline: none;
  border-color: var(--accent);
}
.card {
  padding: 14px 16px;
  background: var(--surface);
  border-radius: 10px;
  margin-bottom: 10px;
  border: 1px solid var(--border);
}
.card h2 {
  font-size: 16px;
  margin: 0 0 6px 0;
  font-weight: 600;
  color: var(--text);
}
.card .row {
  font-size: 14px;
  color: var(--text-muted);
  margin-bottom: 4px;
}
.card .row strong {
  color: var(--text);
  font-weight: 500;
}
.badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 500;
  background: var(--accent-soft);
  color: var(--accent);
  margin-right: 4px;
}
.badge.green { background: var(--green-soft); color: var(--green); }
.badge.yellow { background: var(--yellow-soft); color: var(--yellow); }
.badge.red { background: var(--red-soft); color: var(--red); }
.section-title {
  font-size: 12px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.6px;
  color: var(--text-muted);
  margin: 24px 0 10px 0;
}
a {
  color: var(--accent);
  text-decoration: none;
  font-weight: 500;
}
a:hover { text-decoration: underline; }
.tap-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
  margin: 8px 0;
}
.tap-list a {
  padding: 10px 12px;
  background: var(--bg);
  border: 1px solid var(--border);
  border-radius: 8px;
  display: block;
}
.actions {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  margin-top: 10px;
}
.actions a {
  padding: 8px 12px;
  background: var(--accent-soft);
  color: var(--accent);
  border-radius: 8px;
  font-size: 13px;
  font-weight: 600;
}
.actions a.phone { background: var(--green-soft); color: var(--green); }
.empty {
  padding: 32px 16px;
  text-align: center;
  color: var(--text-muted);
  font-size: 14px;
}
#password-screen {
  padding: 40px 20px;
  max-width: 380px;
  margin: 60px auto;
}
#password-screen h2 {
  margin: 0 0 8px 0;
  font-size: 18px;
}
#password-screen p {
  margin: 0 0 20px 0;
  color: var(--text-muted);
  font-size: 14px;
}
#password-screen input {
  width: 100%;
  padding: 12px 14px;
  font-size: 16px;
  border: 1px solid var(--border);
  background: var(--surface);
  color: var(--text);
  border-radius: 8px;
  margin-bottom: 10px;
}
#password-screen button {
  width: 100%;
  padding: 12px 14px;
  font-size: 16px;
  font-weight: 600;
  background: var(--accent);
  color: #ffffff;
  border: none;
  border-radius: 8px;
  cursor: pointer;
}
#password-screen .error {
  color: var(--red);
  font-size: 13px;
  margin-top: 8px;
  min-height: 18px;
}
.detail-toggle {
  background: none;
  border: none;
  color: var(--accent);
  font-size: 13px;
  font-weight: 600;
  padding: 4px 0;
  cursor: pointer;
}
.detail-content {
  display: none;
  margin-top: 8px;
  padding-top: 8px;
  border-top: 1px solid var(--border);
}
.detail-content.open { display: block; }
table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
  margin: 6px 0;
}
th, td {
  padding: 6px 8px;
  text-align: left;
  border-bottom: 1px solid var(--border);
}
th {
  color: var(--text-muted);
  font-weight: 600;
  font-size: 11px;
  text-transform: uppercase;
}
footer {
  padding: 24px 16px;
  text-align: center;
  color: var(--text-muted);
  font-size: 11px;
}
@media print {
  .tabs, .search, .actions { display: none; }
  .card { break-inside: avoid; }
}
"""


MOBILE_JS = r"""
const PAYLOAD_SALT = "%%SALT%%";
const PAYLOAD_IV = "%%IV%%";
const PAYLOAD_DATA = "%%DATA%%";
const IS_ENCRYPTED = %%ENCRYPTED%%;

let appData = null;

function b64ToBytes(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function deriveKey(password, salt) {
  const enc = new TextEncoder();
  const keyMaterial = await crypto.subtle.importKey(
    "raw", enc.encode(password), "PBKDF2", false, ["deriveKey"]
  );
  return crypto.subtle.deriveKey(
    {
      name: "PBKDF2",
      salt: salt,
      iterations: 100000,
      hash: "SHA-256",
    },
    keyMaterial,
    { name: "AES-GCM", length: 256 },
    false,
    ["decrypt"]
  );
}

async function unlock() {
  const pw = document.getElementById("pw-input").value;
  const errEl = document.getElementById("pw-error");
  errEl.textContent = "";
  try {
    const salt = b64ToBytes(PAYLOAD_SALT);
    const iv = b64ToBytes(PAYLOAD_IV);
    const ciphertext = b64ToBytes(PAYLOAD_DATA);
    const key = await deriveKey(pw, salt);
    const decrypted = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: iv },
      key,
      ciphertext
    );
    const json = new TextDecoder().decode(decrypted);
    appData = JSON.parse(json);
    showApp();
  } catch (e) {
    errEl.textContent = "Incorrect password or corrupted file.";
  }
}

function showApp() {
  document.getElementById("password-screen").style.display = "none";
  document.getElementById("app-shell").style.display = "block";
  renderCurrentTab();
}

function setTab(name) {
  document.querySelectorAll(".tabs button").forEach(function (b) {
    b.classList.toggle("active", b.dataset.tab === name);
  });
  currentTab = name;
  renderCurrentTab();
}

let currentTab = "today";

function escapeHtml(s) {
  if (s === null || s === undefined) return "";
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function badge(text, kind) {
  const k = kind ? " " + kind : "";
  return '<span class="badge' + k + '">' + escapeHtml(text) + "</span>";
}

function statusKind(s) {
  s = (s || "").toLowerCase();
  if (["confirmed","accepted","signed","completed","delivered","done",
       "paid","ready","closed","checked_in"].indexOf(s) >= 0) return "green";
  if (["pending","in_progress","negotiating","prep_in_progress",
       "tentative","under_review","quoted","committed","waitlist"].indexOf(s) >= 0)
    return "yellow";
  if (["declined","cancelled","rejected","no_response","overdue",
       "blocked"].indexOf(s) >= 0) return "red";
  return "";
}

function fmtDate(d) {
  if (!d) return "—";
  try {
    const dt = new Date(String(d).slice(0, 10) + "T00:00:00");
    return dt.toLocaleDateString(undefined, {
      weekday: "short", day: "numeric", month: "short"
    });
  } catch (e) { return d; }
}

function searchableText(item) {
  return JSON.stringify(item).toLowerCase();
}

function getSearchTerm() {
  const el = document.getElementById("search-box");
  return el ? el.value.trim().toLowerCase() : "";
}

function filterBySearch(list) {
  const term = getSearchTerm();
  if (!term) return list;
  return list.filter(function (item) {
    return searchableText(item).indexOf(term) >= 0;
  });
}

function renderToday() {
  const now = new Date();
  const todayIso = now.toISOString().slice(0, 10);
  let items = [];

  (appData.meetings || []).forEach(function (m) {
    if ((m.meeting_date || "").slice(0, 10) === todayIso) {
      items.push({
        kind: "Meeting", icon: "📅", time: m.meeting_time || "",
        title: m.title, sub: m.executive_name, raw: m, status: m.status
      });
    }
  });
  (appData.events || []).forEach(function (e) {
    if ((e.event_date || "").slice(0, 10) === todayIso) {
      items.push({
        kind: "Event", icon: "🎉", time: e.start_time || "",
        title: e.name, sub: e.executive_name, raw: e, status: e.status
      });
    }
  });
  (appData.speaking_engagements || []).forEach(function (s) {
    if ((s.engagement_date || "").slice(0, 10) === todayIso) {
      items.push({
        kind: "Speaking", icon: "🎤", time: s.engagement_time || "",
        title: s.title, sub: s.executive_name, raw: s, status: s.status
      });
    }
  });
  (appData.trips || []).forEach(function (t) {
    const sd = (t.start_date || "").slice(0, 10);
    const ed = (t.end_date || "").slice(0, 10);
    if (sd <= todayIso && ed >= todayIso) {
      items.push({
        kind: "Trip", icon: "✈️", time: "",
        title: t.purpose || t.destination, sub: t.executive_name,
        raw: t, status: t.status
      });
    }
  });

  items.sort(function (a, b) {
    return (a.time || "99:99").localeCompare(b.time || "99:99");
  });
  items = filterBySearch(items);

  if (!items.length) return '<div class="empty">Nothing scheduled today.</div>';

  return items.map(function (i) {
    return card(i.icon + " " + i.title, [
      [i.kind, i.sub],
      [i.time ? "Time" : null, i.time || null],
    ], i.status, null);
  }).join("");
}

function card(title, rows, status, extraHtml) {
  let statusBadge = status ? badge(status.replace(/_/g, " "), statusKind(status)) : "";
  let rowsHtml = "";
  rows.forEach(function (pair) {
    if (!pair[1]) return;
    rowsHtml += '<div class="row"><strong>' + escapeHtml(pair[0]) +
                ':</strong> ' + escapeHtml(pair[1]) + '</div>';
  });
  return '<div class="card">' +
    '<h2>' + escapeHtml(title) + ' ' + statusBadge + '</h2>' +
    rowsHtml +
    (extraHtml || '') +
    '</div>';
}

function renderMeetings() {
  let list = filterBySearch(appData.meetings || []);
  if (!list.length) return '<div class="empty">No meetings.</div>';
  return list.map(function (m) {
    let actions = "";
    const links = [];
    if (m.calendar_event_url)
      links.push('<a href="' + escapeHtml(m.calendar_event_url) +
                 '" target="_blank">📅 Calendar</a>');
    if (m.video_link)
      links.push('<a href="' + escapeHtml(m.video_link) +
                 '" target="_blank">🎥 Video</a>');
    if (m.notes_url)
      links.push('<a href="' + escapeHtml(m.notes_url) +
                 '" target="_blank">📝 ' +
                 escapeHtml(m.notes_label || "Notes") + '</a>');
    if (links.length) actions = '<div class="actions">' + links.join("") + '</div>';

    return card(m.title, [
      ["When", fmtDate(m.meeting_date) + (m.meeting_time ? " · " + m.meeting_time : "")],
      ["With", m.executive_name],
      ["Location", m.location],
    ], m.status, actions);
  }).join("");
}

function renderEvents() {
  let list = filterBySearch(appData.events || []);
  if (!list.length) return '<div class="empty">No events.</div>';
  return list.map(function (e) {
    let actions = "";
    const links = [];
    if (e.calendar_event_url)
      links.push('<a href="' + escapeHtml(e.calendar_event_url) +
                 '" target="_blank">📅 Calendar</a>');
    if (e.rsvp_form_url)
      links.push('<a href="' + escapeHtml(e.rsvp_form_url) +
                 '" target="_blank">📋 RSVP form</a>');
    if (e.run_of_show_url)
      links.push('<a href="' + escapeHtml(e.run_of_show_url) +
                 '" target="_blank">🎬 Run of show</a>');
    if (links.length) actions = '<div class="actions">' + links.join("") + '</div>';
    return card(e.name, [
      ["When", fmtDate(e.event_date) + (e.start_time ? " · " + e.start_time : "")],
      ["With", e.executive_name],
      ["Venue", e.venue_name || e.location_override],
      ["Status", e.status],
    ], null, actions);
  }).join("");
}

function renderSpeaking() {
  let list = filterBySearch(appData.speaking_engagements || []);
  if (!list.length) return '<div class="empty">No speaking engagements.</div>';
  return list.map(function (s) {
    let actions = "";
    const links = [];
    if (s.calendar_event_url)
      links.push('<a href="' + escapeHtml(s.calendar_event_url) +
                 '" target="_blank">📅 Calendar</a>');
    if (s.notes_url)
      links.push('<a href="' + escapeHtml(s.notes_url) +
                 '" target="_blank">📝 Notes</a>');
    if (links.length) actions = '<div class="actions">' + links.join("") + '</div>';

    let feeLine = "";
    if (s.fee_agreed) feeLine = s.fee_agreed + " " + (s.currency || "USD");
    else if (s.fee_offered) feeLine = s.fee_offered + " " + (s.currency || "USD") + " (offered)";

    return card(s.title, [
      ["When", fmtDate(s.engagement_date) + (s.engagement_time ? " · " + s.engagement_time : "")],
      ["With", s.executive_name],
      ["Organizer", s.organizer_name],
      ["Location", s.location],
      ["Fee", feeLine],
    ], s.status, actions);
  }).join("");
}

function renderTrips() {
  let list = filterBySearch(appData.trips || []);
  if (!list.length) return '<div class="empty">No trips.</div>';
  return list.map(function (t) {
    const start = fmtDate(t.start_date);
    const end = fmtDate(t.end_date);
    return card(t.purpose || t.destination, [
      ["Dates", start + " → " + end],
      ["Destination", t.destination],
      ["With", t.executive_name],
    ], t.status, null);
  }).join("");
}

function renderContacts() {
  let list = filterBySearch(appData.contacts || []);
  if (!list.length) return '<div class="empty">No contacts.</div>';
  return list.map(function (c) {
    const actions = [];
    if (c.phone)
      actions.push('<a class="phone" href="tel:' +
                   escapeHtml(c.phone.replace(/\s/g, "")) +
                   '">📞 Call</a>');
    if (c.email)
      actions.push('<a href="mailto:' + escapeHtml(c.email) +
                   '">✉️ Email</a>');
    const actionsHtml = actions.length
      ? '<div class="actions">' + actions.join("") + '</div>' : '';
    return card(c.name, [
      ["Role", c.role],
      ["Company", c.company_id ? "" : "—"],
      ["Phone", c.phone],
      ["Email", c.email],
      ["Location", [c.city, c.country].filter(Boolean).join(", ")],
    ], c.type, actionsHtml);
  }).join("");
}

function renderCurrentTab() {
  const main = document.getElementById("main-content");
  let html = "";
  if (currentTab === "today") html = renderToday();
  else if (currentTab === "meetings") html = renderMeetings();
  else if (currentTab === "events") html = renderEvents();
  else if (currentTab === "speaking") html = renderSpeaking();
  else if (currentTab === "trips") html = renderTrips();
  else if (currentTab === "contacts") html = renderContacts();
  main.innerHTML = html;
}

function attachSearch() {
  const el = document.getElementById("search-box");
  if (el) el.addEventListener("input", renderCurrentTab);
}

function bootstrap() {
  if (!IS_ENCRYPTED) {
    appData = JSON.parse(PAYLOAD_DATA);
    showApp();
  }
  attachSearch();
}

window.addEventListener("DOMContentLoaded", function () {
  document.querySelectorAll(".tabs button").forEach(function (b) {
    b.addEventListener("click", function () {
      setTab(b.dataset.tab);
    });
  });
  bootstrap();
});
"""


# =========================================================
# TEMPLATE ASSEMBLY
# =========================================================


def _build_scope_label(scope, data):
    if scope == "today":
        return "Today only"
    if scope == "event":
        ev = (data.get("events") or [{}])[0]
        return f"Event: {ev.get('name') or '—'}"
    if scope == "trip":
        tr = (data.get("trips") or [{}])[0]
        return f"Trip: {tr.get('purpose') or tr.get('destination') or '—'}"
    return "Full export"


def _render_html(data, password=None):
    payload_json = json.dumps(data, default=str)
    scope = data.get("scope") or "full"
    scope_label = _build_scope_label(scope, data)
    generated = datetime.now().strftime("%d %b %Y · %H:%M")

    if password and CRYPTO_AVAILABLE:
        try:
            ciphertext, salt, iv = encrypt_payload(payload_json, password)
            encrypted = True
        except Exception:
            ciphertext, salt, iv = payload_json, "", ""
            encrypted = False
    else:
        ciphertext = payload_json
        salt = ""
        iv = ""
        encrypted = False

    # Escape closing script tags in payload to avoid breaking the HTML
    def _safe(s):
        return s.replace("</", "<\\/") if isinstance(s, str) else s

    js = MOBILE_JS
    js = js.replace("%%SALT%%", _safe(salt))
    js = js.replace("%%IV%%", _safe(iv))
    js = js.replace("%%DATA%%", _safe(ciphertext))
    js = js.replace("%%ENCRYPTED%%", "true" if encrypted else "false")

    if encrypted:
        password_screen = """
        <div id="password-screen">
          <h2>🔒 Encrypted export</h2>
          <p>Enter the password to view this export.</p>
          <input type="password" id="pw-input" placeholder="Password"
                 autofocus autocomplete="off">
          <button onclick="unlock()">Unlock</button>
          <div class="error" id="pw-error"></div>
        </div>
        """
        app_shell_style = "display:none;"
    else:
        password_screen = ""
        app_shell_style = ""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#2563eb">
<meta name="apple-mobile-web-app-capable" content="yes">
<title>EA Command Center — Mobile</title>
<style>{MOBILE_CSS}</style>
</head>
<body>

{password_screen}

<div id="app-shell" style="{app_shell_style}">
  <header>
    <h1>EA Command Center</h1>
    <p class="meta">Exported {generated} · {_esc(scope_label)}</p>
  </header>

  <nav class="tabs">
    <button data-tab="today" class="active">Today</button>
    <button data-tab="meetings">📅 Meetings</button>
    <button data-tab="events">🎉 Events</button>
    <button data-tab="speaking">🎤 Speaking</button>
    <button data-tab="trips">✈️ Trips</button>
    <button data-tab="contacts">👥 Contacts</button>
  </nav>

  <main>
    <input type="search" class="search" id="search-box"
           placeholder="Search…" autocomplete="off">
    <div id="main-content"></div>
  </main>

  <footer>
    EA Command Center · Read-only mobile export
  </footer>
</div>

<script>
{js}
</script>
</body>
</html>"""
    return html


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


# =========================================================
# PUBLIC API
# =========================================================


def generate_mobile_export(scope="full", scope_id=None, password=None):
    """
    Return the full HTML string for the mobile export.

    scope: "full" | "today" | "event" | "trip"
    scope_id: required for event/trip
    password: if provided and cryptography is available, the export is
              encrypted; otherwise it's plaintext.
    """
    data = build_export_data(scope=scope, scope_id=scope_id)
    if data is None:
        return None
    return _render_html(data, password=password)


def is_encryption_available():
    return CRYPTO_AVAILABLE