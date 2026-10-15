import streamlit as st
import database as db
import doc_generator
import utils
from datetime import datetime, timedelta, date
import csv
import io
import pytz
import os
import json
import re
import pycountry
from openpyxl import Workbook
from excel_export import (
    export_profile_to_excel,
    export_itinerary_to_excel,
    export_expense_to_excel,
    export_spending_to_excel,
    export_company_profile_to_excel,
)
from currency import get_currency_symbol
import currency
import duplicate_detection
import sqlite3
import zipfile
import weather
import streamlit.components.v1 as components


DRESS_CODE_OPTIONS = [
    "No Dress Code", "Business Formal", "Business Casual",
    "Smart Casual", "Casual", "Cocktail", "Black Tie",
    "Formal / Gala", "Other",
]

# Deterministic palette for tag/badge pills.
# Dark saturated hues + white text — passes WCAG AA on both light and
# dark Streamlit themes.
TAG_PALETTE = [
    ("#1e40af", "#ffffff"),  # blue-800
    ("#166534", "#ffffff"),  # green-800
    ("#991b1b", "#ffffff"),  # red-800
    ("#6b21a8", "#ffffff"),  # purple-800
    ("#92400e", "#ffffff"),  # amber-800
    ("#155e75", "#ffffff"),  # cyan-800
    ("#9d174d", "#ffffff"),  # pink-800
    ("#5b21b6", "#ffffff"),  # violet-800
    ("#854d0e", "#ffffff"),  # yellow-800
    ("#065f46", "#ffffff"),  # emerald-800
    ("#334155", "#ffffff"),  # slate-700
    ("#7c2d12", "#ffffff"),  # orange-900
]


def _tag_color(tag):
    """Deterministic (background, foreground) for a given tag string."""
    if not tag:
        return TAG_PALETTE[0]
    h = 0
    for ch in tag.lower():
        h = (h * 31 + ord(ch)) & 0xFFFFFFFF
    return TAG_PALETTE[h % len(TAG_PALETTE)]


def _render_tag_pill(tag):
    """HTML for a single tag pill."""
    bg, fg = _tag_color(tag)
    return (
        f"<span style='background:{bg};color:{fg};"
        f"padding:3px 10px;border-radius:12px;font-size:12px;"
        f"font-weight:500;margin-right:4px;margin-bottom:4px;"
        f"display:inline-block;'>{tag}</span>"
    )


# =========================================================
# HELPERS
# =========================================================

def safe_index(options, value, default="No Preference"):
    if value is None:
        value = default
    try:
        return options.index(value)
    except ValueError:
        return options.index(default)


# Curated common business-travel timezones — surfaced at the top of every
# timezone dropdown so frequent choices are one click away. Entries not
# present in pytz.common_timezones are silently skipped.
COMMON_TIMEZONES = [
    "America/New_York",
    "America/Chicago",
    "America/Denver",
    "America/Los_Angeles",
    "America/Toronto",
    "America/Vancouver",
    "America/Mexico_City",
    "America/Sao_Paulo",
    "Europe/London",
    "Europe/Paris",
    "Europe/Berlin",
    "Europe/Madrid",
    "Europe/Rome",
    "Europe/Istanbul",
    "Europe/Moscow",
    "Africa/Lagos",
    "Africa/Nairobi",
    "Africa/Johannesburg",
    "Asia/Dubai",
    "Asia/Riyadh",
    "Asia/Kolkata",
    "Asia/Singapore",
    "Asia/Hong_Kong",
    "Asia/Shanghai",
    "Asia/Tokyo",
    "Asia/Seoul",
    "Australia/Sydney",
    "Australia/Melbourne",
]


def _tz_display_name(tz):
    """Format a single IANA name as 'Area/City (ABBR)' or fallback to the name."""
    try:
        now = datetime.now(pytz.timezone(tz))
        abbr = now.strftime("%Z") or now.strftime("%z")
        return f"{tz} ({abbr})"
    except Exception:
        return tz


def get_timezone_dropdown_options():
    """
    Return (display_names, tz_map).

    Curated business-travel timezones are surfaced first, followed by the
    rest of pytz.common_timezones so nothing is unreachable. Same signature
    as before — all callers work unchanged.
    """
    display_names = []
    tz_map = {}
    seen_tz = set()

    # 1) Curated shortlist
    for tz in COMMON_TIMEZONES:
        if tz not in pytz.common_timezones or tz in seen_tz:
            continue
        display = _tz_display_name(tz)
        display_names.append(display)
        tz_map[display] = tz
        seen_tz.add(tz)

    # 2) Everything else, alphabetical
    for tz in sorted(pytz.common_timezones):
        if tz in seen_tz:
            continue
        display = _tz_display_name(tz)
        display_names.append(display)
        tz_map[display] = tz
        seen_tz.add(tz)

    return display_names, tz_map


# --- Phase 2 Stage A: standard display formatters ---

def fmt_money(amount, currency="USD"):
    """Format a numeric amount with a currency code."""
    try:
        return f"{float(amount or 0):,.2f} {currency}"
    except (TypeError, ValueError):
        return f"0.00 {currency}"


def fmt_date(d, fmt="%d-%m-%Y"):
    """Format an ISO date string as DD-MM-YYYY (default)."""
    if not d:
        return ""
    try:
        return datetime.fromisoformat(str(d)).strftime(fmt)
    except Exception:
        return str(d)


def fmt_datetime(dt, fmt="%d-%m-%Y %H:%M"):
    """Format an ISO datetime string as DD-MM-YYYY HH:MM (default)."""
    if not dt:
        return ""
    try:
        return datetime.fromisoformat(str(dt)).strftime(fmt)
    except Exception:
        return str(dt)


def _validate_item_times(start_iso, end_iso):
    """
    Return (ok, msg). ok=False when end is strictly before start.
    Equal start and end is allowed (zero-duration item like a status
    marker). Missing or unparseable values pass — those are caught by
    the "Start required" check elsewhere.

    Not a hard block: cross-dateline or cross-timezone entries can
    legitimately produce raw strings where end < start, so callers
    pair this with a "Save anyway?" override.
    """
    if not start_iso or not end_iso:
        return True, ""
    try:
        s = datetime.fromisoformat(start_iso)
        e = datetime.fromisoformat(end_iso)
    except Exception:
        return True, ""
    if e < s:
        return False, (
            f"End time ({e.strftime('%d-%m-%Y %H:%M')}) is before "
            f"start time ({s.strftime('%d-%m-%Y %H:%M')})."
        )
    return True, ""


def _find_duplicate_items(
    existing_items, new_desc, new_start_iso, new_confirmation_code=None, exclude_id=None
):
    """
    Return items on the same trip that look like duplicates of a proposed
    one. Match criteria:
      - same description (case-insensitive) AND same start datetime, OR
      - matching confirmation_code (if both are non-empty).

    `exclude_id` skips an item by DB id, used by the edit handler to
    avoid warning against itself.
    """
    d = (new_desc or "").strip().lower()
    cc = (new_confirmation_code or "").strip().lower()
    matches = []
    for it in existing_items:
        if exclude_id and it.get("id") == exclude_id:
            continue
        same_start = it.get("datetime_start") == new_start_iso
        same_desc = (it.get("description") or "").strip().lower() == d
        same_conf = (
            cc
            and it.get("confirmation_code")
            and it.get("confirmation_code").strip().lower() == cc
        )
        if (same_start and same_desc) or same_conf:
            matches.append(it)
    return matches


def _find_stop_conflicts(
    existing_stops, new_city, new_start_iso, new_end_iso, exclude_idx=None
):
    """
    Return (same_city, other_city) conflicts for a proposed stop.

    same_city  — same city, dates touch or overlap. Likely a duplicate.
    other_city — different city, dates overlap *strictly*. Physically
                 impossible (can't be in two places at once).

    A touching boundary (London ends 17 Oct, Paris starts 17 Oct) is only
    a conflict for the same city — for different cities that's a normal
    travel day. `exclude_idx` skips a stop by its position, used by the
    edit handler to avoid warning against itself.
    """
    cl = new_city.strip().lower()
    same_city = []
    other_city = []
    for i, s in enumerate(existing_stops):
        if exclude_idx is not None and i == exclude_idx:
            continue
        s_city = s["city"].strip().lower()
        if s_city == cl:
            if s["start_date"] <= new_end_iso and s["end_date"] >= new_start_iso:
                same_city.append(s)
        else:
            if s["start_date"] < new_end_iso and s["end_date"] > new_start_iso:
                other_city.append(s)
    return same_city, other_city

def _find_stop_conflicts(existing_stops, new_city, new_start_iso, new_end_iso,
                         exclude_idx=None):
    """
    Return (same_city, other_city) conflicts for a proposed stop.

    same_city  — same city, dates touch or overlap. Likely a duplicate.
    other_city — different city, dates overlap *strictly*. Physically
                 impossible (can't be in two places at once).

    A touching boundary (London ends 17 Oct, Paris starts 17 Oct) is only
    a conflict for the same city — for different cities that's a normal
    travel day. `exclude_idx` skips a stop by its position, used by the
    edit handler to avoid warning against itself.
    """
    cl = new_city.strip().lower()
    same_city = []
    other_city = []
    for i, s in enumerate(existing_stops):
        if exclude_idx is not None and i == exclude_idx:
            continue
        s_city = s["city"].strip().lower()
        if s_city == cl:
            if s["start_date"] <= new_end_iso and s["end_date"] >= new_start_iso:
                same_city.append(s)
        else:
            if s["start_date"] < new_end_iso and s["end_date"] > new_start_iso:
                other_city.append(s)
    return same_city, other_city


def _swap_item_times(a, b):
    """
    Return (a', b') — copies of two item dicts with their datetime_start,
    datetime_end, and cost_date swapped. Everything else is preserved.
    Used by the create-tab ↑/↓ reorder buttons, where items live in
    session state and have no DB IDs yet.
    """
    a2 = dict(a)
    b2 = dict(b)
    a2["datetime_start"], b2["datetime_start"] = (
        b.get("datetime_start"), a.get("datetime_start")
    )
    a2["datetime_end"], b2["datetime_end"] = (
        b.get("datetime_end"), a.get("datetime_end")
    )
    a2["cost_date"], b2["cost_date"] = (
        b.get("cost_date"), a.get("cost_date")
    )
    return a2, b2

# --- Phase 2 Stage C: standard delete-confirmation UI ---


def render_delete_confirmation(flag_key, title, on_confirm, extra_msg=None):
    """
    Render a standard delete confirmation block.

    `on_confirm` may return either:
      - None                      → treated as success, reruns immediately
      - (ok: bool, msg: str)      → message is shown; rerun only on success

    The tuple form exists so callers like delete_company() — which can
    legitimately refuse to delete (e.g. company still has executives) —
    can surface the refusal to the user instead of having the message
    wiped by the immediate rerun.
    """
    st.warning(f"⚠️ {title}")
    if extra_msg:
        st.error(extra_msg)
    col_yes, col_no = st.columns(2)
    with col_yes:
        if st.button("✅ Yes, Delete", key=f"{flag_key}_yes"):
            result = on_confirm()
            st.session_state.pop(flag_key, None)
            if isinstance(result, tuple) and len(result) == 2:
                ok, msg = result
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
                    # Do NOT rerun — let the user read the refusal.
            else:
                st.rerun()
    with col_no:
        if st.button("❌ Cancel", key=f"{flag_key}_no"):
            st.session_state.pop(flag_key, None)
            st.rerun()


def _render_executive_form_fields(
    prefix, defaults, companies, tz_display_names, tz_map, disable_active=False
):
    """
    Renders the shared 9-10 field block used by both the Add Executive
    form and the Edit Executive form. Returns a dict of the current
    values.

    `prefix` — unique string for widget keys (e.g. "add" or "edit").
    `defaults` — dict with the current values, or {} for a fresh form.
    `companies` — list of company dicts for the dropdown.
    `disable_active` — if True, the Active checkbox is hidden (Add form).
    """
    values = {}
    company_options = {c["name"]: c["id"] for c in companies}
    company_names = list(company_options.keys())

    st.markdown("**👤 Identity**")
    # --- Company & identity ---
    current_company_id = defaults.get("company_id")
    curr_company_name = next(
        (n for n, cid in company_options.items() if cid == current_company_id),
        company_names[0] if company_names else "",
    )
    values["company_label"] = st.selectbox(
        "Company*",
        options=company_names,
        index=(
            company_names.index(curr_company_name)
            if curr_company_name in company_names
            else 0
        ),
        key=f"{prefix}_company",
    )
    values["company_id"] = (
        company_options.get(values["company_label"]) if company_names else None
    )

    values["name"] = st.text_input(
        "Full Name*", value=defaults.get("name", ""), key=f"{prefix}_name"
    )
    values["email"] = st.text_input(
        "Email", value=defaults.get("email", ""), key=f"{prefix}_email"
    )

    # --- Core travel preferences ---
    st.markdown("**✈️ Travel Preferences**")
    current_tz = defaults.get("timezone") or _last_used_timezone()
    current_tz_display = next(
        (n for n in tz_display_names if current_tz in n), tz_display_names[0]
    )
    values["tz_display"] = st.selectbox(
        "Timezone",
        options=tz_display_names,
        index=tz_display_names.index(current_tz_display),
        key=f"{prefix}_tz",
    )
    values["timezone"] = tz_map[values["tz_display"]]

    seat_options = ["No Preference", "Aisle", "Window", "Middle"]
    values["seat"] = st.selectbox(
        "Seat Preference",
        options=seat_options,
        index=safe_index(
            seat_options, defaults.get("seat_preference", "No Preference")
        ),
        key=f"{prefix}_seat",
    )

    meal_options = [
        "No Preference",
        "Vegetarian",
        "Vegan",
        "Kosher",
        "Halal",
        "Gluten-Free",
    ]
    values["meal"] = st.selectbox(
        "Meal Preference",
        options=meal_options,
        index=safe_index(
            meal_options, defaults.get("meal_preference", "No Preference")
        ),
        key=f"{prefix}_meal",
    )

    values["diet"] = st.text_input(
        "Dietary Restrictions",
        value=defaults.get("dietary_restrictions", ""),
        key=f"{prefix}_diet",
    )
    values["airline"] = st.text_input(
        "Preferred Airline",
        value=defaults.get("preferred_airline", ""),
        key=f"{prefix}_airline",
    )
    values["tsa"] = st.text_input(
        "TSA PreCheck",
        value=defaults.get("tsa_precheck", ""),
        key=f"{prefix}_tsa",
    )

    # --- Active toggle (Edit only) ---
    if not disable_active:
        st.markdown("**📋 Status**")
        values["active"] = st.checkbox(
            "Active",
            value=bool(defaults.get("is_active", 1)),
            key=f"{prefix}_active",
            help="Uncheck to archive. Archived executives are hidden from "
            "pickers but their data is preserved.",
        )

    return values


_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _is_valid_email(value):
    """Rudimentary check — 'user@host.tld' shape. Empty strings pass
    (email is optional)."""
    if not value:
        return True
    return bool(_EMAIL_RE.match(value.strip()))


def _last_used_timezone():
    """Return the last timezone the user picked for a new exec, or NY."""
    return st.session_state.get("_last_tz", "America/New_York")


def _handle_add_executive_save(vals):
    """Validate and persist a new executive from the shared form renderer."""
    if not vals.get("name"):
        st.error("Full Name is required.")
        return
    if not vals.get("company_id"):
        st.error("Company is required.")
        return
    if not _is_valid_email(vals.get("email")):
        st.error("Email doesn't look valid — check for a missing @ or domain.")
        return

    if vals.get("email"):
        existing = duplicate_detection.find_duplicate_executive(
            vals["email"], vals["name"], vals["company_id"]
        )
        if existing:
            st.warning(
                "⚠️ An executive with the same email or name+company " "already exists:"
            )
            for dup in existing:
                st.write(f"- {dup['name']} (ID: {dup['id']})")
            if not st.checkbox("Add anyway?", key="force_add_exec"):
                st.info(
                    "Tick **Add anyway?** above and click "
                    "**💾 Create Executive** again to proceed."
                )
                return

    db.add_executive(
        vals["company_id"],
        vals["name"],
        vals.get("email") or "",
        vals["timezone"],
        vals["seat"] if vals["seat"] != "No Preference" else "",
        vals.get("diet") or "",
        vals.get("airline") or "",
        vals.get("tsa") or "",
        vals["meal"] if vals["meal"] != "No Preference" else "",
    )

    # Remember timezone for the next new exec
    st.session_state["_last_tz"] = vals["timezone"]

    st.success(
        f"✅ Executive '{vals['name']}' created! "
        "Add passports and memberships in the edit modal."
    )
    st.session_state["show_add_executive"] = False
    st.rerun()


def _handle_edit_executive_save(exec_id, vals):
    """Validate and persist edits to an existing executive."""
    if not vals.get("name"):
        st.error("Full Name is required.")
        return
    if not vals.get("company_id"):
        st.error("Company is required.")
        return
    if not _is_valid_email(vals.get("email")):
        st.error("Email doesn't look valid — check for a missing @ or domain.")
        return

    if vals.get("email"):
        duplicates = duplicate_detection.find_duplicate_executive(
            vals["email"], vals["name"], vals["company_id"]
        )
        # Exclude the exec we're currently editing
        duplicates = [d for d in duplicates if d["id"] != exec_id]
        if duplicates:
            st.warning(
                "⚠️ Another executive with the same email or name+company "
                "already exists:"
            )
            for dup in duplicates:
                st.write(f"- {dup['name']} (ID: {dup['id']})")
            if not st.checkbox("Save anyway?", key="force_save_exec_edit"):
                st.info(
                    "Tick **Save anyway?** above and click "
                    "**💾 Save Changes** again to proceed."
                )
                return

    db.update_executive(
        exec_id,
        vals["company_id"],
        vals["name"],
        vals.get("email") or "",
        vals["timezone"],
        vals["seat"] if vals["seat"] != "No Preference" else "",
        vals.get("diet") or "",
        vals.get("airline") or "",
        vals.get("tsa") or "",
        vals["meal"] if vals["meal"] != "No Preference" else "",
        is_active=1 if vals.get("active", 1) else 0,
    )

    st.success(f"✅ Executive '{vals['name']}' updated!")
    st.session_state["profile_edit_mode"] = False
    st.session_state["show_full_profile"] = False
    st.rerun()


# =========================================================
# DASHBOARD HELPERS
# =========================================================

def _parse_iso_date(s):
    """Return a date object from an ISO string, or None."""
    if not s:
        return None
    try:
        return datetime.fromisoformat(str(s)[:10]).date()
    except Exception:
        return None


def _trip_is_upcoming(trip, today, days_ahead=30):
    start = _parse_iso_date(trip.get("start_date"))
    if not start:
        return False
    return today <= start <= (today + timedelta(days=days_ahead))


def _trip_is_active(trip, today):
    start = _parse_iso_date(trip.get("start_date"))
    end = _parse_iso_date(trip.get("end_date"))
    if not start or not end:
        return False
    return start <= today <= end


def _trip_is_past(trip, today):
    end = _parse_iso_date(trip.get("end_date"))
    if not end:
        return False
    return end < today


def _trips_in_month(trip, year, month):
    """Return True if the trip overlaps the given calendar month."""
    start = _parse_iso_date(trip.get("start_date"))
    end = _parse_iso_date(trip.get("end_date"))
    if not start or not end:
        return False
    month_start = date(year, month, 1)
    if month == 12:
        month_end = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        month_end = date(year, month + 1, 1) - timedelta(days=1)
    return not (end < month_start or start > month_end)


def _trips_in_quarter(trip, year, quarter):
    start_month = (quarter - 1) * 3 + 1
    for m in range(start_month, start_month + 3):
        if _trips_in_month(trip, year, m):
            return True
    return False

def _trip_total_in_usd(trip):
    """
    Return (amount_in_usd, conversion_ok).

    Converts a trip's total_spent from its own base_currency to USD using
    the trip's start date so historical rates are honored. Trips already
    in USD pass through unchanged. If conversion fails, the raw amount is
    returned with ok=False so the caller can decide whether to flag it.
    """
    amt = float(trip.get("total_spent") or 0)
    trip_cur = (trip.get("base_currency") or "USD").upper()
    if trip_cur == "USD":
        return amt, True
    on_date = (trip.get("start_date") or "")[:10] or None
    try:
        return float(currency.convert_amount(amt, trip_cur, "USD", on_date)), True
    except Exception:
        return amt, False
# =========================================================
# BACKUP REMINDER (persistent state)
# =========================================================

BACKUP_STATE_FILE = "ea_app_state.json"


def _load_app_state():
    if os.path.exists(BACKUP_STATE_FILE):
        try:
            with open(BACKUP_STATE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_app_state(state):
    try:
        with open(BACKUP_STATE_FILE, "w") as f:
            json.dump(state, f, default=str)
    except Exception:
        pass


# --- Dismissible warnings state ---

DISMISSED_WARNINGS_PATH = "ea_dismissed_warnings.json"


def _load_dismissed_warnings():
    if not os.path.exists(DISMISSED_WARNINGS_PATH):
        return {}
    try:
        with open(DISMISSED_WARNINGS_PATH, "r") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _save_dismissed_warnings(data):
    try:
        with open(DISMISSED_WARNINGS_PATH, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        st.error(f"Could not save dismissal state: {e}")


def _is_warning_dismissed(key):
    """True if `key` was acknowledged and the dismissal window hasn't lapsed."""
    data = _load_dismissed_warnings()
    entry = data.get(key)
    if not entry:
        return False
    until = entry.get("dismissed_until")
    if not until:
        return False
    try:
        return datetime.now().date() <= datetime.fromisoformat(until).date()
    except Exception:
        return False


def _dismiss_warning(key, days=90):
    """Acknowledge a warning for `days` days."""
    data = _load_dismissed_warnings()
    until = (datetime.now() + timedelta(days=days)).date().isoformat()
    data[key] = {"dismissed_until": until}
    _save_dismissed_warnings(data)


def _mark_backup_done():
    state = _load_app_state()
    state["last_backup_date"] = datetime.now().isoformat()
    _save_app_state(state)


def _days_since_backup():
    state = _load_app_state()
    last = state.get("last_backup_date")
    if not last:
        return None
    try:
        d = datetime.fromisoformat(last).date()
        return (datetime.now().date() - d).days
    except Exception:
        return None


def _in_month_end_window(days_before=5):
    """Return True if today is within the last N days of the current month."""
    today = datetime.now().date()
    next_month_first = (today.replace(day=28) + timedelta(days=4)).replace(day=1)
    days_left = (next_month_first - today).days
    return days_left <= days_before


def _should_show_backup_reminder():
    """Show if we're in the last 5 days of the month AND last backup >25d ago
    (or never backed up)."""
    if not _in_month_end_window(5):
        return False
    days = _days_since_backup()
    if days is None:
        return True
    return days >= 25

# Module-level registry tracking which container_keys have been mounted
# in the current Streamlit script run. Streamlit reruns the whole script
# top-to-bottom on every interaction, so this dict is rebuilt each run —
# it only exists to catch same-run collisions, not cross-run ones.
_BULK_PASTE_MOUNTS_THIS_RUN = set()


def _render_bulk_paste_itinerary(
    container_key, default_timezone, on_import, is_locked=False
):
    """
    Bulk paste UI for itinerary items. `on_import(items)` is called with
    the parsed list when the user confirms.

    `container_key` MUST be unique per mount point. All internal widget
    keys are prefixed with it, so two mounts sharing a key would collide
    inside Streamlit's widget cache and raise a confusing duplicate-key
    error at an arbitrary downstream line. Mounts today:

      - Tab 1 → "create_bulk_paste"
      - Trip modal → f"modal_bulk_paste_{trip_id_modal}"

    The registry below enforces this in development; the mount fails
    loudly here rather than silently misbehaving later.
    """

    with st.expander("📋 Paste from spreadsheet", expanded=False):
        st.caption(
            "Paste rows copied from Excel, Google Sheets, or a CSV file. "
            "A header row is auto-detected. Without a header, columns are "
            "read positionally: **type, description, start, end, location, "
            "cost, currency, confirmation**."
        )
        raw = st.text_area(
            "Paste here",
            height=180,
            key=f"{container_key}_raw",
            placeholder=(
                "type\tdescription\tstart\tend\tlocation\tcost\tcurrency\tconfirmation\n"
                "Flight\tLHR → JFK BA 112\t2025-04-12 08:30\t"
                "2025-04-12 11:45\tHeathrow T5\t0\tUSD\tBA-99812\n"
                "Hotel\tThe Peninsula\t2025-04-12 15:00\t"
                "2025-04-14 11:00\tNYC\t0\tUSD\tPEN-7743"
            ),
        )

        parsed, errors = ([], [])
        if raw and raw.strip():
            parsed, errors = utils.parse_itinerary_paste(
                raw, default_timezone=default_timezone
            )
            # --- Post-parse validation: flag rows where end < start ---
            # Rows are still imported; the warning surfaces in the
            # existing "N warning(s)" expander so the user can spot
            # and fix them afterward.
            for _it in parsed:
                _ok, _msg = _validate_item_times(
                    _it.get("datetime_start"), _it.get("datetime_end")
                )
                if not _ok:
                    _desc = _it.get("description", "") or "(no description)"
                    errors.append(f"'{_desc}': {_msg}")

        if parsed:
            st.success(f"✅ Parsed {len(parsed)} item(s).")
            with st.expander(f"👁️ Preview ({len(parsed)})", expanded=True):
                preview_rows = [
                    {
                        "Type": it["item_type"],
                        "Description": it["description"],
                        "Start": it["datetime_start"][:16].replace("T", " "),
                        "End": it["datetime_end"][:16].replace("T", " "),
                        "Location": it["location"],
                        "Cost": f"{it['cost']:.2f}",
                        "Ccy": it["cost_currency"],
                        "Conf": it["confirmation_code"],
                    }
                    for it in parsed
                ]
                st.dataframe(
                    preview_rows, use_container_width=True, hide_index=True
                )

        if errors:
            with st.expander(f"⚠️ {len(errors)} warning(s)", expanded=False):
                for e in errors:
                    st.caption(f"• {e}")

        if not is_locked:
            import_button_label = (
                f"➕ Import {len(parsed)} item(s)"
                if parsed else "➕ Import"
            )
            if st.button(
                import_button_label,
                key=f"{container_key}_import",
                type="primary",
                disabled=not parsed,
            ):
                on_import(parsed)
                st.session_state.pop(f"{container_key}_raw", None)
                st.rerun()


# --- Existing helpers ---


def get_company_contact_options(company_id):
    contacts = db.get_contacts_for_local_support(company_id)
    options = {}
    for c in contacts:
        role = c.get("role", "") or ""
        suffix = " · freelance" if not c.get("company_id") else ""
        options[f"{c['name']} ({role}){suffix}".strip()] = c["id"]
    return options


def get_local_support_options(company_id):
    """
    Contacts available as Local Support for a trip.
    Company contacts + freelance (no-company) contacts, active only.
    """
    contacts = db.get_contacts_for_local_support(company_id)
    options = {}
    for c in contacts:
        role = c.get("role", "") or ""
        suffix = " · freelance" if not c.get("company_id") else ""
        options[f"{c['name']} ({role}){suffix}".strip()] = c["id"]
    return options


def get_contact_names(contact_ids):
    if not contact_ids:
        return ""
    conn = sqlite3.connect(db.DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        placeholders = ",".join(["?"] * len(contact_ids))
        c.execute(
            f"SELECT name FROM contacts WHERE id IN ({placeholders})",
            contact_ids,
        )
        return ", ".join(row[0] for row in c.fetchall())
    finally:
        conn.close()


def get_trip_delegation_options(trip_id):
    members = db.get_trip_delegation_members(trip_id)
    return {
        f"{m['name']} ({m.get('role', '')})".strip(): m["id"]
        for m in members
    }


def format_item_datetime(item, exec_timezone, display_mode="Home"):
    tz_str = item.get("timezone") or exec_timezone
    display_tz = exec_timezone if display_mode == "Home" else tz_str
    return utils.format_datetime_with_timezone(
        item["datetime_start"], tz_str, display_tz
    )

# =========================================================
# MEETING CONSTANTS
# =========================================================

MEETING_TYPES = [
    "Internal", "Client", "Board", "Candidate", "Personal", "VIP", "Other",
]

MEETING_STATUSES = [
    "draft", "request_sent", "pending", "confirmed", "tentative",
    "declined", "no_response", "rescheduled", "cancelled", "completed",
]

MEETING_STATUS_LABELS = {
    "draft": "📝 Draft",
    "request_sent": "📤 Request Sent",
    "pending": "⏳ Pending",
    "confirmed": "✅ Confirmed",
    "tentative": "❓ Tentative",
    "declined": "❌ Declined",
    "no_response": "🔇 No Response",
    "rescheduled": "🔄 Rescheduled",
    "cancelled": "🚫 Cancelled",
    "completed": "🏁 Completed",
}

ATTENDEE_STATUSES = [
    "pending", "accepted", "declined", "tentative", "no_response", "waitlist",
]

ACTION_ITEM_STATUSES = [
    "not_started", "in_progress", "blocked", "done", "overdue",
]

ACTION_ITEM_STATUS_LABELS = {
    "not_started": "⭕ Not Started",
    "in_progress": "🔄 In Progress",
    "blocked": "🚧 Blocked",
    "done": "✅ Done",
    "overdue": "⚠️ Overdue",
}


def _meeting_status_label(status):
    return MEETING_STATUS_LABELS.get(status, status)


def _action_status_label(status):
    return ACTION_ITEM_STATUS_LABELS.get(status, status)

# =========================================================
# COMMUNICATION CONSTANTS
# =========================================================

COMM_CHANNELS = [
    "email", "phone", "in_person", "video", "sms", "whatsapp", "other",
]

COMM_CHANNEL_LABELS = {
    "email": "✉️ Email",
    "phone": "📞 Phone",
    "in_person": "🤝 In Person",
    "video": "🎥 Video Call",
    "sms": "💬 SMS",
    "whatsapp": "💚 WhatsApp",
    "other": "📌 Other",
}

COMM_DIRECTIONS = ["outbound", "inbound"]

COMM_DIRECTION_LABELS = {
    "outbound": "↗️ Outbound",
    "inbound": "↙️ Inbound",
}


def _channel_label(ch):
    return COMM_CHANNEL_LABELS.get(ch, ch or "—")


def _direction_label(d):
    return COMM_DIRECTION_LABELS.get(d, d or "—")

# =========================================================
# EVENT CONSTANTS
# =========================================================

EVENT_TYPES = [
    "Dinner", "Offsite", "Conference", "Retreat", "Client Party",
    "Board Meeting", "Reception", "Awards", "Launch", "Other",
]

EVENT_STATUSES = [
    "planning", "confirmed", "in_progress", "day_of", "completed",
    "cancelled",
]

EVENT_STATUS_LABELS = {
    "planning": "📝 Planning",
    "confirmed": "✅ Confirmed",
    "in_progress": "🔄 In Progress",
    "day_of": "🎯 Day Of",
    "completed": "🏁 Completed",
    "cancelled": "🚫 Cancelled",
}

GUEST_STATUSES = [
    "not_invited", "invited", "pending", "accepted", "declined",
    "waitlist", "cancelled", "checked_in",
]

GUEST_STATUS_LABELS = {
    "not_invited": "— Not Invited",
    "invited": "📤 Invited",
    "pending": "⏳ Pending",
    "accepted": "✅ Accepted",
    "declined": "❌ Declined",
    "waitlist": "⏸ Waitlist",
    "cancelled": "🚫 Cancelled",
    "checked_in": "🎯 Checked In",
}

GUEST_STATUS_ORDER = [
    "accepted", "pending", "invited", "not_invited",
    "declined", "waitlist", "cancelled", "checked_in",
]


def _event_status_label(s):
    return EVENT_STATUS_LABELS.get(s, s or "—")


def _guest_status_label(s):
    return GUEST_STATUS_LABELS.get(s, s or "—")

# =========================================================
# VENDOR & BUDGET CONSTANTS
# =========================================================

VENDOR_CATEGORIES = [
    "Venue", "Catering", "AV", "Florist", "Photographer",
    "Transport", "Staff", "Printing", "Decor", "Entertainment",
    "Security", "Other",
]

VENDOR_ENGAGEMENT_STATUSES = [
    "shortlisted", "contacted", "quoted", "negotiating",
    "confirmed", "cancelled",
]

VENDOR_ENGAGEMENT_STATUS_LABELS = {
    "shortlisted": "📝 Shortlisted",
    "contacted": "📤 Contacted",
    "quoted": "💬 Quoted",
    "negotiating": "🤝 Negotiating",
    "confirmed": "✅ Confirmed",
    "cancelled": "🚫 Cancelled",
}

CONTRACT_STATUSES = [
    "not_sent", "sent", "signed", "declined",
]

CONTRACT_STATUS_LABELS = {
    "not_sent": "— Not Sent",
    "sent": "📤 Sent",
    "signed": "✍️ Signed",
    "declined": "❌ Declined",
}

PAYMENT_METHODS = [
    "bank_transfer", "card", "cash", "check", "other",
]

PAYMENT_METHOD_LABELS = {
    "bank_transfer": "🏦 Bank Transfer",
    "card": "💳 Card",
    "cash": "💵 Cash",
    "check": "🧾 Check",
    "other": "📌 Other",
}

BUDGET_ITEM_STATUSES = [
    "estimated", "committed", "paid", "reconciled",
]

BUDGET_ITEM_STATUS_LABELS = {
    "estimated": "📝 Estimated",
    "committed": "🤝 Committed",
    "paid": "✅ Paid",
    "reconciled": "🔒 Reconciled",
}


def _vendor_status_label(s):
    return VENDOR_ENGAGEMENT_STATUS_LABELS.get(s, s or "—")


def _contract_status_label(s):
    return CONTRACT_STATUS_LABELS.get(s, s or "—")


def _payment_method_label(m):
    return PAYMENT_METHOD_LABELS.get(m, m or "—")


def _budget_item_status_label(s):
    return BUDGET_ITEM_STATUS_LABELS.get(s, s or "—")

# =========================================================
# SPEAKING ENGAGEMENT CONSTANTS
# =========================================================

SPEAKING_TYPES = [
    "Keynote", "Panel", "Fireside", "Podcast", "Webinar",
    "Workshop", "Guest Lecture", "Interview", "Awards", "Other",
]

SPEAKING_FORMATS = ["In-Person", "Virtual", "Hybrid"]

SPEAKING_STATUSES = [
    "invited", "interested", "declined", "cfp_submitted",
    "under_review", "accepted", "negotiating", "confirmed",
    "prep_in_progress", "ready", "delivered", "follow_up",
    "closed", "cancelled",
]

SPEAKING_STATUS_LABELS = {
    "invited": "📨 Invited",
    "interested": "💭 Interested",
    "declined": "❌ Declined",
    "cfp_submitted": "📤 CFP Submitted",
    "under_review": "🔍 Under Review",
    "accepted": "✅ Accepted",
    "negotiating": "🤝 Negotiating",
    "confirmed": "🎯 Confirmed",
    "prep_in_progress": "📝 Prep",
    "ready": "🚀 Ready",
    "delivered": "🎤 Delivered",
    "follow_up": "📮 Follow-Up",
    "closed": "🔒 Closed",
    "cancelled": "🚫 Cancelled",
}

# Ordered pipeline for Kanban view
SPEAKING_PIPELINE_ORDER = [
    "invited", "interested", "cfp_submitted", "under_review",
    "accepted", "negotiating", "confirmed", "prep_in_progress",
    "ready", "delivered", "follow_up", "closed",
]

SPEAKING_DELIVERABLE_STATUSES = [
    "not_started", "in_progress", "submitted", "done",
]

SPEAKING_DELIVERABLE_STATUS_LABELS = {
    "not_started": "⭕ Not Started",
    "in_progress": "🔄 In Progress",
    "submitted": "📤 Submitted",
    "done": "✅ Done",
}

SPEAKING_CFP_STATUSES = [
    "draft", "submitted", "under_review", "accepted",
    "rejected", "waitlisted",
]

SPEAKING_CFP_STATUS_LABELS = {
    "draft": "📝 Draft",
    "submitted": "📤 Submitted",
    "under_review": "🔍 Under Review",
    "accepted": "✅ Accepted",
    "rejected": "❌ Rejected",
    "waitlisted": "⏸ Waitlisted",
}

RECORDING_RIGHTS_OPTIONS = [
    "Organizer only", "Shared", "Speaker only", "Not specified",
]


def _speaking_status_label(s):
    return SPEAKING_STATUS_LABELS.get(s, s or "—")


def _deliverable_status_label(s):
    return SPEAKING_DELIVERABLE_STATUS_LABELS.get(s, s or "—")


def _cfp_status_label(s):
    return SPEAKING_CFP_STATUS_LABELS.get(s, s or "—")

# =========================================================
# CROSS-LINK HELPERS
# =========================================================


def _suggest_trip_for_date(target_date, exec_id, city_hint=None):
    """
    Return a list of candidate trips whose date range covers target_date,
    for the given executive. If city_hint is provided, trips whose
    destination or any stop city matches are prioritized.
    """
    if not target_date:
        return []
    target_iso = (
        target_date.isoformat()
        if hasattr(target_date, "isoformat")
        else str(target_date)
    )
    candidates = db.get_trips_overlapping_dates(
        target_iso, target_iso, exec_id=exec_id
    )
    if not candidates:
        return []

    if not city_hint:
        return candidates

    # Score candidates: any stop city matching the hint = 0 (best)
    hint = city_hint.strip().lower()
    scored = []
    for t in candidates:
        stops = db.get_trip_stops(t["id"])
        matches = any(
            hint in (s.get("city") or "").lower()
            for s in stops
        )
        scored.append((0 if matches else 1, t))
    scored.sort(key=lambda x: x[0])
    return [t for _, t in scored]


def _trip_dropdown_options(exec_id, current_trip_id=None):
    """
    Return {label: id} for the trip picker on edit forms.
    Format: "— No trip —" then "purpose (start → end)".
    """
    options = {"— No trip —": None}
    if not exec_id:
        return options
    conn = sqlite3.connect(db.DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT id, purpose, destination, start_date, end_date
               FROM trips
               WHERE exec_id = ?
               ORDER BY start_date DESC
               LIMIT 50""",
            (exec_id,),
        )
        for r in c.fetchall():
            label = (
                f"{r['purpose'] or r['destination'] or 'Trip'}"
                f" ({r['start_date']} → {r['end_date']})"
            )
            options[label] = r["id"]
    finally:
        conn.close()
    return options

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(page_title="EA Command Center", layout="wide")

st.markdown("""
<style>
    .stTextInput input:focus, .stNumberInput input:focus, .stDateInput input:focus {
        border-color: #87CEEB !important;
        box-shadow: 0 0 0 0.2rem rgba(135, 206, 235, 0.4) !important;
    }
    .stTextInput input:hover, .stNumberInput input:hover, .stDateInput input:hover {
        border-color: #87CEEB !important;
    }
    .stTextArea textarea:focus {
        border-color: #87CEEB !important;
        box-shadow: 0 0 0 0.2rem rgba(135, 206, 235, 0.4) !important;
    }
    .stTextArea textarea:hover { border-color: #87CEEB !important; }
    .stSelectbox div[data-baseweb="select"]:hover { border-color: #87CEEB !important; }
    .stSelectbox div[data-baseweb="select"]:focus-within {
        border-color: #87CEEB !important;
        box-shadow: 0 0 0 0.2rem rgba(135, 206, 235, 0.4) !important;
    }
</style>
""", unsafe_allow_html=True)


if "upload_counter" not in st.session_state:
    st.session_state.upload_counter = 0

st.title("EA Command Center")
db.init_db()


# =========================================================
# MONTHLY BACKUP REMINDER
# =========================================================

if _should_show_backup_reminder():
    days_since = _days_since_backup()
    if days_since is None:
        msg = "You haven't backed up the database yet."
    else:
        msg = f"Last backup was **{days_since} days ago**."

    st.warning(
        f"📅 **Monthly backup reminder** — {msg} "
        "It's the end of the month. Download a copy of your database "
        "and store it somewhere safe (cloud drive, external disk)."
    )

    col1, col2, col3 = st.columns([1, 1, 3])
    with col1:
        try:
            with open(db.DB_PATH, "rb") as _f:
                _db_bytes = _f.read()
            st.download_button(
                "⬇️ Download Backup (.db)",
                data=_db_bytes,
                file_name=(
                    f"ea_command_center_backup_"
                    f"{datetime.now().strftime('%Y%m%d_%H%M')}.db"
                ),
                mime="application/octet-stream",
                key="backup_db_download_banner",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Could not read database: {e}")
    with col2:
        if st.button(
            "✅ Mark as Backed Up",
            key="mark_backed_up_banner",
            use_container_width=True,
        ):
            _mark_backup_done()
            st.rerun()


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.header("👤 Executive Management")

include_inactive_execs = st.sidebar.checkbox(
    "Show inactive executives",
    value=False,
    key="show_inactive_execs",
    help="Uncheck to hide deactivated executives from this dropdown.",
)

executives = db.get_all_executives(active_only=not include_inactive_execs)
if not executives:
    st.sidebar.warning("No executives yet. Use **➕ Add New Executive** below.")
    exec_id = None
    profile = None
else:
    exec_options = {}
    for e in executives:
        label = f"{e['name']} (ID: {e['id']})"
        if not e.get("is_active", 1):
            label = "⚠️ " + label
        exec_options[label] = e["id"]
    selected_label = st.sidebar.selectbox(
        "Choose Executive", list(exec_options.keys())
    )
    exec_id = exec_options[selected_label]
    profile = db.get_executive_profile(exec_id)


if st.sidebar.button("➕ Add New Executive", use_container_width=True):
    st.session_state["show_add_executive"] = True

if st.session_state.get("show_add_executive", False):
    with st.popover("➕ Add New Executive", use_container_width=True):
        st.subheader("Add New Executive")
        with st.expander("➕ Add New Company"):
            new_comp_name = st.text_input("Company Name", key="add_new_comp_name")
            new_comp_cc = st.text_input("Default Cost Center (optional)", key="add_new_comp_cc")
            new_comp_policy = st.text_area("Policy Notes (optional)", key="add_new_comp_policy")
            if st.button("Add Company", key="add_new_comp_btn"):
                if new_comp_name:
                    db.add_company(new_comp_name, new_comp_cc, new_comp_policy)
                    st.success(f"Company '{new_comp_name}' added!")
                    st.rerun()
                else:
                    st.warning("Company Name required.")

        with st.form("add_executive_form"):
            companies = db.get_all_companies(active_only=True)
            if not companies:
                st.warning("No companies available. Add one using the expander above.")
                st.form_submit_button("💾 Create Executive", disabled=True)
            else:
                tz_display_names, tz_map = get_timezone_dropdown_options()
                vals = _render_executive_form_fields(
                    prefix="add",
                    defaults={},
                    companies=companies,
                    tz_display_names=tz_display_names,
                    tz_map=tz_map,
                    disable_active=True,
                )
                col_save, col_cancel = st.columns(2)
                with col_save:
                    submitted = st.form_submit_button("💾 Create Executive")
                with col_cancel:
                    cancelled = st.form_submit_button("❌ Cancel")

                if cancelled:
                    st.session_state["show_add_executive"] = False
                    st.rerun()

                if submitted:
                    _handle_add_executive_save(vals)
if profile:
    with st.sidebar.expander("📋 Quick Profile", expanded=False):
        st.write(f"**{profile['name']}**")
        if not profile.get("is_active", 1):
            st.caption("⚠️ Inactive")
        st.write(f"🏢 {profile.get('company_name', 'N/A')}")
        st.write(f"🕐 {profile.get('timezone', 'N/A')}")
        st.write(f"💺 {profile.get('seat_preference', 'N/A')}")
        mems = db.get_memberships(exec_id)
        if mems:
            st.caption(f"✈️ {len(mems)} memberships")

        # --- Passport expiry warning ---
        # Most countries require ≥6 months validity at entry. Warn on any
        # passport expiring within 180 days, escalate at 90 and on expiry.
        # Each warning can be acknowledged for 90 days.
        passports = db.get_passports(exec_id)
        today = datetime.now().date()
        for p in passports:
            exp = p.get("expiry_date")
            if not exp:
                continue
            try:
                exp_date = datetime.fromisoformat(exp).date()
            except Exception:
                continue
            days_left = (exp_date - today).days
            country = p.get("country") or "Passport"
            warn_key = f"passport_{p['id']}"

            if _is_warning_dismissed(warn_key):
                continue

            # Compose the warning text (unchanged thresholds)
            if days_left < 0:
                st.error(f"🛂 {country}: **EXPIRED** ({fmt_date(exp)})")
            elif days_left < 90:
                st.error(
                    f"🛂 {country}: expires in **{days_left} days** "
                    f"({fmt_date(exp)})"
                )
            elif days_left < 180:
                st.warning(
                    f"🛂 {country}: expires in {days_left} days "
                    f"({fmt_date(exp)})"
                )
            else:
                continue  # nothing shown, no dismiss button needed

            if st.button(
                "✔ Acknowledge for 90 days",
                key=f"dismiss_passport_warn_{p['id']}",
                use_container_width=True,
            ):
                _dismiss_warning(warn_key, days=90)
                st.rerun()

        if st.button("👤 View Full Profile", use_container_width=True):
            st.session_state["show_full_profile"] = True
            st.session_state["profile_edit_mode"] = False

if st.session_state.get("show_full_profile", False):
    with st.popover("👤 Full Profile", use_container_width=True):
        if st.session_state.get("profile_edit_mode", False):
            st.subheader(f"✏️ Editing: {profile['name']}")
            st.caption(
                f"Company: {profile.get('company_name', '—')}  ·  "
                f"ID: {exec_id}  ·  "
                f"{'Active' if profile.get('is_active', 1) else '⚠️ Inactive'}"
            )
            with st.form("edit_exec_popover"):
                companies = db.get_all_companies(active_only=False)
                tz_display_names, tz_map = get_timezone_dropdown_options()
                vals = _render_executive_form_fields(
                    prefix="edit",
                    defaults=profile,
                    companies=companies,
                    tz_display_names=tz_display_names,
                    tz_map=tz_map,
                    disable_active=False,
                )
                col_save, col_cancel, col_delete = st.columns(3)
                with col_save:
                    submitted = st.form_submit_button("💾 Save Changes")
                with col_cancel:
                    cancelled = st.form_submit_button("❌ Cancel")
                with col_delete:
                    if st.form_submit_button("🗑️ Delete Executive", type="primary"):
                        st.session_state["show_delete_confirmation"] = True

                if cancelled:
                    st.session_state["profile_edit_mode"] = False
                    st.rerun()
                if submitted:
                    _handle_edit_executive_save(exec_id, vals)

            # Passports (edit)
            st.subheader("🛂 Passports")
            passports = db.get_passports(exec_id)
            if passports:
                for p in passports:
                    col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 1, 1])
                    with col1:
                        st.write(f"{p['country']}: {p['passport_number']}")
                    with col2:
                        st.write(f"Exp: {fmt_date(p.get('expiry_date'))}")
                    with col3:
                        st.write((p.get("notes") or "")[:30])
                    with col4:
                        if st.button("✏️", key=f"edit_pass_{p['id']}"):
                            st.session_state[f"editing_passport_{p['id']}"] = True
                    with col5:
                        if st.button("🗑️", key=f"del_pass_edit_{p['id']}"):
                            db.delete_passport(p["id"])
                            st.rerun()

                    if st.session_state.get(f"editing_passport_{p['id']}", False):
                        with st.expander(f"Edit Passport: {p['country']}", expanded=True):
                            with st.form(key=f"edit_pass_form_{p['id']}"):
                                edit_country = st.text_input("Country", value=p["country"], key=f"edit_pass_country_{p['id']}")
                                edit_pass_num = st.text_input("Passport Number", value=p["passport_number"], key=f"edit_pass_num_{p['id']}")
                                edit_expiry = st.date_input("Expiry Date",
                                    value=(datetime.fromisoformat(p["expiry_date"]) if p.get("expiry_date") else None),
                                    key=f"edit_pass_expiry_{p['id']}")
                                edit_issued = st.date_input("Issued Date",
                                    value=(datetime.fromisoformat(p["issued_date"]) if p.get("issued_date") else None),
                                    key=f"edit_pass_issued_{p['id']}")
                                edit_notes = st.text_area("Notes", value=p.get("notes", ""), key=f"edit_pass_notes_{p['id']}")
                                col_save, col_cancel = st.columns(2)
                                with col_save:
                                    if st.form_submit_button("💾 Save"):
                                        if edit_country and edit_pass_num:
                                            db.update_passport(p["id"], edit_country, edit_pass_num,
                                                expiry_date=(edit_expiry.isoformat() if edit_expiry else None),
                                                issued_date=(edit_issued.isoformat() if edit_issued else None),
                                                notes=edit_notes)
                                            st.session_state.pop(f"editing_passport_{p['id']}", None)
                                            st.rerun()
                                        else:
                                            st.warning("Country and Passport Number are required.")
                                with col_cancel:
                                    if st.form_submit_button("❌ Cancel"):
                                        st.session_state.pop(f"editing_passport_{p['id']}", None)
                                        st.rerun()
            else:
                st.caption("No passports yet. Use **➕ Add Passport** below.")

            with st.expander("➕ Add Passport"):
                col_c, col_n = st.columns(2)
                with col_c:
                    new_country = st.text_input("Country", key="edit_pass_country")
                with col_n:
                    new_pass_num = st.text_input("Passport Number", key="edit_pass_num")
                col_e, col_i = st.columns(2)
                with col_e:
                    new_expiry = st.date_input("Expiry Date", value=None, key="edit_pass_expiry")
                with col_i:
                    new_issued = st.date_input("Issued Date", value=None, key="edit_pass_issued")
                new_notes_pass = st.text_area("Notes", key="edit_pass_notes")

                if st.button("➕ Add Passport", key="edit_pass_btn"):
                    if not new_country or not new_pass_num:
                        st.warning("Country and Passport Number required.")
                    else:
                        proceed = True
                        dupes = db.find_duplicate_passports(
                            exec_id, new_country, new_pass_num
                        )
                        if dupes:
                            st.warning(
                                f"⚠️ This executive already has a "
                                f"{new_country} passport on file:"
                            )
                            for d in dupes:
                                st.write(
                                    f"- {d['country']} · {d['passport_number']} "
                                    f"(expires {d.get('expiry_date') or '—'})"
                                )
                            if not st.checkbox(
                                "Add anyway?", key="force_add_passport"
                            ):
                                proceed = False
                                st.info(
                                    "Tick **Add anyway?** above and click "
                                    "**➕ Add Passport** again to save the "
                                    "additional passport."
                                )
                        if proceed:
                            db.add_passport(
                                exec_id, new_country, new_pass_num,
                                expiry_date=new_expiry.isoformat() if new_expiry else None,
                                issued_date=new_issued.isoformat() if new_issued else None,
                                notes=new_notes_pass,
                            )
                            st.session_state.pop("force_add_passport", None)
                            st.rerun()

            # Memberships (edit)
            st.subheader("✈️ Memberships")
            mems = db.get_memberships(exec_id)
            if mems:
                for m in mems:
                    col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 1, 1])
                    with col1:
                        emoji = ("✈️" if m["category"] == "airline" else
                                 "🏨" if m["category"] == "hotel" else
                                 "🚗" if m["category"] == "car rental" else
                                 "🛋️" if m["category"] == "lounge" else
                                 "🚄" if m["category"] == "rail" else
                                 "⛴️" if m["category"] == "ferry" else
                                 "🚗" if m["category"] == "ride-share" else
                                 "💳" if m["category"] == "credit card" else "📌")
                        st.write(f"{emoji} {m['program_name']}: {m['membership_number']}")
                    with col2:
                        details = [d for d in [m.get("tier"), m.get("alliance"), m.get("airport_code")] if d]
                        st.write(", ".join(details) if details else "")
                    with col3:
                        st.write((m.get("notes") or "")[:30])
                    with col4:
                        if st.button("✏️", key=f"edit_mem_{m['id']}"):
                            st.session_state[f"editing_membership_{m['id']}"] = True
                    with col5:
                        if st.button("🗑️", key=f"del_mem_edit_{m['id']}"):
                            db.delete_membership(m["id"])
                            st.rerun()

                    if st.session_state.get(f"editing_membership_{m['id']}", False):
                        with st.expander(f"Edit Membership: {m['program_name']}", expanded=True):
                            with st.form(key=f"edit_mem_form_{m['id']}"):
                                edit_cat = st.selectbox("Category",
                                    ["Airline", "Hotel", "Car Rental", "Lounge", "Rail", "Ferry", "Ride-Share", "Credit Card"],
                                    index=["airline", "hotel", "car rental", "lounge", "rail", "ferry", "ride-share", "credit card"].index(m["category"]),
                                    key=f"edit_mem_cat_{m['id']}")
                                edit_prog = st.text_input("Program Name", value=m["program_name"], key=f"edit_mem_prog_{m['id']}")
                                edit_num = st.text_input("Membership Number", value=m["membership_number"], key=f"edit_mem_num_{m['id']}")

                                with st.expander("More details (optional)"):
                                    edit_tier = edit_alliance = edit_airport = edit_notes = None
                                    if edit_cat == "Airline":
                                        edit_tier = st.text_input("Tier", value=m.get("tier") or "", key=f"edit_mem_tier_{m['id']}")
                                        edit_alliance = st.text_input("Alliance", value=m.get("alliance") or "", key=f"edit_mem_alliance_{m['id']}")
                                        edit_airport = st.text_input("Airport Code", value=m.get("airport_code") or "", key=f"edit_mem_airport_{m['id']}")
                                        edit_notes = st.text_area("Notes", value=m.get("notes") or "", key=f"edit_mem_notes_{m['id']}")
                                    elif edit_cat in ("Hotel", "Car Rental", "Lounge", "Rail", "Ferry", "Ride-Share", "Credit Card"):
                                        edit_tier = st.text_input("Status/Tier", value=m.get("tier") or "", key=f"edit_mem_tier_{m['id']}")
                                        edit_notes = st.text_area("Notes", value=m.get("notes") or "", key=f"edit_mem_notes_{m['id']}")
                                    else:
                                        edit_notes = st.text_area("Notes", value=m.get("notes") or "", key=f"edit_mem_notes_{m['id']}")

                                col_save, col_cancel = st.columns(2)
                                with col_save:
                                    if st.form_submit_button("💾 Save"):
                                        if edit_prog and edit_num:
                                            db.update_membership(m["id"], edit_cat.lower(), edit_prog, edit_num,
                                                tier=edit_tier, alliance=edit_alliance,
                                                airport_code=edit_airport, notes=edit_notes)
                                            st.session_state.pop(f"editing_membership_{m['id']}", None)
                                            st.rerun()
                                        else:
                                            st.warning("Program Name and Membership Number are required.")
                                with col_cancel:
                                    if st.form_submit_button("❌ Cancel"):
                                        st.session_state.pop(f"editing_membership_{m['id']}", None)
                                        st.rerun()
            else:
                st.caption("No memberships yet. Use **➕ Add Membership** below.")

            with st.expander("➕ Add Membership"):
                col_cat, col_name, col_num = st.columns(3)
                with col_cat:
                    new_cat = st.selectbox("Category",
                        ["Airline", "Hotel", "Car Rental", "Lounge", "Rail", "Ferry", "Ride-Share", "Credit Card"],
                        key="edit_mem_cat")
                with col_name:
                    new_name_mem = st.text_input("Program Name", key="edit_mem_name")
                with col_num:
                    new_num_mem = st.text_input("Membership Number", key="edit_mem_num")

                with st.expander("➕ More details (optional)"):
                    col_extra1, col_extra2 = st.columns(2)
                    new_tier = new_alliance = new_airport = new_notes_mem = None
                    if new_cat == "Airline":
                        with col_extra1:
                            new_tier = st.text_input("Tier", key="edit_mem_tier")
                            new_alliance = st.text_input("Alliance", key="edit_mem_alliance")
                        with col_extra2:
                            new_airport = st.text_input("Airport Code", key="edit_mem_airport")
                            new_notes_mem = st.text_area("Notes", key="edit_mem_notes")
                    elif new_cat in ("Hotel", "Car Rental", "Lounge", "Rail", "Ferry", "Ride-Share", "Credit Card"):
                        with col_extra1:
                            new_tier = st.text_input("Status/Tier", key="edit_mem_tier")
                        with col_extra2:
                            new_notes_mem = st.text_area("Notes", key="edit_mem_notes")
                    else:
                        with col_extra1:
                            new_notes_mem = st.text_area("Notes", key="edit_mem_notes")

                if st.button("➕ Add Membership", key="edit_mem_add_btn"):
                    if not new_name_mem or not new_num_mem:
                        st.warning("Program Name and Membership Number required.")
                    else:
                        proceed = True
                        dupes = db.find_duplicate_memberships(
                            exec_id, new_name_mem
                        )
                        if dupes:
                            st.warning(
                                f"⚠️ This executive already has a "
                                f"'{new_name_mem}' membership on file:"
                            )
                            for d in dupes:
                                st.write(
                                    f"- {d['program_name']} · "
                                    f"{d['membership_number']} "
                                    f"({d.get('category') or '—'})"
                                )
                            if not st.checkbox(
                                "Add anyway?", key="force_add_membership"
                            ):
                                proceed = False
                                st.info(
                                    "Tick **Add anyway?** above and click "
                                    "**➕ Add Membership** again to save the "
                                    "additional membership."
                                )
                        if proceed:
                            db.add_membership(
                                exec_id, new_cat.lower(),
                                new_name_mem, new_num_mem,
                                tier=new_tier, alliance=new_alliance,
                                airport_code=new_airport, notes=new_notes_mem,
                            )
                            st.session_state.pop("force_add_membership", None)
                            st.rerun()

        else:
            # Read-only
            profile_data = db.get_full_executive_profile(exec_id)
            if profile_data:
                for key, value in profile_data.items():
                    st.write(f"**{key}:** {value}")
            if not profile.get("is_active", 1):
                st.caption("⚠️ Inactive")

            st.divider()
            st.subheader("🛂 Passports")
            passports = db.get_passports(exec_id)
            if passports:
                for p in passports:
                    col1, col2, col3 = st.columns([2, 2, 2])
                    with col1:
                        st.write(f"{p['country']}: {p['passport_number']}")
                    with col2:
                        st.write(f"Exp: {fmt_date(p.get('expiry_date'))}")
                    with col3:
                        st.write(p.get("notes") or "")
            else:
                st.caption("No passports added.")

            st.subheader("✈️ Memberships")
            mems = db.get_memberships(exec_id)
            if mems:
                for m in mems:
                    col1, col2, col3 = st.columns([2, 2, 2])
                    with col1:
                        emoji = ("✈️" if m["category"] == "airline" else
                                 "🏨" if m["category"] == "hotel" else
                                 "🚗" if m["category"] == "car rental" else
                                 "🛋️" if m["category"] == "lounge" else
                                 "🚄" if m["category"] == "rail" else
                                 "⛴️" if m["category"] == "ferry" else
                                 "🚗" if m["category"] == "ride-share" else
                                 "💳" if m["category"] == "credit card" else "📌")
                        st.write(f"{emoji} {m['program_name']}: {m['membership_number']}")
                    with col2:
                        details = [d for d in [m.get("tier"), m.get("alliance"), m.get("airport_code")] if d]
                        st.write(", ".join(details) if details else "")
                    with col3:
                        st.write(m.get("notes") or "")
            else:
                st.caption("No memberships added.")

            st.divider()
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("✏️ Edit Executive", use_container_width=True):
                    st.session_state["profile_edit_mode"] = True
                    st.rerun()
            with col2:
                if st.button("🗑️ Delete Executive", use_container_width=True):
                    st.session_state["confirm_delete_from_view"] = True
            with col3:
                if st.button("❌ Close Profile", use_container_width=True):
                    st.session_state["show_full_profile"] = False
                    st.session_state["profile_edit_mode"] = False
                    st.session_state["confirm_delete_from_view"] = False
                    st.rerun()

            if st.session_state.get("confirm_delete_from_view", False):
                trip_count = db.get_executive_trip_count(exec_id)
                extra = (f"This executive has {trip_count} trip(s). They will also be deleted."
                         if trip_count > 0 else None)
                render_delete_confirmation(
                    flag_key="confirm_delete_from_view",
                    title=f"Permanently delete executive '{profile['name']}'?",
                    on_confirm=lambda: db.delete_executive(exec_id, force=True),
                    extra_msg=extra,
                )


# Export Profile
if profile:
    with st.sidebar.expander("📤 Export Profile", expanded=False):
        col_csv, col_doc, col_excel = st.columns(3)
        with col_csv:
            if st.button("📊 CSV", key="export_profile_csv_btn"):
                profile_data = db.get_full_executive_profile(exec_id)
                if profile_data:
                    mems = db.get_memberships(exec_id)
                    profile_data["Memberships"] = "; ".join(
                        f"{m['program_name']}: {m['membership_number']}" for m in mems)
                    output = io.StringIO()
                    writer = csv.DictWriter(output, fieldnames=profile_data.keys())
                    writer.writeheader()
                    writer.writerow(profile_data)
                    st.download_button("⬇️ Download", data=output.getvalue(),
                        file_name=f"{profile_data['Name']}_profile.csv",
                        mime="text/csv", key="csv_download_side")
        with col_doc:
            if st.button("📄 Word", key="export_profile_docx_btn"):
                profile_data = db.get_full_executive_profile(exec_id)
                if profile_data:
                    mems = db.get_memberships(exec_id)
                    profile_data["Memberships"] = "; ".join(
                        f"{m['program_name']}: {m['membership_number']}" for m in mems)
                    doc_stream = doc_generator.generate_executive_profile_doc(
                        profile_data, exec_id, get_currency_symbol("USD"))
                    st.download_button("⬇️ Download", data=doc_stream,
                        file_name=f"{profile_data['Name']}_Profile.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml",
                        key="docx_download_side")
        with col_excel:
            if st.button("📊 Excel", key="export_profile_xlsx_btn"):
                profile_data = db.get_full_executive_profile(exec_id)
                if profile_data:
                    excel_stream = export_profile_to_excel(exec_id, get_currency_symbol("USD"))
                    if excel_stream:
                        st.download_button("⬇️ Download", data=excel_stream,
                            file_name=f"{profile_data['Name']}_Profile.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_download_side")


with st.sidebar.expander("📤 Export All Profiles", expanded=False):
    all_profiles = db.get_all_executive_profiles(active_only=False)
    if all_profiles:
        col_aw, col_ac, col_ae = st.columns(3)
        with col_aw:
            if st.button("📄 Word (All)", key="export_all_docx_btn"):
                doc_stream = doc_generator.generate_all_executive_profiles_doc(all_profiles)
                st.download_button("⬇️ Download Word", data=doc_stream,
                    file_name=f"All_Executives_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml",
                    key="all_doc_download")
        with col_ac:
            if st.button("📊 CSV (All)", key="export_all_csv_btn"):
                output = io.StringIO()
                headers = list(all_profiles[0].keys())
                writer = csv.DictWriter(output, fieldnames=headers)
                writer.writeheader()
                for p in all_profiles:
                    writer.writerow(p)
                st.download_button("⬇️ Download CSV", data=output.getvalue().encode("utf-8"),
                    file_name=f"All_Executives_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv", key="all_csv_download")
        with col_ae:
            if st.button("📊 Excel (All)", key="export_all_xlsx_btn"):
                wb = Workbook()
                ws = wb.active
                ws.title = "All Executives"
                headers = list(all_profiles[0].keys())
                for col_idx, header in enumerate(headers, 1):
                    ws.cell(row=1, column=col_idx, value=header)
                for row_idx, p in enumerate(all_profiles, 2):
                    for col_idx, key in enumerate(headers, 1):
                        ws.cell(row=row_idx, column=col_idx, value=p.get(key, ""))
                excel_stream = io.BytesIO()
                wb.save(excel_stream)
                excel_stream.seek(0)
                st.download_button("⬇️ Download Excel", data=excel_stream,
                    file_name=f"All_Executives_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="all_excel_download")
    else:
        st.caption("No executives available to export.")


with st.sidebar.expander("💾 Import / Restore Database"):
    st.caption("Upload a file to restore or merge data.")
    import_mode = st.radio("Import Mode",
        options=["Merge (Add to existing)", "Replace (Full restore)"],
        key="import_mode",
        help="Merge adds new data; Replace overwrites everything.")
    uploaded_file = st.file_uploader("Choose a file", type=["db", "json", "csv"],
        key="import_file",
        help="Supported: .db (replace), .json (merge), .csv (merge executives).")
    if uploaded_file is not None:
        st.info(f"📄 {uploaded_file.name} ({uploaded_file.size / 1024:.1f} KB)")
        if st.button("🚀 Start Import", type="primary", key="start_import_btn"):
            try:
                if import_mode == "Replace (Full restore)":
                    if uploaded_file.name.endswith(".db"):
                        with open(db.DB_PATH, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        st.success("✅ Database replaced successfully! Refreshing...")
                        st.rerun()
                    else:
                        st.warning("Replace mode only accepts .db files.")
                else:
                    if uploaded_file.name.endswith(".json"):
                        data = json.load(uploaded_file)
                        result = db.merge_database_data(data)
                        st.success(result)
                    elif uploaded_file.name.endswith(".csv"):
                        content = uploaded_file.getvalue().decode("utf-8").splitlines()
                        reader = csv.DictReader(content)
                        result = db.import_executives_from_csv(reader)
                        st.success(result)
                    else:
                        st.warning("Merge mode accepts .json or .csv files.")
            except Exception as e:
                st.error(f"Import failed: {e}")


with st.sidebar.expander("📤 Export Database", expanded=False):
    # ---- Backup status + one-click .db download ----
    days = _days_since_backup()
    if days is None:
        st.caption("⚠️ No backup recorded yet.")
    else:
        st.caption(f"Last backup: {days} day(s) ago.")

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        try:
            with open(db.DB_PATH, "rb") as _f:
                _db_bytes = _f.read()
            st.download_button(
                "⬇️ Backup .db",
                data=_db_bytes,
                file_name=(
                    f"ea_command_center_backup_"
                    f"{datetime.now().strftime('%Y%m%d_%H%M')}.db"
                ),
                mime="application/octet-stream",
                key="sidebar_db_backup_download",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Could not read DB: {e}")
    with col_b2:
        if st.button(
            "✅ Mark Done",
            key="sidebar_mark_backed_up",
            use_container_width=True,
        ):
            _mark_backup_done()
            st.rerun()

    st.divider()

    # ---- Existing JSON / CSV (ZIP) exports ----
    st.caption("Export all data as JSON or CSV (ZIP).")
    col_exp_json, col_exp_csv = st.columns(2)
    with col_exp_json:
        if st.button("📄 JSON", use_container_width=True, key="export_db_json_btn"):
            data = db.export_all_data()
            json_str = json.dumps(data, indent=2, default=str)
            st.download_button(
                "⬇️ Download JSON",
                data=json_str,
                file_name=f"database_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                key="export_json",
            )
    with col_exp_csv:
        if st.button("📊 CSV (ZIP)", use_container_width=True, key="export_db_zip_btn"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                tables = [
                    "companies", "executives", "contacts", "trip_delegation",
                    "trips", "itinerary_items", "trip_stops", "trip_templates",
                    "categories", "executive_memberships", "executive_passports",
                    "item_delegation", "item_contacts", "packing_lists",
                    "packing_items", "trip_checklists", "trip_checklist_items",
                    "hospitals", "exchange_rates",
                ]
                for table in tables:
                    rows = db.get_all_rows(table)
                    if rows:
                        csv_buffer = io.StringIO()
                        writer = csv.DictWriter(csv_buffer, fieldnames=rows[0].keys())
                        writer.writeheader()
                        writer.writerows(rows)
                        zipf.writestr(f"{table}.csv", csv_buffer.getvalue())
                    else:
                        conn = sqlite3.connect(db.DB_PATH, timeout=30)
                        try:
                            cur = conn.cursor()
                            cur.execute(f"PRAGMA table_info({table})")
                            cols = [r[1] for r in cur.fetchall()]
                        finally:
                            conn.close()
                        if cols:
                            csv_buffer = io.StringIO()
                            writer = csv.DictWriter(csv_buffer, fieldnames=cols)
                            writer.writeheader()
                            zipf.writestr(f"{table}.csv", csv_buffer.getvalue())
            zip_buffer.seek(0)
            st.download_button(
                "⬇️ Download ZIP",
                data=zip_buffer,
                file_name=f"database_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                key="export_csv_zip",
            )


with st.sidebar.expander("💱 Exchange Rates", expanded=False):
    st.caption("Fetch latest rates and store in the database. "
               "Rates are cached and reused for historical accuracy.")
    total_rates = 0
    last_fetch = None
    conn = None
    try:
        conn = sqlite3.connect(db.DB_PATH, timeout=30)
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM exchange_rates")
        total_rates = c.fetchone()[0]
        c.execute("SELECT MAX(fetched_at) FROM exchange_rates")
        last_fetch = c.fetchone()[0]
    except Exception:
        total_rates = 0
        last_fetch = None
    finally:
        if conn:
            conn.close()

    st.write(f"**{total_rates}** rate row(s) stored.")
    if last_fetch:
        st.caption(f"Last fetched: {fmt_datetime(last_fetch)}")

    if st.button("🔄 Refresh Rates Now", use_container_width=True, key="refresh_rates_btn"):
        with st.spinner("Fetching rates from API..."):
            result = currency.seed_rates_for_date()
        if result.get("error"):
            st.error(f"Failed: {result['error']}")
        else:
            st.success(f"✅ Stored {result['stored']} rate(s) for {result['date']} (skipped {result['skipped']}).")
            st.rerun()


# =========================================================
# TRIP EDIT MODAL (defined before tabs)
# =========================================================


def _clear_dialog_trip_id():
    st.session_state.pop("dialog_trip_id", None)


@st.dialog(
    "✈️ Edit Trip",
    width="large",
    on_dismiss=_clear_dialog_trip_id,
)
def _render_trip_edit_modal(trip_id_modal, country_list):
    trip_modal_data = db.get_trip(trip_id_modal)
    if not trip_modal_data:
        st.warning("Trip data not found.")
        return

    is_locked = trip_modal_data.get("status") in ["approved", "final"]

    if f"modal_stops_{trip_id_modal}" not in st.session_state:
        st.session_state[f"modal_stops_{trip_id_modal}"] = db.get_trip_stops(trip_id_modal)
    if f"modal_items_{trip_id_modal}" not in st.session_state:
        st.session_state[f"modal_items_{trip_id_modal}"] = db.get_items_for_trip(trip_id_modal)

    st.subheader(f"✈️ Edit Trip: {trip_modal_data.get('purpose', 'Untitled')}")
    if is_locked:
        st.info("🔒 This trip is Approved or Final – read-only view.")

    with st.form(key=f"edit_trip_form_{trip_id_modal}"):
        new_purpose = st.text_input("Trip Name", value=trip_modal_data.get("purpose", ""),
            key=f"modal_purpose_{trip_id_modal}", disabled=is_locked)

        # --- Trip dates (parity with the create form) ---
        col_date1, col_date2 = st.columns(2)
        with col_date1:
            try:
                current_start = (
                    datetime.fromisoformat(trip_modal_data["start_date"]).date()
                    if trip_modal_data.get("start_date")
                    else datetime.now().date()
                )
            except Exception:
                current_start = datetime.now().date()
            new_start_date = st.date_input(
                "Start Date*",
                value=current_start,
                key=f"modal_start_date_{trip_id_modal}",
                disabled=is_locked,
            )
        with col_date2:
            try:
                current_end = (
                    datetime.fromisoformat(trip_modal_data["end_date"]).date()
                    if trip_modal_data.get("end_date")
                    else datetime.now().date()
                )
            except Exception:
                current_end = datetime.now().date()
            new_end_date = st.date_input(
                "End Date*",
                value=current_end,
                key=f"modal_end_date_{trip_id_modal}",
                disabled=is_locked,
            )
        if new_start_date and new_end_date and new_end_date < new_start_date:
            st.warning("End date must be on or after start date.")

        col_dep1, col_dep2 = st.columns(2)
        with col_dep1:
            new_dep_city = st.text_input("Departure City", value=trip_modal_data.get("departure_city", ""),
                key=f"modal_dep_city_{trip_id_modal}", disabled=is_locked)
        with col_dep2:
            new_dep_region = st.text_input("Departure Region", value=trip_modal_data.get("departure_region", ""),
                key=f"modal_dep_region_{trip_id_modal}", disabled=is_locked)
        new_dep_country = st.selectbox("Departure Country",
            options=[""] + country_list,
            index=(([""] + country_list).index(trip_modal_data.get("departure_country", ""))
                   if trip_modal_data.get("departure_country") in ([""] + country_list) else 0),
            key=f"modal_dep_country_{trip_id_modal}", disabled=is_locked)

        new_receipt_folder = st.text_input("📥 Receipt Upload Folder URL (optional)",
            value=trip_modal_data.get("receipt_upload_folder_url", "") or "",
            placeholder="https://drive.google.com/drive/folders/…",
            key=f"modal_receipt_folder_{trip_id_modal}", disabled=is_locked,
            help="Shared folder for delegation members to drop receipt photos.")

        st.write("**Budget**")
        col_bud1, col_bud2 = st.columns(2)
        with col_bud1:
            base_cur = trip_modal_data.get("base_currency", "USD")
            budget_base = trip_modal_data.get("budget", 0.0)
            new_budget = st.number_input("Budget Amount (in Base Currency)",
                min_value=0.0, step=100.0, value=float(budget_base),
                key=f"modal_budget_{trip_id_modal}", disabled=is_locked)
        with col_bud2:
            st.write(f"Base Currency: {base_cur}")

        base_currency_options = currency.CURRENCIES_SUPPORTED
        new_base_currency = st.selectbox("Base Currency", options=base_currency_options,
            index=(base_currency_options.index(trip_modal_data.get("base_currency", "USD"))
                   if trip_modal_data.get("base_currency") in base_currency_options else 0),
            key=f"modal_base_currency_{trip_id_modal}", disabled=is_locked)

        current_status = trip_modal_data.get("status", "draft")
        new_status = st.selectbox("Status", options=["draft", "approved", "final"],
            index=(["draft", "approved", "final"].index(current_status)
                   if current_status in ["draft", "approved", "final"] else 0),
            key=f"modal_status_{trip_id_modal}", disabled=is_locked)

        st.write("**Timezone Display**")
        modal_tz_display = st.radio(
            "Show times in:",
            options=["Home", "Destination"],
            index=0,
            key=f"modal_tz_display_radio_{trip_id_modal}",
        )
        # Mirror the radio's value into the key the item list reads from,
        # so display_mode_modal reflects the user's current choice.
        st.session_state[f"modal_tz_display_{trip_id_modal}"] = modal_tz_display

        st.caption(
            "**Home** converts every item to the executive's home timezone. "
            "**Destination** shows each item in its own assigned timezone — "
            "if all your items use the same timezone, the two modes will "
            "look identical."
        )

        st.write("**Contacts**")
        exec_profile_modal = db.get_executive_profile(trip_modal_data["exec_id"])
        company_id_modal = exec_profile_modal.get("company_id") if exec_profile_modal else None
        if company_id_modal:
            contact_options_modal = get_local_support_options(company_id_modal)
            existing_contact_ids = (json.loads(trip_modal_data.get("trip_contacts", "[]"))
                if trip_modal_data.get("trip_contacts") else [])
            existing_labels = [label for label, cid in contact_options_modal.items()
                if cid in existing_contact_ids]
            selected_contacts_modal = st.multiselect("Select local support contacts for this trip",
                options=list(contact_options_modal.keys()), default=existing_labels,
                key=f"modal_contacts_{trip_id_modal}", disabled=is_locked)
            selected_contact_ids_modal = [contact_options_modal[label] for label in selected_contacts_modal]
        else:
            selected_contact_ids_modal = []

        st.write("**Travelers**")
        if company_id_modal:
            all_company_contacts = db.get_contacts_for_delegation(company_id_modal)
            traveler_options = {}
            for c in all_company_contacts:
                role = c.get("role", "") or ""
                suffix = " · freelance" if not c.get("company_id") else ""
                label = f"{c['name']} ({role}){suffix}".strip()
                traveler_options[label] = c["id"]

            existing_members = db.get_trip_delegation_members(trip_id_modal)
            existing_ids = [m["id"] for m in existing_members]
            existing_labels = [
                label for label, cid in traveler_options.items() if cid in existing_ids
            ]
            if not existing_labels:
                default_ids = set(db.get_company_default_contacts(company_id_modal))
                existing_labels = [
                    label
                    for label, cid in traveler_options.items()
                    if cid in default_ids
                ]

            selected_travelers = st.multiselect(
                "Travelers on this trip",
                options=list(traveler_options.keys()),
                default=existing_labels,
                key=f"modal_travelers_{trip_id_modal}",
                disabled=is_locked,
            )
            selected_traveler_ids = [traveler_options[l] for l in selected_travelers]
        else:
            selected_travelers = []
            selected_traveler_ids = []

        # Summary caption — reads AFTER the picker has run
        if selected_traveler_ids:
            preview = ", ".join(selected_travelers[:3])
            suffix = " ..." if len(selected_travelers) > 3 else ""
            st.caption(f"{len(selected_traveler_ids)} traveler(s): {preview}{suffix}")
        else:
            st.caption("No travelers selected yet.")

        if not is_locked:
            submitted = st.form_submit_button("💾 Save Changes")
            if submitted:
                db.update_trip_purpose(trip_id_modal, new_purpose)
                # Persist date changes; reuse the existing destination since
                # update_trip_dates() also writes to the destination column.
                if new_start_date and new_end_date:
                    db.update_trip_dates(
                        trip_id_modal,
                        new_start_date.isoformat(),
                        new_end_date.isoformat(),
                        trip_modal_data.get("destination", ""),
                    )
                db.update_trip_budget(trip_id_modal, new_budget)
                db.update_trip_departure_details(
                    trip_id_modal, new_dep_city, new_dep_region, new_dep_country
                )
                db.update_trip_base_currency(trip_id_modal, new_base_currency)
                db.update_trip_display_currency(trip_id_modal, new_base_currency)
                db.update_trip_status(trip_id_modal, new_status)
                db.update_trip_contacts(trip_id_modal, selected_contact_ids_modal)
                db.update_trip_receipt_folder(trip_id_modal, new_receipt_folder)
                db.set_trip_delegation(trip_id_modal, selected_traveler_ids)
                st.success("✅ Trip header updated.")
                st.rerun()

    # Weather — one expander per stop, date-aligned to the stop's range
    st.write("**🌦️ Weather**")
    trip_stops_weather = st.session_state.get(f"modal_stops_{trip_id_modal}", [])
    if not trip_stops_weather:
        trip_stops_weather = db.get_trip_stops(trip_id_modal)
    if not trip_stops_weather:
        st.caption("Add at least one stop to see weather.")
    else:
        for s_idx, stop in enumerate(trip_stops_weather):
            stop_city = (stop.get("city") or "").strip()
            stop_start = stop.get("start_date")
            stop_end = stop.get("end_date")
            if not (stop_city and stop_start and stop_end):
                continue

            header = (
                f"🌦️ {stop_city} · " f"{fmt_date(stop_start)} → {fmt_date(stop_end)}"
            )
            cache_key = f"weather_modal_{trip_id_modal}_{s_idx}"

            with st.expander(header, expanded=False):
                if st.button(
                    "Fetch weather",
                    key=f"fetch_weather_modal_{trip_id_modal}_{s_idx}",
                    help=f"Fetch weather for {stop_city} across the trip dates",
                ):
                    with st.spinner(f"Fetching weather for {stop_city}..."):
                        w = weather.get_weather_for_range(
                            stop_city, stop_start, stop_end
                        )
                    st.session_state[cache_key] = w
                    st.rerun()

                # Three distinct states: never fetched, fetched-empty,
                # fetched-with-data. Showing different messages for each
                # avoids the "did the button do anything?" confusion.
                if cache_key not in st.session_state:
                    st.caption(
                        "Click **Fetch weather** to load the forecast "
                        "for these dates."
                    )
                    continue

                cached = st.session_state.get(cache_key)
                if not cached:
                    st.info(
                        f"No weather data available for {stop_city} on "
                        f"{fmt_date(stop_start)} → {fmt_date(stop_end)}. "
                        "Forecasts typically cover up to 16 days ahead."
                    )
                    continue

                if cached.get("is_historical_fallback"):
                    st.warning(
                        "📅 Forecast not yet available for these dates. "
                        "Showing typical conditions from the same dates "
                        f"last year ({fmt_date(cached.get('start_date'))} → "
                        f"{fmt_date(cached.get('end_date'))})."
                    )
                else:
                    st.caption(
                        f"📍 {cached['location']} · "
                        f"fetched {fmt_datetime(cached['fetched_at'])}"
                    )

                cur = cached.get("current")
                if cur:
                    col_w1, col_w2, col_w3, col_w4 = st.columns(4)
                    with col_w1:
                        st.metric("Now", f"{cur['temp']}°C", cur["icon"])
                    with col_w2:
                        st.metric("Humidity", f"{cur['humidity']}%")
                    with col_w3:
                        st.metric("Wind", f"{cur['wind']} km/h")
                    with col_w4:
                        st.caption(cur["desc"])

                daily = cached.get("daily") or []
                if daily:
                    for d in daily:
                        col_d1, col_d2, col_d3, col_d4 = st.columns([2, 1, 1, 1])
                        with col_d1:
                            st.write(f"{d['icon']} {fmt_date(d['date'])}")
                        with col_d2:
                            st.write(f"**{d['max']}°** / {d['min']}°")
                        with col_d3:
                            st.write(f"☔ {d['precip']} mm")
                        with col_d4:
                            st.write(f"UV {d['uv']}")
                else:
                    st.caption("No daily records in this range.")
    # Stops
    st.write("**📍 Stops**")
    stops = st.session_state[f"modal_stops_{trip_id_modal}"]
    for idx, stop in enumerate(stops):
        (
            col_s1, col_s2, col_s3, col_s4,
            col_up, col_dn, col_s5, col_s6,
        ) = st.columns([2, 2, 2, 2, 0.4, 0.4, 0.5, 0.5])
        with col_s1:
            st.write(f"**{idx+1}.** {stop['city']}")
        with col_s2:
            loc_parts = [p for p in [stop.get("region"), stop.get("country")] if p]
            st.write(", ".join(loc_parts) if loc_parts else "")
        with col_s3:
            st.write(f"{fmt_date(stop['start_date'])} → {fmt_date(stop['end_date'])}")
        with col_s4:
            st.write(stop.get("notes", "")[:30])
        with col_up:
            if not is_locked and idx > 0:
                if st.button("↑", key=f"modal_stop_up_{trip_id_modal}_{idx}",
                             help="Move this stop up"):
                    # Swap in the in-memory list, then persist
                    stops[idx - 1], stops[idx] = stops[idx], stops[idx - 1]
                    st.session_state[f"modal_stops_{trip_id_modal}"] = stops
                    ordered_ids = [
                        s.get("id") for s in stops if s.get("id")
                    ]
                    db.reorder_trip_stops(trip_id_modal, ordered_ids)
                    st.rerun()
        with col_dn:
            if not is_locked and idx < len(stops) - 1:
                if st.button("↓", key=f"modal_stop_dn_{trip_id_modal}_{idx}",
                             help="Move this stop down"):
                    stops[idx + 1], stops[idx] = stops[idx], stops[idx + 1]
                    st.session_state[f"modal_stops_{trip_id_modal}"] = stops
                    ordered_ids = [
                        s.get("id") for s in stops if s.get("id")
                    ]
                    db.reorder_trip_stops(trip_id_modal, ordered_ids)
                    st.rerun()
        with col_s5:
            if not is_locked:
                if st.button("✏️", key=f"modal_edit_stop_{trip_id_modal}_{idx}"):
                    st.session_state[f"modal_editing_stop_{trip_id_modal}_{idx}"] = True
        with col_s6:
            if not is_locked:
                if st.button("🗑️", key=f"modal_del_stop_{trip_id_modal}_{idx}"):
                    stop_id_to_delete = stop.get("id")
                    if stop_id_to_delete:
                        db.delete_trip_stop(stop_id_to_delete)
                    st.session_state[f"modal_stops_{trip_id_modal}"] = (
                        db.get_trip_stops(trip_id_modal)
                    )
                    st.rerun()

        # --- Edit Stop expander ---
        if (
            st.session_state.get(f"modal_editing_stop_{trip_id_modal}_{idx}", False)
            and not is_locked
        ):
            with st.expander(f"Edit Stop: {stop['city']}", expanded=True):
                with st.form(key=f"modal_edit_stop_form_{trip_id_modal}_{idx}"):
                    es_city = st.text_input(
                        "City*",
                        value=stop["city"],
                        key=f"modal_e_stop_city_{trip_id_modal}_{idx}",
                    )
                    es_country = st.selectbox(
                        "Country (optional)",
                        options=[""] + country_list,
                        index=(
                            ([""] + country_list).index(stop.get("country", ""))
                            if stop.get("country", "") in country_list
                            else 0
                        ),
                        key=f"modal_e_stop_country_{trip_id_modal}_{idx}",
                    )
                    es_region = st.text_input(
                        "Region / State",
                        value=stop.get("region", ""),
                        key=f"modal_e_stop_region_{trip_id_modal}_{idx}",
                    )
                    es_notes = st.text_input(
                        "Notes",
                        value=stop.get("notes", ""),
                        key=f"modal_e_stop_notes_{trip_id_modal}_{idx}",
                    )
                    col_a, col_b = st.columns(2)
                    with col_a:
                        try:
                            default_start = datetime.fromisoformat(
                                stop["start_date"]
                            ).date()
                        except Exception:
                            default_start = datetime.now().date()
                        es_start = st.date_input(
                            "Start Date*",
                            value=default_start,
                            key=f"modal_e_stop_start_{trip_id_modal}_{idx}",
                        )
                    with col_b:
                        try:
                            default_end = datetime.fromisoformat(
                                stop["end_date"]
                            ).date()
                        except Exception:
                            default_end = datetime.now().date()
                        es_end = st.date_input(
                            "End Date*",
                            value=default_end,
                            key=f"modal_e_stop_end_{trip_id_modal}_{idx}",
                        )
                    sub_save, sub_cancel = st.columns(2)
                    with sub_save:
                        if st.form_submit_button("💾 Update Stop"):
                            if not (es_city and es_start and es_end):
                                st.warning(
                                    "City, Start Date, and End Date are required."
                                )
                            else:
                                es_start_iso = es_start.isoformat()
                                es_end_iso = es_end.isoformat()
                                same_city, other_city = _find_stop_conflicts(
                                    st.session_state[f"modal_stops_{trip_id_modal}"],
                                    es_city, es_start_iso, es_end_iso,
                                    exclude_idx=idx,
                                )

                                proceed = True

                                if other_city:
                                    st.error(
                                        "⚠️ These dates overlap with stops "
                                        "in a **different city** — you "
                                        "can't be in two places at once:"
                                    )
                                    for c in other_city:
                                        st.write(
                                            f"- {c['city']} · "
                                            f"{fmt_date(c['start_date'])} → "
                                            f"{fmt_date(c['end_date'])}"
                                        )
                                    if not st.checkbox(
                                        "Update anyway? (overlaps across cities)",
                                        key=(
                                            f"force_edit_stop_other_city_"
                                            f"{trip_id_modal}_{idx}"
                                        ),
                                    ):
                                        proceed = False
                                        st.info(
                                            "Tick **Update anyway?** above "
                                            "and click **💾 Update Stop** "
                                            "again — only do this if the "
                                            "overlap is intentional."
                                        )

                                elif same_city:
                                    st.warning(
                                        f"⚠️ Another stop on this trip is "
                                        f"already in **{es_city}** with "
                                        "overlapping dates:"
                                    )
                                    for c in same_city:
                                        st.write(
                                            f"- {c['city']} · "
                                            f"{fmt_date(c['start_date'])} → "
                                            f"{fmt_date(c['end_date'])}"
                                        )
                                    if not st.checkbox(
                                        "Update anyway?",
                                        key=(
                                            f"force_edit_stop_"
                                            f"{trip_id_modal}_{idx}"
                                        ),
                                    ):
                                        proceed = False
                                        st.info(
                                            "Tick **Update anyway?** above "
                                            "and click **💾 Update Stop** "
                                            "again to save."
                                        )

                                if proceed:
                                    stop_id = stop.get("id")
                                    if stop_id:
                                        db.update_trip_stop(
                                            stop_id,
                                            es_city,
                                            es_country,
                                            es_region,
                                            es_start_iso,
                                            es_end_iso,
                                            es_notes,
                                        )
                                    # Refresh session state from the DB so
                                    # the in-memory list matches reality.
                                    st.session_state[
                                        f"modal_stops_{trip_id_modal}"
                                    ] = db.get_trip_stops(trip_id_modal)
                                    st.session_state.pop(
                                        f"force_edit_stop_{trip_id_modal}_{idx}",
                                        None,
                                    )
                                    st.session_state.pop(
                                        f"force_edit_stop_other_city_"
                                        f"{trip_id_modal}_{idx}",
                                        None,
                                    )
                                    st.session_state[
                                        f"modal_editing_stop_{trip_id_modal}_{idx}"
                                    ] = False
                                    st.rerun()
                    with sub_cancel:
                        if st.form_submit_button("❌ Cancel"):
                            st.session_state[
                                f"modal_editing_stop_{trip_id_modal}_{idx}"
                            ] = False
                            st.rerun()

    with st.expander("➕ Add Stop"):
        col_sc1, col_sc2 = st.columns(2)
        with col_sc1:
            new_stop_city = st.text_input("City*", key=f"modal_new_stop_city_{trip_id_modal}", disabled=is_locked)
        with col_sc2:
            new_stop_country = st.selectbox("Country", options=[""] + country_list,
                key=f"modal_new_stop_country_{trip_id_modal}", disabled=is_locked)
        col_sr, col_sn = st.columns(2)
        with col_sr:
            new_stop_region = st.text_input("Region", key=f"modal_new_stop_region_{trip_id_modal}", disabled=is_locked)
        with col_sn:
            new_stop_notes = st.text_input("Notes", key=f"modal_new_stop_notes_{trip_id_modal}", disabled=is_locked)
        col_ss, col_se = st.columns(2)
        with col_ss:
            new_stop_start = st.date_input("Start Date*", value=datetime.now(),
                key=f"modal_new_stop_start_{trip_id_modal}", disabled=is_locked)
        with col_se:
            new_stop_end = st.date_input("End Date*", value=datetime.now(),
                key=f"modal_new_stop_end_{trip_id_modal}", disabled=is_locked)
        if not is_locked:
            if st.button("➕ Add Stop", key=f"modal_add_stop_{trip_id_modal}"):
                if not (new_stop_city and new_stop_start and new_stop_end):
                    st.warning("City, Start Date, and End Date are required.")
                else:
                    new_start_iso = new_stop_start.isoformat()
                    new_end_iso = new_stop_end.isoformat()
                    same_city, other_city = _find_stop_conflicts(
                        st.session_state[f"modal_stops_{trip_id_modal}"],
                        new_stop_city, new_start_iso, new_end_iso,
                    )

                    proceed = True

                    if other_city:
                        st.error(
                            "⚠️ These dates overlap with stops in a "
                            "**different city** — you can't be in two "
                            "places at once:"
                        )
                        for c in other_city:
                            st.write(
                                f"- {c['city']} · "
                                f"{fmt_date(c['start_date'])} → "
                                f"{fmt_date(c['end_date'])}"
                            )
                        if not st.checkbox(
                            "Add anyway? (dates overlap across cities)",
                            key=f"force_add_modal_stop_other_city_{trip_id_modal}",
                        ):
                            proceed = False
                            st.info(
                                "Tick **Add anyway?** above and click "
                                "**➕ Add Stop** again — only do this if "
                                "the overlap is intentional."
                            )

                    elif same_city:
                        st.warning(
                            f"⚠️ This trip already has a stop in "
                            f"**{new_stop_city}** with overlapping dates:"
                        )
                        for c in same_city:
                            st.write(
                                f"- {c['city']} · "
                                f"{fmt_date(c['start_date'])} → "
                                f"{fmt_date(c['end_date'])}"
                            )
                        if not st.checkbox(
                            "Add anyway?",
                            key=f"force_add_modal_stop_{trip_id_modal}",
                        ):
                            proceed = False
                            st.info(
                                "Tick **Add anyway?** above and click "
                                "**➕ Add Stop** again — this is normal "
                                "for a city you revisit on a separate "
                                "leg of the trip."
                            )

                    if proceed:
                        next_order = (
                            len(st.session_state[f"modal_stops_{trip_id_modal}"]) + 1
                        )
                        db.add_trip_stop(
                            trip_id_modal,
                            next_order,
                            new_stop_city,
                            new_stop_country,
                            new_stop_region,
                            new_start_iso,
                            new_end_iso,
                            new_stop_notes,
                        )
                        st.session_state[f"modal_stops_{trip_id_modal}"] = (
                            db.get_trip_stops(trip_id_modal)
                        )
                        st.session_state.pop(
                            f"force_add_modal_stop_{trip_id_modal}", None
                        )
                        st.session_state.pop(
                            f"force_add_modal_stop_other_city_{trip_id_modal}", None
                        )
                        st.success("✅ Stop added.")
                        st.rerun()
    # Items
    st.write("**📋 Itinerary Items**")
    base_currency_options = currency.CURRENCIES_SUPPORTED
    items = st.session_state[f"modal_items_{trip_id_modal}"]

    # ---- Bulk-assign trip travelers to all items ----
    if not is_locked and items:
        trip_travelers = db.get_trip_delegation_members(trip_id_modal)
        if trip_travelers:
            if st.button(
                f"➕ Apply {len(trip_travelers)} traveler(s) to all "
                f"{len(items)} item(s)",
                key=f"apply_travelers_all_{trip_id_modal}",
                help="Assign every traveler on this trip to every item. "
                "Useful when the whole delegation attends everything.",
            ):
                traveler_ids = [m["id"] for m in trip_travelers]
                for it in items:
                    if it.get("id"):
                        db.set_item_delegation_members(it["id"], traveler_ids)
                st.session_state[f"modal_items_{trip_id_modal}"] = (
                    db.get_items_for_trip(trip_id_modal)
                )
                st.success(
                    f"Assigned {len(traveler_ids)} traveler(s) to "
                    f"{len(items)} item(s)."
                )
                st.rerun()
        else:
            st.caption(
                "Add travelers in the trip header above to enable " "bulk assignment."
            )

    sel_key = f"selected_item_ids_{trip_id_modal}"
    if sel_key not in st.session_state:
        st.session_state[sel_key] = set()

    if items and not is_locked:
        col_selall, col_delsel, _spacer = st.columns([1.2, 1.6, 4])
        with col_selall:
            all_ids = [it["id"] for it in items if it.get("id")]
            all_selected = bool(all_ids) and all(i in st.session_state[sel_key] for i in all_ids)
            toggle = st.checkbox("Select all", value=all_selected,
                                 key=f"select_all_items_{trip_id_modal}")
            if toggle != all_selected:
                st.session_state[sel_key] = set(all_ids) if toggle else set()
                st.rerun()
        with col_delsel:
            n_sel = len(st.session_state[sel_key])
            if n_sel:
                if st.button(f"🗑️ Delete {n_sel} selected", type="primary",
                             key=f"bulk_del_items_{trip_id_modal}"):
                    st.session_state[
                        f"confirm_bulk_del_items_{trip_id_modal}"
                    ] = True
                    st.rerun()
            else:
                st.caption("Tick items to enable bulk delete.")

        if st.session_state.get(f"confirm_bulk_del_items_{trip_id_modal}", False):
            def _bulk_del_items():
                for iid in list(st.session_state[sel_key]):
                    db.delete_itinerary_item(iid)
                st.session_state[sel_key] = set()
                st.session_state.pop(f"select_all_items_{trip_id_modal}", None)
                st.session_state[f"modal_items_{trip_id_modal}"] = (
                    db.get_items_for_trip(trip_id_modal)
                )

            render_delete_confirmation(
                flag_key=f"confirm_bulk_del_items_{trip_id_modal}",
                title=(
                    f"Permanently delete "
                    f"{len(st.session_state[sel_key])} selected item(s)?"
                ),
                on_confirm=_bulk_del_items,
            )

    exec_tz_modal = (exec_profile_modal.get("timezone", "America/New_York")
                     if exec_profile_modal else "America/New_York")
    display_mode_modal = st.session_state.get(f"modal_tz_display_{trip_id_modal}", "Home")

    for idx, item in enumerate(items):
        dt_display = format_item_datetime(item, exec_tz_modal, display_mode_modal)
        col_cb, col_i1, col_i2, col_up, col_dn, col_i3, col_i4, col_i5 = st.columns(
            [0.4, 3, 2, 0.4, 0.4, 1, 1, 1]
        )
        with col_cb:
            if not is_locked and item.get("id"):
                is_checked = item["id"] in st.session_state[sel_key]
                new_state = st.checkbox("", value=is_checked,
                                        key=f"modal_sel_item_{trip_id_modal}_{item['id']}")
                if new_state != is_checked:
                    if new_state:
                        st.session_state[sel_key].add(item["id"])
                    else:
                        st.session_state[sel_key].discard(item["id"])
                    # Reset the Select-All checkbox so its stored True
                    # doesn't re-select what the user just unticked.
                    st.session_state.pop(f"select_all_items_{trip_id_modal}", None)
        with col_i1:
            st.write(f"{item['description']} ({item['item_type']})")
            st.caption(f"🕐 {dt_display}")
            if item.get("delegation_ids"):
                d_names = get_contact_names(item["delegation_ids"])
                if d_names:
                    st.caption(f"👥 {d_names}")
            if item.get("contact_ids"):
                c_names = get_contact_names(item["contact_ids"])
                if c_names:
                    st.caption(f"📞 {c_names}")
        with col_i2:
            st.write(fmt_money(item.get("cost", 0), item.get("cost_currency", "USD")))
        with col_up:
            if not is_locked and idx > 0:
                if st.button("↑", key=f"modal_item_up_{trip_id_modal}_{idx}",
                             help="Swap times with the item above"):
                    prev = items[idx - 1]
                    cur_start, cur_end = item["datetime_start"], item["datetime_end"]
                    prev_start, prev_end = prev["datetime_start"], prev["datetime_end"]
                    if item.get("id") and prev.get("id"):
                        db.update_itinerary_item(item["id"], item["item_type"], item["description"],
                            prev_start, prev_end, item.get("location", ""), item.get("cost", 0),
                            item.get("confirmation_code", ""), item.get("notes", ""),
                            item.get("is_confirmed", 0), item.get("cost_currency", "USD"),
                            timezone=item.get("timezone"), venue_id=item.get("venue_id"),
                            cost_date=item.get("cost_date"))
                        db.update_itinerary_item(prev["id"], prev["item_type"], prev["description"],
                            cur_start, cur_end, prev.get("location", ""), prev.get("cost", 0),
                            prev.get("confirmation_code", ""), prev.get("notes", ""),
                            prev.get("is_confirmed", 0), prev.get("cost_currency", "USD"),
                            timezone=prev.get("timezone"), venue_id=prev.get("venue_id"),
                            cost_date=prev.get("cost_date"))
                    st.session_state[f"modal_items_{trip_id_modal}"] = db.get_items_for_trip(trip_id_modal)
                    st.rerun()
        with col_dn:
            if not is_locked and idx < len(items) - 1:
                if st.button("↓", key=f"modal_item_dn_{trip_id_modal}_{idx}",
                             help="Swap times with the item below"):
                    nxt = items[idx + 1]
                    cur_start, cur_end = item["datetime_start"], item["datetime_end"]
                    nxt_start, nxt_end = nxt["datetime_start"], nxt["datetime_end"]
                    if item.get("id") and nxt.get("id"):
                        db.update_itinerary_item(item["id"], item["item_type"], item["description"],
                            nxt_start, nxt_end, item.get("location", ""), item.get("cost", 0),
                            item.get("confirmation_code", ""), item.get("notes", ""),
                            item.get("is_confirmed", 0), item.get("cost_currency", "USD"),
                            timezone=item.get("timezone"), venue_id=item.get("venue_id"),
                            cost_date=item.get("cost_date"))
                        db.update_itinerary_item(nxt["id"], nxt["item_type"], nxt["description"],
                            cur_start, cur_end, nxt.get("location", ""), nxt.get("cost", 0),
                            nxt.get("confirmation_code", ""), nxt.get("notes", ""),
                            nxt.get("is_confirmed", 0), nxt.get("cost_currency", "USD"),
                            timezone=nxt.get("timezone"), venue_id=nxt.get("venue_id"),
                            cost_date=nxt.get("cost_date"))
                    st.session_state[f"modal_items_{trip_id_modal}"] = db.get_items_for_trip(trip_id_modal)
                    st.rerun()
        with col_i3:
            if not is_locked:
                if st.button("✏️", key=f"modal_edit_item_{trip_id_modal}_{idx}",
                             help="Edit this item"):
                    st.session_state[f"modal_editing_item_{trip_id_modal}_{idx}"] = True
        with col_i4:
            if not is_locked:
                if st.button("📄", key=f"modal_dup_item_{trip_id_modal}_{idx}",
                             help="Duplicate this item (+1 day)"):
                    # full Phase 1 duplicate body
                    try:
                        orig_start = datetime.fromisoformat(item["datetime_start"])
                        orig_end = datetime.fromisoformat(item["datetime_end"]) if item.get("datetime_end") else None
                    except Exception:
                        orig_start, orig_end = datetime.now(), None
                    cost_date_val = item.get("cost_date")
                    if cost_date_val:
                        try:
                            cost_date_val = (datetime.fromisoformat(cost_date_val).date() + timedelta(days=1)).isoformat()
                        except Exception:
                            cost_date_val = None
                    new_item_id = db.add_itinerary_item(
                        trip_id_modal, item["item_type"],
                        (item.get("description", "") + " (copy)").strip(),
                        (orig_start + timedelta(days=1)).isoformat(),
                        ((orig_end + timedelta(days=1)).isoformat() if orig_end else None),
                        item.get("location", ""), item.get("cost", 0),
                        item.get("confirmation_code", ""), item.get("notes", ""),
                        item.get("is_confirmed", 0), item.get("cost_currency", "USD"),
                        timezone=item.get("timezone"), venue_id=item.get("venue_id"),
                        cost_date=cost_date_val)
                    if item.get("delegation_ids"):
                        db.set_item_delegation_members(new_item_id, item["delegation_ids"])
                    if item.get("contact_ids"):
                        db.set_item_contacts(new_item_id, item["contact_ids"])
                    st.session_state[f"modal_items_{trip_id_modal}"] = db.get_items_for_trip(trip_id_modal)
                    st.rerun()
        with col_i5:
            if not is_locked:
                if st.button("🗑️", key=f"modal_del_item_{trip_id_modal}_{idx}",
                             help="Delete this item"):
                    st.session_state[
                        f"confirm_del_item_{trip_id_modal}_{item['id']}"
                    ] = True
                    st.rerun()

        # --- Delete confirmation (renders below the row) ---
        if (
            not is_locked
            and item.get("id")
            and st.session_state.get(
                f"confirm_del_item_{trip_id_modal}_{item['id']}", False
            )
        ):
            def _del_item(iid=item["id"]):
                db.delete_itinerary_item(iid)
                st.session_state[sel_key].discard(iid)
                st.session_state[f"modal_items_{trip_id_modal}"] = (
                    db.get_items_for_trip(trip_id_modal)
                )

            render_delete_confirmation(
                flag_key=f"confirm_del_item_{trip_id_modal}_{item['id']}",
                title=f"Permanently delete item '{item['description']}'?",
                on_confirm=_del_item,
            )

        if (
            st.session_state.get(f"modal_editing_item_{trip_id_modal}_{idx}", False)
            and not is_locked
        ):
            with st.expander(f"Edit Item: {item['description']}", expanded=True):

                # ─────────────────────────────────────────────────────
                # Inline venue add — sibling form, not nested
                # ─────────────────────────────────────────────────────
                with st.expander(
                    "➕ Need a new venue? Add it here first", expanded=False
                ):
                    with st.form(key=f"modal_inline_add_venue_{trip_id_modal}_{idx}"):
                        col_va, col_vb = st.columns(2)
                        with col_va:
                            mv_name = st.text_input(
                                "Venue Name*",
                                key=f"modal_v_name_{trip_id_modal}_{idx}",
                            )
                            mv_address = st.text_input(
                                "Address",
                                key=f"modal_v_address_{trip_id_modal}_{idx}",
                            )
                            mv_city = st.text_input(
                                "City",
                                key=f"modal_v_city_{trip_id_modal}_{idx}",
                            )
                        with col_vb:
                            mv_country = st.selectbox(
                                "Country",
                                options=[""] + country_list,
                                key=f"modal_v_country_{trip_id_modal}_{idx}",
                            )
                            mv_wifi_ssid = st.text_input(
                                "WiFi SSID",
                                key=f"modal_v_wifi_ssid_{trip_id_modal}_{idx}",
                            )
                            mv_wifi_pw = st.text_input(
                                "WiFi Password",
                                key=f"modal_v_wifi_pw_{trip_id_modal}_{idx}",
                            )
                        mv_notes = st.text_area(
                            "Notes",
                            key=f"modal_v_notes_{trip_id_modal}_{idx}",
                            height=80,
                        )
                        if st.form_submit_button("➕ Add venue"):
                            if not mv_name:
                                st.warning("Venue Name is required.")
                            else:
                                db.add_venue(
                                    name=mv_name,
                                    address=mv_address,
                                    city=mv_city,
                                    country=mv_country,
                                    wifi_ssid=mv_wifi_ssid,
                                    wifi_password=mv_wifi_pw,
                                    notes=mv_notes,
                                )
                                st.success(f"✅ Venue '{mv_name}' added.")
                                st.rerun()

                # ─────────────────────────────────────────────────────
                # Item edit form
                # ─────────────────────────────────────────────────────
                with st.form(key=f"edit_item_form_{trip_id_modal}_{idx}"):
                    e_type = st.selectbox(
                        "Type",
                        options=(
                            [cat["name"] for cat in db.get_all_categories()]
                            if db.get_all_categories()
                            else ["Flight", "Hotel", "Meeting", "Transport"]
                        ),
                        index=0,
                        key=f"modal_e_type_{trip_id_modal}_{idx}",
                    )
                    e_desc = st.text_input(
                        "Description",
                        value=item["description"],
                        key=f"modal_e_desc_{trip_id_modal}_{idx}",
                    )
                    e_start = st.datetime_input(
                        "Start",
                        value=datetime.fromisoformat(item["datetime_start"]),
                        key=f"modal_e_start_{trip_id_modal}_{idx}",
                    )
                    e_end = st.datetime_input(
                        "End",
                        value=(
                            datetime.fromisoformat(item["datetime_end"])
                            if item["datetime_end"]
                            else datetime.now()
                        ),
                        key=f"modal_e_end_{trip_id_modal}_{idx}",
                    )
                    e_loc = st.text_input(
                        "Location",
                        value=item.get("location", ""),
                        key=f"modal_e_loc_{trip_id_modal}_{idx}",
                    )
                    e_cost = st.number_input(
                        "Cost",
                        value=float(item.get("cost", 0)),
                        key=f"modal_e_cost_{trip_id_modal}_{idx}",
                    )
                    e_currency = st.selectbox(
                        "Currency",
                        options=base_currency_options,
                        index=(
                            base_currency_options.index(
                                item.get("cost_currency", "USD")
                            )
                            if item.get("cost_currency", "USD") in base_currency_options
                            else 0
                        ),
                        key=f"modal_e_currency_{trip_id_modal}_{idx}",
                    )

                    # Cost date — falls back to the item's start date if unset
                    default_cost_date_modal = item.get("cost_date")
                    if default_cost_date_modal:
                        try:
                            default_cost_date_modal = datetime.fromisoformat(
                                default_cost_date_modal
                            ).date()
                        except Exception:
                            default_cost_date_modal = datetime.fromisoformat(
                                item["datetime_start"]
                            ).date()
                    else:
                        default_cost_date_modal = datetime.fromisoformat(
                            item["datetime_start"]
                        ).date()
                    e_cost_date_modal = st.date_input(
                        "Cost Date",
                        value=default_cost_date_modal,
                        key=f"modal_e_cost_date_{trip_id_modal}_{idx}",
                        help="Date the cost was incurred — used for accurate "
                        "currency conversion.",
                    )

                    # Timezone
                    tz_display_names, tz_map = get_timezone_dropdown_options()
                    current_tz_modal = item.get("timezone") or exec_tz_modal
                    current_tz_display_modal = next(
                        (n for n in tz_display_names if current_tz_modal in n),
                        tz_display_names[0],
                    )
                    e_timezone_modal = st.selectbox(
                        "Time Zone",
                        options=tz_display_names,
                        index=tz_display_names.index(current_tz_display_modal),
                        key=f"modal_e_timezone_{trip_id_modal}_{idx}",
                    )
                    e_timezone_value_modal = tz_map[e_timezone_modal]

                    e_confirmed = st.checkbox(
                        "Confirmed",
                        value=bool(item.get("is_confirmed", 0)),
                        key=f"modal_e_confirmed_{trip_id_modal}_{idx}",
                    )
                    e_notes = st.text_area(
                        "Notes",
                        value=item.get("notes", ""),
                        key=f"modal_e_notes_{trip_id_modal}_{idx}",
                    )

                    # Venue dropdown — only for session-type items
                    session_types = [
                        "Meeting",
                        "Conference",
                        "Dinner",
                        "Site Visit",
                        "Tour",
                        "Activity",
                    ]
                    e_venue_id_modal = item.get("venue_id")
                    if e_type in session_types:
                        venue_options_modal = db.get_venues(active_only=True)
                        venue_labels_modal = {
                            f"{v['name']}"
                            + (f" — {v['city']}" if v.get("city") else ""): v["id"]
                            for v in venue_options_modal
                        }
                        venue_labels_modal["(No venue)"] = None
                        current_venue_label_modal = "(No venue)"
                        for lbl, vid in venue_labels_modal.items():
                            if vid == e_venue_id_modal:
                                current_venue_label_modal = lbl
                                break
                        selected_venue_label_modal = st.selectbox(
                            "Venue (optional)",
                            options=list(venue_labels_modal.keys()),
                            index=list(venue_labels_modal.keys()).index(
                                current_venue_label_modal
                            ),
                            key=f"modal_e_venue_{trip_id_modal}_{idx}",
                        )
                        e_venue_id_modal = venue_labels_modal[
                            selected_venue_label_modal
                        ]

                    # Delegation members
                    delegation_options_modal = {}
                    if company_id_modal:
                        delegation_options_modal = get_trip_delegation_options(
                            trip_id_modal
                        )
                    if delegation_options_modal:
                        current_delegation_ids = item.get("delegation_ids", [])
                        current_labels = [
                            label
                            for label, did in delegation_options_modal.items()
                            if did in current_delegation_ids
                        ]
                        selected_delegation_modal = st.multiselect(
                            "Assign Delegation Members",
                            options=list(delegation_options_modal.keys()),
                            default=current_labels,
                            key=f"modal_item_delegation_{trip_id_modal}_{idx}",
                        )
                        selected_delegation_ids_modal = [
                            delegation_options_modal[l]
                            for l in selected_delegation_modal
                        ]
                    else:
                        st.info(
                            "No travelers on this trip yet. Add them in the "
                            "**Travelers** section of the trip header first, "
                            "then come back here to assign them to this item."
                        )
                        selected_delegation_ids_modal = []

                    # Local support contacts
                    if company_id_modal:
                        contact_options_modal = get_company_contact_options(
                            company_id_modal
                        )
                        current_contact_ids = item.get("contact_ids", [])
                        current_labels = [
                            label
                            for label, cid in contact_options_modal.items()
                            if cid in current_contact_ids
                        ]
                        selected_contacts_modal = st.multiselect(
                            "Assign Local Support Contacts",
                            options=list(contact_options_modal.keys()),
                            default=current_labels,
                            key=f"modal_item_contacts_{trip_id_modal}_{idx}",
                        )
                        selected_contact_ids_modal = [
                            contact_options_modal[l] for l in selected_contacts_modal
                        ]
                    else:
                        selected_contact_ids_modal = []

                    # ---- Actions ----
                    # ---- Actions ----
                    if st.form_submit_button("💾 Update Item"):
                        item_id_to_update = item.get("id")
                        if not item_id_to_update:
                            st.warning("Cannot update — item has no DB id.")
                        else:
                            proceed = True

                            # --- Time sanity check ---
                            time_ok, time_msg = _validate_item_times(
                                e_start.isoformat(),
                                e_end.isoformat() if e_end else None,
                            )
                            if not time_ok:
                                st.warning(f"⚠️ {time_msg}")
                                if not st.checkbox(
                                    "Save anyway? (end before start)",
                                    key=(
                                        f"force_edit_item_time_"
                                        f"{trip_id_modal}_{idx}"
                                    ),
                                ):
                                    proceed = False
                                    st.info(
                                        "Tick **Save anyway?** above and "
                                        "click **💾 Update Item** again "
                                        "to save."
                                    )

                            # --- Duplicate check ---
                            dupes = _find_duplicate_items(
                                st.session_state[
                                    f"modal_items_{trip_id_modal}"
                                ],
                                e_desc,
                                e_start.isoformat(),
                                exclude_id=item_id_to_update,
                            )
                            if dupes:
                                st.warning(
                                    "⚠️ Another item on this trip has the "
                                    "same description and start time:"
                                )
                                for d in dupes:
                                    st.write(
                                        f"- {d['description']} · "
                                        f"{d['datetime_start'][:16].replace('T', ' ')}"
                                    )
                                if not st.checkbox(
                                    "Update anyway?",
                                    key=f"force_edit_item_{trip_id_modal}_{idx}",
                                ):
                                    proceed = False
                                    st.info(
                                        "Tick **Update anyway?** above and "
                                        "click **💾 Update Item** again to "
                                        "save."
                                    )

                            if proceed:
                                db.update_itinerary_item(
                                    item_id_to_update,
                                    e_type,
                                    e_desc,
                                    e_start.isoformat(),
                                    e_end.isoformat() if e_end else None,
                                    e_loc,
                                    e_cost,
                                    item.get("confirmation_code", ""),
                                    e_notes,
                                    1 if e_confirmed else 0,
                                    e_currency,
                                    timezone=e_timezone_value_modal,
                                    venue_id=e_venue_id_modal,
                                    cost_date=e_cost_date_modal.isoformat(),
                                )
                                db.set_item_delegation_members(
                                    item_id_to_update,
                                    selected_delegation_ids_modal,
                                )
                                db.set_item_contacts(
                                    item_id_to_update,
                                    selected_contact_ids_modal,
                                )
                                st.session_state.pop(
                                    f"force_edit_item_{trip_id_modal}_{idx}",
                                    None,
                                )
                                st.session_state.pop(
                                    f"force_edit_item_time_"
                                    f"{trip_id_modal}_{idx}",
                                    None,
                                )
                                st.session_state[f"modal_items_{trip_id_modal}"] = (
                                    db.get_items_for_trip(trip_id_modal)
                                )
                                st.session_state[
                                    f"modal_editing_item_{trip_id_modal}_{idx}"
                                ] = False
                                st.success("✅ Item updated.")
                                st.rerun()

                    if st.form_submit_button("❌ Cancel"):
                        st.session_state.pop(
                            f"force_edit_item_{trip_id_modal}_{idx}", None
                        )
                        st.session_state.pop(
                            f"force_edit_item_time_{trip_id_modal}_{idx}", None
                        )
                        st.session_state[
                            f"modal_editing_item_{trip_id_modal}_{idx}"
                        ] = False
                        st.rerun()

    # ---- Bulk paste (modal) ----
    def _import_into_trip(items):
        for it in items:
            db.add_itinerary_item(
                trip_id_modal,
                it["item_type"],
                it["description"],
                it["datetime_start"],
                it["datetime_end"],
                it["location"],
                it["cost"],
                it["confirmation_code"],
                it["notes"],
                it["is_confirmed"],
                it["cost_currency"],
                timezone=it["timezone"],
                venue_id=it["venue_id"],
                cost_date=it["cost_date"],
            )
        st.session_state[f"modal_items_{trip_id_modal}"] = (
            db.get_items_for_trip(trip_id_modal)
        )

    _render_bulk_paste_itinerary(
        container_key=f"modal_bulk_paste_{trip_id_modal}",
        default_timezone=exec_tz_modal,
        on_import=_import_into_trip,
        is_locked=is_locked,
    )

    with st.expander("➕ Add Item"):
        with st.form(key=f"add_item_form_{trip_id_modal}"):
            n_type = st.selectbox("Type",
                options=([cat["name"] for cat in db.get_all_categories()]
                         if db.get_all_categories() else ["Flight", "Hotel", "Meeting", "Transport"]),
                key=f"modal_n_type_{trip_id_modal}", disabled=is_locked)
            n_desc = st.text_input("Description", key=f"modal_n_desc_{trip_id_modal}", disabled=is_locked)
            n_start = st.datetime_input("Start", value=datetime.now(),
                key=f"modal_n_start_{trip_id_modal}", disabled=is_locked)
            n_end = st.datetime_input("End", value=datetime.now(),
                key=f"modal_n_end_{trip_id_modal}", disabled=is_locked)
            n_loc = st.text_input("Location", key=f"modal_n_loc_{trip_id_modal}", disabled=is_locked)
            n_cost = st.number_input("Cost", min_value=0.0, value=0.0,
                key=f"modal_n_cost_{trip_id_modal}", disabled=is_locked)
            n_currency = st.selectbox("Currency", options=base_currency_options,
                key=f"modal_n_currency_{trip_id_modal}", disabled=is_locked)
            n_cost_date = st.date_input("Cost Date", value=datetime.now(),
                key=f"modal_n_cost_date_{trip_id_modal}", disabled=is_locked)
            tz_display_names, tz_map = get_timezone_dropdown_options()
            if exec_profile_modal and exec_profile_modal.get("timezone"):
                default_tz_display_modal = next(
                    (n for n in tz_display_names if exec_profile_modal["timezone"] in n),
                    tz_display_names[0])
            else:
                default_tz_display_modal = tz_display_names[0]
            n_timezone_modal = st.selectbox("Time Zone (for this event)",
                options=tz_display_names, index=tz_display_names.index(default_tz_display_modal),
                key=f"modal_n_timezone_{trip_id_modal}", disabled=is_locked)
            n_timezone_value_modal = tz_map[n_timezone_modal]
            n_confirmed = st.checkbox("Confirmed", key=f"modal_n_confirmed_{trip_id_modal}", disabled=is_locked)
            n_notes = st.text_area("Notes", key=f"modal_n_notes_{trip_id_modal}", disabled=is_locked)
            if not is_locked and company_id_modal:
                delegation_options_modal = get_trip_delegation_options(trip_id_modal)
                selected_delegation_modal_new = st.multiselect("Assign Delegation Members",
                    options=list(delegation_options_modal.keys()),
                    key=f"modal_new_item_delegation_{trip_id_modal}")
                selected_delegation_ids_modal_new = [delegation_options_modal[l]
                    for l in selected_delegation_modal_new]
            else:
                selected_delegation_ids_modal_new = []
            if not is_locked and company_id_modal:
                contact_options_modal = get_company_contact_options(company_id_modal)
                selected_contacts_modal_new = st.multiselect("Assign Local Support Contacts",
                    options=list(contact_options_modal.keys()),
                    key=f"modal_new_item_contacts_{trip_id_modal}")
                selected_contact_ids_modal_new = [contact_options_modal[l]
                    for l in selected_contacts_modal_new]
            else:
                selected_contact_ids_modal_new = []

            if not is_locked:
                if st.form_submit_button("➕ Add Item"):
                    if not (n_desc and n_start):
                        st.warning("Description and Start Time are required.")
                    else:
                        proceed = True

                        # --- Time sanity check ---
                        time_ok, time_msg = _validate_item_times(
                            n_start.isoformat(),
                            n_end.isoformat() if n_end else None,
                        )
                        if not time_ok:
                            st.warning(f"⚠️ {time_msg}")
                            if not st.checkbox(
                                "Save anyway? (end before start)",
                                key=f"force_add_modal_item_time_{trip_id_modal}",
                            ):
                                proceed = False
                                st.info(
                                    "Tick **Save anyway?** above and click "
                                    "**➕ Add Item** again to save."
                                )

                        # --- Duplicate check ---
                        dupes = _find_duplicate_items(
                            st.session_state[f"modal_items_{trip_id_modal}"],
                            n_desc,
                            n_start.isoformat(),
                        )
                        if dupes:
                            st.warning(
                                "⚠️ This trip already has an item with the "
                                "same description and start time:"
                            )
                            for d in dupes:
                                st.write(
                                    f"- {d['description']} · "
                                    f"{d['datetime_start'][:16].replace('T', ' ')}"
                                )
                            if not st.checkbox(
                                "Add anyway?",
                                key=f"force_add_modal_item_{trip_id_modal}",
                            ):
                                proceed = False
                                st.info(
                                    "Tick **Add anyway?** above and click "
                                    "**➕ Add Item** again to save the "
                                    "duplicate."
                                )

                        if proceed:
                            new_item_id = db.add_itinerary_item(
                                trip_id_modal,
                                n_type,
                                n_desc,
                                n_start.isoformat(),
                                n_end.isoformat() if n_end else None,
                                n_loc,
                                n_cost,
                                "",
                                n_notes,
                                1 if n_confirmed else 0,
                                n_currency,
                                timezone=n_timezone_value_modal,
                                venue_id=None,
                                cost_date=n_cost_date.isoformat(),
                            )
                            if selected_delegation_ids_modal_new:
                                db.set_item_delegation_members(
                                    new_item_id, selected_delegation_ids_modal_new
                                )
                            if selected_contact_ids_modal_new:
                                db.set_item_contacts(
                                    new_item_id, selected_contact_ids_modal_new
                                )
                            st.session_state.pop(
                                f"force_add_modal_item_{trip_id_modal}", None
                            )
                            st.session_state.pop(
                                f"force_add_modal_item_time_{trip_id_modal}", None
                            )
                            st.session_state[f"modal_items_{trip_id_modal}"] = (
                                db.get_items_for_trip(trip_id_modal)
                            )
                            st.success("✅ Item added.")
                            st.rerun()


    # ---- Linked Activities ----
    st.write("**🔗 Linked Activities**")

    linked_meetings = db.get_meetings_for_trip(trip_id_modal)
    linked_events = db.get_events_for_trip(trip_id_modal)
    linked_speaking = db.get_speaking_for_trip(trip_id_modal)

    total_linked = (
        len(linked_meetings) + len(linked_events) + len(linked_speaking)
    )

    if total_linked == 0:
        st.caption(
            "No meetings, events, or speaking engagements linked to "
            "this trip yet. Use the 🔗 Linked Trip field on any record "
            "to link it here."
        )
    else:
        if linked_meetings:
            st.markdown(f"**📅 Meetings ({len(linked_meetings)})**")
            for m in linked_meetings:
                col1, col2 = st.columns([5, 1])
                with col1:
                    st.markdown(f"• {m['title']}")
                    st.caption(
                        f"📅 {fmt_date(m.get('meeting_date'))}"
                        + (
                            f"  ·  🕐 {m['meeting_time']}"
                            if m.get("meeting_time") else ""
                        )
                        + f"  ·  {_meeting_status_label(m.get('status'))}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"trip_link_mtg_{trip_id_modal}_{m['id']}",
                    ):
                        st.session_state["dialog_meeting_id"] = m["id"]
                        st.rerun()

        if linked_events:
            st.markdown(f"**🎉 Events ({len(linked_events)})**")
            for ev in linked_events:
                col1, col2 = st.columns([5, 1])
                with col1:
                    st.markdown(f"• {ev['name']}")
                    st.caption(
                        f"📅 {fmt_date(ev.get('event_date'))}"
                        + (
                            f"  ·  🕐 {ev['start_time']}"
                            if ev.get("start_time") else ""
                        )
                        + f"  ·  {_event_status_label(ev.get('status'))}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"trip_link_ev_{trip_id_modal}_{ev['id']}",
                    ):
                        st.session_state["dialog_event_id"] = ev["id"]
                        st.rerun()

        if linked_speaking:
            st.markdown(
                f"**🎤 Speaking Engagements ({len(linked_speaking)})**"
            )
            for eng in linked_speaking:
                col1, col2 = st.columns([5, 1])
                with col1:
                    st.markdown(f"• {eng['title']}")
                    st.caption(
                        f"📅 {fmt_date(eng.get('engagement_date'))}"
                        + f"  ·  {_speaking_status_label(eng.get('status'))}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"trip_link_sp_{trip_id_modal}_{eng['id']}",
                    ):
                        st.session_state["dialog_speaking_id"] = eng["id"]
                        st.rerun()

    st.divider()

    # ---- Emergency Info ----
    st.write("**🚨 Emergency Info**")

    # Emergency contacts for this executive's company
    emergency_contacts = []
    if company_id_modal:
        all_c = db.get_contacts(company_id_modal, active_only=True)
        emergency_contacts = [c for c in all_c if c.get("type") == "Emergency"]

    # Emergency numbers from destination guides
    trip_emergency_numbers = {}
    for stop in stops:
        country = (stop.get("country") or "").strip()
        if country and country not in trip_emergency_numbers:
            dg = db.get_destination_guide_by_country(country)
            if dg:
                trip_emergency_numbers[country] = {
                    "police": dg.get("emergency_police"),
                    "ambulance": dg.get("emergency_ambulance"),
                    "fire": dg.get("emergency_fire"),
                }

    # Hospitals for the trip's stop cities
    trip_hospitals = db.get_hospitals_for_trip(trip_id_modal)

    # Embassies matching (host country of stop) + (exec passport countries)
    passport_countries = db.get_executive_passport_countries(trip_modal_data["exec_id"])
    trip_embassies = db.get_embassies_for_trip(
        trip_id_modal, representing_countries=passport_countries
    )

    with st.expander("View Emergency Details", expanded=False):

        if emergency_contacts:
            st.markdown("**🚨 Emergency Contacts**")
            for ec in emergency_contacts:
                col_a, col_b = st.columns([2, 3])
                with col_a:
                    st.write(f"**{ec['name']}**")
                    if ec.get("role"):
                        st.caption(ec["role"])
                with col_b:
                    if ec.get("phone"):
                        st.write(f"📞 {ec['phone']}")
                    if ec.get("email"):
                        st.write(f"✉️ {ec['email']}")
                st.divider()

        if trip_emergency_numbers:
            st.markdown("**☎️ Emergency Numbers**")
            for loc, nums in trip_emergency_numbers.items():
                st.write(f"**{loc}**")
                parts = []
                if nums.get("police"):
                    parts.append(f"🚓 Police: {nums['police']}")
                if nums.get("ambulance"):
                    parts.append(f"🚑 Ambulance: {nums['ambulance']}")
                if nums.get("fire"):
                    parts.append(f"🚒 Fire: {nums['fire']}")
                if parts:
                    st.caption(" · ".join(parts))
            st.divider()

        if trip_hospitals:
            st.markdown("**🏥 Hospitals**")
            for h in trip_hospitals:
                st.write(f"**{h['name']}** — {h.get('city', '')}")
                if h.get("phone"):
                    st.caption(f"📞 {h['phone']}")
                if h.get("maps_url"):
                    st.caption(f"[Open in Maps]({h['maps_url']})")
            st.divider()

        if trip_embassies:
            st.markdown("**🏛️ Embassies / Consulates**")
            for emb in trip_embassies:
                st.write(
                    f"**{emb.get('name') or emb['representing_country'] + ' Embassy'}** "
                    f"in {emb['host_country']}"
                )
                if emb.get("address"):
                    st.caption(f"🏠 {emb['address']}")
                if emb.get("phone"):
                    st.caption(f"📞 {emb['phone']}")
                if emb.get("hours"):
                    st.caption(f"🕐 {emb['hours']}")
            st.divider()

        if not (
            emergency_contacts
            or trip_emergency_numbers
            or trip_hospitals
            or trip_embassies
        ):
            st.caption(
                "No emergency info available for this trip yet. "
                "Add hospitals, embassies, or emergency contacts "
                "to the Library or Contacts tab."
            )

    # Per Diem & Expenses
    st.write("**💰 Per Diem & Expenses**")
    trip_members_for_expenses = db.get_trip_delegation_members(trip_id_modal)
    if not trip_members_for_expenses:
        st.caption("No delegation members on this trip yet. Select travellers above.")
    else:
        with st.expander("⚙️ Per Diem Settings", expanded=False):
            for member in trip_members_for_expenses:
                existing_pd = db.get_per_diem(trip_id_modal, member["id"])
                col_a, col_b, col_c, col_d = st.columns([2, 1, 1, 1])
                with col_a:
                    st.write(f"**{member['name']}** ({member.get('role', '')})")
                with col_b:
                    new_rate = st.number_input("Daily Rate", min_value=0.0, step=10.0,
                        value=(float(existing_pd["daily_rate"]) if existing_pd else 0.0),
                        key=f"pd_rate_{trip_id_modal}_{member['id']}", disabled=is_locked)
                with col_c:
                    new_days = st.number_input("Days", min_value=0, step=1,
                        value=(int(existing_pd["days"]) if existing_pd else 0),
                        key=f"pd_days_{trip_id_modal}_{member['id']}", disabled=is_locked)
                with col_d:
                    st.write("")
                    st.write("")
                    if not is_locked:
                        if st.button("💾 Save", key=f"pd_save_{trip_id_modal}_{member['id']}"):
                            db.set_per_diem(trip_id_modal, member["id"], new_rate, new_days,
                                currency=trip_modal_data.get("base_currency", "USD"))
                            st.success(f"Per diem saved for {member['name']}.")
                            st.rerun()

        base_cur_modal_exp = trip_modal_data.get("base_currency", "USD")
        # ---- Bulk import expenses ----
        with st.expander("📋 Bulk import expenses (paste or CSV)", expanded=False):
            st.caption(
                "Paste rows copied from a spreadsheet, or a CSV file. "
                "A header row is required. Columns: **traveler, date, "
                "category, description, amount, currency, reimbursable, "
                "notes**. `traveler` must match a delegation member's "
                "name exactly."
            )

            # Build a name → contact_id map for validation + resolution
            member_name_map = {
                m["name"].strip().lower(): m["id"] for m in trip_members_for_expenses
            }

            raw_exp = st.text_area(
                "Paste here",
                height=180,
                key=f"exp_bulk_raw_{trip_id_modal}",
                placeholder=(
                    "traveler\tdate\tcategory\tdescription\tamount\t"
                    "currency\treimbursable\tnotes\n"
                    "Jane Doe\t2025-04-12\tMeals\tTeam dinner\t145.50\t"
                    "USD\tyes\tClient meeting\n"
                    "Jane Doe\t2025-04-13\tTransport\tTaxi to airport\t"
                    "62.00\tUSD\tyes\t"
                ),
            )

            parsed_exp, exp_errors = ([], [])
            if raw_exp and raw_exp.strip():
                parsed_exp, exp_errors = utils.parse_expense_paste(
                    raw_exp, default_currency=base_cur_modal_exp
                )

            # Validate traveler names against the delegation roster
            unknown_names = sorted(
                {
                    e["traveler"]
                    for e in parsed_exp
                    if e["traveler"].lower() not in member_name_map
                }
            )
            if unknown_names:
                st.error(
                    "🚫 Unknown traveler name(s): "
                    + ", ".join(unknown_names)
                    + ". Add them to the trip delegation first."
                )
                # Drop unresolvable rows so Import is only enabled when all
                # rows map cleanly
                parsed_exp = [
                    e for e in parsed_exp if e["traveler"].lower() in member_name_map
                ]

            if parsed_exp:
                st.success(f"✅ Parsed {len(parsed_exp)} expense row(s).")
                with st.expander(f"👁️ Preview ({len(parsed_exp)})", expanded=True):
                    preview_rows = [
                        {
                            "Traveler": e["traveler"],
                            "Date": e["expense_date"],
                            "Category": e["category"],
                            "Description": e["description"],
                            "Amount": f"{e['amount']:.2f}",
                            "Ccy": e["currency"],
                            "Reim": "Yes" if e["is_reimbursable"] else "No",
                            "Notes": e["notes"],
                        }
                        for e in parsed_exp
                    ]
                    st.dataframe(
                        preview_rows,
                        use_container_width=True,
                        hide_index=True,
                    )

            if exp_errors:
                with st.expander(f"⚠️ {len(exp_errors)} warning(s)", expanded=False):
                    for e in exp_errors:
                        st.caption(f"• {e}")

            if not is_locked:
                import_label = (
                    f"➕ Import {len(parsed_exp)} expense(s)"
                    if parsed_exp
                    else "➕ Import"
                )
                if st.button(
                    import_label,
                    key=f"exp_bulk_import_{trip_id_modal}",
                    type="primary",
                    disabled=not parsed_exp,
                ):
                    for e in parsed_exp:
                        contact_id = member_name_map[e["traveler"].lower()]
                        db.add_expense(
                            trip_id=trip_id_modal,
                            contact_id=contact_id,
                            expense_date=e["expense_date"],
                            category=e["category"],
                            description=e["description"],
                            amount=e["amount"],
                            currency=e["currency"],
                            receipt_path=e["receipt_path"],
                            notes=e["notes"],
                            is_reimbursable=e["is_reimbursable"],
                        )
                    st.session_state.pop(f"exp_bulk_raw_{trip_id_modal}", None)
                    st.success(f"Imported {len(parsed_exp)} expense(s).")
                    st.rerun()
        with st.expander("➕ Add Expense", expanded=False):
            with st.form(key=f"add_expense_form_{trip_id_modal}"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    member_options_exp = {m["name"]: m["id"] for m in trip_members_for_expenses}
                    exp_member_label = st.selectbox("Traveler*", options=list(member_options_exp.keys()),
                        key=f"exp_member_{trip_id_modal}")
                    exp_member_id = member_options_exp[exp_member_label]
                with col2:
                    exp_date = st.date_input("Date*", value=datetime.now(), key=f"exp_date_{trip_id_modal}")
                    exp_category = st.selectbox("Category",
                        options=["Meals", "Transport", "Lodging", "Incidentals",
                                 "Entertainment", "Communication", "Other"],
                        key=f"exp_category_{trip_id_modal}")
                with col3:
                    exp_amount = st.number_input("Amount", min_value=0.0, step=10.0, value=0.0,
                        key=f"exp_amount_{trip_id_modal}")
                    exp_reimbursable = st.checkbox("Reimbursable", value=True,
                        key=f"exp_reim_{trip_id_modal}")
                exp_description = st.text_input("Description", key=f"exp_desc_{trip_id_modal}")
                exp_notes = st.text_area("Notes", key=f"exp_notes_{trip_id_modal}", height=80)
                exp_receipt = st.file_uploader("Attach Receipt (image or PDF)",
                    type=["png", "jpg", "jpeg", "pdf"], key=f"exp_receipt_{trip_id_modal}")
                if st.form_submit_button("➕ Add Expense"):
                    if exp_amount <= 0:
                        st.warning("Amount must be greater than zero.")
                    else:
                        receipt_path_val = None
                        if exp_receipt is not None:
                            folder = f"ea_receipts/trip_{trip_id_modal}/expenses"
                            os.makedirs(folder, exist_ok=True)
                            fname = f"exp_{int(datetime.now().timestamp())}_{exp_receipt.name}"
                            receipt_path_val = f"{folder}/{fname}"
                            with open(receipt_path_val, "wb") as f:
                                f.write(exp_receipt.getbuffer())
                        db.add_expense(trip_id=trip_id_modal, contact_id=exp_member_id,
                            expense_date=exp_date.isoformat(), category=exp_category,
                            description=exp_description, amount=exp_amount,
                            currency=base_cur_modal_exp, receipt_path=receipt_path_val,
                            notes=exp_notes, is_reimbursable=1 if exp_reimbursable else 0)
                        st.success("Expense added!")
                        st.rerun()

        summary_exp = db.get_expense_summary(trip_id_modal)
        if summary_exp:
            st.write("**📊 Delegation Expense Summary**")
            total_allowance_all = sum(s["allowance"] for s in summary_exp)
            total_spent_all = sum(s["spent"] for s in summary_exp)
            total_remaining_all = total_allowance_all - total_spent_all

            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                st.metric("Delegation Allowance", fmt_money(total_allowance_all, base_cur_modal_exp))
            with col_s2:
                st.metric("Total Spent", fmt_money(total_spent_all, base_cur_modal_exp))
            with col_s3:
                st.metric("Remaining", fmt_money(total_remaining_all, base_cur_modal_exp))

            st.write("**Per‑Traveler Breakdown**")
            for s in summary_exp:
                with st.container():
                    col_a, col_b, col_c, col_d = st.columns([2, 1.2, 1.2, 1.2])
                    with col_a:
                        st.write(f"**{s['name']}** ({s.get('role','')})")
                        st.caption(f"{s['entry_count']} entr{'y' if s['entry_count']==1 else 'ies'}")
                    with col_b:
                        st.write(f"Allowance: **{fmt_money(s['allowance'], base_cur_modal_exp)}**")
                    with col_c:
                        st.write(f"Spent: **{fmt_money(s['spent'], base_cur_modal_exp)}**")
                    with col_d:
                        color = "🟢" if s["remaining"] >= 0 else "🔴"
                        st.write(f"{color} Remaining: **{fmt_money(s['remaining'], base_cur_modal_exp)}**")
                    st.divider()

        all_expenses = db.get_expenses(trip_id_modal)
        if all_expenses:
            with st.expander(f"📋 All Expenses ({len(all_expenses)})", expanded=False):
                member_name_lookup = {m["id"]: m["name"] for m in trip_members_for_expenses}
                for ex in all_expenses:
                    ex_id = ex["id"]
                    col_a, col_b, col_c, col_d, col_e = st.columns([1, 2, 2, 1, 1])
                    with col_a:
                        st.write(fmt_date(ex["expense_date"]))
                    with col_b:
                        st.write(f"**{member_name_lookup.get(ex['contact_id'], '—')}**")
                        if ex.get("category"):
                            st.caption(ex["category"])
                    with col_c:
                        st.write(ex.get("description") or "—")
                        if ex.get("receipt_path"):
                            st.caption("🧾 Receipt attached")
                    with col_d:
                        st.write(f"**{fmt_money(ex['amount'], ex.get('currency','USD'))}**")
                    with col_e:
                        if not is_locked:
                            if st.button("🗑️", key=f"del_exp_{ex_id}_{trip_id_modal}"):
                                db.delete_expense(ex_id)
                                st.rerun()
                    st.divider()
        else:
            st.caption("No expenses yet.")

    # Packing Lists
    st.write("**🎒 Packing Lists**")
    trip_members_for_packing = db.get_trip_delegation_members(trip_id_modal)
    if not trip_members_for_packing:
        st.caption("No delegation members on this trip yet.")
    else:
        with st.expander("⚙️ Apply Packing Template to All Travelers", expanded=False):
            packing_tpls = db.get_packing_templates(active_only=True)
            if packing_tpls:
                tpl_labels = {f"{t['name']}" + (f" ({t['category']})" if t.get("category") else ""): t["id"]
                    for t in packing_tpls}
                selected_tpl_label = st.selectbox("Select Packing Template",
                    options=list(tpl_labels.keys()), key=f"packing_tpl_apply_{trip_id_modal}")
                if not is_locked:
                    if st.button("Apply to All Travelers", key=f"apply_packing_all_{trip_id_modal}"):
                        tpl_id = tpl_labels[selected_tpl_label]
                        total_added = 0
                        for member in trip_members_for_packing:
                            total_added += db.apply_packing_template(trip_id_modal, member["id"], tpl_id)
                        st.success(f"Added {total_added} item(s) across all travelers.")
                        st.rerun()
            else:
                st.caption("No packing templates available. Create one in the Library tab.")

        for member in trip_members_for_packing:
            plist = db.get_packing_list(trip_id_modal, member["id"])
            if not plist:
                if not is_locked:
                    if st.button(f"➕ Create packing list for {member['name']}",
                        key=f"create_packing_{trip_id_modal}_{member['id']}"):
                        db.create_packing_list(trip_id_modal, member["id"])
                        st.rerun()
                else:
                    st.caption(f"No packing list for {member['name']}.")
                continue

            list_id = plist["id"]
            packing_items_local = db.get_packing_items(list_id)
            total = len(packing_items_local)
            packed = sum(1 for it in packing_items_local if it["packed"])

            with st.expander(
                f"🎒 {member['name']} ({packed} of {total} packed)", expanded=False
            ):
                progress_val = packed / total if total else 0
                st.progress(progress_val)
                categories = {}
                for it in packing_items_local:
                    cat = it.get("category") or "Other"
                    categories.setdefault(cat, []).append(it)
                for cat_name, cat_items in categories.items():
                    st.caption(f"**{cat_name}**")
                    for it in cat_items:
                        col_cb, col_name, col_del = st.columns([0.5, 6, 1])
                        with col_cb:
                            checked = st.checkbox("", value=bool(it["packed"]),
                                key=f"pack_chk_{trip_id_modal}_{list_id}_{it['id']}",
                                disabled=is_locked)
                            if checked != bool(it["packed"]) and not is_locked:
                                db.toggle_packing_item(it["id"], checked)
                                st.rerun()
                        with col_name:
                            if it["packed"]:
                                st.markdown(f"~~{it['item_name']}~~")
                            else:
                                st.write(it["item_name"])
                        with col_del:
                            if not is_locked:
                                if st.button("🗑️", key=f"pack_del_{trip_id_modal}_{list_id}_{it['id']}"):
                                    db.delete_packing_item(it["id"])
                                    st.rerun()

                if not is_locked:
                    with st.form(key=f"add_packing_item_{trip_id_modal}_{list_id}"):
                        col_a, col_b = st.columns([3, 2])
                        with col_a:
                            new_pi_name = st.text_input("Add item", key=f"new_pi_name_{trip_id_modal}_{list_id}")
                        with col_b:
                            new_pi_cat = st.selectbox("Category",
                                options=["Documents & Badges", "Devices & Chargers", "Clothing",
                                         "Formalwear", "Presentation Materials", "Toiletries",
                                         "Personal", "Other"],
                                key=f"new_pi_cat_{trip_id_modal}_{list_id}")
                        if st.form_submit_button("➕ Add Item"):
                            if new_pi_name:
                                db.add_packing_item(list_id, new_pi_name, category=new_pi_cat)
                                st.rerun()
                            else:
                                st.warning("Item name required.")

                if not is_locked:
                    if st.button("🗑️ Delete Entire List", key=f"pack_del_list_{trip_id_modal}_{list_id}"):
                        db.delete_packing_list(list_id)
                        st.rerun()

    # Trip Checklists
    st.write("**✅ Trip Checklists**")
    trip_checklists = db.get_trip_checklists(trip_id_modal)
    with st.expander("⚙️ Apply Checklist Template", expanded=False):
        cl_tpls = db.get_checklist_templates(active_only=True)
        if cl_tpls:
            cl_tpl_labels = {t["name"]: t["id"] for t in cl_tpls}
            selected_cl_tpl = st.selectbox("Select Checklist Template",
                options=list(cl_tpl_labels.keys()), key=f"cl_tpl_apply_{trip_id_modal}")
            if not is_locked:
                if st.button("Apply Template", key=f"apply_cl_tpl_{trip_id_modal}"):
                    new_cl_id = db.apply_checklist_template(trip_id_modal, cl_tpl_labels[selected_cl_tpl])
                    if new_cl_id:
                        st.success("Checklist added.")
                        st.rerun()
        else:
            st.caption("No checklist templates available. Create one in the Library tab.")

    if trip_checklists:
        for cl in trip_checklists:
            cl_id = cl["id"]
            cl_items = db.get_checklist_items(cl_id)
            cl_total = len(cl_items)
            cl_done = sum(1 for it in cl_items if it["is_done"])
            with st.expander(f"✅ {cl['name']} ({cl_done} of {cl_total} done)", expanded=False):
                if cl.get("description"):
                    st.caption(cl["description"])
                p_val = cl_done / cl_total if cl_total else 0
                st.progress(p_val)
                for it in cl_items:
                    col_cb, col_text, col_del = st.columns([0.5, 6, 1])
                    with col_cb:
                        is_checked = st.checkbox("", value=bool(it["is_done"]),
                            key=f"cl_chk_{trip_id_modal}_{cl_id}_{it['id']}", disabled=is_locked)
                        if is_checked != bool(it["is_done"]) and not is_locked:
                            db.toggle_checklist_item(it["id"], is_checked)
                            st.rerun()
                    with col_text:
                        if it["is_done"]:
                            st.markdown(f"~~{it['item_text']}~~")
                        else:
                            st.write(it["item_text"])
                    with col_del:
                        if not is_locked:
                            if st.button("🗑️", key=f"cl_del_{trip_id_modal}_{cl_id}_{it['id']}"):
                                db.delete_checklist_item(it["id"])
                                st.rerun()
                if not is_locked:
                    with st.form(key=f"add_cl_item_{trip_id_modal}_{cl_id}"):
                        new_cl_text = st.text_input("Add checklist item",
                            key=f"new_cl_text_{trip_id_modal}_{cl_id}")
                        if st.form_submit_button("➕ Add"):
                            if new_cl_text:
                                db.add_checklist_item(cl_id, new_cl_text)
                                st.rerun()
                            else:
                                st.warning("Item text required.")
                if not is_locked:
                    if st.button("🗑️ Delete Checklist", key=f"del_cl_{trip_id_modal}_{cl_id}"):
                        db.delete_trip_checklist(cl_id)
                        st.rerun()
    else:
        st.caption("No checklists on this trip yet.")

    # Actions
    st.divider()
    col_left, col_mid, col_right, col_dup, col_travel = st.columns(5)
    with col_left:
        if not is_locked:
            if st.button("🗑️ Delete This Trip", type="primary", use_container_width=True,
                key=f"delete_trip_modal_{trip_id_modal}"):
                st.session_state[f"confirm_del_modal_{trip_id_modal}"] = True
    with col_mid:
        if current_status in ["approved", "final"]:
            if st.button("↩️ Revert to Draft", use_container_width=True,
                key=f"revert_trip_modal_{trip_id_modal}"):
                db.update_trip_status(trip_id_modal, "draft")
                st.success("Trip reverted to Draft status.")
                st.session_state.pop(f"modal_stops_{trip_id_modal}", None)
                st.session_state.pop(f"modal_items_{trip_id_modal}", None)
                st.rerun()
    with col_right:
        if st.button("📋 Save as Template", use_container_width=True,
            key=f"save_template_modal_{trip_id_modal}"):
            st.session_state[f"show_save_template_modal_{trip_id_modal}"] = True
    with col_dup:
        if st.button("📄 Duplicate Trip", use_container_width=True,
            key=f"duplicate_trip_modal_{trip_id_modal}",
            help="Create a copy of this trip — stops, items, delegation, and "
                 "checklists are carried over. You can edit dates on the copy."):
            new_trip_id = db.duplicate_trip(
                trip_id_modal, trip_modal_data["exec_id"]
            )
            if new_trip_id:
                st.success(
                    f"✅ Trip duplicated (new ID: {new_trip_id}). "
                    "Find it in **✈️ All Trips**."
                )
                st.rerun()
            else:
                st.error("Failed to duplicate trip — see logs for details.")
    with col_travel:
        if st.button("📦 Travel Pack", use_container_width=True,
            key=f"travel_pack_modal_{trip_id_modal}"):
            st.session_state[f"show_travel_pack_modal_{trip_id_modal}"] = True

    if st.session_state.get(f"show_save_template_modal_{trip_id_modal}", False):
        st.info("Save this trip as a reusable template.")
        template_name_modal = st.text_input("Template Name*",
            value=f"{trip_modal_data.get('purpose', '')} Template",
            key=f"template_name_modal_{trip_id_modal}")
        template_desc_modal = st.text_input("Description (optional)",
            key=f"template_desc_modal_{trip_id_modal}")
        col_y, col_n = st.columns(2)
        with col_y:
            if st.button("💾 Save", key=f"confirm_save_template_modal_{trip_id_modal}"):
                if template_name_modal:
                    new_id = db.save_trip_as_template(trip_id_modal, template_name_modal, template_desc_modal)
                    if new_id:
                        st.success(f"✅ Template '{template_name_modal}' saved!")
                        st.session_state.pop(f"show_save_template_modal_{trip_id_modal}", None)
                        st.rerun()
                    else:
                        st.error("Failed to save template.")
                else:
                    st.warning("Template Name is required.")
        with col_n:
            if st.button("Cancel", key=f"cancel_save_template_modal_{trip_id_modal}"):
                st.session_state.pop(f"show_save_template_modal_{trip_id_modal}", None)
                st.rerun()
    if st.session_state.get(f"show_travel_pack_modal_{trip_id_modal}", False):
        st.info("Generate a self-contained Travel Pack in your preferred format.")
        html_content = doc_generator.generate_travel_pack_html(
            trip_id_modal, exec_tz_modal, display_mode_modal
        )
        if html_content:
            # ---- Preview toggle ----
            preview_key = f"travel_pack_preview_{trip_id_modal}"
            col_prev, col_html, col_pdf, col_word = st.columns([1.2, 1, 1, 1])
            with col_prev:
                if st.button(
                    "👁️ Preview",
                    key=f"toggle_travel_pack_preview_{trip_id_modal}",
                    use_container_width=True,
                    help="Render the HTML travel pack inline without downloading.",
                ):
                    st.session_state[preview_key] = not st.session_state.get(
                        preview_key, False
                    )
                    st.rerun()
            with col_html:
                st.download_button(
                    label="🌐 HTML",
                    data=html_content,
                    file_name=f"TravelPack_{trip_modal_data.get('purpose', 'trip')}.html",
                    mime="text/html",
                    key=f"download_travel_pack_html_{trip_id_modal}",
                    use_container_width=True,
                )
            with col_pdf:
                pdf_stream = doc_generator.generate_travel_pack_pdf(
                    trip_id_modal, exec_tz_modal, display_mode_modal
                )
                if pdf_stream:
                    st.download_button(
                        label="📄 PDF",
                        data=pdf_stream,
                        file_name=f"TravelPack_{trip_modal_data.get('purpose', 'trip')}.pdf",
                        mime="application/pdf",
                        key=f"download_travel_pack_pdf_{trip_id_modal}",
                        use_container_width=True,
                    )
            with col_word:
                docx_stream = doc_generator.generate_travel_pack_docx(
                    trip_id_modal, exec_tz_modal, display_mode_modal
                )
                if docx_stream:
                    st.download_button(
                        label="📄 Word",
                        data=docx_stream,
                        file_name=f"TravelPack_{trip_modal_data.get('purpose', 'trip')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml",
                        key=f"download_travel_pack_docx_{trip_id_modal}",
                        use_container_width=True,
                    )

            # ---- Inline preview ----
            if st.session_state.get(preview_key, False):
                with st.container(border=True):
                    components.html(html_content, height=800, scrolling=True)
                    st.caption(
                        "ℹ️ Links inside the preview may not open — "
                        "download the HTML to use them."
                    )

            st.divider()
            col_pack_hide, col_pack_close = st.columns(2)
            with col_pack_hide:
                if st.button(
                    "🙈 Hide Travel Pack",
                    key=f"hide_travel_pack_{trip_id_modal}",
                    use_container_width=True,
                    help="Collapse the travel pack panel but keep the trip modal open.",
                ):
                    st.session_state.pop(
                        f"show_travel_pack_modal_{trip_id_modal}", None
                    )
                    st.session_state.pop(preview_key, None)
                    st.rerun()
            with col_pack_close:
                if st.button(
                    "✅ Done — Close Trip",
                    key=f"close_modal_from_pack_{trip_id_modal}",
                    type="primary",
                    use_container_width=True,
                    help="Close the trip modal entirely.",
                ):
                    st.session_state.pop(f"modal_stops_{trip_id_modal}", None)
                    st.session_state.pop(f"modal_items_{trip_id_modal}", None)
                    st.session_state.pop("dialog_trip_id", None)
                    st.rerun()
        else:
            st.error("Failed to generate travel pack.")

    if st.session_state.get(f"confirm_del_modal_{trip_id_modal}", False):

        def _delete_and_close():
            db.delete_trip(trip_id_modal)
            st.session_state.pop("dialog_trip_id", None)

        render_delete_confirmation(
            flag_key=f"confirm_del_modal_{trip_id_modal}",
            title="Permanently delete this trip?",
            on_confirm=_delete_and_close,
        )

    if st.button("Close", key=f"close_modal_{trip_id_modal}"):
        st.session_state.pop(f"modal_stops_{trip_id_modal}", None)
        st.session_state.pop(f"modal_items_{trip_id_modal}", None)
        st.session_state.pop("dialog_trip_id", None)   # ← new
        st.rerun()

# =========================================================
# MEETING DETAIL DIALOG
# =========================================================


def _clear_dialog_meeting_id():
    st.session_state.pop("dialog_meeting_id", None)


@st.dialog(
    "📅 Meeting",
    width="large",
    on_dismiss=_clear_dialog_meeting_id,
)

# =========================================================
# COMMUNICATIONS DIALOG
# =========================================================


def _clear_dialog_comm_contact_id():
    st.session_state.pop("dialog_comm_contact_id", None)


@st.dialog(
    "📞 Communications",
    width="large",
    on_dismiss=_clear_dialog_comm_contact_id,
)
def _render_communications_dialog(contact_id):
    contact = db.get_contact(contact_id)
    if not contact:
        st.warning("Contact not found.")
        return

    st.subheader(f"📞 {contact['name']}")
    if contact.get("role"):
        st.caption(contact["role"])

    # Summary row
    last_contacted = contact.get("last_contacted_at") or "—"
    last_channel = _channel_label(contact.get("last_channel"))
    next_action = contact.get("next_action") or "—"
    next_action_date = contact.get("next_action_date") or "—"

    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("Last Contacted", last_contacted)
    with col_b:
        st.metric("Channel", last_channel)

    col_c, col_d = st.columns(2)
    with col_c:
        st.metric("Next Action", next_action)
    with col_d:
        st.metric("Due", next_action_date)

    st.divider()

    # ================================================
    # LOG NEW COMMUNICATION
    # ================================================
    with st.expander("📝 Log a Communication", expanded=False):
        execs_active = db.get_all_executives(active_only=True)
        with st.form(key=f"comm_log_form_{contact_id}"):
            col1, col2 = st.columns(2)
            with col1:
                comm_date = st.date_input(
                    "Date", value=datetime.now().date(),
                    key=f"comm_date_{contact_id}",
                )
                channel = st.selectbox(
                    "Channel",
                    options=COMM_CHANNELS,
                    format_func=_channel_label,
                    key=f"comm_channel_{contact_id}",
                )
            with col2:
                direction = st.selectbox(
                    "Direction",
                    options=COMM_DIRECTIONS,
                    format_func=_direction_label,
                    key=f"comm_direction_{contact_id}",
                )
                exec_label = "—"
                if execs_active:
                    exec_options = {"—": None}
                    exec_options.update(
                        {e["name"]: e["id"] for e in execs_active}
                    )
                    exec_label = st.selectbox(
                        "Executive (optional)",
                        options=list(exec_options.keys()),
                        key=f"comm_exec_{contact_id}",
                    )
                    exec_id_val = exec_options[exec_label]
                else:
                    exec_id_val = None

            summary = st.text_input(
                "Summary (one line)",
                placeholder="e.g. Called re: Q3 board prep",
                key=f"comm_summary_{contact_id}",
            )
            col_a, col_b = st.columns([3, 2])
            with col_a:
                next_action_text = st.text_input(
                    "Next Action (optional)",
                    placeholder="e.g. Follow up on budget",
                    key=f"comm_next_action_{contact_id}",
                )
            with col_b:
                next_action_date_input = st.date_input(
                    "Next Action Date (optional)",
                    value=None,
                    key=f"comm_next_action_date_{contact_id}",
                )

            if st.form_submit_button("💾 Save"):
                if not summary:
                    st.error("Please enter a short summary.")
                else:
                    db.add_communication(
                        contact_id=contact_id,
                        comm_date=comm_date.isoformat(),
                        channel=channel,
                        direction=direction,
                        summary=summary,
                        next_action=next_action_text or None,
                        next_action_date=(
                            next_action_date_input.isoformat()
                            if next_action_date_input else None
                        ),
                        exec_id=exec_id_val,
                        linked_to_type="standalone",
                        linked_to_id=None,
                    )
                    st.success("✅ Logged.")
                    st.rerun()

    st.divider()

    # ================================================
    # TIMELINE
    # ================================================
    st.markdown("**📜 Timeline**")
    communications = db.get_communications_for_contact(contact_id)
    if not communications:
        st.caption("No communications logged yet.")
    else:
        for comm in communications:
            col1, col2, col3 = st.columns([5, 1, 0.5])
            with col1:
                st.markdown(
                    f"**{fmt_date(comm['comm_date'])}** · "
                    f"{_channel_label(comm.get('channel'))} · "
                    f"{_direction_label(comm.get('direction'))}"
                )
                if comm.get("summary"):
                    st.write(comm["summary"])
                if comm.get("next_action"):
                    due = comm.get("next_action_date") or "—"
                    st.caption(f"→ Next: {comm['next_action']} (due {due})")
                if comm.get("executive_name"):
                    st.caption(f"On behalf of: {comm['executive_name']}")
            with col2:
                if comm.get("linked_to_type") and comm.get("linked_to_id"):
                    st.caption(
                        f"🔗 {comm['linked_to_type']} #{comm['linked_to_id']}"
                    )
            with col3:
                if st.button(
                    "🗑️", key=f"comm_del_{comm['id']}",
                    help="Delete this entry",
                ):
                    db.delete_communication(comm["id"])
                    st.rerun()
            st.divider()

    # Danger zone
    if st.button("Close", key=f"close_comm_dialog_{contact_id}"):
        st.session_state.pop("dialog_comm_contact_id", None)
        st.rerun()

# =========================================================
# EVENT DETAIL DIALOG
# =========================================================


def _clear_dialog_event_id():
    st.session_state.pop("dialog_event_id", None)


@st.dialog(
    "🎉 Event",
    width="large",
    on_dismiss=_clear_dialog_event_id,
)
def _render_event_detail_dialog(event_id):
    event = db.get_event(event_id)
    if not event:
        st.warning("Event not found.")
        return

    exec_id = event["exec_id"]
    st.subheader(f"🎉 {event['name']}")
    st.caption(
        f"**{event.get('executive_name') or '—'}** · "
        f"{_event_status_label(event.get('status'))}  ·  "
        f"📅 {fmt_date(event.get('event_date'))}"
    )

    # ================================================
    # SECTION: Edit Event
    # ================================================
    with st.expander("✏️ Edit Event", expanded=False):
        with st.form(key=f"ev_edit_form_{event_id}"):
            col1, col2 = st.columns(2)
            with col1:
                edit_name = st.text_input(
                    "Name*", value=event.get("name") or "",
                    key=f"ev_edit_name_{event_id}",
                )
                edit_type = st.selectbox(
                    "Type",
                    options=[""] + EVENT_TYPES,
                    index=(
                        ([""] + EVENT_TYPES).index(event["event_type"])
                        if event.get("event_type") in EVENT_TYPES
                        else 0
                    ),
                    key=f"ev_edit_type_{event_id}",
                )
                edit_status = st.selectbox(
                    "Status",
                    options=EVENT_STATUSES,
                    index=(
                        EVENT_STATUSES.index(event["status"])
                        if event.get("status") in EVENT_STATUSES
                        else 0
                    ),
                    format_func=_event_status_label,
                    key=f"ev_edit_status_{event_id}",
                )
            with col2:
                try:
                    default_date = (
                        datetime.fromisoformat(event["event_date"]).date()
                        if event.get("event_date")
                        else datetime.now().date()
                    )
                except Exception:
                    default_date = datetime.now().date()
                edit_date = st.date_input(
                    "Date", value=default_date,
                    key=f"ev_edit_date_{event_id}",
                )
                edit_start = st.text_input(
                    "Start Time (e.g. 19:00)",
                    value=event.get("start_time") or "",
                    key=f"ev_edit_start_{event_id}",
                )
                edit_end = st.text_input(
                    "End Time (e.g. 22:00)",
                    value=event.get("end_time") or "",
                    key=f"ev_edit_end_{event_id}",
                )

            edit_objective = st.text_area(
                "Objective", value=event.get("objective") or "",
                height=80, key=f"ev_edit_obj_{event_id}",
            )

            col_a, col_b = st.columns(2)
            with col_a:
                edit_headcount = st.number_input(
                    "Expected Headcount", min_value=0, step=5,
                    value=int(event.get("headcount") or 0),
                    key=f"ev_edit_headcount_{event_id}",
                )
                edit_budget = st.number_input(
                    "Budget Total", min_value=0.0, step=100.0,
                    value=float(event.get("budget_total") or 0),
                    key=f"ev_edit_budget_{event_id}",
                )
            with col_b:
                currency_options = ["USD", "EUR", "GBP", "NGN", "ZAR", "KES"]
                edit_currency = st.selectbox(
                    "Currency",
                    options=currency_options,
                    index=(
                        currency_options.index(event["currency"])
                        if event.get("currency") in currency_options
                        else 0
                    ),
                    key=f"ev_edit_currency_{event_id}",
                )
                edit_location = st.text_input(
                    "Location Override (if no venue)",
                    value=event.get("location_override") or "",
                    key=f"ev_edit_loc_{event_id}",
                )

            # Venue picker
            venue_options = db.get_venues(active_only=True)
            venue_labels = {"—": None}
            for v in venue_options:
                label = v["name"]
                if v.get("city"):
                    label += f" — {v['city']}"
                venue_labels[label] = v["id"]
            current_venue_label = "—"
            for lbl, vid in venue_labels.items():
                if vid == event.get("venue_id"):
                    current_venue_label = lbl
                    break
            edit_venue_label = st.selectbox(
                "Venue",
                options=list(venue_labels.keys()),
                index=list(venue_labels.keys()).index(current_venue_label),
                key=f"ev_edit_venue_{event_id}",
            )
            edit_venue_id = venue_labels[edit_venue_label]

            st.markdown("**🔗 Links**")
            edit_cal_url = st.text_input(
                "Calendar Event URL",
                value=event.get("calendar_event_url") or "",
                key=f"ev_edit_cal_{event_id}",
            )
            edit_rsvp_url = st.text_input(
                "Google Form RSVP URL",
                value=event.get("rsvp_form_url") or "",
                key=f"ev_edit_rsvp_{event_id}",
                help="The Form that guests fill out to RSVP.",
            )
            col_x, col_y = st.columns(2)
            with col_x:
                edit_run_url = st.text_input(
                    "Run of Show Doc URL",
                    value=event.get("run_of_show_url") or "",
                    key=f"ev_edit_run_{event_id}",
                )
            with col_y:
                edit_logistics_url = st.text_input(
                    "Logistics Doc URL",
                    value=event.get("logistics_url") or "",
                    key=f"ev_edit_log_url_{event_id}",
                )

            edit_notes = st.text_area(
                "Quick Notes",
                value=event.get("notes_native") or "",
                height=100, key=f"ev_edit_notes_{event_id}",
            )

            # Linked Trip picker
            trip_options = _trip_dropdown_options(exec_id)
            trip_option_keys = list(trip_options.keys())
            current_trip_label = "— No trip —"
            for lbl, tid in trip_options.items():
                if tid == event.get("linked_trip_id"):
                    current_trip_label = lbl
                    break
            edit_linked_trip_label = st.selectbox(
                "🔗 Linked Trip",
                options=trip_option_keys,
                index=(
                    trip_option_keys.index(current_trip_label)
                    if current_trip_label in trip_option_keys else 0
                ),
                key=f"ev_edit_linked_trip_{event_id}",
            )
            edit_linked_trip_id = trip_options[edit_linked_trip_label]

            col_save, col_cancel = st.columns(2)
            with col_save:
                save_clicked = st.form_submit_button("💾 Save")
            with col_cancel:
                cancel_clicked = st.form_submit_button("❌ Cancel")

            if cancel_clicked:
                st.rerun()

            if save_clicked:
                if not edit_name:
                    st.error("Name is required.")
                else:
                    db.update_event(
                        event_id,
                        name=edit_name,
                        event_type=edit_type or None,
                        status=edit_status,
                        event_date=(
                            edit_date.isoformat() if edit_date else None
                        ),
                        start_time=edit_start or None,
                        end_time=edit_end or None,
                        objective=edit_objective or None,
                        headcount=int(edit_headcount) or None,
                        budget_total=edit_budget or 0,
                        currency=edit_currency,
                        venue_id=edit_venue_id,
                        location_override=edit_location or None,
                        calendar_event_url=edit_cal_url or None,
                        rsvp_form_url=edit_rsvp_url or None,
                        run_of_show_url=edit_run_url or None,
                        logistics_url=edit_logistics_url or None,
                        notes_native=edit_notes or None,
                        linked_trip_id=edit_linked_trip_id,
                    )
                    
                    st.success("✅ Event updated.")
                    st.rerun()

    st.divider()

    # ================================================
    # SECTION: Guests
    # ================================================
    st.markdown("**👥 Guests**")
    counts = db.get_event_guest_status_counts(event_id)
    accepted = counts.get("accepted", 0)
    pending = counts.get("pending", 0) + counts.get("invited", 0)
    declined = counts.get("declined", 0)
    waitlist = counts.get("waitlist", 0)
    not_invited = counts.get("not_invited", 0)

    col_a, col_b, col_c, col_d, col_e = st.columns(5)
    with col_a:
        st.metric("Accepted", accepted)
    with col_b:
        st.metric("Pending", pending)
    with col_c:
        st.metric("Declined", declined)
    with col_d:
        st.metric("Waitlist", waitlist)
    with col_e:
        st.metric("Not Invited", not_invited)

    guests = db.get_event_guests(event_id)

    if guests:
        # Sort by status priority
        def _guest_sort(g):
            return (
                GUEST_STATUS_ORDER.index(g.get("invite_status") or "pending")
                if g.get("invite_status") in GUEST_STATUS_ORDER
                else 99,
                (g.get("guest_name") or g.get("contact_name") or "").lower(),
            )
        guests = sorted(guests, key=_guest_sort)

        for g in guests:
            display_name = (
                g.get("guest_name")
                or g.get("contact_name")
                or "—"
            )
            col1, col2, col3, col4 = st.columns([3, 2, 2, 0.5])
            with col1:
                st.markdown(f"**{display_name}**")
                email = g.get("guest_email") or g.get("contact_email")
                if email:
                    st.caption(email)
                if g.get("plus_ones"):
                    st.caption(f"+{g['plus_ones']} guest(s)")
            with col2:
                current_status = g.get("invite_status") or "pending"
                new_status = st.selectbox(
                    "Status",
                    options=GUEST_STATUSES,
                    index=(
                        GUEST_STATUSES.index(current_status)
                        if current_status in GUEST_STATUSES
                        else 0
                    ),
                    format_func=_guest_status_label,
                    key=f"ev_guest_status_{g['id']}",
                    label_visibility="collapsed",
                )
                if new_status != current_status:
                    db.update_event_guest(
                        g["id"], invite_status=new_status
                    )
                    st.rerun()
            with col3:
                if g.get("dietary"):
                    st.caption(f"🍽 {g['dietary']}")
                if g.get("accessibility"):
                    st.caption(f"♿ {g['accessibility']}")
            with col4:
                if st.button(
                    "🗑️", key=f"ev_guest_del_{g['id']}"
                ):
                    db.delete_event_guest(g["id"])
                    st.rerun()
    else:
        st.caption("No guests yet.")

    with st.expander("➕ Add Guest", expanded=False):
        with st.form(key=f"ev_add_guest_form_{event_id}"):
            contacts = db.get_contacts(active_only=True)
            source_options = ["Existing contact", "New guest"]
            source = st.radio(
                "Source", source_options, horizontal=True,
                key=f"ev_guest_source_{event_id}",
            )

            contact_id_val = None
            guest_name_val = None
            guest_email_val = None

            if source == "Existing contact":
                if contacts:
                    contact_options = {
                        f"{c['name']} ({c.get('role') or '—'})": c["id"]
                        for c in contacts
                    }
                    sel_label = st.selectbox(
                        "Contact*",
                        options=list(contact_options.keys()),
                        key=f"ev_guest_contact_{event_id}",
                    )
                    contact_id_val = contact_options[sel_label]
                else:
                    st.info(
                        "No contacts yet. Add one in the Contacts tab "
                        "or switch to 'New guest'."
                    )
            else:
                guest_name_val = st.text_input(
                    "Guest Name*", key=f"ev_guest_name_{event_id}"
                )
                guest_email_val = st.text_input(
                    "Guest Email", key=f"ev_guest_email_{event_id}"
                )

            col_a, col_b, col_c = st.columns(3)
            with col_a:
                initial_status = st.selectbox(
                    "Status",
                    options=GUEST_STATUSES,
                    index=GUEST_STATUSES.index("invited"),
                    format_func=_guest_status_label,
                    key=f"ev_guest_init_status_{event_id}",
                )
            with col_b:
                plus_ones = st.number_input(
                    "Plus Ones", min_value=0, step=1, value=0,
                    key=f"ev_guest_plus_{event_id}",
                )
            with col_c:
                badge_name = st.text_input(
                    "Badge Name", key=f"ev_guest_badge_{event_id}"
                )

            col_d, col_e = st.columns(2)
            with col_d:
                dietary = st.text_input(
                    "Dietary", key=f"ev_guest_diet_{event_id}"
                )
                accessibility = st.text_input(
                    "Accessibility",
                    key=f"ev_guest_access_{event_id}",
                )
            with col_e:
                allergy = st.text_input(
                    "Allergies", key=f"ev_guest_allergy_{event_id}"
                )
                seating = st.text_input(
                    "Seating Preference",
                    key=f"ev_guest_seating_{event_id}",
                )

            if st.form_submit_button("➕ Add Guest"):
                if source == "Existing contact":
                    if not contact_id_val:
                        st.error("Select a contact.")
                    else:
                        existing = [
                            g["contact_id"] for g in guests
                            if g.get("contact_id")
                        ]
                        if contact_id_val in existing:
                            st.warning(
                                "This contact is already a guest."
                            )
                        else:
                            db.add_event_guest(
                                event_id,
                                contact_id=contact_id_val,
                                invite_status=initial_status,
                                plus_ones=int(plus_ones),
                                badge_name=badge_name or None,
                                dietary=dietary or None,
                                accessibility=accessibility or None,
                                allergy_details=allergy or None,
                                seating_preference=seating or None,
                            )
                            st.rerun()
                else:
                    if not guest_name_val:
                        st.error("Guest name is required.")
                    else:
                        db.add_event_guest(
                            event_id,
                            guest_name=guest_name_val,
                            guest_email=guest_email_val or None,
                            invite_status=initial_status,
                            plus_ones=int(plus_ones),
                            badge_name=badge_name or None,
                            dietary=dietary or None,
                            accessibility=accessibility or None,
                            allergy_details=allergy or None,
                            seating_preference=seating or None,
                        )
                        st.rerun()
    st.divider()

    # ================================================
    # SECTION: Vendors
    # ================================================
    st.markdown("**🏢 Vendors**")
    engagements = db.get_vendor_engagements(event_id)

    if engagements:
        for eng in engagements:
            col1, col2, col3, col4 = st.columns([3, 2, 2, 0.5])
            with col1:
                st.markdown(f"**{eng['vendor_name']}**")
                if eng.get("category"):
                    st.caption(eng["category"])
                if eng.get("quote_amount") and eng.get("agreed_amount"):
                    st.caption(
                        f"Quote: {fmt_money(eng['quote_amount'], eng.get('currency') or 'USD')}  ·  "
                        f"Agreed: {fmt_money(eng['agreed_amount'], eng.get('currency') or 'USD')}"
                    )
                elif eng.get("agreed_amount"):
                    st.caption(
                        f"Agreed: {fmt_money(eng['agreed_amount'], eng.get('currency') or 'USD')}"
                    )
            with col2:
                current = eng.get("status") or "shortlisted"
                new_status = st.selectbox(
                    "Status",
                    options=VENDOR_ENGAGEMENT_STATUSES,
                    index=(
                        VENDOR_ENGAGEMENT_STATUSES.index(current)
                        if current in VENDOR_ENGAGEMENT_STATUSES else 0
                    ),
                    format_func=_vendor_status_label,
                    key=f"eng_status_list_{eng['id']}",
                    label_visibility="collapsed",
                )
                if new_status != current:
                    db.update_vendor_engagement(
                        eng["id"], status=new_status
                    )
                    if new_status == "confirmed":
                        db.recompute_budget_after_engagement_change(
                            eng["id"]
                        )
                    st.rerun()
            with col3:
                contract = eng.get("contract_status") or "not_sent"
                st.caption(_contract_status_label(contract))
                if eng.get("contract_url"):
                    st.markdown(f"[Open contract]({eng['contract_url']})")
            with col4:
                if st.button(
                    "📂", key=f"eng_open_{eng['id']}",
                    help="Open engagement detail / payments",
                ):
                    st.session_state["dialog_engagement_id"] = eng["id"]
    else:
        st.caption("No vendors engaged yet.")

    with st.expander("➕ Engage a Vendor", expanded=False):
        with st.form(key=f"eng_add_{event_id}"):
            vendors_list = db.get_vendors(active_only=True)
            if not vendors_list:
                st.info(
                    "No vendors in the directory yet. Add one in the "
                    "Vendors tab first."
                )
            else:
                vendor_options = {
                    f"{v['name']} ({v.get('category') or '—'})": v["id"]
                    for v in vendors_list
                }
                sel_label = st.selectbox(
                    "Vendor*",
                    options=list(vendor_options.keys()),
                    key=f"eng_add_vendor_{event_id}",
                )
                col_a, col_b = st.columns(2)
                with col_a:
                    eng_cat = st.selectbox(
                        "Category (for this engagement)",
                        options=VENDOR_CATEGORIES,
                        key=f"eng_add_cat_{event_id}",
                    )
                    eng_quote = st.number_input(
                        "Quote Amount", min_value=0.0, step=50.0,
                        value=0.0, key=f"eng_add_quote_{event_id}",
                    )
                with col_b:
                    eng_status = st.selectbox(
                        "Status",
                        options=VENDOR_ENGAGEMENT_STATUSES,
                        format_func=_vendor_status_label,
                        key=f"eng_add_status_{event_id}",
                    )
                    eng_agreed = st.number_input(
                        "Agreed Amount", min_value=0.0, step=50.0,
                        value=0.0, key=f"eng_add_agreed_{event_id}",
                    )
                eng_currency = st.text_input(
                    "Currency", value=event.get("currency") or "USD",
                    key=f"eng_add_currency_{event_id}",
                )
                eng_notes = st.text_area(
                    "Notes", key=f"eng_add_notes_{event_id}", height=60
                )
                if st.form_submit_button("➕ Engage"):
                    vendor_id = vendor_options[sel_label]
                    new_eng_id = db.add_vendor_engagement(
                        event_id=event_id,
                        vendor_id=vendor_id,
                        category=eng_cat,
                        status=eng_status,
                        quote_amount=eng_quote or None,
                        agreed_amount=eng_agreed or None,
                        currency=eng_currency,
                        notes=eng_notes or None,
                    )
                    if eng_status == "confirmed" and eng_agreed > 0:
                        db.recompute_budget_after_engagement_change(
                            new_eng_id
                        )
                    st.success("✅ Vendor engaged.")
                    st.rerun()

    st.divider()

    # ================================================
    # SECTION: Budget
    # ================================================
    st.markdown("**💰 Budget**")
    summary = db.get_budget_summary(event_id)
    budget_total = event.get("budget_total") or 0
    currency = event.get("currency") or "USD"

    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.metric(
            "Budget Total",
            fmt_money(budget_total, currency),
        )
    with col_b:
        st.metric(
            "Estimated",
            fmt_money(summary["estimated"], currency),
        )
    with col_c:
        st.metric(
            "Actual",
            fmt_money(summary["actual"], currency),
        )
    with col_d:
        st.metric(
            "Paid",
            fmt_money(summary["paid"], currency),
        )

    # Over-budget alert
    if budget_total > 0:
        ratio = summary["actual"] / budget_total
        if ratio >= 1:
            st.error(
                f"🔴 Over budget by "
                f"{fmt_money(summary['actual'] - budget_total, currency)}"
            )
        elif ratio >= 0.9:
            st.warning(
                f"⚠️ At {ratio*100:.0f}% of budget "
                f"({fmt_money(summary['actual'], currency)} of "
                f"{fmt_money(budget_total, currency)})"
            )

    items = db.get_budget_items(event_id)
    if items:
        for item in items:
            col1, col2, col3, col4 = st.columns([3, 2, 2, 0.5])
            with col1:
                st.markdown(
                    f"**{item.get('description') or '—'}**"
                )
                if item.get("category"):
                    st.caption(item["category"])
                if item.get("vendor_name"):
                    st.caption(f"🏢 {item['vendor_name']}")
            with col2:
                st.caption(
                    f"Est: {fmt_money(item.get('estimated_amount') or 0, currency)}"
                )
                st.caption(
                    f"Act: {fmt_money(item.get('actual_amount') or 0, currency)}"
                )
            with col3:
                st.caption(
                    f"Paid: {fmt_money(item.get('paid_amount') or 0, currency)}"
                )
                st.caption(
                    _budget_item_status_label(item.get("status"))
                )
            with col4:
                if st.button(
                    "🗑️", key=f"bi_del_{item['id']}",
                    help="Delete this line item",
                ):
                    db.delete_budget_item(item["id"])
                    st.rerun()
    else:
        st.caption("No budget line items yet.")

    # Apply template button
    with st.expander("➕ Add Line Item", expanded=False):
        if st.button(
            "📋 Apply a Template",
            key=f"apply_tpl_{event_id}",
        ):
            st.session_state[f"show_apply_tpl_{event_id}"] = True

        if st.session_state.get(f"show_apply_tpl_{event_id}", False):
            templates = db.get_budget_templates()
            if templates:
                tpl_options = {t["name"]: t["id"] for t in templates}
                tpl_sel = st.selectbox(
                    "Template",
                    options=list(tpl_options.keys()),
                    key=f"tpl_pick_{event_id}",
                )
                if st.button(
                    "Apply", key=f"tpl_apply_btn_{event_id}"
                ):
                    n = db.apply_budget_template_to_event(
                        event_id, tpl_options[tpl_sel]
                    )
                    st.success(f"✅ Added {n} line items.")
                    st.session_state.pop(
                        f"show_apply_tpl_{event_id}", None
                    )
                    st.rerun()
            else:
                st.caption("No templates available.")

        with st.form(key=f"bi_add_{event_id}"):
            col_a, col_b = st.columns(2)
            with col_a:
                bi_cat = st.selectbox(
                    "Category",
                    options=VENDOR_CATEGORIES,
                    key=f"bi_cat_{event_id}",
                )
                bi_est = st.number_input(
                    "Estimated", min_value=0.0, step=50.0, value=0.0,
                    key=f"bi_est_{event_id}",
                )
            with col_b:
                bi_desc = st.text_input(
                    "Description", key=f"bi_desc_{event_id}"
                )
                bi_act = st.number_input(
                    "Actual", min_value=0.0, step=50.0, value=0.0,
                    key=f"bi_act_{event_id}",
                )
            bi_paid = st.number_input(
                "Paid", min_value=0.0, step=50.0, value=0.0,
                key=f"bi_paid_{event_id}",
            )
            bi_notes = st.text_area(
                "Notes", key=f"bi_notes_{event_id}", height=60
            )
            if st.form_submit_button("➕ Add Line Item"):
                if not bi_desc:
                    st.error("Description is required.")
                else:
                    db.add_budget_item(
                        event_id=event_id,
                        category=bi_cat,
                        description=bi_desc,
                        estimated_amount=bi_est,
                        actual_amount=bi_act,
                        paid_amount=bi_paid,
                        status="estimated",
                    )
                    st.rerun()
                    
    st.divider()

    # ================================================
    # SECTION: Delete Event
    # ================================================
    if st.button(
        "🗑️ Delete This Event", type="primary",
        key=f"ev_del_btn_{event_id}",
    ):
        st.session_state[f"confirm_del_ev_{event_id}"] = True

    if st.session_state.get(f"confirm_del_ev_{event_id}", False):
        st.warning(f"⚠️ Permanently delete event '{event['name']}'?")
        col_yes, col_no = st.columns(2)
        with col_yes:
            if st.button(
                "✅ Yes, Delete", key=f"ev_del_yes_{event_id}"
            ):
                db.delete_event(event_id)
                st.session_state.pop(f"confirm_del_ev_{event_id}", None)
                st.session_state.pop("dialog_event_id", None)
                st.rerun()
        with col_no:
            if st.button(
                "❌ Cancel", key=f"ev_del_no_{event_id}"
            ):
                st.session_state.pop(f"confirm_del_ev_{event_id}", None)
                st.rerun()

# =========================================================
# RSVP IMPORT DIALOG (Google Forms CSV)
# =========================================================


def _clear_dialog_rsvp_import():
    st.session_state.pop("dialog_rsvp_import_event_id", None)


@st.dialog(
    "📥 Import RSVPs from Google Forms CSV",
    width="large",
    on_dismiss=_clear_dialog_rsvp_import,
)
def _render_rsvp_import_dialog(event_id):
    event = db.get_event(event_id)
    if not event:
        st.warning("Event not found.")
        return

    st.subheader(f"📥 Import RSVPs — {event['name']}")
    st.caption(
        "Export responses from the Google Sheet linked to your Form, "
        "then upload the CSV here."
    )

    uploaded = st.file_uploader(
        "CSV file",
        type=["csv"],
        key=f"rsvp_csv_{event_id}",
    )
    if not uploaded:
        st.info("Choose a CSV to continue.")
        return

    import csv as _csv
    import io as _io

    content = uploaded.getvalue().decode("utf-8-sig")
    reader = _csv.DictReader(_io.StringIO(content))
    rows = list(reader)
    if not rows:
        st.error("CSV is empty.")
        return

    headers = reader.fieldnames or []
    st.write(f"**{len(rows)} row(s)** detected")
    st.caption(f"Columns: {', '.join(headers)}")

    # ---- Column mapping ----
    st.markdown("**Column Mapping**")

    def _find_header(candidates):
        lowered = {h.lower().strip(): h for h in headers}
        for c in candidates:
            if c in lowered:
                return lowered[c]
        return None

    auto_email = _find_header([
        "email", "email address", "e-mail", "guest email",
        "your email", "email address*",
    ])
    auto_name = _find_header([
        "name", "full name", "guest name", "your name",
        "attendee name",
    ])
    auto_attend = _find_header([
        "attending", "will you attend", "rsvp", "will you attend?",
        "attending?", "coming?",
    ])
    auto_plus = _find_header([
        "plus ones", "guests", "number of guests",
        "additional guests", "plus ones (including you)",
    ])
    auto_diet = _find_header([
        "dietary", "dietary requirements", "dietary restrictions",
        "food restrictions",
    ])
    auto_access = _find_header([
        "accessibility", "accessibility needs", "accessibility requirements",
    ])
    auto_ts = _find_header(["timestamp", "submitted at", "date", "time"])

    def _select_column(label, auto_value, optional=True):
        options = ["—"] + list(headers)
        default_idx = 0
        if auto_value and auto_value in options:
            default_idx = options.index(auto_value)
        sel = st.selectbox(
            label,
            options=options,
            index=default_idx,
            key=f"rsvp_map_{event_id}_{label}",
        )
        if sel == "—":
            if not optional:
                st.warning(f"{label} is required.")
            return None
        return sel

    col1, col2 = st.columns(2)
    with col1:
        col_email = _select_column("Email", auto_email, optional=False)
        col_attend = _select_column("Attending", auto_attend, optional=False)
        col_diet = _select_column("Dietary", auto_diet)
    with col2:
        col_name = _select_column("Name", auto_name)
        col_plus = _select_column("Plus Ones", auto_plus)
        col_access = _select_column("Accessibility", auto_access)

    col_ts = _select_column(
        "Timestamp (optional)", auto_ts, optional=True
    )

    # ---- Status mapping ----
    st.markdown("**Status Mapping**")
    st.caption(
        "Map the values in your Attending column to app statuses."
    )
    attend_values = set()
    if col_attend:
        for r in rows:
            val = (r.get(col_attend) or "").strip().lower()
            if val:
                attend_values.add(val)
    attend_values = sorted(attend_values)

    status_map = {}
    for val in attend_values:
        default_status = "accepted"
        if val in ("yes", "y", "true", "attending", "accept", "accepted"):
            default_status = "accepted"
        elif val in ("no", "n", "false", "decline", "declined"):
            default_status = "declined"
        elif val in ("maybe", "tentative", "perhaps"):
            default_status = "pending"
        status_map[val] = st.selectbox(
            f"'{val}' →",
            options=GUEST_STATUSES,
            index=(
                GUEST_STATUSES.index(default_status)
                if default_status in GUEST_STATUSES
                else 0
            ),
            format_func=_guest_status_label,
            key=f"rsvp_map_status_{event_id}_{val}",
        )

    # ---- Preview ----
    st.markdown("**Preview**")
    matched = 0
    new = 0
    ambiguous = 0
    skipped = 0
    errors = []
    preview_rows = []

    for i, r in enumerate(rows, start=2):
        email = (r.get(col_email) or "").strip() if col_email else ""
        name = (r.get(col_name) or "").strip() if col_name else ""
        attend_val = (
            (r.get(col_attend) or "").strip().lower()
            if col_attend else ""
        )

        if not email and not name:
            skipped += 1
            continue
        if not attend_val or attend_val not in status_map:
            errors.append(f"Row {i}: no status for '{attend_val}'")
            continue

        target_status = status_map[attend_val]

        matched_guest = None
        if email:
            matched_guest = db.find_event_guest_by_email(event_id, email)
        if not matched_guest and name:
            matched_guest = db.find_event_guest_by_name(event_id, name)

        if matched_guest:
            matched += 1
            action = "UPDATE"
        else:
            new += 1
            action = "NEW"

        preview_rows.append({
            "row": i,
            "action": action,
            "name": name or "—",
            "email": email or "—",
            "status": _guest_status_label(target_status),
        })

    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.metric("Matched", matched)
    with col_b:
        st.metric("New", new)
    with col_c:
        st.metric("Skipped", skipped)
    with col_d:
        st.metric("Errors", len(errors))

    if preview_rows:
        st.dataframe(
            preview_rows[:20],
            use_container_width=True,
            hide_index=True,
        )
        if len(preview_rows) > 20:
            st.caption(f"… and {len(preview_rows) - 20} more rows")

    if errors:
        with st.expander(f"⚠️ {len(errors)} error(s)"):
            for e in errors:
                st.caption(f"• {e}")

    # ---- Import ----
    st.divider()
    col_do, col_cancel = st.columns(2)
    with col_do:
        if st.button(
            "📥 Import RSVPs",
            type="primary",
            disabled=(matched + new == 0),
            key=f"rsvp_import_btn_{event_id}",
        ):
            created = 0
            updated = 0
            for r in rows:
                email = (r.get(col_email) or "").strip() if col_email else ""
                name = (r.get(col_name) or "").strip() if col_name else ""
                attend_val = (
                    (r.get(col_attend) or "").strip().lower()
                    if col_attend else ""
                )
                if not attend_val or attend_val not in status_map:
                    continue
                if not email and not name:
                    continue

                target_status = status_map[attend_val]
                plus_ones = 0
                if col_plus:
                    try:
                        plus_ones = int(
                            (r.get(col_plus) or "0").strip() or 0
                        )
                    except Exception:
                        plus_ones = 0
                dietary = (
                    (r.get(col_diet) or "").strip() if col_diet else ""
                )
                accessibility = (
                    (r.get(col_access) or "").strip()
                    if col_access else ""
                )
                ts = (r.get(col_ts) or "").strip() if col_ts else None

                matched_guest = None
                if email:
                    matched_guest = db.find_event_guest_by_email(
                        event_id, email
                    )
                if not matched_guest and name:
                    matched_guest = db.find_event_guest_by_name(
                        event_id, name
                    )

                if matched_guest:
                    db.update_event_guest(
                        matched_guest["id"],
                        invite_status=target_status,
                        plus_ones=plus_ones,
                        dietary=dietary or None,
                        accessibility=accessibility or None,
                        rsvp_timestamp=ts,
                        rsvp_source="google_form",
                    )
                    updated += 1
                else:
                    db.add_event_guest(
                        event_id,
                        guest_name=name or None,
                        guest_email=email or None,
                        invite_status=target_status,
                        plus_ones=plus_ones,
                        dietary=dietary or None,
                        accessibility=accessibility or None,
                        rsvp_timestamp=ts,
                        rsvp_source="google_form",
                    )
                    created += 1

            st.success(
                f"✅ Import complete — "
                f"{updated} updated, {created} created."
            )
            st.session_state.pop(f"rsvp_csv_{event_id}", None)
            st.rerun()

    with col_cancel:
        if st.button(
            "Cancel", key=f"rsvp_import_cancel_{event_id}"
        ):
            st.session_state.pop("dialog_rsvp_import_event_id", None)
            st.rerun()

# =========================================================
# VENDOR DETAIL DIALOG
# =========================================================


def _clear_dialog_vendor_id():
    st.session_state.pop("dialog_vendor_id", None)


@st.dialog(
    "🏢 Vendor",
    width="large",
    on_dismiss=_clear_dialog_vendor_id,
)
def _render_vendor_detail_dialog(vendor_id):
    vendor = db.get_vendor(vendor_id)
    if not vendor:
        st.warning("Vendor not found.")
        return

    st.subheader(f"🏢 {vendor['name']}")
    if vendor.get("category"):
        st.caption(vendor["category"])

    with st.form(key=f"vendor_edit_{vendor_id}"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input(
                "Name*", value=vendor.get("name") or "",
                key=f"ve_name_{vendor_id}",
            )
            category = st.selectbox(
                "Category",
                options=[""] + VENDOR_CATEGORIES,
                index=(
                    ([""] + VENDOR_CATEGORIES).index(vendor["category"])
                    if vendor.get("category") in VENDOR_CATEGORIES
                    else 0
                ),
                key=f"ve_cat_{vendor_id}",
            )
            email = st.text_input(
                "Email", value=vendor.get("email") or "",
                key=f"ve_email_{vendor_id}",
            )
            phone = st.text_input(
                "Phone", value=vendor.get("phone") or "",
                key=f"ve_phone_{vendor_id}",
            )
        with col2:
            website = st.text_input(
                "Website", value=vendor.get("website") or "",
                key=f"ve_website_{vendor_id}",
            )
            city = st.text_input(
                "City", value=vendor.get("city") or "",
                key=f"ve_city_{vendor_id}",
            )
            country = st.text_input(
                "Country", value=vendor.get("country") or "",
                key=f"ve_country_{vendor_id}",
            )
            rating_val = st.number_input(
                "Rating (1–5)",
                min_value=0, max_value=5, step=1,
                value=int(vendor.get("rating") or 0),
                key=f"ve_rating_{vendor_id}",
            )

        address = st.text_input(
            "Address", value=vendor.get("address") or "",
            key=f"ve_address_{vendor_id}",
        )
        notes = st.text_area(
            "Notes", value=vendor.get("notes") or "", height=80,
            key=f"ve_notes_{vendor_id}",
        )
        col_a, col_b = st.columns(2)
        with col_a:
            preferred = st.checkbox(
                "⭐ Preferred",
                value=bool(vendor.get("preferred")),
                key=f"ve_pref_{vendor_id}",
            )
        with col_b:
            is_active = st.checkbox(
                "Active",
                value=bool(vendor.get("is_active", 1)),
                key=f"ve_active_{vendor_id}",
            )

        col_save, col_del = st.columns(2)
        with col_save:
            save_clicked = st.form_submit_button("💾 Save")
        with col_del:
            delete_clicked = st.form_submit_button(
                "🗑️ Delete Vendor", type="secondary"
            )

        if save_clicked:
            if not name:
                st.error("Name is required.")
            else:
                db.update_vendor(
                    vendor_id,
                    name=name,
                    category=category or None,
                    email=email or None,
                    phone=phone or None,
                    website=website or None,
                    city=city or None,
                    country=country or None,
                    address=address or None,
                    notes=notes or None,
                    rating=int(rating_val) if rating_val else None,
                    preferred=1 if preferred else 0,
                    is_active=1 if is_active else 0,
                )
                st.success("✅ Saved.")
                st.rerun()

        if delete_clicked:
            st.session_state[f"confirm_del_vendor_{vendor_id}"] = True

    if st.session_state.get(f"confirm_del_vendor_{vendor_id}", False):
        st.warning("⚠️ Delete this vendor? Engagements referencing it will also be removed.")
        col_y, col_n = st.columns(2)
        with col_y:
            if st.button("Yes, Delete", key=f"vd_yes_{vendor_id}"):
                db.delete_vendor(vendor_id)
                st.session_state.pop(f"confirm_del_vendor_{vendor_id}", None)
                st.session_state.pop("dialog_vendor_id", None)
                st.rerun()
        with col_n:
            if st.button("Cancel", key=f"vd_no_{vendor_id}"):
                st.session_state.pop(f"confirm_del_vendor_{vendor_id}", None)
                st.rerun()

    # ---- History: engagements this vendor appears in ----
    st.divider()
    st.markdown("**📜 Engagement History**")
    conn = sqlite3.connect(db.DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ve.*, ev.name AS event_name, ev.event_date
               FROM vendor_engagements ve
               JOIN events ev ON ve.event_id = ev.id
               WHERE ve.vendor_id = ?
               ORDER BY ev.event_date DESC""",
            (vendor_id,),
        )
        rows = [dict(r) for r in c.fetchall()]
    finally:
        conn.close()

    if not rows:
        st.caption("No engagements yet.")
    else:
        total_spent = 0.0
        for r in rows:
            st.markdown(
                f"**{r['event_name']}** — "
                f"{_vendor_status_label(r.get('status'))}"
            )
            agreed = r.get("agreed_amount") or 0
            total_spent += agreed
            st.caption(
                f"📅 {fmt_date(r.get('event_date'))}  ·  "
                f"Agreed: {fmt_money(agreed, r.get('currency') or 'USD')}"
            )
            st.divider()
        st.caption(f"**Total agreed across all events:** "
                   f"{fmt_money(total_spent, 'USD')}")

    if st.button("Close", key=f"close_vendor_{vendor_id}"):
        st.session_state.pop("dialog_vendor_id", None)
        st.rerun()


# =========================================================
# VENDOR ENGAGEMENT DIALOG
# =========================================================


def _clear_dialog_engagement_id():
    st.session_state.pop("dialog_engagement_id", None)


@st.dialog(
    "🤝 Vendor Engagement",
    width="large",
    on_dismiss=_clear_dialog_engagement_id,
)
def _render_engagement_detail_dialog(engagement_id):
    eng = db.get_vendor_engagement(engagement_id)
    if not eng:
        st.warning("Engagement not found.")
        return

    st.subheader(f"🤝 {eng['vendor_name']}")
    st.caption(
        f"Status: {_vendor_status_label(eng.get('status'))}  ·  "
        f"Contract: {_contract_status_label(eng.get('contract_status'))}"
    )

    # Edit engagement
    with st.expander("✏️ Edit Engagement", expanded=False):
        with st.form(key=f"eng_edit_{engagement_id}"):
            col1, col2 = st.columns(2)
            with col1:
                status = st.selectbox(
                    "Status",
                    options=VENDOR_ENGAGEMENT_STATUSES,
                    index=(
                        VENDOR_ENGAGEMENT_STATUSES.index(eng["status"])
                        if eng.get("status") in VENDOR_ENGAGEMENT_STATUSES
                        else 0
                    ),
                    format_func=_vendor_status_label,
                    key=f"eng_status_{engagement_id}",
                )
                quote_amount = st.number_input(
                    "Quote Amount", min_value=0.0, step=50.0,
                    value=float(eng.get("quote_amount") or 0),
                    key=f"eng_quote_{engagement_id}",
                )
                agreed_amount = st.number_input(
                    "Agreed Amount", min_value=0.0, step=50.0,
                    value=float(eng.get("agreed_amount") or 0),
                    key=f"eng_agreed_{engagement_id}",
                )
            with col2:
                currency = st.text_input(
                    "Currency", value=eng.get("currency") or "USD",
                    key=f"eng_cur_{engagement_id}",
                )
                contract_status = st.selectbox(
                    "Contract Status",
                    options=CONTRACT_STATUSES,
                    index=(
                        CONTRACT_STATUSES.index(eng["contract_status"])
                        if eng.get("contract_status") in CONTRACT_STATUSES
                        else 0
                    ),
                    format_func=_contract_status_label,
                    key=f"eng_contract_status_{engagement_id}",
                )
                contract_url = st.text_input(
                    "Contract URL",
                    value=eng.get("contract_url") or "",
                    key=f"eng_contract_url_{engagement_id}",
                )

            payment_terms = st.text_input(
                "Payment Terms",
                value=eng.get("payment_terms") or "",
                key=f"eng_terms_{engagement_id}",
            )
            notes = st.text_area(
                "Notes", value=eng.get("notes") or "", height=80,
                key=f"eng_notes_{engagement_id}",
            )

            if st.form_submit_button("💾 Save"):
                db.update_vendor_engagement(
                    engagement_id,
                    status=status,
                    quote_amount=quote_amount or None,
                    agreed_amount=agreed_amount or None,
                    currency=currency,
                    contract_status=contract_status,
                    contract_url=contract_url or None,
                    payment_terms=payment_terms or None,
                    notes=notes or None,
                )
                # If confirmed, sync to budget
                if status == "confirmed" and agreed_amount > 0:
                    db.recompute_budget_after_engagement_change(
                        engagement_id
                    )
                st.success("✅ Engagement updated.")
                st.rerun()

    st.divider()

    # Payments section
    st.markdown("**💳 Payments**")
    payments = db.get_vendor_payments(engagement_id)
    if payments:
        for p in payments:
            col1, col2, col3, col4 = st.columns([3, 2, 2, 0.5])
            with col1:
                status_icon = "✅" if p.get("paid_date") else "⏳"
                st.markdown(
                    f"{status_icon} **{fmt_money(p.get('amount') or 0, p.get('currency') or 'USD')}**"
                )
                if p.get("method"):
                    st.caption(_payment_method_label(p.get("method")))
            with col2:
                st.caption(f"Due: {p.get('due_date') or '—'}")
            with col3:
                st.caption(f"Paid: {p.get('paid_date') or '—'}")
            with col4:
                if st.button("🗑️", key=f"pay_del_{p['id']}"):
                    db.delete_vendor_payment(p["id"])
                    st.rerun()
    else:
        st.caption("No payments recorded.")

    with st.expander("➕ Add Payment", expanded=False):
        with st.form(key=f"add_pay_{engagement_id}"):
            col1, col2 = st.columns(2)
            with col1:
                amount = st.number_input(
                    "Amount", min_value=0.0, step=50.0, value=0.0,
                    key=f"pay_amount_{engagement_id}",
                )
                currency = st.text_input(
                    "Currency", value=eng.get("currency") or "USD",
                    key=f"pay_currency_{engagement_id}",
                )
                due_date_input = st.date_input(
                    "Due Date", value=None,
                    key=f"pay_due_{engagement_id}",
                )
            with col2:
                paid_date_input = st.date_input(
                    "Paid Date (leave blank if not yet paid)",
                    value=None, key=f"pay_paid_{engagement_id}",
                )
                method = st.selectbox(
                    "Method",
                    options=[""] + PAYMENT_METHODS,
                    format_func=lambda m: (
                        "—" if m == "" else _payment_method_label(m)
                    ),
                    key=f"pay_method_{engagement_id}",
                )
                reference = st.text_input(
                    "Reference (invoice #, etc.)",
                    key=f"pay_ref_{engagement_id}",
                )
            notes = st.text_area(
                "Notes", key=f"pay_notes_{engagement_id}", height=60
            )
            if st.form_submit_button("➕ Add Payment"):
                if amount <= 0:
                    st.error("Amount must be greater than zero.")
                else:
                    db.add_vendor_payment(
                        vendor_engagement_id=engagement_id,
                        amount=amount,
                        currency=currency,
                        due_date=(
                            due_date_input.isoformat()
                            if due_date_input else None
                        ),
                        paid_date=(
                            paid_date_input.isoformat()
                            if paid_date_input else None
                        ),
                        method=method or None,
                        reference=reference or None,
                        notes=notes or None,
                    )
                    st.success("✅ Payment added.")
                    st.rerun()

    # Summary
    total_paid = sum(
        float(p.get("amount") or 0)
        for p in payments if p.get("paid_date")
    )
    total_due = sum(
        float(p.get("amount") or 0)
        for p in payments if not p.get("paid_date")
    )
    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("Paid", fmt_money(total_paid, eng.get("currency") or "USD"))
    with col_b:
        st.metric("Due", fmt_money(total_due, eng.get("currency") or "USD"))

    st.divider()

    if st.button("Close", key=f"eng_close_{engagement_id}"):
        st.session_state.pop("dialog_engagement_id", None)
        st.rerun()

# =========================================================
# SPEAKING ENGAGEMENT DETAIL DIALOG
# =========================================================


def _clear_dialog_speaking_id():
    st.session_state.pop("dialog_speaking_id", None)


@st.dialog(
    "🎤 Speaking Engagement",
    width="large",
    on_dismiss=_clear_dialog_speaking_id,
)
def _render_speaking_detail_dialog(engagement_id):
    eng = db.get_speaking_engagement(engagement_id)
    if not eng:
        st.warning("Engagement not found.")
        return

    st.subheader(f"🎤 {eng['title']}")
    st.caption(
        f"**{eng.get('executive_name') or '—'}**  ·  "
        f"{_speaking_status_label(eng.get('status'))}  ·  "
        f"📅 {fmt_date(eng.get('engagement_date'))}"
    )

    # ================================================
    # SECTION: Edit Engagement
    # ================================================
    with st.expander("✏️ Edit Engagement", expanded=False):
        with st.form(key=f"speak_edit_{engagement_id}"):
            col1, col2 = st.columns(2)
            with col1:
                edit_title = st.text_input(
                    "Title*", value=eng.get("title") or "",
                    key=f"speak_edit_title_{engagement_id}",
                )
                edit_type = st.selectbox(
                    "Type",
                    options=[""] + SPEAKING_TYPES,
                    index=(
                        ([""] + SPEAKING_TYPES).index(
                            eng["engagement_type"]
                        )
                        if eng.get("engagement_type") in SPEAKING_TYPES
                        else 0
                    ),
                    key=f"speak_edit_type_{engagement_id}",
                )
                edit_format = st.selectbox(
                    "Format",
                    options=[""] + SPEAKING_FORMATS,
                    index=(
                        ([""] + SPEAKING_FORMATS).index(eng["format"])
                        if eng.get("format") in SPEAKING_FORMATS
                        else 0
                    ),
                    key=f"speak_edit_format_{engagement_id}",
                )
                edit_status = st.selectbox(
                    "Status",
                    options=SPEAKING_STATUSES,
                    index=(
                        SPEAKING_STATUSES.index(eng["status"])
                        if eng.get("status") in SPEAKING_STATUSES
                        else 0
                    ),
                    format_func=_speaking_status_label,
                    key=f"speak_edit_status_{engagement_id}",
                )
            with col2:
                try:
                    default_date = (
                        datetime.fromisoformat(
                            eng["engagement_date"]
                        ).date()
                        if eng.get("engagement_date")
                        else datetime.now().date()
                    )
                except Exception:
                    default_date = datetime.now().date()
                edit_date = st.date_input(
                    "Date", value=default_date,
                    key=f"speak_edit_date_{engagement_id}",
                )
                edit_time = st.text_input(
                    "Time (e.g. 14:00)",
                    value=eng.get("engagement_time") or "",
                    key=f"speak_edit_time_{engagement_id}",
                )
                edit_duration = st.text_input(
                    "Duration (e.g. 45 min)",
                    value=eng.get("duration") or "",
                    key=f"speak_edit_dur_{engagement_id}",
                )

            # Organizer
            st.markdown("**Organizer**")
            col_a, col_b = st.columns(2)
            with col_a:
                edit_organizer_name = st.text_input(
                    "Organizer Name",
                    value=eng.get("organizer_name") or "",
                    key=f"speak_edit_org_name_{engagement_id}",
                )
            with col_b:
                contacts = db.get_contacts(active_only=True)
                contact_options = {"—": None}
                contact_options.update(
                    {c["name"]: c["id"] for c in contacts}
                )
                current_org = "—"
                for lbl, cid in contact_options.items():
                    if cid == eng.get("organizer_contact_id"):
                        current_org = lbl
                        break
                edit_organizer_label = st.selectbox(
                    "Organizer Contact",
                    options=list(contact_options.keys()),
                    index=list(contact_options.keys()).index(current_org),
                    key=f"speak_edit_org_contact_{engagement_id}",
                )
                edit_organizer_contact_id = contact_options[
                    edit_organizer_label
                ]

            st.markdown("**Location & Audience**")
            col_c, col_d = st.columns(2)
            with col_c:
                edit_location = st.text_input(
                    "Location",
                    value=eng.get("location") or "",
                    key=f"speak_edit_loc_{engagement_id}",
                )
                edit_timezone = st.text_input(
                    "Timezone (for virtual)",
                    value=eng.get("timezone") or "",
                    key=f"speak_edit_tz_{engagement_id}",
                )
            with col_d:
                edit_audience = st.text_input(
                    "Audience",
                    value=eng.get("audience") or "",
                    key=f"speak_edit_audience_{engagement_id}",
                )
                edit_topic = st.text_input(
                    "Topic",
                    value=eng.get("topic") or "",
                    key=f"speak_edit_topic_{engagement_id}",
                )

            st.markdown("**Fee & Terms**")
            col_e, col_f = st.columns(2)
            with col_e:
                edit_fee_offered = st.number_input(
                    "Fee Offered", min_value=0.0, step=500.0,
                    value=float(eng.get("fee_offered") or 0),
                    key=f"speak_edit_fee_offered_{engagement_id}",
                )
                edit_fee_agreed = st.number_input(
                    "Fee Agreed", min_value=0.0, step=500.0,
                    value=float(eng.get("fee_agreed") or 0),
                    key=f"speak_edit_fee_agreed_{engagement_id}",
                )
            with col_f:
                edit_currency = st.text_input(
                    "Currency", value=eng.get("currency") or "USD",
                    key=f"speak_edit_currency_{engagement_id}",
                )
                edit_recording_rights = st.selectbox(
                    "Recording Rights",
                    options=[""] + RECORDING_RIGHTS_OPTIONS,
                    index=(
                        ([""] + RECORDING_RIGHTS_OPTIONS).index(
                            eng["recording_rights"]
                        )
                        if eng.get("recording_rights")
                            in RECORDING_RIGHTS_OPTIONS
                        else 0
                    ),
                    key=f"speak_edit_rights_{engagement_id}",
                )

            col_g, col_h = st.columns(2)
            with col_g:
                edit_travel_covered = st.checkbox(
                    "Travel Covered",
                    value=bool(eng.get("travel_covered")),
                    key=f"speak_edit_travel_{engagement_id}",
                )
            with col_h:
                edit_accom_covered = st.checkbox(
                    "Accommodation Covered",
                    value=bool(eng.get("accommodation_covered")),
                    key=f"speak_edit_accom_{engagement_id}",
                )

            edit_exclusivity = st.text_input(
                "Exclusivity",
                value=eng.get("exclusivity") or "",
                key=f"speak_edit_excl_{engagement_id}",
            )

            st.markdown("**🔗 Links**")
            edit_cal_url = st.text_input(
                "Calendar Event URL",
                value=eng.get("calendar_event_url") or "",
                key=f"speak_edit_cal_{engagement_id}",
            )
            edit_notes_url = st.text_input(
                "Notes Doc URL",
                value=eng.get("notes_url") or "",
                key=f"speak_edit_notes_url_{engagement_id}",
            )
            edit_notes_native = st.text_area(
                "Quick Notes",
                value=eng.get("notes_native") or "",
                height=80,
                key=f"speak_edit_notes_{engagement_id}",
            )

            col_save, col_cancel = st.columns(2)
            with col_save:
                save_clicked = st.form_submit_button("💾 Save")
            with col_cancel:
                cancel_clicked = st.form_submit_button("❌ Cancel")

            if cancel_clicked:
                st.rerun()

            if save_clicked:
                if not edit_title:
                    st.error("Title is required.")
                else:
                    db.update_speaking_engagement(
                        engagement_id,
                        title=edit_title,
                        engagement_type=edit_type or None,
                        format=edit_format or None,
                        status=edit_status,
                        engagement_date=(
                            edit_date.isoformat() if edit_date else None
                        ),
                        engagement_time=edit_time or None,
                        duration=edit_duration or None,
                        organizer_name=(
                            edit_organizer_name or None
                        ),
                        organizer_contact_id=(
                            edit_organizer_contact_id
                        ),
                        location=edit_location or None,
                        timezone=edit_timezone or None,
                        audience=edit_audience or None,
                        topic=edit_topic or None,
                        fee_offered=(
                            edit_fee_offered
                            if edit_fee_offered > 0 else None
                        ),
                        fee_agreed=(
                            edit_fee_agreed
                            if edit_fee_agreed > 0 else None
                        ),
                        currency=edit_currency,
                        recording_rights=(
                            edit_recording_rights or None
                        ),
                        travel_covered=(
                            1 if edit_travel_covered else 0
                        ),
                        accommodation_covered=(
                            1 if edit_accom_covered else 0
                        ),
                        exclusivity=edit_exclusivity or None,
                        calendar_event_url=edit_cal_url or None,
                        notes_url=edit_notes_url or None,
                        notes_native=edit_notes_native or None,
                        linked_trip_id=edit_linked_trip_id,
                    )
                    st.success("✅ Engagement updated.")
                    st.rerun()

    st.divider()

    # ================================================
    # SECTION: Deliverables
    # ================================================
    st.markdown("**📦 Deliverables**")
    deliverables = db.get_speaking_deliverables(engagement_id)

    if not deliverables:
        st.caption("No deliverables yet.")
        if st.button(
            "📋 Add 5 Default Deliverables",
            key=f"speak_seed_deliv_{engagement_id}",
        ):
            n = db.seed_default_deliverables(engagement_id)
            st.success(f"✅ Added {n} deliverables.")
            st.rerun()
    else:
        for d in deliverables:
            col1, col2, col3, col4 = st.columns([3, 2, 2, 0.5])
            with col1:
                st.markdown(f"**{d['name']}**")
                bits = []
                if d.get("owner"):
                    bits.append(f"👤 {d['owner']}")
                if d.get("due_date"):
                    bits.append(f"📅 {fmt_date(d['due_date'])}")
                if bits:
                    st.caption(" · ".join(bits))
                if d.get("file_url"):
                    st.markdown(f"[Open file]({d['file_url']})")
                if d.get("submitted_to_organizer"):
                    st.caption("📤 Submitted to organizer")
                if d.get("organizer_feedback"):
                    st.caption(f"💬 {d['organizer_feedback']}")
            with col2:
                current = d.get("status") or "not_started"
                new_status = st.selectbox(
                    "Status",
                    options=SPEAKING_DELIVERABLE_STATUSES,
                    index=(
                        SPEAKING_DELIVERABLE_STATUSES.index(current)
                        if current in SPEAKING_DELIVERABLE_STATUSES
                        else 0
                    ),
                    format_func=_deliverable_status_label,
                    key=f"deliv_status_{d['id']}",
                    label_visibility="collapsed",
                )
                if new_status != current:
                    db.update_speaking_deliverable(
                        d["id"], status=new_status
                    )
                    st.rerun()
            with col3:
                if d.get("due_date"):
                    today_iso = datetime.now().date().isoformat()
                    if (
                        d["due_date"] < today_iso
                        and d.get("status") != "done"
                    ):
                        st.caption("🔴 Overdue")
            with col4:
                if st.button("🗑️", key=f"deliv_del_{d['id']}"):
                    db.delete_speaking_deliverable(d["id"])
                    st.rerun()

    with st.expander("➕ Add Deliverable", expanded=False):
        with st.form(key=f"add_deliv_{engagement_id}"):
            col1, col2 = st.columns(2)
            with col1:
                d_name = st.text_input(
                    "Name*", key=f"new_deliv_name_{engagement_id}"
                )
                d_owner = st.text_input(
                    "Owner", key=f"new_deliv_owner_{engagement_id}"
                )
            with col2:
                d_due = st.date_input(
                    "Due Date", value=None,
                    key=f"new_deliv_due_{engagement_id}",
                )
                d_url = st.text_input(
                    "File URL", key=f"new_deliv_url_{engagement_id}"
                )
            if st.form_submit_button("➕ Add"):
                if not d_name:
                    st.error("Name is required.")
                else:
                    db.add_speaking_deliverable(
                        engagement_id,
                        name=d_name,
                        owner=d_owner or None,
                        due_date=(
                            d_due.isoformat() if d_due else None
                        ),
                        file_url=d_url or None,
                    )
                    st.rerun()

    st.divider()

    # ================================================
    # SECTION: CFP Submissions
    # ================================================
    st.markdown("**📤 CFP Submissions**")
    cfps = db.get_speaking_cfps(engagement_id)

    if cfps:
        for cf in cfps:
            col1, col2, col3 = st.columns([4, 2, 0.5])
            with col1:
                st.markdown(
                    f"**{cf.get('proposal_title') or 'Untitled'}**  ·  "
                    f"{_cfp_status_label(cf.get('status'))}"
                )
                if cf.get("submission_deadline"):
                    st.caption(
                        f"Deadline: {fmt_date(cf['submission_deadline'])}"
                    )
                if cf.get("submission_date"):
                    st.caption(
                        f"Submitted: {fmt_date(cf['submission_date'])}"
                    )
                if cf.get("cfp_url"):
                    st.markdown(f"[CFP page]({cf['cfp_url']})")
                if cf.get("proposal_doc_url"):
                    st.markdown(f"[Proposal doc]({cf['proposal_doc_url']})")
            with col2:
                current = cf.get("status") or "draft"
                new_status = st.selectbox(
                    "Status",
                    options=SPEAKING_CFP_STATUSES,
                    index=(
                        SPEAKING_CFP_STATUSES.index(current)
                        if current in SPEAKING_CFP_STATUSES else 0
                    ),
                    format_func=_cfp_status_label,
                    key=f"cfp_status_{cf['id']}",
                    label_visibility="collapsed",
                )
                if new_status != current:
                    db.update_speaking_cfp(cf["id"], status=new_status)
                    st.rerun()
            with col3:
                if st.button("🗑️", key=f"cfp_del_{cf['id']}"):
                    db.delete_speaking_cfp(cf["id"])
                    st.rerun()
    else:
        st.caption("No CFP submissions yet.")

    with st.expander("➕ Add CFP Submission", expanded=False):
        with st.form(key=f"add_cfp_{engagement_id}"):
            col1, col2 = st.columns(2)
            with col1:
                cfp_url = st.text_input(
                    "CFP URL", key=f"new_cfp_url_{engagement_id}"
                )
                deadline = st.date_input(
                    "Submission Deadline", value=None,
                    key=f"new_cfp_deadline_{engagement_id}",
                )
                prop_title = st.text_input(
                    "Proposal Title",
                    key=f"new_cfp_title_{engagement_id}",
                )
            with col2:
                submitted_date = st.date_input(
                    "Submitted On (leave blank if not yet)",
                    value=None,
                    key=f"new_cfp_submitted_{engagement_id}",
                )
                prop_doc = st.text_input(
                    "Proposal Doc URL",
                    key=f"new_cfp_doc_{engagement_id}",
                )
                initial_status = st.selectbox(
                    "Status",
                    options=SPEAKING_CFP_STATUSES,
                    format_func=_cfp_status_label,
                    key=f"new_cfp_status_{engagement_id}",
                )
            abstract = st.text_area(
                "Abstract",
                key=f"new_cfp_abstract_{engagement_id}",
                height=100,
            )
            if st.form_submit_button("➕ Add"):
                db.add_speaking_cfp(
                    engagement_id,
                    cfp_url=cfp_url or None,
                    submission_deadline=(
                        deadline.isoformat() if deadline else None
                    ),
                    submission_date=(
                        submitted_date.isoformat()
                        if submitted_date else None
                    ),
                    proposal_title=prop_title or None,
                    abstract=abstract or None,
                    proposal_doc_url=prop_doc or None,
                    status=initial_status,
                )
                st.rerun()

    st.divider()

    # ================================================
    # SECTION: Logistics
    # ================================================
    with st.expander("✈️ Logistics", expanded=False):
        with st.form(key=f"speak_logistics_{engagement_id}"):
            col1, col2 = st.columns(2)
            with col1:
                log_travel_required = st.checkbox(
                    "Travel Required",
                    value=bool(eng.get("travel_covered")),
                    key=f"log_travel_{engagement_id}",
                )
                log_flight_links = st.text_area(
                    "Flight links / confirmations",
                    key=f"log_flights_{engagement_id}", height=60,
                )
                log_accom = st.text_input(
                    "Accommodation",
                    key=f"log_accom_{engagement_id}",
                )
                log_ground = st.text_input(
                    "Ground transport",
                    key=f"log_ground_{engagement_id}",
                )
            with col2:
                log_venue = st.text_input(
                    "Venue address",
                    key=f"log_venue_{engagement_id}",
                )
                log_green_room = st.text_input(
                    "Green room",
                    key=f"log_green_{engagement_id}",
                )
                log_dress = st.text_input(
                    "Dress code",
                    key=f"log_dress_{engagement_id}",
                )
                log_onsite = st.text_input(
                    "On-site contact",
                    key=f"log_onsite_{engagement_id}",
                )
            log_rehearsal = st.date_input(
                "Rehearsal Date", value=None,
                key=f"log_rehearsal_{engagement_id}",
            )
            log_tech = st.date_input(
                "Tech Check Date", value=None,
                key=f"log_tech_{engagement_id}",
            )
            log_slides_due = st.date_input(
                "Slides Due to Organizer", value=None,
                key=f"log_slides_{engagement_id}",
            )
            if st.form_submit_button("💾 Save Logistics"):
                # Store logistics in the notes field for now —
                # a dedicated table would be Phase 8+ if needed
                logistics_text = (
                    f"LOGISTICS\n"
                    f"Travel required: {log_travel_required}\n"
                    f"Flight links: {log_flight_links}\n"
                    f"Accommodation: {log_accom}\n"
                    f"Ground transport: {log_ground}\n"
                    f"Venue: {log_venue}\n"
                    f"Green room: {log_green_room}\n"
                    f"Dress code: {log_dress}\n"
                    f"On-site contact: {log_onsite}\n"
                    f"Rehearsal: {log_rehearsal}\n"
                    f"Tech check: {log_tech}\n"
                    f"Slides due: {log_slides_due}"
                )
                db.update_speaking_engagement(
                    engagement_id,
                    notes_native=logistics_text,
                )
                st.success("✅ Logistics saved to Quick Notes.")
                st.rerun()

    st.divider()

    # ================================================
    # SECTION: Post-Engagement
    # ================================================
    with st.expander("📮 Post-Engagement", expanded=False):
        with st.form(key=f"speak_post_{engagement_id}"):
            col1, col2 = st.columns(2)
            with col1:
                thankyou = st.checkbox(
                    "Thank-you sent",
                    value=bool(eng.get("thank_you_sent")),
                    key=f"post_thankyou_{engagement_id}",
                )
                try:
                    default_thanks_date = (
                        datetime.fromisoformat(
                            eng["thank_you_date"]
                        ).date()
                        if eng.get("thank_you_date")
                        else None
                    )
                except Exception:
                    default_thanks_date = None
                thanks_date = st.date_input(
                    "Thank-you date",
                    value=default_thanks_date,
                    key=f"post_thanks_date_{engagement_id}",
                )
                feedback = st.text_area(
                    "Feedback received",
                    value=eng.get("feedback_received") or "",
                    height=80,
                    key=f"post_feedback_{engagement_id}",
                )
            with col2:
                rating = st.number_input(
                    "Rating (1–5)",
                    min_value=0, max_value=5, step=1,
                    value=int(eng.get("feedback_rating") or 0),
                    key=f"post_rating_{engagement_id}",
                )
                testimonial = st.text_area(
                    "Testimonial",
                    value=eng.get("testimonial") or "",
                    height=80,
                    key=f"post_testimonial_{engagement_id}",
                )
                testimonial_perm = st.checkbox(
                    "Permission to use testimonial",
                    value=bool(eng.get("testimonial_permission")),
                    key=f"post_perm_{engagement_id}",
                )

            col_c, col_d = st.columns(2)
            with col_c:
                recording_recv = st.checkbox(
                    "Recording received",
                    value=bool(eng.get("recording_received")),
                    key=f"post_recv_{engagement_id}",
                )
                recording_url = st.text_input(
                    "Recording URL",
                    value=eng.get("recording_url") or "",
                    key=f"post_url_{engagement_id}",
                )
            with col_d:
                would_repeat = st.selectbox(
                    "Would repeat?",
                    options=["", "Yes", "No", "Maybe"],
                    index=(
                        ["", "Yes", "No", "Maybe"].index(
                            {
                                1: "Yes", 0: "No", 2: "Maybe"
                            }.get(eng.get("would_repeat"), "")
                        )
                        if eng.get("would_repeat") is not None
                        else 0
                    ),
                    key=f"post_repeat_{engagement_id}",
                )

            if st.form_submit_button("💾 Save Post-Engagement"):
                repeat_map = {"Yes": 1, "No": 0, "Maybe": 2, "": None}
                db.update_speaking_engagement(
                    engagement_id,
                    thank_you_sent=1 if thankyou else 0,
                    thank_you_date=(
                        thanks_date.isoformat() if thanks_date else None
                    ),
                    feedback_received=feedback or None,
                    feedback_rating=(
                        int(rating) if rating else None
                    ),
                    testimonial=testimonial or None,
                    testimonial_permission=(
                        1 if testimonial_perm else 0
                    ),
                    recording_received=1 if recording_recv else 0,
                    recording_url=recording_url or None,
                    would_repeat=repeat_map.get(would_repeat),
                )
                st.success("✅ Post-engagement saved.")
                st.rerun()

    st.divider()

    # ================================================
    # SECTION: Delete
    # ================================================
    if st.button(
        "🗑️ Delete This Engagement", type="primary",
        key=f"speak_del_btn_{engagement_id}",
    ):
        st.session_state[f"confirm_del_speak_{engagement_id}"] = True

    if st.session_state.get(f"confirm_del_speak_{engagement_id}", False):
        st.warning(f"⚠️ Permanently delete '{eng['title']}'?")
        col_y, col_n = st.columns(2)
        with col_y:
            if st.button(
                "Yes, Delete", key=f"speak_del_yes_{engagement_id}"
            ):
                db.delete_speaking_engagement(engagement_id)
                st.session_state.pop(
                    f"confirm_del_speak_{engagement_id}", None
                )
                st.session_state.pop("dialog_speaking_id", None)
                st.rerun()
        with col_n:
            if st.button(
                "Cancel", key=f"speak_del_no_{engagement_id}"
            ):
                st.session_state.pop(
                    f"confirm_del_speak_{engagement_id}", None
                )
                st.rerun()

def _render_meeting_detail_dialog(meeting_id):
    meeting = db.get_meeting(meeting_id)
    if not meeting:
        st.warning("Meeting not found.")
        return

    exec_id = meeting["exec_id"]
    st.subheader(f"📅 {meeting['title']}")
    st.caption(
        f"**{meeting.get('executive_name') or '—'}** · "
        f"{_meeting_status_label(meeting.get('status') or 'draft')}"
    )

    # ================================================
    # SECTION: Edit Meeting
    # ================================================
    with st.expander("✏️ Edit Meeting", expanded=False):
        with st.form(key=f"mtg_edit_form_{meeting_id}"):
            col1, col2 = st.columns(2)
            with col1:
                edit_title = st.text_input(
                    "Title*",
                    value=meeting.get("title") or "",
                    key=f"mtg_edit_title_{meeting_id}",
                )
                edit_type = st.selectbox(
                    "Type",
                    options=[""] + MEETING_TYPES,
                    index=(
                        ([""] + MEETING_TYPES).index(meeting["meeting_type"])
                        if meeting.get("meeting_type") in MEETING_TYPES
                        else 0
                    ),
                    key=f"mtg_edit_type_{meeting_id}",
                )
                edit_status = st.selectbox(
                    "Status",
                    options=MEETING_STATUSES,
                    index=(
                        MEETING_STATUSES.index(meeting["status"])
                        if meeting.get("status") in MEETING_STATUSES
                        else 0
                    ),
                    format_func=_meeting_status_label,
                    key=f"mtg_edit_status_{meeting_id}",
                )
            with col2:
                try:
                    default_date = (
                        datetime.fromisoformat(meeting["meeting_date"]).date()
                        if meeting.get("meeting_date")
                        else datetime.now().date()
                    )
                except Exception:
                    default_date = datetime.now().date()
                edit_date = st.date_input(
                    "Date",
                    value=default_date,
                    key=f"mtg_edit_date_{meeting_id}",
                )
                edit_time = st.text_input(
                    "Time (e.g. 14:30)",
                    value=meeting.get("meeting_time") or "",
                    key=f"mtg_edit_time_{meeting_id}",
                )
                edit_duration = st.number_input(
                    "Duration (minutes)",
                    min_value=0,
                    step=15,
                    value=int(meeting.get("duration_minutes") or 0),
                    key=f"mtg_edit_dur_{meeting_id}",
                )

            edit_location = st.text_input(
                "Location",
                value=meeting.get("location") or "",
                key=f"mtg_edit_loc_{meeting_id}",
            )
            edit_video = st.text_input(
                "Video Link",
                value=meeting.get("video_link") or "",
                key=f"mtg_edit_video_{meeting_id}",
            )
            edit_cal_url = st.text_input(
                "🔗 Calendar Event URL (Google Calendar / Outlook)",
                value=meeting.get("calendar_event_url") or "",
                key=f"mtg_edit_cal_{meeting_id}",
                help="Paste the link to the calendar event. Opens in a new tab.",
            )
            edit_agenda = st.text_area(
                "Agenda",
                value=meeting.get("agenda") or "",
                height=100,
                key=f"mtg_edit_agenda_{meeting_id}",
            )
            edit_notes_native = st.text_area(
                "Quick Notes (plain text)",
                value=meeting.get("notes_native") or "",
                height=120,
                key=f"mtg_edit_notes_{meeting_id}",
            )

            col_a, col_b = st.columns(2)
            with col_a:
                edit_notes_url = st.text_input(
                    "🔗 Notes Doc URL",
                    value=meeting.get("notes_url") or "",
                    key=f"mtg_edit_notes_url_{meeting_id}",
                )
            with col_b:
                edit_notes_label = st.text_input(
                    "Notes Label",
                    value=meeting.get("notes_label") or "",
                    placeholder="e.g. Board meeting notes",
                    key=f"mtg_edit_notes_label_{meeting_id}",
                )

            try:
                default_follow = (
                    datetime.fromisoformat(meeting["follow_up_date"]).date()
                    if meeting.get("follow_up_date")
                    else None
                )
            except Exception:
                default_follow = None
            edit_followup = st.date_input(
                "Follow-up Date",
                value=default_follow,
                key=f"mtg_edit_followup_{meeting_id}",
            )

            # Linked Trip picker
            trip_options = _trip_dropdown_options(exec_id)
            trip_option_keys = list(trip_options.keys())
            current_trip_label = "— No trip —"
            for lbl, tid in trip_options.items():
                if tid == meeting.get("linked_trip_id"):
                    current_trip_label = lbl
                    break
            edit_linked_trip_label = st.selectbox(
                "🔗 Linked Trip",
                options=trip_option_keys,
                index=(
                    trip_option_keys.index(current_trip_label)
                    if current_trip_label in trip_option_keys else 0
                ),
                key=f"mtg_edit_linked_trip_{meeting_id}",
            )
            edit_linked_trip_id = trip_options[edit_linked_trip_label]

            col_save, col_cancel = st.columns(2)
            with col_save:
                save_clicked = st.form_submit_button("💾 Save")
            with col_cancel:
                cancel_clicked = st.form_submit_button("❌ Cancel")

            if cancel_clicked:
                st.rerun()

            if save_clicked:
                if not edit_title:
                    st.error("Title is required.")
                else:
                    old_status = meeting.get("status") or "draft"
                    status_is_transitioning_to_completed = (
                        edit_status == "completed" and old_status != "completed"
                    )

                    db.update_meeting(
                        meeting_id,
                        title=edit_title,
                        meeting_type=edit_type or None,
                        status=edit_status,
                        meeting_date=edit_date.isoformat() if edit_date else None,
                        meeting_time=edit_time or None,
                        duration_minutes=int(edit_duration) or None,
                        location=edit_location or None,
                        video_link=edit_video or None,
                        calendar_event_url=edit_cal_url or None,
                        agenda=edit_agenda or None,
                        notes_native=edit_notes_native or None,
                        notes_url=edit_notes_url or None,
                        notes_label=edit_notes_label or None,
                        follow_up_date=(
                            edit_followup.isoformat() if edit_followup else None
                        ),
                        linked_trip_id=edit_linked_trip_id,
                    )

                    # Auto-log attendance if the meeting just completed
                    if status_is_transitioning_to_completed:
                        logged = db.auto_log_meeting_completion(
                            meeting_id, exec_id
                        )
                        if logged > 0:
                            st.success(
                                f"✅ Meeting updated. "
                                f"Logged {logged} attendee communication(s)."
                            )
                        else:
                            st.success("✅ Meeting updated.")
                    else:
                        st.success("✅ Meeting updated.")

                    st.rerun()

    st.divider()

    # ================================================
    # SECTION: Attendees
    # ================================================
    st.markdown("**👥 Attendees**")
    attendees = db.get_meeting_attendees(meeting_id)

    if attendees:
        for att in attendees:
            col1, col2, col3, col4 = st.columns([3, 2, 2, 0.5])
            with col1:
                vip = " ⭐" if att.get("is_vip") else ""
                internal = " · Internal" if att.get("is_internal") else " · External"
                st.markdown(f"**{att['contact_name'] or '—'}**{vip}")
                if att.get("role"):
                    st.caption(att["role"])
                st.caption(internal)
            with col2:
                current_status = att.get("status") or "pending"
                new_status = st.selectbox(
                    "Status",
                    options=ATTENDEE_STATUSES,
                    index=(
                        ATTENDEE_STATUSES.index(current_status)
                        if current_status in ATTENDEE_STATUSES
                        else 0
                    ),
                    key=f"mtg_att_status_{att['id']}",
                    label_visibility="collapsed",
                )
                if new_status != current_status:
                    db.update_meeting_attendee(att["id"], status=new_status)
                    st.rerun()
            with col3:
                if att.get("response_notes"):
                    st.caption(att["response_notes"])
            with col4:
                if st.button("🗑️", key=f"mtg_att_del_{att['id']}"):
                    db.remove_meeting_attendee(att["id"])
                    st.rerun()
    else:
        st.caption("No attendees yet.")

    with st.expander("➕ Add Attendee", expanded=False):
        with st.form(key=f"mtg_add_att_form_{meeting_id}"):
            contacts = db.get_contacts(active_only=True)
            if not contacts:
                st.info("No contacts yet. Add contacts in the Contacts tab.")
            else:
                contact_options = {
                    f"{c['name']} ({c.get('role') or '—'})": c["id"] for c in contacts
                }
                selected_label = st.selectbox(
                    "Contact*",
                    options=list(contact_options.keys()),
                    key=f"mtg_add_att_contact_{meeting_id}",
                )
                col_a, col_b, col_c = st.columns(3)
                with col_a:
                    is_internal = st.checkbox(
                        "Internal",
                        key=f"mtg_add_att_internal_{meeting_id}",
                    )
                with col_b:
                    is_vip = st.checkbox("VIP", key=f"mtg_add_att_vip_{meeting_id}")
                with col_c:
                    att_status = st.selectbox(
                        "Status",
                        options=ATTENDEE_STATUSES,
                        key=f"mtg_add_att_status_{meeting_id}",
                    )
                role = st.text_input(
                    "Role at meeting",
                    key=f"mtg_add_att_role_{meeting_id}",
                )
                if st.form_submit_button("➕ Add"):
                    contact_id = contact_options[selected_label]
                    # Prevent exact duplicate
                    existing_ids = [a["contact_id"] for a in attendees]
                    if contact_id in existing_ids:
                        st.warning("This contact is already an attendee.")
                    else:
                        db.add_meeting_attendee(
                            meeting_id,
                            contact_id,
                            is_internal=1 if is_internal else 0,
                            is_vip=1 if is_vip else 0,
                            role=role or None,
                            status=att_status,
                        )
                        st.rerun()

    st.divider()

    # ================================================
    # SECTION: Action Items
    # ================================================
    st.markdown("**✅ Action Items**")
    action_items = db.get_meeting_action_items(meeting_id)

    if action_items:
        for item in action_items:
            col1, col2, col3, col4 = st.columns([4, 2, 2, 0.5])
            with col1:
                st.markdown(f"**{item['description']}**")
                owner = (
                    item.get("owner_contact_name") or item.get("owner_exec_name") or "—"
                )
                due = item.get("due_date") or "—"
                st.caption(f"Owner: {owner} · Due: {due}")
            with col2:
                current_status = item.get("status") or "not_started"
                new_status = st.selectbox(
                    "Status",
                    options=ACTION_ITEM_STATUSES,
                    index=(
                        ACTION_ITEM_STATUSES.index(current_status)
                        if current_status in ACTION_ITEM_STATUSES
                        else 0
                    ),
                    format_func=_action_status_label,
                    key=f"mtg_act_status_{item['id']}",
                    label_visibility="collapsed",
                )
                if new_status != current_status:
                    db.update_meeting_action_item(item["id"], status=new_status)
                    st.rerun()
            with col3:
                st.caption("")
            with col4:
                if st.button("🗑️", key=f"mtg_act_del_{item['id']}"):
                    db.delete_meeting_action_item(item["id"])
                    st.rerun()
    else:
        st.caption("No action items yet.")

    with st.expander("➕ Add Action Item", expanded=False):
        with st.form(key=f"mtg_add_action_form_{meeting_id}"):
            description = st.text_input(
                "Description*", key=f"mtg_add_action_desc_{meeting_id}"
            )
            contacts = db.get_contacts(active_only=True)
            owner_type = st.radio(
                "Owner type",
                options=["Contact", "Executive", "None"],
                horizontal=True,
                key=f"mtg_add_action_owner_type_{meeting_id}",
            )
            owner_contact_id = None
            owner_exec_id = None
            if owner_type == "Contact" and contacts:
                contact_options = {f"{c['name']}": c["id"] for c in contacts}
                selected_owner_label = st.selectbox(
                    "Owner contact",
                    options=list(contact_options.keys()),
                    key=f"mtg_add_action_owner_{meeting_id}",
                )
                owner_contact_id = contact_options[selected_owner_label]
            elif owner_type == "Executive":
                execs = db.get_all_executives(active_only=False)
                if execs:
                    exec_options = {e["name"]: e["id"] for e in execs}
                    selected_exec_label = st.selectbox(
                        "Owner executive",
                        options=list(exec_options.keys()),
                        key=f"mtg_add_action_owner_exec_{meeting_id}",
                    )
                    owner_exec_id = exec_options[selected_exec_label]

            col_a, col_b = st.columns(2)
            with col_a:
                due_date_input = st.date_input(
                    "Due Date",
                    value=None,
                    key=f"mtg_add_action_due_{meeting_id}",
                )
            with col_b:
                initial_status = st.selectbox(
                    "Status",
                    options=ACTION_ITEM_STATUSES,
                    format_func=_action_status_label,
                    key=f"mtg_add_action_status_{meeting_id}",
                )

            if st.form_submit_button("➕ Add"):
                if not description:
                    st.error("Description is required.")
                else:
                    db.add_meeting_action_item(
                        meeting_id,
                        description,
                        owner_contact_id=owner_contact_id,
                        owner_exec_id=owner_exec_id,
                        due_date=(
                            due_date_input.isoformat() if due_date_input else None
                        ),
                        status=initial_status,
                    )
                    st.rerun()

    st.divider()

    # ================================================
    # SECTION: Decisions
    # ================================================
    st.markdown("**📌 Decisions**")
    decisions = db.get_meeting_decisions(meeting_id)

    if decisions:
        for dec in decisions:
            col1, col2 = st.columns([5, 0.5])
            with col1:
                st.markdown(f"• {dec['decision']}")
                if dec.get("made_by"):
                    st.caption(f"Made by: {dec['made_by']}")
            with col2:
                if st.button("🗑️", key=f"mtg_dec_del_{dec['id']}"):
                    db.delete_meeting_decision(dec["id"])
                    st.rerun()
    else:
        st.caption("No decisions logged yet.")

    with st.form(key=f"mtg_add_decision_form_{meeting_id}"):
        col_a, col_b = st.columns([3, 2])
        with col_a:
            decision_text = st.text_input(
                "Decision", key=f"mtg_add_dec_text_{meeting_id}"
            )
        with col_b:
            made_by = st.text_input("Made by", key=f"mtg_add_dec_by_{meeting_id}")
        if st.form_submit_button("➕ Log Decision"):
            if decision_text:
                db.add_meeting_decision(meeting_id, decision_text, made_by or None)
                st.rerun()

    st.divider()

    # ================================================
    # SECTION: Danger Zone
    # ================================================
    if st.button(
        "🗑️ Delete This Meeting", type="primary", key=f"mtg_del_btn_{meeting_id}"
    ):
        st.session_state[f"confirm_del_mtg_{meeting_id}"] = True

    if st.session_state.get(f"confirm_del_mtg_{meeting_id}", False):
        st.warning(f"⚠️ Permanently delete meeting '{meeting['title']}'?")
        col_yes, col_no = st.columns(2)
        with col_yes:
            if st.button("✅ Yes, Delete", key=f"mtg_del_yes_{meeting_id}"):
                db.delete_meeting(meeting_id)
                st.session_state.pop(f"confirm_del_mtg_{meeting_id}", None)
                st.session_state.pop("dialog_meeting_id", None)
                st.rerun()
        with col_no:
            if st.button("❌ Cancel", key=f"mtg_del_no_{meeting_id}"):
                st.session_state.pop(f"confirm_del_mtg_{meeting_id}", None)
                st.rerun()


# =========================================================
# TABS
# =========================================================

# Module-level country list — used by both Tab 1 and the trip edit modal
country_list = sorted([c.name for c in pycountry.countries])

tab_names = [
    "🏠 Dashboard",
    "📅 Meetings",
    "🎉 Events",
    "🎤 Speaking",
    "✈️ Trip Planner",
    "✈️ All Trips",
    "👥 Contacts",
    "🏢 Companies",
    "🏢 Venues",
    "🏢 Vendors",
    "📚 Library",
]
(
    tab_dashboard,
    tab_meetings,
    tab_events,
    tab_speaking,
    tab_trip_planner,
    tab_all_trips,
    tab_contacts,
    tab_companies,
    tab_venues,
    tab_vendors,
    tab_library,
) = st.tabs(tab_names)


# =========================================================
# TAB: MEETINGS
# =========================================================

with tab_meetings:
    st.header("📅 Meetings")

    # ---- Filters row ----
    col1, col2, col3, col4 = st.columns([2, 1.5, 1.8, 2.5])
    with col1:
        exec_filter_options = ["All Executives"] + [
            f"{e['name']} (ID: {e['id']})"
            for e in db.get_all_executives(active_only=False)
        ]
        exec_filter_label = st.selectbox(
            "Executive", exec_filter_options, key="mtg_filter_exec"
        )
        exec_filter_id = (
            None
            if exec_filter_label == "All Executives"
            else int(exec_filter_label.split("(ID: ")[1].rstrip(")"))
        )
    with col2:
        status_filter = st.selectbox(
            "Status",
            ["Active", "All"] + MEETING_STATUSES,
            format_func=lambda s: {
                "Active": "Active",
                "All": "All",
            }.get(s, _meeting_status_label(s)),
            key="mtg_filter_status",
        )
    with col3:
        view_filter = st.selectbox(
            "View",
            [
                "All meetings",
                "Awaiting response",
                "Without notes",
            ],
            key="mtg_filter_view",
        )
    with col4:
        search_query = st.text_input(
            "Search", placeholder="Title or agenda...",
            key="mtg_search",
        )

    # ---- New button ----
    col_new, _sp = st.columns([1.2, 5])
    with col_new:
        if st.button("➕ New Meeting", type="primary",
                     key="mtg_new_btn"):
            st.session_state["mtg_show_create"] = True

    # ---- Create form ----
    if st.session_state.get("mtg_show_create", False):
        with st.expander("➕ New Meeting", expanded=True):
            execs_active = db.get_all_executives(active_only=True)
            if not execs_active:
                st.warning(
                    "No active executives. Add one in the sidebar first."
                )
            else:
                with st.form("mtg_create_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        exec_options = {
                            e["name"]: e["id"] for e in execs_active
                        }
                        new_exec_label = st.selectbox(
                            "Executive*",
                            options=list(exec_options.keys()),
                            key="mtg_new_exec",
                        )
                        new_title = st.text_input(
                            "Title*", key="mtg_new_title"
                        )
                        new_type = st.selectbox(
                            "Type",
                            options=[""] + MEETING_TYPES,
                            key="mtg_new_type",
                        )
                        new_status = st.selectbox(
                            "Status",
                            options=MEETING_STATUSES,
                            format_func=_meeting_status_label,
                            key="mtg_new_status",
                        )
                    with col2:
                        new_date = st.date_input(
                            "Date", value=datetime.now().date(),
                            key="mtg_new_date",
                        )
                        new_time = st.text_input(
                            "Time (e.g. 14:30)", key="mtg_new_time"
                        )
                        new_duration = st.number_input(
                            "Duration (minutes)", min_value=0, step=15,
                            value=30, key="mtg_new_dur",
                        )

                    new_location = st.text_input(
                        "Location", key="mtg_new_loc"
                    )
                    new_video = st.text_input(
                        "Video Link", key="mtg_new_video"
                    )
                    new_cal_url = st.text_input(
                        "🔗 Calendar Event URL (Google Calendar / Outlook)",
                        key="mtg_new_cal",
                    )
                    new_agenda = st.text_area(
                        "Agenda", height=100, key="mtg_new_agenda"
                    )
                    new_notes_native = st.text_area(
                        "Quick Notes (plain text)",
                        height=100, key="mtg_new_notes",
                    )
                    col_a, col_b = st.columns(2)
                    with col_a:
                        new_notes_url = st.text_input(
                            "🔗 Notes Doc URL", key="mtg_new_notes_url"
                        )
                    with col_b:
                        new_notes_label = st.text_input(
                            "Notes Label",
                            placeholder="e.g. Board meeting notes",
                            key="mtg_new_notes_label",
                        )
                    new_followup = st.date_input(
                        "Follow-up Date", value=None,
                        key="mtg_new_followup",
                    )

                    col_submit, col_cancel = st.columns(2)
                    with col_submit:
                        submitted = st.form_submit_button("💾 Create")
                    with col_cancel:
                        cancelled = st.form_submit_button("❌ Cancel")

                    if cancelled:
                        st.session_state["mtg_show_create"] = False
                        st.rerun()

                    if submitted:
                        if not new_title:
                            st.error("Title is required.")
                        else:
                            new_exec_id = exec_options[new_exec_label]
                            # Duplicate check
                            dupes = db.find_duplicate_meetings(
                                new_exec_id, new_title,
                                new_date.isoformat() if new_date else None,
                            )
                            if dupes:
                                st.warning(
                                    "⚠️ A meeting with the same title and "
                                    "date already exists for this executive."
                                )
                                for d in dupes:
                                    st.caption(
                                        f"- {d['title']} "
                                        f"({d.get('meeting_date')})"
                                    )
                                st.info(
                                    "Change the title or date, or edit the "
                                    "existing meeting instead."
                                )
                            else:
                                db.add_meeting(
                                    new_exec_id,
                                    new_title,
                                    meeting_type=new_type or None,
                                    status=new_status,
                                    meeting_date=(
                                        new_date.isoformat()
                                        if new_date else None
                                    ),
                                    meeting_time=new_time or None,
                                    duration_minutes=(
                                        int(new_duration)
                                        if new_duration else None
                                    ),
                                    location=new_location or None,
                                    video_link=new_video or None,
                                    calendar_event_url=new_cal_url or None,
                                    agenda=new_agenda or None,
                                    notes_native=new_notes_native or None,
                                    notes_url=new_notes_url or None,
                                    notes_label=new_notes_label or None,
                                    follow_up_date=(
                                        new_followup.isoformat()
                                        if new_followup else None
                                    ),
                                )
                                st.session_state["mtg_show_create"] = False
                                st.success("✅ Meeting created.")
                                st.rerun()

    st.divider()

    # ---- Build query filters ----
    status_param = None
    if status_filter == "Active":
        status_param = [
            "draft", "request_sent", "pending", "confirmed",
            "tentative", "rescheduled",
        ]
    elif status_filter != "All":
        status_param = status_filter

    without_notes = view_filter == "Without notes"
    awaiting = view_filter == "Awaiting response"

    meetings = db.get_meetings(
        exec_id=exec_filter_id,
        status=status_param,
        without_notes=without_notes,
        awaiting_response=awaiting,
        search=search_query or None,
    )

    # ---- List header ----
    if not meetings:
        st.info("No meetings match these filters.")
    else:
        st.write(f"**{len(meetings)} meeting(s)**")

        # ---- Render each meeting as a compact row ----
        for m in meetings:
            with st.container():
                col1, col2, col3 = st.columns([6, 1, 1])
                with col1:
                    status_label = _meeting_status_label(
                        m.get("status") or "draft"
                    )
                    st.markdown(
                        f"### {m['title']}  ·  {status_label}"
                    )
                    st.caption(
                        f"👤 {m.get('executive_name') or '—'}  ·  "
                        f"📅 {fmt_date(m.get('meeting_date'))}"
                        + (
                            f"  ·  🕐 {m['meeting_time']}"
                            if m.get("meeting_time") else ""
                        )
                    )
                    # Attendee + action item counts (lightweight counts)
                    atts = db.get_meeting_attendees(m["id"])
                    acts = db.get_meeting_action_items(m["id"])
                    notes_flag = (
                        "📝 Notes linked"
                        if (m.get("notes_url") or m.get("notes_native"))
                        else "⚠️ No notes"
                    )
                    st.caption(
                        f"👥 {len(atts)} attendee(s)  ·  "
                        f"✅ {len(acts)} action item(s)  ·  {notes_flag}"
                    )
                with col2:
                    if st.button(
                        "📂 Open",
                        key=f"mtg_open_{m['id']}",
                        use_container_width=True,
                    ):
                        st.session_state["dialog_meeting_id"] = m["id"]
                with col3:
                    if st.button(
                        "🗑️",
                        key=f"mtg_list_del_{m['id']}",
                    ):
                        st.session_state[
                            f"confirm_list_del_mtg_{m['id']}"
                        ] = True

                # ---- Delete confirmation (inline) ----
                if st.session_state.get(
                    f"confirm_list_del_mtg_{m['id']}", False
                ):
                    st.warning(
                        f"⚠️ Permanently delete '{m['title']}'?"
                    )
                    col_yes, col_no = st.columns(2)
                    with col_yes:
                        if st.button(
                            "✅ Yes, Delete",
                            key=f"mtg_list_del_yes_{m['id']}",
                        ):
                            db.delete_meeting(m["id"])
                            st.session_state.pop(
                                f"confirm_list_del_mtg_{m['id']}", None
                            )
                            st.rerun()
                    with col_no:
                        if st.button(
                            "❌ Cancel",
                            key=f"mtg_list_del_no_{m['id']}",
                        ):
                            st.session_state.pop(
                                f"confirm_list_del_mtg_{m['id']}", None
                            )
                            st.rerun()

                st.divider()


# =========================================================
# TAB: EVENTS
# =========================================================

with tab_events:
    st.header("🎉 Events")

    # ---- Filters ----
    col1, col2, col3, col4 = st.columns([2, 1.6, 2, 2.4])
    with col1:
        exec_filter_options = ["All Executives"] + [
            f"{e['name']} (ID: {e['id']})"
            for e in db.get_all_executives(active_only=False)
        ]
        exec_filter_label = st.selectbox(
            "Executive", exec_filter_options, key="ev_filter_exec"
        )
        exec_filter_id = (
            None
            if exec_filter_label == "All Executives"
            else int(exec_filter_label.split("(ID: ")[1].rstrip(")"))
        )
    with col2:
        status_filter = st.selectbox(
            "Status",
            ["Active", "All"] + EVENT_STATUSES,
            format_func=lambda s: {
                "Active": "Active",
                "All": "All",
            }.get(s, _event_status_label(s)),
            key="ev_filter_status",
        )
    with col3:
        view_filter = st.selectbox(
            "View",
            [
                "All events",
                "Upcoming (30 days)",
                "Upcoming (90 days)",
                "Overdue RSVPs",
            ],
            key="ev_filter_view",
        )
    with col4:
        search_query = st.text_input(
            "Search", placeholder="Name or objective...",
            key="ev_search",
        )

    # ---- New button ----
    col_new, _sp = st.columns([1.2, 5])
    with col_new:
        if st.button(
            "➕ New Event", type="primary", key="ev_new_btn"
        ):
            st.session_state["ev_show_create"] = True

    # ---- Create form ----
    if st.session_state.get("ev_show_create", False):
        with st.expander("➕ New Event", expanded=True):
            execs_active = db.get_all_executives(active_only=True)
            if not execs_active:
                st.warning(
                    "No active executives. Add one in the sidebar first."
                )
            else:
                with st.form("ev_create_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        exec_options = {
                            e["name"]: e["id"] for e in execs_active
                        }
                        new_exec_label = st.selectbox(
                            "Executive*",
                            options=list(exec_options.keys()),
                            key="ev_new_exec",
                        )
                        new_name = st.text_input(
                            "Event Name*", key="ev_new_name"
                        )
                        new_type = st.selectbox(
                            "Type",
                            options=[""] + EVENT_TYPES,
                            key="ev_new_type",
                        )
                        new_status = st.selectbox(
                            "Status",
                            options=EVENT_STATUSES,
                            format_func=_event_status_label,
                            key="ev_new_status",
                        )
                    with col2:
                        new_date = st.date_input(
                            "Date", value=datetime.now().date(),
                            key="ev_new_date",
                        )
                        new_start = st.text_input(
                            "Start Time (e.g. 19:00)",
                            key="ev_new_start",
                        )
                        new_end = st.text_input(
                            "End Time (e.g. 22:00)",
                            key="ev_new_end",
                        )
                        new_headcount = st.number_input(
                            "Expected Headcount", min_value=0, step=5,
                            value=0, key="ev_new_headcount",
                        )

                    new_objective = st.text_area(
                        "Objective", height=80, key="ev_new_obj"
                    )

                    col_a, col_b = st.columns(2)
                    with col_a:
                        new_budget = st.number_input(
                            "Budget Total", min_value=0.0, step=100.0,
                            value=0.0, key="ev_new_budget",
                        )
                    with col_b:
                        new_currency = st.selectbox(
                            "Currency",
                            options=["USD", "EUR", "GBP", "NGN",
                                     "ZAR", "KES"],
                            key="ev_new_currency",
                        )

                    # Venue picker
                    venue_options = db.get_venues(active_only=True)
                    venue_labels = {"—": None}
                    for v in venue_options:
                        label = v["name"]
                        if v.get("city"):
                            label += f" — {v['city']}"
                        venue_labels[label] = v["id"]
                    new_venue_label = st.selectbox(
                        "Venue",
                        options=list(venue_labels.keys()),
                        key="ev_new_venue",
                    )
                    new_venue_id = venue_labels[new_venue_label]

                    new_location = st.text_input(
                        "Location Override (if no venue)",
                        key="ev_new_loc",
                    )

                    st.markdown("**🔗 Links**")
                    new_cal_url = st.text_input(
                        "Calendar Event URL", key="ev_new_cal"
                    )
                    new_rsvp_url = st.text_input(
                        "Google Form RSVP URL", key="ev_new_rsvp"
                    )
                    col_x, col_y = st.columns(2)
                    with col_x:
                        new_run_url = st.text_input(
                            "Run of Show Doc URL", key="ev_new_run"
                        )
                    with col_y:
                        new_logistics_url = st.text_input(
                            "Logistics Doc URL", key="ev_new_log"
                        )

                    new_notes = st.text_area(
                        "Quick Notes", height=100, key="ev_new_notes"
                    )

                    col_submit, col_cancel = st.columns(2)
                    with col_submit:
                        submitted = st.form_submit_button("💾 Create")
                    with col_cancel:
                        cancelled = st.form_submit_button("❌ Cancel")

                    if cancelled:
                        st.session_state["ev_show_create"] = False
                        st.rerun()

                    if submitted:
                        if not new_name:
                            st.error("Event Name is required.")
                        else:
                            new_exec_id = exec_options[new_exec_label]
                            dupes = db.find_duplicate_events(
                                new_exec_id, new_name,
                                new_date.isoformat() if new_date else None,
                            )
                            if dupes:
                                st.warning(
                                    "⚠️ An event with the same name and "
                                    "date already exists for this executive."
                                )
                                for d in dupes:
                                    st.caption(
                                        f"- {d['name']} "
                                        f"({d.get('event_date')})"
                                    )
                                st.info(
                                    "Change the name or date, or edit "
                                    "the existing event instead."
                                )
                            else:
                                db.add_event(
                                    new_exec_id,
                                    new_name,
                                    objective=new_objective or None,
                                    event_type=new_type or None,
                                    status=new_status,
                                    event_date=(
                                        new_date.isoformat()
                                        if new_date else None
                                    ),
                                    start_time=new_start or None,
                                    end_time=new_end or None,
                                    venue_id=new_venue_id,
                                    location_override=(
                                        new_location or None
                                    ),
                                    headcount=(
                                        int(new_headcount)
                                        if new_headcount else None
                                    ),
                                    budget_total=new_budget or 0,
                                    currency=new_currency,
                                    calendar_event_url=(
                                        new_cal_url or None
                                    ),
                                    rsvp_form_url=new_rsvp_url or None,
                                    run_of_show_url=new_run_url or None,
                                    logistics_url=(
                                        new_logistics_url or None
                                    ),
                                    notes_native=new_notes or None,
                                )
                                st.session_state["ev_show_create"] = False
                                st.success("✅ Event created.")
                                st.rerun()

    st.divider()

    # ---- Build filters ----
    status_param = None
    if status_filter == "Active":
        status_param = [
            "planning", "confirmed", "in_progress", "day_of",
        ]
    elif status_filter != "All":
        status_param = status_filter

    upcoming_only = view_filter in (
        "Upcoming (30 days)", "Upcoming (90 days)"
    )
    end_date_param = None
    if view_filter == "Upcoming (30 days)":
        end_date_param = (
            datetime.now().date() + timedelta(days=30)
        ).isoformat()
    elif view_filter == "Upcoming (90 days)":
        end_date_param = (
            datetime.now().date() + timedelta(days=90)
        ).isoformat()

    overdue_rsvps = view_filter == "Overdue RSVPs"

    events = db.get_events(
        exec_id=exec_filter_id,
        status=status_param,
        upcoming_only=upcoming_only,
        end_date=end_date_param,
        overdue_rsvps=overdue_rsvps,
        search=search_query or None,
    )

    # ---- List header ----
    if not events:
        st.info("No events match these filters.")
    else:
        st.write(f"**{len(events)} event(s)**")

        for ev in events:
            with st.container():
                col1, col2, col3 = st.columns([6, 1, 1])
                with col1:
                    status_label = _event_status_label(
                        ev.get("status") or "planning"
                    )
                    st.markdown(f"### {ev['name']}  ·  {status_label}")
                    venue_display = (
                        ev.get("venue_name")
                        or ev.get("location_override")
                        or "—"
                    )
                    st.caption(
                        f"👤 {ev.get('executive_name') or '—'}  ·  "
                        f"📅 {fmt_date(ev.get('event_date'))}"
                        + (
                            f"  ·  🕐 {ev['start_time']}"
                            if ev.get("start_time") else ""
                        )
                        + f"  ·  📍 {venue_display}"
                    )
                    counts = db.get_event_guest_status_counts(ev["id"])
                    accepted = counts.get("accepted", 0)
                    pending = counts.get("pending", 0) + counts.get(
                        "invited", 0
                    )
                    total = sum(counts.values())
                    st.caption(
                        f"👥 {total} guest(s)  ·  "
                        f"✅ {accepted} accepted  ·  "
                        f"⏳ {pending} pending"
                    )
                with col2:
                    if st.button(
                        "📂 Open",
                        key=f"ev_open_{ev['id']}",
                        use_container_width=True,
                    ):
                        st.session_state["dialog_event_id"] = ev["id"]
                with col3:
                    if st.button(
                        "📥 RSVP",
                        key=f"ev_rsvp_{ev['id']}",
                        use_container_width=True,
                        help="Import RSVPs from a Google Forms CSV",
                    ):
                        st.session_state[
                            "dialog_rsvp_import_event_id"
                        ] = ev["id"]

                st.divider()

    # ---- Nothing more inside tab_events ----


# =========================================================
# TAB: SPEAKING
# =========================================================

with tab_speaking:
    st.header("🎤 Speaking Engagements")

    # ---- View toggle ----
    col1, col2, col3 = st.columns([1.5, 3, 3])
    with col1:
        view_mode = st.radio(
            "View",
            ["Pipeline", "List"],
            horizontal=True,
            key="speak_view_mode",
        )
    with col2:
        exec_filter_options = ["All Executives"] + [
            f"{e['name']} (ID: {e['id']})"
            for e in db.get_all_executives(active_only=False)
        ]
        exec_filter_label = st.selectbox(
            "Executive", exec_filter_options, key="speak_filter_exec"
        )
        exec_filter_id = (
            None
            if exec_filter_label == "All Executives"
            else int(exec_filter_label.split("(ID: ")[1].rstrip(")"))
        )
    with col3:
        search_q = st.text_input(
            "Search", placeholder="Title, organizer, topic...",
            key="speak_search",
        )

    col_new, _sp = st.columns([1.4, 5])
    with col_new:
        if st.button(
            "➕ New Engagement", type="primary", key="speak_new_btn"
        ):
            st.session_state["speak_show_create"] = True

    # ---- Create form ----
    if st.session_state.get("speak_show_create", False):
        with st.expander("➕ New Engagement", expanded=True):
            execs_active = db.get_all_executives(active_only=True)
            if not execs_active:
                st.warning(
                    "No active executives. Add one in the sidebar first."
                )
            else:
                with st.form("speak_create_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        exec_options = {
                            e["name"]: e["id"] for e in execs_active
                        }
                        new_exec_label = st.selectbox(
                            "Executive*",
                            options=list(exec_options.keys()),
                            key="speak_new_exec",
                        )
                        new_title = st.text_input(
                            "Title*", key="speak_new_title"
                        )
                        new_type = st.selectbox(
                            "Type",
                            options=[""] + SPEAKING_TYPES,
                            key="speak_new_type",
                        )
                        new_format = st.selectbox(
                            "Format",
                            options=[""] + SPEAKING_FORMATS,
                            key="speak_new_format",
                        )
                        new_status = st.selectbox(
                            "Status",
                            options=SPEAKING_STATUSES,
                            format_func=_speaking_status_label,
                            key="speak_new_status",
                        )
                    with col2:
                        new_date = st.date_input(
                            "Date", value=datetime.now().date(),
                            key="speak_new_date",
                        )
                        new_time = st.text_input(
                            "Time (e.g. 14:00)", key="speak_new_time"
                        )
                        new_duration = st.text_input(
                            "Duration (e.g. 45 min)",
                            key="speak_new_duration",
                        )
                        new_organizer_name = st.text_input(
                            "Organizer Name", key="speak_new_org"
                        )

                    new_location = st.text_input(
                        "Location", key="speak_new_loc"
                    )
                    new_audience = st.text_input(
                        "Audience", key="speak_new_audience"
                    )
                    new_topic = st.text_input(
                        "Topic", key="speak_new_topic"
                    )

                    col_a, col_b = st.columns(2)
                    with col_a:
                        new_fee_offered = st.number_input(
                            "Fee Offered", min_value=0.0, step=500.0,
                            value=0.0, key="speak_new_fee_offered",
                        )
                        new_currency = st.text_input(
                            "Currency", value="USD",
                            key="speak_new_currency",
                        )
                    with col_b:
                        new_fee_agreed = st.number_input(
                            "Fee Agreed", min_value=0.0, step=500.0,
                            value=0.0, key="speak_new_fee_agreed",
                        )
                        new_cal_url = st.text_input(
                            "Calendar Event URL",
                            key="speak_new_cal",
                        )

                    new_notes = st.text_area(
                        "Quick Notes", height=80,
                        key="speak_new_notes",
                    )

                    col_sub, col_cancel = st.columns(2)
                    with col_sub:
                        submitted = st.form_submit_button("💾 Create")
                    with col_cancel:
                        cancelled = st.form_submit_button("❌ Cancel")

                    if cancelled:
                        st.session_state["speak_show_create"] = False
                        st.rerun()

                    if submitted:
                        if not new_title:
                            st.error("Title is required.")
                        else:
                            new_exec_id = exec_options[new_exec_label]
                            dupes = db.find_duplicate_speaking_engagements(
                                new_exec_id, new_title,
                                new_date.isoformat() if new_date else None,
                            )
                            if dupes:
                                st.warning(
                                    "⚠️ An engagement with the same title "
                                    "and date already exists."
                                )
                                for d in dupes:
                                    st.caption(
                                        f"- {d['title']} "
                                        f"({d.get('engagement_date')})"
                                    )
                                st.info("Edit the existing one instead.")
                            else:
                                new_id = db.add_speaking_engagement(
                                    new_exec_id,
                                    new_title,
                                    organizer_name=(
                                        new_organizer_name or None
                                    ),
                                    engagement_type=new_type or None,
                                    format=new_format or None,
                                    status=new_status,
                                    engagement_date=(
                                        new_date.isoformat()
                                        if new_date else None
                                    ),
                                    engagement_time=new_time or None,
                                    duration=new_duration or None,
                                    location=new_location or None,
                                    audience=new_audience or None,
                                    topic=new_topic or None,
                                    fee_offered=(
                                        new_fee_offered
                                        if new_fee_offered > 0 else None
                                    ),
                                    fee_agreed=(
                                        new_fee_agreed
                                        if new_fee_agreed > 0 else None
                                    ),
                                    currency=new_currency,
                                    calendar_event_url=(
                                        new_cal_url or None
                                    ),
                                    notes_native=new_notes or None,
                                )
                                st.session_state["speak_show_create"] = False
                                st.success("✅ Engagement created.")
                                st.rerun()

    st.divider()

    # ---- Build query ----
    engagements = db.get_speaking_engagements(
        exec_id=exec_filter_id,
        search=search_q or None,
    )

    # ---- Render ----
    if not engagements:
        st.info("No engagements match these filters.")
    else:
        if view_mode == "Pipeline":
            st.write("**Pipeline View** — counts by status")
            counts = db.get_speaking_pipeline_counts()
            # Compact pipeline summary
            pipeline_cols = st.columns(len(SPEAKING_PIPELINE_ORDER[:6]))
            for i, status_key in enumerate(SPEAKING_PIPELINE_ORDER[:6]):
                with pipeline_cols[i]:
                    st.metric(
                        _speaking_status_label(status_key),
                        counts.get(status_key, 0),
                    )
            pipeline_cols_2 = st.columns(len(SPEAKING_PIPELINE_ORDER[6:]))
            for i, status_key in enumerate(SPEAKING_PIPELINE_ORDER[6:]):
                with pipeline_cols_2[i]:
                    st.metric(
                        _speaking_status_label(status_key),
                        counts.get(status_key, 0),
                    )

            st.divider()

            # Group by status, show in pipeline order
            by_status = {}
            for e in engagements:
                by_status.setdefault(
                    e.get("status") or "invited", []
                ).append(e)

            for status_key in SPEAKING_PIPELINE_ORDER:
                group = by_status.get(status_key, [])
                if not group:
                    continue
                with st.expander(
                    f"{_speaking_status_label(status_key)} · {len(group)}",
                    expanded=status_key in (
                        "confirmed", "prep_in_progress", "ready"
                    ),
                ):
                    for eng in group:
                        col1, col2 = st.columns([5, 1])
                        with col1:
                            st.markdown(f"**{eng['title']}**")
                            st.caption(
                                f"👤 {eng.get('executive_name') or '—'}  ·  "
                                f"📅 {fmt_date(eng.get('engagement_date'))}"
                                + (
                                    f"  ·  {eng['engagement_type']}"
                                    if eng.get("engagement_type") else ""
                                )
                                + (
                                    f"  ·  💰 {fmt_money(eng['fee_agreed'], eng.get('currency') or 'USD')}"
                                    if eng.get("fee_agreed") else ""
                                )
                            )
                        with col2:
                            if st.button(
                                "📂 Open",
                                key=f"speak_pipe_open_{eng['id']}",
                                use_container_width=True,
                            ):
                                st.session_state[
                                    "dialog_speaking_id"
                                ] = eng["id"]

        else:
            # List view
            st.write(f"**{len(engagements)} engagement(s)**")
            for eng in engagements:
                col1, col2 = st.columns([6, 1])
                with col1:
                    st.markdown(
                        f"### {eng['title']}  ·  "
                        f"{_speaking_status_label(eng.get('status'))}"
                    )
                    st.caption(
                        f"👤 {eng.get('executive_name') or '—'}  ·  "
                        f"📅 {fmt_date(eng.get('engagement_date'))}"
                        + (
                            f"  ·  📍 {eng['location']}"
                            if eng.get("location") else ""
                        )
                        + (
                            f"  ·  💰 {fmt_money(eng['fee_agreed'], eng.get('currency') or 'USD')}"
                            if eng.get("fee_agreed") else ""
                        )
                    )
                    delivs = db.get_speaking_deliverables(eng["id"])
                    if delivs:
                        done = sum(
                            1 for d in delivs
                            if d.get("status") == "done"
                        )
                        st.caption(
                            f"📦 {done}/{len(delivs)} deliverables done"
                        )
                with col2:
                    if st.button(
                        "📂 Open",
                        key=f"speak_list_open_{eng['id']}",
                        use_container_width=True,
                    ):
                        st.session_state["dialog_speaking_id"] = eng["id"]

                st.divider()


# =========================================================
# TAB: VENDORS
# =========================================================

with tab_vendors:
    st.header("🏢 Vendors")

    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        cat_filter = st.selectbox(
            "Category",
            ["All"] + VENDOR_CATEGORIES,
            key="vendor_filter_cat",
        )
    with col2:
        pref_filter = st.checkbox(
            "Preferred only", key="vendor_filter_pref"
        )
    with col3:
        search_q = st.text_input(
            "Search", placeholder="Name, email, city...",
            key="vendor_search",
        )

    col_new, _sp = st.columns([1.2, 5])
    with col_new:
        if st.button(
            "➕ New Vendor", type="primary", key="vendor_new_btn"
        ):
            st.session_state["vendor_show_create"] = True

    if st.session_state.get("vendor_show_create", False):
        with st.expander("➕ New Vendor", expanded=True):
            with st.form("vendor_create_form"):
                col1, col2 = st.columns(2)
                with col1:
                    v_name = st.text_input("Name*", key="new_v_name")
                    v_cat = st.selectbox(
                        "Category",
                        options=[""] + VENDOR_CATEGORIES,
                        key="new_v_cat",
                    )
                    v_email = st.text_input("Email", key="new_v_email")
                    v_phone = st.text_input("Phone", key="new_v_phone")
                with col2:
                    v_website = st.text_input(
                        "Website", key="new_v_website"
                    )
                    v_city = st.text_input("City", key="new_v_city")
                    v_country = st.text_input(
                        "Country", key="new_v_country"
                    )
                    v_rating = st.number_input(
                        "Rating", min_value=0, max_value=5, step=1,
                        value=0, key="new_v_rating",
                    )
                v_address = st.text_input("Address", key="new_v_addr")
                v_notes = st.text_area(
                    "Notes", key="new_v_notes", height=80
                )
                v_pref = st.checkbox("⭐ Preferred", key="new_v_pref")

                col_sub, col_cancel = st.columns(2)
                with col_sub:
                    submitted = st.form_submit_button("💾 Create")
                with col_cancel:
                    cancelled = st.form_submit_button("❌ Cancel")

                if cancelled:
                    st.session_state["vendor_show_create"] = False
                    st.rerun()

                if submitted:
                    if not v_name:
                        st.error("Name is required.")
                    else:
                        dupes = db.find_duplicate_vendors(v_name)
                        if dupes:
                            st.warning(
                                f"⚠️ A vendor named '{v_name}' already exists."
                            )
                            for d in dupes:
                                st.caption(f"- {d['name']}")
                            st.info(
                                "Edit the existing vendor instead, or "
                                "change the name."
                            )
                        else:
                            db.add_vendor(
                                name=v_name,
                                category=v_cat or None,
                                email=v_email or None,
                                phone=v_phone or None,
                                website=v_website or None,
                                city=v_city or None,
                                country=v_country or None,
                                address=v_address or None,
                                notes=v_notes or None,
                                rating=(
                                    int(v_rating) if v_rating else None
                                ),
                                preferred=1 if v_pref else 0,
                            )
                            st.session_state["vendor_show_create"] = False
                            st.success("✅ Vendor created.")
                            st.rerun()

    st.divider()

    vendors = db.get_vendors(
        category=cat_filter if cat_filter != "All" else None,
        preferred_only=pref_filter,
        search=search_q or None,
    )

    if not vendors:
        st.info("No vendors match these filters.")
    else:
        st.write(f"**{len(vendors)} vendor(s)**")
        for v in vendors:
            col1, col2, col3 = st.columns([5, 1, 1])
            with col1:
                pref = "⭐ " if v.get("preferred") else ""
                st.markdown(f"### {pref}{v['name']}")
                info_bits = []
                if v.get("category"):
                    info_bits.append(v["category"])
                if v.get("city") or v.get("country"):
                    loc = ", ".join(
                        x for x in [v.get("city"), v.get("country")] if x
                    )
                    info_bits.append(loc)
                if v.get("rating"):
                    info_bits.append("★" * int(v["rating"]))
                if info_bits:
                    st.caption(" · ".join(info_bits))
                contact_bits = []
                if v.get("email"):
                    contact_bits.append(f"✉️ {v['email']}")
                if v.get("phone"):
                    contact_bits.append(f"📞 {v['phone']}")
                if contact_bits:
                    st.caption(" · ".join(contact_bits))
            with col2:
                if st.button(
                    "📂 Open", key=f"vendor_open_{v['id']}",
                    use_container_width=True,
                ):
                    st.session_state["dialog_vendor_id"] = v["id"]
            with col3:
                if st.button("🗑️", key=f"vendor_del_{v['id']}"):
                    st.session_state[
                        f"confirm_list_del_vendor_{v['id']}"
                    ] = True

            if st.session_state.get(
                f"confirm_list_del_vendor_{v['id']}", False
            ):
                st.warning(f"⚠️ Delete '{v['name']}'?")
                col_y, col_n = st.columns(2)
                with col_y:
                    if st.button(
                        "Yes, Delete", key=f"vendor_list_yes_{v['id']}"
                    ):
                        db.delete_vendor(v["id"])
                        st.session_state.pop(
                            f"confirm_list_del_vendor_{v['id']}", None
                        )
                        st.rerun()
                with col_n:
                    if st.button(
                        "Cancel", key=f"vendor_list_no_{v['id']}"
                    ):
                        st.session_state.pop(
                            f"confirm_list_del_vendor_{v['id']}", None
                        )
                        st.rerun()

            st.divider()

# =========================================================
# TAB: DASHBOARD
# =========================================================

with tab_dashboard:
    st.header("🏠 Dashboard")

    # ================================================
    # UNIFIED TIMELINE (next 14 days)
    # ================================================
    timeline_start = datetime.now().date()
    timeline_end = timeline_start + timedelta(days=14)
    timeline_items = db.get_upcoming_activities(
        timeline_start.isoformat(),
        timeline_end.isoformat(),
    )
    timeline_trips = db.get_upcoming_trips(
        timeline_start.isoformat(),
        timeline_end.isoformat(),
    )

    # Merge and group by date
    all_timeline = timeline_items + timeline_trips
    all_timeline.sort(
        key=lambda x: (
            x.get("date") or "9999-12-31",
            x.get("time") or "99:99",
        )
    )

    with st.expander(
        f"📆 What's On (next 14 days) — {len(all_timeline)} item(s)",
        expanded=bool(all_timeline),
    ):
        if not all_timeline:
            st.caption("Nothing scheduled in the next two weeks.")
        else:
            # Group by date
            by_date = {}
            for item in all_timeline:
                d = item.get("date") or "Unknown"
                by_date.setdefault(d, []).append(item)

            for d, items in by_date.items():
                try:
                    nice_date = datetime.fromisoformat(d).strftime(
                        "%A, %d %b %Y"
                    )
                except Exception:
                    nice_date = d
                st.markdown(f"**{nice_date}**")
                for item in items:
                    kind = item.get("kind")
                    kind_icon = {
                        "meeting": "📅",
                        "event": "🎉",
                        "speaking": "🎤",
                        "trip": "✈️",
                    }.get(kind, "•")
                    time_str = item.get("time") or ""
                    time_part = f"  ·  {time_str}" if time_str else ""
                    exec_name = item.get("executive_name") or "—"
                    st.markdown(
                        f"{kind_icon}  {item.get('title') or '—'}"
                        f"{time_part}"
                    )
                    extra = item.get("extra") or ""
                    caption_bits = [f"👤 {exec_name}"]
                    if extra:
                        caption_bits.append(f"📍 {extra}")
                    if kind == "trip" and item.get("status"):
                        caption_bits.append(item["status"])
                    st.caption("  ·  ".join(caption_bits))
                st.divider()

    st.divider()
    # ================================================
    # END UNIFIED TIMELINE
    # ================================================

    # ================================================
    # GLOBAL SEARCH
    # ================================================
    search_q = st.text_input(
        "🔍 Search everything (meetings, events, engagements, trips, "
        "contacts, vendors)",
        key="global_search_q",
        placeholder="Type a keyword and press Enter",
    )
    if search_q:
        results = db.global_search(search_q, limit_per_kind=8)
        total = sum(len(v) for v in results.values())
        if total == 0:
            st.caption("No matches.")
        else:
            st.caption(f"**{total}** result(s)")
            cols = st.columns(3)
            col_map = {
                "meetings": (cols[0], "📅 Meetings", "dialog_meeting_id"),
                "events": (cols[0], "🎉 Events", "dialog_event_id"),
                "speaking": (cols[1], "🎤 Speaking", "dialog_speaking_id"),
                "trips": (cols[1], "✈️ Trips", None),
                "contacts": (cols[2], "👥 Contacts", "dialog_comm_contact_id"),
                "vendors": (cols[2], "🏢 Vendors", "dialog_vendor_id"),
            }
            for kind, (col, label, dialog_key) in col_map.items():
                rows = results.get(kind) or []
                if not rows:
                    continue
                with col:
                    st.markdown(f"**{label}** ({len(rows)})")
                    for r in rows:
                        if st.button(
                            f"{r['title'] or '—'}",
                            key=f"gs_{kind}_{r['id']}",
                            use_container_width=True,
                        ):
                            if dialog_key:
                                st.session_state[dialog_key] = r["id"]
                                st.rerun()
                            elif kind == "trips":
                                # Open trip edit modal
                                st.session_state["dialog_trip_id"] = r["id"]
                                st.rerun()

    st.divider()
    # ================================================
    # END GLOBAL SEARCH
    # ================================================

    # ================================================
    # MEETINGS WIDGETS
    # ================================================
    today_iso = datetime.now().date().isoformat()
    week_end_iso = (datetime.now().date() + timedelta(days=7)).isoformat()

    todays_meetings = db.get_meetings(
        start_date=today_iso,
        end_date=today_iso,
    )
    awaiting_meetings = db.get_meetings(awaiting_response=True)
    no_notes_meetings = db.get_meetings(without_notes=True)
    actions_due = db.get_action_items_due(
        start_date=today_iso,
        end_date=week_end_iso,
    )

    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.metric("📅 Today's Meetings", len(todays_meetings))
    with col_b:
        st.metric("⏳ Awaiting Response", len(awaiting_meetings))
    with col_c:
        st.metric("⚠️ Without Notes", len(no_notes_meetings))
    with col_d:
        st.metric("✅ Action Items Due (7d)", len(actions_due))

    # ---- Expanded widgets ----
    with st.expander("📅 Today's Meetings", expanded=bool(todays_meetings)):
        if not todays_meetings:
            st.caption("No meetings today.")
        else:
            for m in todays_meetings:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{m['title']}**")
                    st.caption(
                        f"👤 {m.get('executive_name') or '—'}  ·  "
                        f"🕐 {m.get('meeting_time') or '—'}  ·  "
                        f"{_meeting_status_label(m.get('status') or 'draft')}"
                    )
                with col2:
                    if st.button("Open", key=f"dash_mtg_open_{m['id']}"):
                        st.session_state["dialog_meeting_id"] = m["id"]
                        st.rerun()

    with st.expander("⏳ Meetings Awaiting Response", expanded=False):
        if not awaiting_meetings:
            st.caption("Nothing pending.")
        else:
            for m in awaiting_meetings:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{m['title']}**")
                    st.caption(
                        f"👤 {m.get('executive_name') or '—'}  ·  "
                        f"📅 {fmt_date(m.get('meeting_date'))}  ·  "
                        f"{_meeting_status_label(m.get('status') or 'draft')}"
                    )
                with col2:
                    if st.button("Open", key=f"dash_mtg_awaiting_open_{m['id']}"):
                        st.session_state["dialog_meeting_id"] = m["id"]
                        st.rerun()

    with st.expander("✅ Action Items Due (next 7 days)", expanded=bool(actions_due)):
        if not actions_due:
            st.caption("Nothing due.")
        else:
            for item in actions_due:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{item['description']}**")
                    owner = (
                        item.get("owner_contact_name")
                        or item.get("owner_exec_name")
                        or "—"
                    )
                    st.caption(
                        f"📅 {fmt_date(item.get('due_date'))}  ·  "
                        f"👤 {owner}  ·  "
                        f"from *{item.get('meeting_title')}*"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_action_open_{item['id']}",
                    ):
                        st.session_state["dialog_meeting_id"] = item["meeting_id"]
                        st.rerun()

    st.divider()
    # ================================================
    # END MEETINGS WIDGETS
    # ================================================

    # ================================================
    # COMMUNICATIONS WIDGETS
    # ================================================
    next_actions_this_week = db.get_next_actions_due(
        start_date=today_iso, end_date=week_end_iso,
    )
    next_actions_overdue = [
        c for c in db.get_next_actions_due()
        if c.get("next_action_date")
        and c["next_action_date"] < today_iso
    ]
    gone_quiet_widget = db.get_contacts_gone_quiet(days=60)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("📞 Next Actions (7d)", len(next_actions_this_week))
    with col_b:
        st.metric("🔴 Overdue Follow-ups", len(next_actions_overdue))
    with col_c:
        st.metric("💤 Gone Quiet (60d+)", len(gone_quiet_widget))

    with st.expander(
        "🔴 Overdue Follow-ups", expanded=bool(next_actions_overdue)
    ):
        if not next_actions_overdue:
            st.caption("None overdue.")
        else:
            for c in next_actions_overdue[:10]:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{c['name']}** — {c['next_action']}")
                    st.caption(
                        f"Due: {c.get('next_action_date') or '—'}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_comm_overdue_{c['id']}",
                    ):
                        st.session_state[
                            "dialog_comm_contact_id"
                        ] = c["id"]
                        st.rerun()

    with st.expander(
        "📞 Next Actions This Week", expanded=False
    ):
        if not next_actions_this_week:
            st.caption("Nothing due this week.")
        else:
            for c in next_actions_this_week:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{c['name']}** — {c['next_action']}")
                    st.caption(
                        f"Due: {c.get('next_action_date') or '—'}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_comm_week_{c['id']}",
                    ):
                        st.session_state[
                            "dialog_comm_contact_id"
                        ] = c["id"]
                        st.rerun()

    st.divider()
    # ================================================
    # END COMMUNICATIONS WIDGETS
    # ================================================

    # ================================================
    # EVENTS WIDGETS
    # ================================================
    upcoming_30 = db.get_events(
        upcoming_only=True,
        end_date=(
            datetime.now().date() + timedelta(days=30)
        ).isoformat(),
        status=["planning", "confirmed", "in_progress", "day_of"],
    )
    overdue_rsvp_events = db.get_events_with_overdue_rsvps(days_ahead=14)
    upcoming_90 = db.get_events(
        upcoming_only=True,
        end_date=(
            datetime.now().date() + timedelta(days=90)
        ).isoformat(),
        status=["planning", "confirmed", "in_progress", "day_of"],
    )

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("🎉 Events (30d)", len(upcoming_30))
    with col_b:
        st.metric("⏳ Overdue RSVPs", len(overdue_rsvp_events))
    with col_c:
        st.metric("📅 Events (90d)", len(upcoming_90))

    with st.expander(
        "🎉 Upcoming Events (30 days)",
        expanded=bool(upcoming_30),
    ):
        if not upcoming_30:
            st.caption("No upcoming events.")
        else:
            for ev in upcoming_30:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{ev['name']}**")
                    st.caption(
                        f"👤 {ev.get('executive_name') or '—'}  ·  "
                        f"📅 {fmt_date(ev.get('event_date'))}  ·  "
                        f"{_event_status_label(ev.get('status'))}"
                    )
                with col2:
                    if st.button(
                        "Open", key=f"dash_ev_open_{ev['id']}"
                    ):
                        st.session_state["dialog_event_id"] = ev["id"]
                        st.rerun()

    with st.expander(
        "⏳ Overdue RSVPs",
        expanded=bool(overdue_rsvp_events),
    ):
        if not overdue_rsvp_events:
            st.caption("No pending RSVPs for upcoming events.")
        else:
            for ev in overdue_rsvp_events:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{ev['name']}**")
                    st.caption(
                        f"📅 {fmt_date(ev.get('event_date'))}  ·  "
                        f"⏳ {ev.get('pending_count', 0)} pending"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_ev_rsvp_open_{ev['id']}",
                    ):
                        st.session_state["dialog_event_id"] = ev["id"]
                        st.rerun()

    st.divider()
    # ================================================
    # END EVENTS WIDGETS
    # ================================================

    # ================================================
    # VENDOR & BUDGET WIDGETS
    # ================================================
    payments_due = db.get_payments_due(
        start_date=today_iso, end_date=week_end_iso,
    )
    over_budget_events = db.get_events_over_budget(threshold=0.9)
    unsigned_contracts = db.get_unsigned_contracts_within_days(days=7)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("💳 Payments Due (7d)", len(payments_due))
    with col_b:
        st.metric("💰 Events Over 90%", len(over_budget_events))
    with col_c:
        st.metric("✍️ Unsigned Contracts", len(unsigned_contracts))

    with st.expander(
        "💳 Payments Due (next 7 days)", expanded=bool(payments_due)
    ):
        if not payments_due:
            st.caption("No payments due.")
        else:
            for p in payments_due:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(
                        f"**{p['vendor_name']}** — "
                        f"{fmt_money(p.get('amount') or 0, p.get('currency') or 'USD')}"
                    )
                    st.caption(
                        f"📅 {fmt_date(p.get('due_date'))}  ·  "
                        f"from *{p['event_name']}*"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_pay_open_{p['id']}",
                    ):
                        st.session_state["dialog_event_id"] = p["event_id"]
                        st.rerun()

    with st.expander(
        "💰 Events Over 90% Budget",
        expanded=bool(over_budget_events),
    ):
        if not over_budget_events:
            st.caption("All events within budget.")
        else:
            for ev in over_budget_events:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{ev['name']}**")
                    ratio = (
                        ev["spent_actual"] / ev["budget_total"]
                        if ev["budget_total"] else 0
                    )
                    st.caption(
                        f"📅 {fmt_date(ev.get('event_date'))}  ·  "
                        f"{ratio*100:.0f}% of "
                        f"{fmt_money(ev['budget_total'], ev.get('currency') or 'USD')}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_ob_open_{ev['id']}",
                    ):
                        st.session_state["dialog_event_id"] = ev["id"]
                        st.rerun()

    with st.expander(
        "✍️ Unsigned Contracts (7 days)",
        expanded=bool(unsigned_contracts),
    ):
        if not unsigned_contracts:
            st.caption("No unsigned contracts coming up.")
        else:
            for c in unsigned_contracts:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{c['vendor_name']}**")
                    st.caption(
                        f"Event: {c['event_name']}  ·  "
                        f"📅 {fmt_date(c.get('event_date'))}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_uc_open_{c['id']}",
                    ):
                        st.session_state["dialog_event_id"] = c["event_id"]
                        st.rerun()

    st.divider()
    # ================================================
    # END VENDOR & BUDGET WIDGETS
    # ================================================

    # ================================================
    # SPEAKING ENGAGEMENT WIDGETS
    # ================================================
    today_iso_speak = today_iso
    week_end_iso_speak = week_end_iso

    upcoming_speak_30 = db.get_speaking_engagements(
        upcoming_only=True,
        end_date=(
            datetime.now().date() + timedelta(days=30)
        ).isoformat(),
    )
    deliverables_due = db.get_deliverables_due(
        start_date=today_iso_speak, end_date=week_end_iso_speak,
    )
    cfps_closing = db.get_cfps_closing_within_days(days=14)
    pipeline_value = db.get_speaking_pipeline_value()
    thankyous_overdue = db.get_thank_yous_overdue(days_since_delivery=7)

    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.metric("🎤 Upcoming (30d)", len(upcoming_speak_30))
    with col_b:
        st.metric("📦 Deliverables Due (7d)", len(deliverables_due))
    with col_c:
        st.metric("📤 CFPs Closing (14d)", len(cfps_closing))
    with col_d:
        st.metric(
            "💰 Confirmed Value",
            fmt_money(pipeline_value, "USD"),
        )

    with st.expander(
        "🎤 Upcoming Engagements (30 days)",
        expanded=bool(upcoming_speak_30),
    ):
        if not upcoming_speak_30:
            st.caption("No upcoming engagements.")
        else:
            for eng in upcoming_speak_30:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{eng['title']}**")
                    st.caption(
                        f"👤 {eng.get('executive_name') or '—'}  ·  "
                        f"📅 {fmt_date(eng.get('engagement_date'))}  ·  "
                        f"{_speaking_status_label(eng.get('status'))}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_speak_open_{eng['id']}",
                    ):
                        st.session_state["dialog_speaking_id"] = eng["id"]
                        st.rerun()

    with st.expander(
        "📦 Deliverables Due (7 days)",
        expanded=bool(deliverables_due),
    ):
        if not deliverables_due:
            st.caption("No deliverables due this week.")
        else:
            for d in deliverables_due:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{d['name']}**")
                    st.caption(
                        f"📅 {fmt_date(d.get('due_date'))}  ·  "
                        f"from *{d.get('engagement_title')}*"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_deliv_open_{d['id']}",
                    ):
                        st.session_state["dialog_speaking_id"] = d[
                            "engagement_id"
                        ]
                        st.rerun()

    with st.expander(
        "📤 CFPs Closing (14 days)",
        expanded=bool(cfps_closing),
    ):
        if not cfps_closing:
            st.caption("No CFPs closing soon.")
        else:
            for cf in cfps_closing:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(
                        f"**{cf.get('proposal_title') or 'Untitled'}**"
                    )
                    st.caption(
                        f"📅 Deadline: {fmt_date(cf.get('submission_deadline'))}  ·  "
                        f"from *{cf.get('engagement_title')}*"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_cfp_open_{cf['id']}",
                    ):
                        st.session_state["dialog_speaking_id"] = cf[
                            "engagement_id"
                        ]
                        st.rerun()

    with st.expander(
        "📮 Thank-Yous Overdue",
        expanded=bool(thankyous_overdue),
    ):
        if not thankyous_overdue:
            st.caption("All thank-yous sent.")
        else:
            for eng in thankyous_overdue:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{eng['title']}**")
                    st.caption(
                        f"📅 {fmt_date(eng.get('engagement_date'))}  ·  "
                        f"👤 {eng.get('executive_name') or '—'}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"dash_ty_open_{eng['id']}",
                    ):
                        st.session_state["dialog_speaking_id"] = eng["id"]
                        st.rerun()

    st.divider()
    # ================================================
    # END SPEAKING ENGAGEMENT WIDGETS
    # ================================================

    # ---- Load everything once ----
    all_execs = db.get_all_executives(active_only=False)
    active_execs = [e for e in all_execs if e.get("is_active", 1)]
    all_companies = db.get_all_companies(active_only=False)
    active_companies = [c for c in all_companies if c.get("is_active", 1)]
    all_contacts = db.get_contacts(active_only=False)
    active_contacts = [c for c in all_contacts if c.get("is_active", 1)]
    all_venues = db.get_venues(active_only=False)
    active_venues = [v for v in all_venues if v.get("is_active", 1)]
    summaries = db.get_spending_summary()

    today = datetime.now().date()

    upcoming_trips = [t for t in summaries if _trip_is_upcoming(t, today)]
    active_trips = [t for t in summaries if _trip_is_active(t, today)]
    past_trips = [t for t in summaries if _trip_is_past(t, today)]

    # ---- Top metric row ----
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(
            "Executives",
            len(active_execs),
            delta=f"{len(all_execs) - len(active_execs)} inactive"
            if len(all_execs) > len(active_execs) else None,
        )
    with col2:
        st.metric(
            "Companies",
            len(active_companies),
            delta=f"{len(all_companies) - len(active_companies)} inactive"
            if len(all_companies) > len(active_companies) else None,
        )
    with col3:
        st.metric("Contacts", len(active_contacts))
    with col4:
        st.metric("Total Trips", len(summaries))

    # ---- Second metric row: trip status split ----
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Upcoming (30 days)", len(upcoming_trips))
    with col2:
        st.metric("In Progress", len(active_trips))
    with col3:
        st.metric("Past Trips", len(past_trips))
    with col4:
        # Total spend this calendar month, converted to USD
        this_month_spend = 0.0
        for t in summaries:
            if _trips_in_month(t, today.year, today.month):
                amt_usd, _ok = _trip_total_in_usd(t)
                this_month_spend += amt_usd
        st.metric(
            f"Spend · {today.strftime('%b %Y')}",
            fmt_money(this_month_spend, "USD"),
        )

    st.divider()

    # ---- Upcoming trips table ----
    st.subheader("📅 Upcoming Trips (next 30 days)")
    if not upcoming_trips:
        st.info("No trips scheduled in the next 30 days.")
    else:
        for t in sorted(upcoming_trips, key=lambda x: x.get("start_date", "")):
            col_a, col_b, col_c, col_d, col_e = st.columns([2, 2, 1.5, 1.5, 1])
            with col_a:
                st.write(f"**{t.get('destination', '—')}**")
            with col_b:
                st.caption(t.get("executive_name", "—"))
            with col_c:
                st.write(fmt_date(t.get("start_date")))
            with col_d:
                st.write(fmt_money(t.get("budget") or 0, t.get("base_currency", "USD")))
            with col_e:
                status = t.get("status", "draft")
                st.caption(
                    "Draft" if status == "draft"
                    else "Approved" if status == "approved"
                    else "Final" if status == "final" else status
                )
            st.divider()

    # ---- Spending snapshot ----
    st.subheader("💰 Spending Snapshot")
    st.caption(
        "All figures converted to **USD**. Each trip is converted at its "
        "own start-date exchange rate, so historical costs stay accurate."
    )
    col1, col2, col3 = st.columns(3)

    def _sum_spend(trips):
        """Sum total_spent across trips, each converted to USD."""
        total = 0.0
        failed = 0
        for t in trips:
            amt_usd, ok = _trip_total_in_usd(t)
            total += amt_usd
            if not ok:
                failed += 1
        if failed:
            st.caption(
                f"⚠️ {failed} trip(s) had no exchange rate available and "
                "were included at face value."
            )
        return total

    month_trips = [t for t in summaries
                   if _trips_in_month(t, today.year, today.month)]
    quarter = (today.month - 1) // 3 + 1
    quarter_trips = [t for t in summaries
                     if _trips_in_quarter(t, today.year, quarter)]
    year_trips = [
        t
        for t in summaries
        if any(_trips_in_month(t, today.year, m) for m in range(1, 13))
    ]

    with col1:
        st.metric(f"This Month ({today.strftime('%b')})", fmt_money(_sum_spend(month_trips), "USD"))
    with col2:
        st.metric(f"Q{quarter} {today.year}", fmt_money(_sum_spend(quarter_trips), "USD"))
    with col3:
        st.metric(f"{today.year} YTD", fmt_money(_sum_spend(year_trips), "USD"))

    # ---- Per-executive spend breakdown ----
    if summaries:
        st.subheader("👤 Spend by Executive")
        st.caption("All figures converted to USD at each trip's start-date rate.")
        by_exec = {}
        for t in summaries:
            name = t.get("executive_name", "—")
            amt_usd, _ok = _trip_total_in_usd(t)
            by_exec[name] = by_exec.get(name, 0.0) + amt_usd
        rows = sorted(by_exec.items(), key=lambda kv: -kv[1])
        for name, total in rows:
            col_a, col_b = st.columns([3, 2])
            with col_a:
                st.write(f"**{name}**")
            with col_b:
                st.write(fmt_money(total, "USD"))
        st.divider()

    # ---- Library content counts ----
    st.subheader("📚 Library Content")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Guides", len(db.get_destination_guides(active_only=True)))
        st.metric("Visa Rules", len(db.get_visa_rules(active_only=True)))
    with col2:
        st.metric("Checklist Tpls", len(db.get_checklist_templates(active_only=True)))
        st.metric("Packing Tpls", len(db.get_packing_templates(active_only=True)))
    with col3:
        st.metric("Trip Tpls", len(db.get_trip_templates(active_only=True)))
        st.metric("Hospitals", len(db.get_hospitals(active_only=True)))
    with col4:
        st.metric("Embassies", len(db.get_embassies(active_only=True)))
        st.metric("Venues", len(active_venues))

    st.divider()

    # ---- Recently created trips ----
    st.subheader("🕒 Recently Created Trips")
    recent = sorted(
        summaries,
        key=lambda t: t.get("start_date", ""),
        reverse=True,
    )[:5]
    if not recent:
        st.info("No trips yet.")
    else:
        for t in recent:
            col_a, col_b, col_c = st.columns([2.5, 2, 1.5])
            with col_a:
                st.write(f"**{t.get('destination', '—')}**")
            with col_b:
                st.caption(t.get("executive_name", "—"))
            with col_c:
                st.caption(
                    f"{fmt_date(t.get('start_date'))} → "
                    f"{fmt_date(t.get('end_date'))}"
                )


# =========================================================
# TAB: TRIP PLANNER
# =========================================================

with tab_trip_planner:
    if not profile:
        st.info("👤 **Add an executive in the sidebar** to start planning a trip. "
                "The other tabs (Companies, Contacts, Venues, Library) work without one.")
    else:
        exec_dropdown_options = {f"{e['name']} (ID: {e['id']})": e["id"]
            for e in db.get_all_executives(active_only=True)}
        trip_exec_label = st.selectbox("👤 Executive for this Trip",
            options=list(exec_dropdown_options.keys()), key="create_trip_exec")
        trip_exec_id = exec_dropdown_options[trip_exec_label]
        # Resolve company for the chosen trip executive. Used by the
        # local-support and delegation pickers below (previously this name
        # leaked in from other scopes — NameError at runtime).
        _trip_exec_profile = db.get_executive_profile(trip_exec_id)
        company_id = (
            _trip_exec_profile.get("company_id") if _trip_exec_profile else None
        )
        trip_purpose = st.text_input("Trip Name / Purpose (e.g., 'Q3 Sales Tour')",
            key="create_trip_purpose")

        col_start, col_end = st.columns(2)
        with col_start:
            overall_start = st.date_input("Start Date*",
                value=st.session_state.get("create_overall_start", datetime.now()),
                key="create_overall_start")
        with col_end:
            overall_end = st.date_input("End Date*",
                value=st.session_state.get("create_overall_end", datetime.now() + timedelta(days=1)),
                key="create_overall_end")
        if overall_start and overall_start < datetime.now().date() - timedelta(days=1):
            st.warning(
                f"⚠️ The start date ({fmt_date(overall_start.isoformat())}) is in the past. "
                "This is fine for logging a past trip — but double-check the year."
            )
        if overall_start and overall_end and overall_end >= overall_start:
            duration = (overall_end - overall_start).days
            st.caption(f"⏱️ Duration: {duration} day(s)")
        elif overall_start and overall_end:
            st.warning("End date must be after start date.")

        col_dep_city, col_dep_region = st.columns(2)
        with col_dep_city:
            departure_city = st.text_input("Departure City*", key="create_departure_city")
        with col_dep_region:
            departure_region = st.text_input("Region / State (optional)", key="create_departure_region")
        country_list = sorted([c.name for c in pycountry.countries])
        departure_country = st.selectbox(
            "Country (optional)",
            options=[""] + country_list,
            key="create_departure_country",
        )

        receipt_upload_folder = st.text_input(
            "📥 Receipt Upload Folder URL (optional)",
            placeholder="https://drive.google.com/drive/folders/…",
            key="create_receipt_folder",
            help="Shared folder for delegation members to drop receipt photos. "
            "Will appear in the travel pack.",
        )

        col_budget, col_currency = st.columns(2)
        with col_budget:
            budget = st.number_input("Budget Amount (in Base Currency)",
                min_value=0.0, step=100.0, value=0.0, key="create_trip_budget")
        with col_currency:
            base_currency_options = currency.CURRENCIES_SUPPORTED
            trip_base_currency = st.selectbox("Base Currency", options=base_currency_options,
                index=0, key="create_base_currency")

        trip_status = st.selectbox("Trip Status", options=["draft", "approved", "final"],
            index=0, key="create_trip_status",
            help="Draft = editable, Approved = locked, Final = locked.")

        tz_display_mode = st.radio("Show times in:", options=["Home", "Destination"],
            index=0, key="create_tz_display_mode", horizontal=True)
        st.caption(
            "**Home** converts every item to the executive's home timezone. "
            "**Destination** shows each item in its own assigned timezone."
        )
        # -------here------
        if "create_trip_stops" not in st.session_state:
            st.session_state["create_trip_stops"] = []

        if st.session_state["create_trip_stops"]:
            for idx, stop in enumerate(st.session_state["create_trip_stops"]):
                (
                    col1, col2, col3, col4,
                    col_up, col_dn, col5, col6,
                ) = st.columns([2, 2, 2, 2, 0.4, 0.4, 0.5, 0.5])
                with col1:
                    st.write(f"**{idx + 1}.** {stop['city']}")
                with col2:
                    loc_parts = [p for p in [stop.get("region"), stop.get("country")] if p]
                    st.write(", ".join(loc_parts) if loc_parts else "")
                with col3:
                    st.write(f"{fmt_date(stop['start_date'])} → {fmt_date(stop['end_date'])}")
                with col4:
                    st.write(stop.get("notes", "")[:30])
                with col_up:
                    if idx > 0:
                        if st.button("↑", key=f"create_stop_up_{idx}",
                                     help="Move this stop up"):
                            lst = st.session_state["create_trip_stops"]
                            lst[idx - 1], lst[idx] = lst[idx], lst[idx - 1]
                            st.rerun()
                with col_dn:
                    lst_len = len(st.session_state["create_trip_stops"])
                    if idx < lst_len - 1:
                        if st.button("↓", key=f"create_stop_dn_{idx}",
                                     help="Move this stop down"):
                            lst = st.session_state["create_trip_stops"]
                            lst[idx + 1], lst[idx] = lst[idx], lst[idx + 1]
                            st.rerun()
                with col5:
                    if st.button("✏️", key=f"create_edit_stop_{idx}"):
                        st.session_state[f"create_editing_stop_{idx}"] = True
                with col6:
                    if st.button("🗑️", key=f"create_del_stop_{idx}"):
                        st.session_state["create_trip_stops"].pop(idx)
                        st.rerun()

                # --- Edit Stop expander ---
                if st.session_state.get(f"create_editing_stop_{idx}", False):
                    with st.expander(f"Edit Stop: {stop['city']}", expanded=True):
                        with st.form(key=f"create_edit_stop_form_{idx}"):
                            e_city = st.text_input("City*", value=stop["city"],
                                                   key=f"create_e_stop_city_{idx}")
                            e_country = st.selectbox(
                                "Country (optional)",
                                options=[""] + country_list,
                                index=([""] + country_list).index(stop.get("country", ""))
                                      if stop.get("country", "") in country_list else 0,
                                key=f"create_e_stop_country_{idx}")
                            e_region = st.text_input("Region / State",
                                                     value=stop.get("region", ""),
                                                     key=f"create_e_stop_region_{idx}")
                            e_notes = st.text_input("Notes",
                                                    value=stop.get("notes", ""),
                                                    key=f"create_e_stop_notes_{idx}")
                            col_a, col_b = st.columns(2)
                            with col_a:
                                try:
                                    default_start = datetime.fromisoformat(stop["start_date"]).date()
                                except Exception:
                                    default_start = datetime.now().date()
                                e_start = st.date_input("Start Date*", value=default_start,
                                                        key=f"create_e_stop_start_{idx}")
                            with col_b:
                                try:
                                    default_end = datetime.fromisoformat(stop["end_date"]).date()
                                except Exception:
                                    default_end = datetime.now().date()
                                e_end = st.date_input("End Date*", value=default_end,
                                                      key=f"create_e_stop_end_{idx}")
                            sub_save, sub_cancel = st.columns(2)
                            with sub_save:
                                if st.form_submit_button("💾 Update Stop"):
                                    if e_city and e_start and e_end:
                                        st.session_state["create_trip_stops"][idx] = {
                                            "city": e_city, "country": e_country,
                                            "region": e_region,
                                            "start_date": e_start.isoformat(),
                                            "end_date": e_end.isoformat(),
                                            "notes": e_notes,
                                        }
                                        st.session_state[f"create_editing_stop_{idx}"] = False
                                        st.rerun()
                                    else:
                                        st.warning("City, Start Date, and End Date are required.")
                            with sub_cancel:
                                if st.form_submit_button("❌ Cancel"):
                                    st.session_state[f"create_editing_stop_{idx}"] = False
                                    st.rerun()

        # --- Add Destination Stop expander (always visible) ---
        with st.expander("➕ Add Destination Stop"):
            col_city, col_country = st.columns(2)
            with col_city:
                new_city = st.text_input("City*", key="create_new_stop_city")
            with col_country:
                new_country = st.selectbox("Country (optional)", options=[""] + country_list,
                    key="create_new_stop_country")
            col_region, col_notes = st.columns(2)
            with col_region:
                new_region = st.text_input("Region / State (optional)", key="create_new_stop_region")
            with col_notes:
                new_stop_notes = st.text_input("Notes (optional)", key="create_new_stop_notes")
            col_start, col_end = st.columns(2)
            with col_start:
                new_start = st.date_input("Start Date*", value=datetime.now(), key="create_new_stop_start")
            with col_end:
                new_end = st.date_input("End Date*", value=datetime.now(), key="create_new_stop_end")

            if st.button("➕ Add Stop", key="create_add_stop_button"):
                if not (new_city and new_start and new_end):
                    st.warning("City, Start Date, and End Date are required.")
                else:
                    new_start_iso = new_start.isoformat()
                    new_end_iso = new_end.isoformat()
                    same_city, other_city = _find_stop_conflicts(
                        st.session_state["create_trip_stops"],
                        new_city,
                        new_start_iso,
                        new_end_iso,
                    )

                    proceed = True

                    if other_city:
                        st.error(
                            "⚠️ These dates overlap with stops in a "
                            "**different city** — you can't be in two "
                            "places at once:"
                        )
                        for c in other_city:
                            st.write(
                                f"- {c['city']} · "
                                f"{fmt_date(c['start_date'])} → "
                                f"{fmt_date(c['end_date'])}"
                            )
                        if not st.checkbox(
                            "Add anyway? (dates overlap across cities)",
                            key="force_add_create_stop_other_city",
                        ):
                            proceed = False
                            st.info(
                                "Tick **Add anyway?** above and click "
                                "**➕ Add Stop** again — only do this if "
                                "the overlap is intentional."
                            )

                    elif same_city:
                        st.warning(
                            f"⚠️ You already have a stop in **{new_city}** "
                            "with overlapping dates:"
                        )
                        for c in same_city:
                            st.write(
                                f"- {c['city']} · "
                                f"{fmt_date(c['start_date'])} → "
                                f"{fmt_date(c['end_date'])}"
                            )
                        if not st.checkbox("Add anyway?", key="force_add_create_stop"):
                            proceed = False
                            st.info(
                                "Tick **Add anyway?** above and click "
                                "**➕ Add Stop** again — normal if you "
                                "revisit this city on a separate leg."
                            )

                    if proceed:
                        st.session_state["create_trip_stops"].append(
                            {
                                "city": new_city,
                                "country": new_country,
                                "region": new_region,
                                "start_date": new_start_iso,
                                "end_date": new_end_iso,
                                "notes": new_stop_notes,
                            }
                        )
                        st.session_state.pop("force_add_create_stop", None)
                        st.session_state.pop("force_add_create_stop_other_city", None)
                        st.success(
                            f"Added: {new_city}"
                            + (f", {new_country}" if new_country else "")
                        )
                        st.rerun()

        # ---- Weather preview for planned stops ----
        _create_stops_wx = st.session_state.get("create_trip_stops", [])
        if _create_stops_wx:
            st.write("**🌦️ Weather Preview**")
            st.caption(
                "Weather for the dates you've planned at each stop. "
                "Forecasts cover up to 16 days ahead; trips further out "
                "fall back to last year's data for the same dates."
            )
            for s_idx, stop in enumerate(_create_stops_wx):
                stop_city = (stop.get("city") or "").strip()
                stop_start = stop.get("start_date")
                stop_end = stop.get("end_date")
                if not (stop_city and stop_start and stop_end):
                    continue

                header = (
                    f"🌦️ {stop_city} · "
                    f"{fmt_date(stop_start)} → {fmt_date(stop_end)}"
                )
                cache_key = f"weather_create_{s_idx}"

                with st.expander(header, expanded=False):
                    if st.button(
                        "Fetch weather",
                        key=f"fetch_weather_create_{s_idx}",
                        help=f"Fetch weather for {stop_city} across the planned dates",
                    ):
                        with st.spinner(f"Fetching weather for {stop_city}..."):
                            w = weather.get_weather_for_range(
                                stop_city, stop_start, stop_end
                            )
                        st.session_state[cache_key] = w
                        st.rerun()

                    if cache_key not in st.session_state:
                        st.caption(
                            "Click **Fetch weather** to load the forecast "
                            "for these dates."
                        )
                        continue

                    cached = st.session_state.get(cache_key)
                    if not cached:
                        st.info(
                            f"No weather data available for {stop_city} on "
                            f"{fmt_date(stop_start)} → {fmt_date(stop_end)}."
                        )
                        continue

                    if cached.get("is_historical_fallback"):
                        st.warning(
                            "📅 Forecast not yet available for these dates. "
                            "Showing typical conditions from the same dates "
                            f"last year ({fmt_date(cached.get('start_date'))} → "
                            f"{fmt_date(cached.get('end_date'))})."
                        )
                    else:
                        st.caption(
                            f"📍 {cached['location']} · "
                            f"fetched {fmt_datetime(cached['fetched_at'])}"
                        )

                    cur = cached.get("current")
                    if cur:
                        col_w1, col_w2, col_w3, col_w4 = st.columns(4)
                        with col_w1:
                            st.metric("Now", f"{cur['temp']}°C", cur["icon"])
                        with col_w2:
                            st.metric("Humidity", f"{cur['humidity']}%")
                        with col_w3:
                            st.metric("Wind", f"{cur['wind']} km/h")
                        with col_w4:
                            st.caption(cur["desc"])

                    daily = cached.get("daily") or []
                    if daily:
                        for d in daily:
                            col_d1, col_d2, col_d3, col_d4 = st.columns([2, 1, 1, 1])
                            with col_d1:
                                st.write(f"{d['icon']} {fmt_date(d['date'])}")
                            with col_d2:
                                st.write(f"**{d['max']}°** / {d['min']}°")
                            with col_d3:
                                st.write(f"☔ {d['precip']} mm")
                            with col_d4:
                                st.write(f"UV {d['uv']}")
                    else:
                        st.caption("No daily records in this range.")

        st.subheader("📋 Local Support for This Trip")
        st.caption(
            "People who help at each stop — drivers, fixers, local assistants. "
            "They appear in the travel pack's Local Support section."
        )

        # Country filter still narrows the list
        conn = sqlite3.connect(db.DB_PATH, timeout=30)
        try:
            c = conn.cursor()
            c.execute(
                "SELECT DISTINCT country FROM contacts "
                "WHERE country IS NOT NULL AND country != '' "
                "ORDER BY country"
            )
            countries = [row[0] for row in c.fetchall()]
        finally:
            conn.close()
        country_options = ["All Countries"] + countries
        selected_country = st.selectbox(
            "Filter by country",
            options=country_options,
            index=0,
            key="create_contact_country_filter",
        )
        filter_country = (
            None if selected_country == "All Countries" else selected_country
        )

        # Build the options list from the shared function
        if company_id:
            contact_options = get_local_support_options(company_id)
            # If a country filter is active, narrow the list
            if filter_country:
                allowed_ids = {
                    c["id"]
                    for c in db.get_contacts_for_local_support(company_id)
                    if c.get("country") == filter_country
                }
                contact_options = {
                    label: cid
                    for label, cid in contact_options.items()
                    if cid in allowed_ids
                }
        else:
            contact_options = {}

        if contact_options:
            # Pre-select company defaults
            default_ids = set(db.get_company_default_contacts(company_id))
            current_ids = st.session_state.get("create_trip_contact_ids", [])
            if current_ids:
                pre_selected = [
                    label
                    for label, cid in contact_options.items()
                    if cid in current_ids
                ]
            else:
                pre_selected = [
                    label
                    for label, cid in contact_options.items()
                    if cid in default_ids
                ]

            selected_contact_labels = st.multiselect(
                "Local support contacts for this trip",
                options=list(contact_options.keys()),
                default=pre_selected,
                key="create_trip_contacts",
            )
            st.session_state["create_trip_contact_ids"] = [
                contact_options[label] for label in selected_contact_labels
            ]
        else:
            if filter_country:
                st.warning(
                    f"No local support contacts match {filter_country}. "
                    "Try **All Countries**, or add contacts in the "
                    "**👥 Contacts** tab."
                )
            else:
                st.info(
                    "No contacts available for this company yet. "
                    "Add them in the **👥 Contacts** tab, or use the "
                    "inline add below."
                )
            st.session_state["create_trip_contact_ids"] = []

        st.subheader("👥 Delegation")
        if company_id:

            # ---- Inline add: new person → company roster + auto-select ----
            with st.expander(
                "➕ Add a new person to this company's roster", expanded=False
            ):
                with st.form(key="create_inline_add_delegation_form"):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        new_dm_name = st.text_input("Name*", key="create_dm_name")
                        new_dm_role = st.text_input("Role", key="create_dm_role")
                        new_dm_phone = st.text_input("Phone", key="create_dm_phone")
                    with col_b:
                        new_dm_email = st.text_input("Email", key="create_dm_email")
                        new_dm_country = st.selectbox(
                            "Country (optional)",
                            options=[""] + country_list,
                            key="create_dm_country",
                        )
                        new_dm_type = st.selectbox(
                            "Type",
                            options=[
                                "Staff",
                                "Partner",
                                "Local Support",
                                "Emergency",
                                "Other",
                            ],
                            index=0,
                            key="create_dm_type",
                        )
                    if st.form_submit_button("➕ Add to roster"):
                        if not new_dm_name:
                            st.warning("Name is required.")
                        else:
                            dupes = db.find_duplicate_contacts(
                                company_id,
                                name=new_dm_name,
                                email=new_dm_email or None,
                                phone=new_dm_phone or None,
                            )
                            if dupes:
                                st.warning(
                                    "⚠️ A contact with the same name, email, "
                                    "or phone already exists in this company:"
                                )
                                for d in dupes:
                                    st.write(f"- {d['name']} ({d.get('role','')})")
                                st.info(
                                    "Add them from the list below instead. "
                                    "To force-add a duplicate, use the "
                                    "**👥 Contacts** tab."
                                )
                            else:
                                new_id = db.add_contact(
                                    company_id=company_id,
                                    name=new_dm_name,
                                    role=new_dm_role or None,
                                    phone=new_dm_phone or None,
                                    email=new_dm_email or None,
                                    country=new_dm_country or None,
                                    type=new_dm_type,
                                )
                                # Auto-select the newly created contact
                                current = st.session_state.get(
                                    "create_trip_delegation_ids", []
                                )
                                if new_id not in current:
                                    current.append(new_id)
                                st.session_state["create_trip_delegation_ids"] = current
                                # Force the multiselect to re-read its default
                                st.session_state.pop(
                                    "create_trip_delegation_multiselect", None
                                )
                                st.success(f"✅ {new_dm_name} added and pre-selected.")
                                st.rerun()

            all_company_contacts = db.get_contacts_for_delegation(company_id)
            contact_options = {}
            for c in all_company_contacts:
                role = c.get("role", "") or ""
                # Tag unaffiliated contacts so it's obvious in the dropdown
                suffix = " · freelance" if not c.get("company_id") else ""
                label = f"{c['name']} ({role}){suffix}".strip()
                contact_options[label] = c["id"]
            current_ids = st.session_state.get("create_trip_delegation_ids", [])
            if not current_ids:
                # Nothing picked yet — pre-select company defaults
                default_ids = set(db.get_company_default_contacts(company_id))
                current_labels = [
                    label
                    for label, cid in contact_options.items()
                    if cid in default_ids
                ]
            else:
                current_labels = [
                    label
                    for label, cid in contact_options.items()
                    if cid in current_ids
                ]
            selected_labels = st.multiselect(
                "Travelers on this trip",
                options=list(contact_options.keys()),
                default=current_labels,
                key="create_trip_delegation_multiselect",
            )
            st.session_state["create_trip_delegation_ids"] = [
                contact_options[label] for label in selected_labels]
            if not contact_options:
                st.caption("No contacts for this company yet — add them in the **👥 Contacts** tab first.")
        else:
            st.warning("No company selected – cannot add delegation members.")

        if "create_trip_items" not in st.session_state:
            st.session_state["create_trip_items"] = []

        if st.session_state["create_trip_items"]:
            # --- Batch selection state (indices) ---
            if "create_selected_item_indices" not in st.session_state:
                st.session_state["create_selected_item_indices"] = set()

            # --- Self-healing: drop any stale indices ---
            # The selection set holds positions, not IDs. If the item list
            # ever shrinks without the mutation site resetting the set, any
            # out-of-range index would silently point at the wrong row on
            # the next render. Filtering here keeps the invariant without
            # relying on every future edit site remembering to reset.
            _n_items = len(st.session_state["create_trip_items"])
            st.session_state["create_selected_item_indices"] = {
                _i
                for _i in st.session_state["create_selected_item_indices"]
                if 0 <= _i < _n_items
            }

            # --- Batch toolbar ---
            col_selall, col_delsel, _spacer = st.columns([1.2, 1.6, 4])
            with col_selall:
                total = len(st.session_state["create_trip_items"])
                all_selected = (
                    total > 0
                    and len(st.session_state["create_selected_item_indices"]) == total
                )
                toggle = st.checkbox(
                    "Select all",
                    value=all_selected,
                    key="create_select_all_items",
                )
                if toggle != all_selected:
                    st.session_state["create_selected_item_indices"] = (
                        set(range(total)) if toggle else set()
                    )
                    st.rerun()
            with col_delsel:
                n_sel = len(st.session_state["create_selected_item_indices"])
                if n_sel:
                    if st.button(
                        f"🗑️ Delete {n_sel} selected",
                        type="primary",
                        key="create_bulk_del_items",
                    ):
                        # Pop by index, descending so positions stay valid
                        for i in sorted(
                            st.session_state["create_selected_item_indices"],
                            reverse=True,
                        ):
                            st.session_state["create_trip_items"].pop(i)
                        st.session_state["create_selected_item_indices"] = set()
                        st.session_state.pop("create_select_all_items", None)   # ← new
                        st.success(f"Deleted {n_sel} item(s).")
                        st.rerun()
                else:
                    st.caption("Tick items to enable bulk delete.")
            exec_tz = profile.get("timezone", "America/New_York")
            display_mode = st.session_state.get("create_tz_display_mode", "Home")

            for idx, item in enumerate(st.session_state["create_trip_items"]):
                dt_display = format_item_datetime(item, exec_tz, display_mode)
                (
                    col_cb, col_i1, col_i2,
                    col_up, col_dn,
                    col_i3, col_i4, col_i5,
                ) = st.columns([0.4, 2, 2, 0.4, 0.4, 1, 1, 1])
                with col_cb:
                    is_checked = idx in st.session_state["create_selected_item_indices"]
                    new_state = st.checkbox(
                        "",
                        value=is_checked,
                        key=f"create_sel_item_{idx}",
                    )
                    if new_state != is_checked:
                        if new_state:
                            st.session_state["create_selected_item_indices"].add(idx)
                        else:
                            st.session_state["create_selected_item_indices"].discard(idx)
                        st.session_state.pop("create_select_all_items", None)
                with col_i1:
                    st.write(f"{item['description']} ({item['item_type']})")
                    st.caption(f"🕐 {dt_display}")
                with col_i2:
                    st.write(
                        fmt_money(item.get("cost", 0), item.get("cost_currency", "USD"))
                    )
                    if item.get("delegation_ids"):
                        names = get_contact_names(item["delegation_ids"])
                        if names:
                            st.caption(f"👥 {names}")
                    if item.get("contact_ids"):
                        contact_names = get_contact_names(item["contact_ids"])
                        if contact_names:
                            st.caption(f"📞 {contact_names}")
                with col_up:
                    if idx > 0:
                        if st.button(
                            "↑",
                            key=f"create_item_up_{idx}",
                            help="Swap times with the item above",
                        ):
                            lst = st.session_state["create_trip_items"]
                            lst[idx - 1], lst[idx] = _swap_item_times(
                                lst[idx - 1], lst[idx]
                            )
                            st.rerun()
                with col_dn:
                    lst_len = len(st.session_state["create_trip_items"])
                    if idx < lst_len - 1:
                        if st.button(
                            "↓",
                            key=f"create_item_dn_{idx}",
                            help="Swap times with the item below",
                        ):
                            lst = st.session_state["create_trip_items"]
                            lst[idx], lst[idx + 1] = _swap_item_times(
                                lst[idx], lst[idx + 1]
                            )
                            st.rerun()
                with col_i3:
                    if st.button(
                        "✏️", key=f"create_edit_item_{idx}", help="Edit this item"
                    ):
                        st.session_state[f"create_editing_item_{idx}"] = True
                with col_i4:
                    if st.button(
                        "📄",
                        key=f"create_dup_item_{idx}",
                        help="Duplicate this item (+1 day)",
                    ):
                        original = st.session_state["create_trip_items"][idx]
                        try:
                            orig_start = datetime.fromisoformat(
                                original["datetime_start"]
                            )
                            orig_end = (
                                datetime.fromisoformat(original["datetime_end"])
                                if original.get("datetime_end")
                                else None
                            )
                        except Exception:
                            orig_start, orig_end = datetime.now(), None

                        duplicate = dict(original)
                        duplicate["description"] = (
                            original.get("description", "") + " (copy)"
                        ).strip()
                        duplicate["datetime_start"] = (
                            orig_start + timedelta(days=1)
                        ).isoformat()
                        duplicate["datetime_end"] = (
                            (orig_end + timedelta(days=1)).isoformat()
                            if orig_end
                            else None
                        )
                        # Copy assignment lists (safe: they're lists of ints)
                        duplicate["delegation_ids"] = list(
                            original.get("delegation_ids", [])
                        )
                        duplicate["contact_ids"] = list(original.get("contact_ids", []))
                        st.session_state["create_trip_items"].append(duplicate)
                        st.success(f"✅ Duplicated: {duplicate['description']}")
                        st.rerun()
                with col_i5:
                    if st.button("🗑️", key=f"create_del_item_{idx}",
                                 help="Delete this item"):
                        st.session_state[
                            f"confirm_del_create_item_{idx}"
                        ] = True
                        st.rerun()

                # --- Delete confirmation for this row ---
                if st.session_state.get(
                    f"confirm_del_create_item_{idx}", False
                ):
                    def _del_create_item(i=idx):
                        st.session_state["create_trip_items"].pop(i)
                        st.session_state["create_selected_item_indices"] = set()

                    render_delete_confirmation(
                        flag_key=f"confirm_del_create_item_{idx}",
                        title=f"Remove item '{item['description']}'?",
                        on_confirm=_del_create_item,
                    )

                if st.session_state.get(f"create_editing_item_{idx}", False):
                    with st.expander(f"Edit Item: {item['description']}", expanded=True):

                        with st.expander("➕ Need a new venue? Add it here first",
                                         expanded=False):
                            with st.form(
                                key=f"create_edit_inline_add_venue_form_{idx}"
                            ):
                                col_va, col_vb = st.columns(2)
                                with col_va:
                                    ce_v_name = st.text_input(
                                        "Venue Name*",
                                        key=f"create_e_v_name_{idx}",
                                    )
                                    ce_v_address = st.text_input(
                                        "Address",
                                        key=f"create_e_v_address_{idx}",
                                    )
                                    ce_v_city = st.text_input(
                                        "City",
                                        key=f"create_e_v_city_{idx}",
                                    )
                                with col_vb:
                                    ce_v_country = st.selectbox(
                                        "Country",
                                        options=[""] + country_list,
                                        key=f"create_e_v_country_{idx}",
                                    )
                                    ce_v_wifi_ssid = st.text_input(
                                        "WiFi SSID",
                                        key=f"create_e_v_wifi_ssid_{idx}",
                                    )
                                    ce_v_wifi_pw = st.text_input(
                                        "WiFi Password",
                                        key=f"create_e_v_wifi_pw_{idx}",
                                    )
                                ce_v_notes = st.text_area(
                                    "Notes",
                                    key=f"create_e_v_notes_{idx}",
                                    height=80,
                                )
                                if st.form_submit_button("➕ Add venue"):
                                    if not ce_v_name:
                                        st.warning("Venue Name is required.")
                                    else:
                                        db.add_venue(
                                            name=ce_v_name,
                                            address=ce_v_address,
                                            city=ce_v_city,
                                            country=ce_v_country,
                                            wifi_ssid=ce_v_wifi_ssid,
                                            wifi_password=ce_v_wifi_pw,
                                            notes=ce_v_notes,
                                        )
                                        st.success(
                                            f"✅ Venue '{ce_v_name}' added."
                                        )
                                        st.rerun()

                        with st.form(key=f"create_edit_item_form_{idx}"):
                            e_type = st.selectbox("Type",
                                options=([cat["name"] for cat in db.get_all_categories()]
                                         if db.get_all_categories() else ["Flight", "Hotel", "Meeting", "Transport"]),
                                index=0, key=f"create_e_type_{idx}")
                            e_desc = st.text_input("Description", value=item["description"], key=f"create_e_desc_{idx}")
                            e_start = st.datetime_input("Start",
                                value=datetime.fromisoformat(item["datetime_start"]), key=f"create_e_start_{idx}")
                            e_end = st.datetime_input("End",
                                value=(datetime.fromisoformat(item["datetime_end"]) if item["datetime_end"] else datetime.now()),
                                key=f"create_e_end_{idx}")
                            e_loc = st.text_input("Location", value=item.get("location", ""), key=f"create_e_loc_{idx}")
                            e_cost = st.number_input("Cost", value=float(item.get("cost", 0)), key=f"create_e_cost_{idx}")
                            base_currency_options = currency.CURRENCIES_SUPPORTED
                            e_currency = st.selectbox("Currency", options=base_currency_options,
                                index=(base_currency_options.index(item.get("cost_currency", "USD"))
                                       if item.get("cost_currency", "USD") in base_currency_options else 0),
                                key=f"create_e_currency_{idx}")
                            default_cost_date = item.get("cost_date")
                            if default_cost_date:
                                try:
                                    default_cost_date = datetime.fromisoformat(default_cost_date).date()
                                except Exception:
                                    default_cost_date = datetime.fromisoformat(item["datetime_start"]).date()
                            else:
                                default_cost_date = datetime.fromisoformat(item["datetime_start"]).date()
                            e_cost_date = st.date_input("Cost Date", value=default_cost_date,
                                key=f"create_e_cost_date_{idx}")
                            tz_display_names, tz_map = get_timezone_dropdown_options()
                            current_tz = item.get("timezone") or profile.get("timezone", "America/New_York")
                            current_tz_display = next(
                                (n for n in tz_display_names if current_tz in n), tz_display_names[0])
                            e_timezone = st.selectbox("Time Zone", options=tz_display_names,
                                index=tz_display_names.index(current_tz_display), key=f"create_e_timezone_{idx}")
                            e_timezone_value = tz_map[e_timezone]
                            e_confirmed = st.checkbox("Confirmed", value=bool(item.get("is_confirmed", 0)),
                                key=f"create_e_confirmed_{idx}")
                            e_notes = st.text_area("Notes", value=item.get("notes", ""), key=f"create_e_notes_{idx}")
                            session_types = ["Meeting", "Conference", "Dinner", "Site Visit", "Tour", "Activity"]
                            e_venue_id = item.get("venue_id")
                            if e_type in session_types:
                                venue_options = db.get_venues(active_only=True)
                                venue_labels = {f"{v['name']}" + (f" — {v['city']}" if v.get("city") else ""): v["id"]
                                    for v in venue_options}
                                venue_labels["(No venue)"] = None
                                current_venue_label = "(No venue)"
                                for lbl, vid in venue_labels.items():
                                    if vid == e_venue_id:
                                        current_venue_label = lbl
                                        break
                                selected_venue_label = st.selectbox("Venue (optional)",
                                    options=list(venue_labels.keys()),
                                    index=list(venue_labels.keys()).index(current_venue_label),
                                    key=f"create_e_venue_{idx}")
                                e_venue_id = venue_labels[selected_venue_label]

                            if st.session_state.get("create_trip_delegation_ids"):
                                delegation_options = {}
                                for c in db.get_contacts_for_delegation(company_id):
                                    if c["id"] not in st.session_state["create_trip_delegation_ids"]:
                                        continue
                                    role = c.get("role", "") or ""
                                    suffix = " · freelance" if not c.get("company_id") else ""
                                    label = f"{c['name']} ({role}){suffix}".strip()
                                    delegation_options[label] = c["id"]
                                current_delegation = item.get("delegation_ids", [])
                                current_labels = [label for label, cid in delegation_options.items()
                                    if cid in current_delegation]
                                selected_parts = st.multiselect("Assign Delegation Members",
                                    options=list(delegation_options.keys()), default=current_labels,
                                    key=f"create_e_delegation_{idx}")
                                e_delegation_ids = [delegation_options[l] for l in selected_parts]
                            else:
                                e_delegation_ids = []

                            if company_id:
                                contact_options = get_company_contact_options(company_id)
                                current_contact_ids = item.get("contact_ids", [])
                                current_labels = [label for label, cid in contact_options.items()
                                    if cid in current_contact_ids]
                                selected_contacts = st.multiselect("Assign Local Support Contacts",
                                    options=list(contact_options.keys()), default=current_labels,
                                    key=f"create_e_contacts_{idx}")
                                e_contact_ids = [contact_options[l] for l in selected_contacts]
                            else:
                                e_contact_ids = []

                            if st.form_submit_button("💾 Update Item"):
                                proceed = True

                                # --- Time sanity check ---
                                time_ok, time_msg = _validate_item_times(
                                    e_start.isoformat(),
                                    e_end.isoformat() if e_end else None,
                                )
                                if not time_ok:
                                    st.warning(f"⚠️ {time_msg}")
                                    if not st.checkbox(
                                        "Save anyway? (end before start)",
                                        key=f"force_edit_create_item_time_{idx}",
                                    ):
                                        proceed = False
                                        st.info(
                                            "Tick **Save anyway?** above "
                                            "and click **💾 Update Item** "
                                            "again to save."
                                        )

                                # --- Duplicate check against other rows ---
                                dupes = _find_duplicate_items(
                                    [
                                        it
                                        for j, it in enumerate(
                                            st.session_state["create_trip_items"]
                                        )
                                        if j != idx
                                    ],
                                    e_desc,
                                    e_start.isoformat(),
                                )
                                if dupes:
                                    st.warning(
                                        "⚠️ Another item in this trip has "
                                        "the same description and start "
                                        "time:"
                                    )
                                    for d in dupes:
                                        st.write(
                                            f"- {d['description']} · "
                                            f"{d['datetime_start'][:16].replace('T', ' ')}"
                                        )
                                    if not st.checkbox(
                                        "Update anyway?",
                                        key=f"force_edit_create_item_{idx}",
                                    ):
                                        proceed = False
                                        st.info(
                                            "Tick **Update anyway?** above "
                                            "and click **💾 Update Item** "
                                            "again to save."
                                        )

                                if proceed:
                                    st.session_state["create_trip_items"][idx] = {
                                        "item_type": e_type,
                                        "description": e_desc,
                                        "datetime_start": e_start.isoformat(),
                                        "datetime_end": (
                                            e_end.isoformat() if e_end else None
                                        ),
                                        "location": e_loc,
                                        "cost": e_cost,
                                        "cost_currency": e_currency,
                                        "is_confirmed": 1 if e_confirmed else 0,
                                        "confirmation_code": item.get(
                                            "confirmation_code", ""
                                        ),
                                        "notes": e_notes,
                                        "delegation_ids": e_delegation_ids,
                                        "contact_ids": e_contact_ids,
                                        "timezone": e_timezone_value,
                                        "venue_id": e_venue_id,
                                        "cost_date": e_cost_date.isoformat(),
                                    }
                                    st.session_state.pop(
                                        f"force_edit_create_item_time_{idx}", None
                                    )
                                    st.session_state.pop(
                                        f"force_edit_create_item_{idx}", None
                                    )
                                    st.session_state[
                                        f"create_editing_item_{idx}"
                                    ] = False
                                    st.rerun()


                            if st.form_submit_button("❌ Cancel"):
                                st.session_state[f"create_editing_item_{idx}"] = False
                                st.rerun()




        # ---- Inline add: new venue → library ----
        with st.expander("➕ Add a new venue to the library", expanded=False):
            with st.form(key="create_inline_add_venue_form"):
                col_va, col_vb = st.columns(2)
                with col_va:
                    new_v_name = st.text_input(
                        "Venue Name*", key="create_v_name"
                    )
                    new_v_address = st.text_input(
                        "Address", key="create_v_address"
                    )
                    new_v_city = st.text_input(
                        "City", key="create_v_city"
                    )
                with col_vb:
                    new_v_country = st.selectbox(
                        "Country",
                        options=[""] + country_list,
                        key="create_v_country",
                    )
                    new_v_wifi_ssid = st.text_input(
                        "WiFi SSID", key="create_v_wifi_ssid"
                    )
                    new_v_wifi_pw = st.text_input(
                        "WiFi Password", key="create_v_wifi_pw"
                    )
                new_v_notes = st.text_area(
                    "Notes", key="create_v_notes", height=80
                )
                if st.form_submit_button("➕ Add venue"):
                    if not new_v_name:
                        st.warning("Venue Name is required.")
                    else:
                        db.add_venue(
                            name=new_v_name,
                            address=new_v_address,
                            city=new_v_city,
                            country=new_v_country,
                            wifi_ssid=new_v_wifi_ssid,
                            wifi_password=new_v_wifi_pw,
                            notes=new_v_notes,
                        )
                        st.success(f"✅ Venue '{new_v_name}' added.")
                        st.rerun()

        # ---- Bulk paste (create form) ----
        def _import_into_create_session(items):
            st.session_state["create_trip_items"].extend(items)
            st.success(f"Added {len(items)} item(s).")

        _render_bulk_paste_itinerary(
            container_key="create_bulk_paste",
            default_timezone=(
                profile.get("timezone", "America/New_York")
                if profile else "UTC"
            ),
            on_import=_import_into_create_session,
        )

        with st.expander("➕ Add Itinerary Item"):
            with st.form(key="create_add_item_form"):
                n_type = st.selectbox("Type",
                    options=([cat["name"] for cat in db.get_all_categories()]
                             if db.get_all_categories() else ["Flight", "Hotel", "Meeting", "Transport"]),
                    key="create_n_type")
                n_desc = st.text_input("Description", key="create_n_desc")
                n_start = st.datetime_input("Start", value=datetime.now(), key="create_n_start")
                n_end = st.datetime_input("End", value=datetime.now(), key="create_n_end")
                n_loc = st.text_input("Location", key="create_n_loc")
                n_cost = st.number_input("Cost", min_value=0.0, value=0.0, key="create_n_cost")
                base_currency_options = currency.CURRENCIES_SUPPORTED
                n_currency = st.selectbox("Currency", options=base_currency_options, key="create_n_currency")
                n_cost_date = st.date_input("Cost Date", value=datetime.now(), key="create_n_cost_date")
                tz_display_names, tz_map = get_timezone_dropdown_options()
                if profile and profile.get("timezone"):
                    default_tz_display = next(
                        (n for n in tz_display_names if profile["timezone"] in n), tz_display_names[0])
                else:
                    default_tz_display = tz_display_names[0]
                n_timezone = st.selectbox("Time Zone (for this event)",
                    options=tz_display_names, index=tz_display_names.index(default_tz_display),
                    key="create_n_timezone")
                n_timezone_value = tz_map[n_timezone]
                n_confirmed = st.checkbox("Confirmed", key="create_n_confirmed")
                n_notes = st.text_area("Notes", key="create_n_notes")
                session_types = ["Meeting", "Conference", "Dinner", "Site Visit", "Tour", "Activity"]
                n_venue_id = None
                if n_type in session_types:
                    venue_options = db.get_venues(active_only=True)
                    venue_labels = {f"{v['name']}" + (f" — {v['city']}" if v.get("city") else ""): v["id"]
                        for v in venue_options}
                    venue_labels["(No venue)"] = None
                    selected_venue_label = st.selectbox("Venue (optional)",
                        options=list(venue_labels.keys()),
                        index=list(venue_labels.keys()).index("(No venue)"), key="create_n_venue")
                    n_venue_id = venue_labels[selected_venue_label]

                if company_id and st.session_state.get("create_trip_delegation_ids"):
                    delegation_options = {}
                    for c in db.get_contacts_for_delegation(company_id):
                        if c["id"] not in st.session_state["create_trip_delegation_ids"]:
                            continue
                        role = c.get("role", "") or ""
                        suffix = " · freelance" if not c.get("company_id") else ""
                        label = f"{c['name']} ({role}){suffix}".strip()
                        delegation_options[label] = c["id"]
                    selected_delegation = st.multiselect("Assign Delegation Members",
                        options=list(delegation_options.keys()), key="create_item_delegation")
                    selected_delegation_ids = [delegation_options[l] for l in selected_delegation]
                else:
                    selected_delegation_ids = []

                if company_id:
                    contact_options = get_company_contact_options(company_id)
                    selected_contacts = st.multiselect("Assign Local Support Contacts",
                        options=list(contact_options.keys()), key="create_item_contacts")
                    selected_contact_ids = [contact_options[l] for l in selected_contacts]
                else:
                    selected_contact_ids = []

                if st.form_submit_button("➕ Add Item"):
                    if not (n_desc and n_start):
                        st.warning("Description and Start Time are required.")
                    else:
                        proceed = True

                        # --- Time sanity check ---
                        time_ok, time_msg = _validate_item_times(
                            n_start.isoformat(),
                            n_end.isoformat() if n_end else None,
                        )
                        if not time_ok:
                            st.warning(f"⚠️ {time_msg}")
                            if not st.checkbox(
                                "Save anyway? (end before start)",
                                key="force_add_create_item_time",
                            ):
                                proceed = False
                                st.info(
                                    "Tick **Save anyway?** above and click "
                                    "**➕ Add Item** again to save."
                                )

                        # --- Duplicate check ---
                        dupes = _find_duplicate_items(
                            st.session_state["create_trip_items"],
                            n_desc,
                            n_start.isoformat(),
                        )
                        if dupes:
                            st.warning(
                                "⚠️ You already have an item with the same "
                                "description and start time:"
                            )
                            for d in dupes:
                                st.write(
                                    f"- {d['description']} · "
                                    f"{d['datetime_start'][:16].replace('T', ' ')}"
                                )
                            if not st.checkbox(
                                "Add anyway?",
                                key="force_add_create_item",
                            ):
                                proceed = False
                                st.info(
                                    "Tick **Add anyway?** above and click "
                                    "**➕ Add Item** again to save the "
                                    "duplicate."
                                )

                        if proceed:
                            st.session_state["create_trip_items"].append(
                                {
                                    "item_type": n_type,
                                    "description": n_desc,
                                    "datetime_start": n_start.isoformat(),
                                    "datetime_end": (
                                        n_end.isoformat() if n_end else None
                                    ),
                                    "location": n_loc,
                                    "cost": n_cost,
                                    "cost_currency": n_currency,
                                    "is_confirmed": 1 if n_confirmed else 0,
                                    "confirmation_code": "",
                                    "notes": n_notes,
                                    "delegation_ids": selected_delegation_ids,
                                    "contact_ids": selected_contact_ids,
                                    "timezone": n_timezone_value,
                                    "venue_id": n_venue_id,
                                    "cost_date": n_cost_date.isoformat(),
                                }
                            )
                            st.session_state.pop("force_add_create_item", None)
                            st.session_state.pop("force_add_create_item_time", None)
                            st.rerun()

        col_clear, col_create = st.columns(2)
        with col_clear:
            if st.button("🗑️ Clear Form", key="clear_create_form"):
                st.session_state["create_trip_stops"] = []
                st.session_state["create_trip_items"] = []
                st.session_state["create_selected_item_indices"] = set()   # ← new
                st.session_state["create_trip_delegation_ids"] = []
                st.session_state["create_trip_contact_ids"] = []
                for key in [
                    "create_trip_purpose",
                    "create_departure_city",
                    "create_departure_region",
                    "create_departure_country",
                    "create_trip_budget",
                    "create_base_currency",
                    "create_trip_status",
                    "create_overall_start",
                    "create_overall_end",
                    "create_receipt_folder",
                ]:
                    st.session_state.pop(key, None)
                st.rerun()

        with col_create:
            if st.button("🚀 Create Trip", key="create_trip_button"):
                if trip_purpose and st.session_state["create_trip_stops"]:
                    overall_start_iso = overall_start.isoformat()
                    overall_end_iso = overall_end.isoformat()
                    stop_cities = [stop["city"] for stop in st.session_state["create_trip_stops"]]
                    dest_summary = " → ".join(stop_cities)

                    proceed = True
                    existing_trips = duplicate_detection.find_duplicate_trips(
                        trip_exec_id, trip_purpose, overall_start_iso, overall_end_iso)
                    if existing_trips:
                        st.warning("⚠️ You already have a trip with the same purpose and overlapping dates:")
                        for dup in existing_trips:
                            st.write(f"- {dup['destination']} ({fmt_date(dup['start_date'])} to {fmt_date(dup['end_date'])})")
                        if not st.checkbox("Proceed anyway?", key="force_trip_create"):
                            proceed = False
                            st.info("Tick **Proceed anyway?** above and click **🚀 Create Trip** again to save the duplicate.")

                    if proceed:
                        budget_base = budget
                        selected_contact_ids = st.session_state.get("create_trip_contact_ids", [])
                        delegation_contact_ids = st.session_state.get("create_trip_delegation_ids", [])

                        trip_id = db.create_or_get_trip(trip_exec_id, dest_summary,
                            overall_start_iso, overall_end_iso, trip_purpose,
                            trip_base_currency, trip_base_currency, trip_status,
                            trip_contacts=selected_contact_ids)
                        db.update_trip_budget(trip_id, budget_base)
                        db.update_trip_departure_details(trip_id, departure_city, departure_region, departure_country)
                        db.update_trip_status(trip_id, trip_status)
                        db.update_trip_receipt_folder(trip_id, receipt_upload_folder)
                        db.set_trip_delegation(trip_id, delegation_contact_ids)

                        db.delete_all_trip_stops(trip_id)
                        for idx, stop in enumerate(st.session_state["create_trip_stops"]):
                            db.add_trip_stop(trip_id, idx + 1, stop["city"], stop.get("country", ""),
                                stop.get("region", ""), stop["start_date"], stop["end_date"],
                                stop.get("notes", ""))

                        for item in st.session_state["create_trip_items"]:
                            item_id = db.add_itinerary_item(trip_id, item["item_type"], item["description"],
                                item["datetime_start"], item["datetime_end"], item.get("location", ""),
                                item.get("cost", 0), item.get("confirmation_code", ""), item.get("notes", ""),
                                item.get("is_confirmed", 0), item["cost_currency"],
                                timezone=item.get("timezone"), venue_id=item.get("venue_id"),
                                cost_date=item.get("cost_date"))
                            if item.get("delegation_ids"):
                                db.set_item_delegation_members(item_id, item["delegation_ids"])
                            if item.get("contact_ids"):
                                db.set_item_contacts(item_id, item["contact_ids"])

                        st.session_state["create_trip_stops"] = []
                        st.session_state["create_trip_items"] = []
                        st.session_state["create_selected_item_indices"] = set()   # ← new
                        st.session_state["create_trip_delegation_ids"] = []
                        st.session_state["create_trip_contact_ids"] = []
                        for key in ["create_trip_purpose", "create_departure_city",
                                    "create_departure_region", "create_departure_country",
                                    "create_trip_budget", "create_base_currency",
                                    "create_trip_status", "create_overall_start",
                                    "create_overall_end", "create_receipt_folder"]:
                            st.session_state.pop(key, None)

                        st.success(f"✅ Trip '{trip_purpose}' created successfully with status '{trip_status}'!")
                        st.session_state["last_created_trip_id"] = trip_id
                        st.session_state["last_created_trip_name"] = trip_purpose

                        col_st, col_cont = st.columns(2)
                        with col_st:
                            if st.button("📋 Save as Template", key="save_template_after_create"):
                                st.session_state["show_save_template_after_create"] = True
                        with col_cont:
                            if st.button("Continue", key="continue_after_create"):
                                st.session_state.pop("last_created_trip_id", None)
                                st.session_state.pop("last_created_trip_name", None)
                                st.session_state.pop("show_save_template_after_create", None)
                                st.rerun()

                        if st.session_state.get("show_save_template_after_create", False):
                            st.info("Save this trip as a reusable template.")
                            template_name = st.text_input("Template Name*",
                                value=f"{trip_purpose} Template", key="template_name_after_create")
                            template_desc = st.text_input("Description (optional)", key="template_desc_after_create")
                            cy, cn = st.columns(2)
                            with cy:
                                if st.button("💾 Save", key="confirm_save_after_create"):
                                    if template_name:
                                        new_id = db.save_trip_as_template(trip_id, template_name, template_desc)
                                        if new_id:
                                            st.success(f"✅ Template '{template_name}' saved!")
                                            st.session_state.pop("show_save_template_after_create", None)
                                            st.session_state.pop("last_created_trip_id", None)
                                            st.session_state.pop("last_created_trip_name", None)
                                            st.rerun()
                                        else:
                                            st.error("Failed to save template.")
                                    else:
                                        st.warning("Template Name is required.")
                            with cn:
                                if st.button("Cancel", key="cancel_save_after_create"):
                                    st.session_state.pop("show_save_template_after_create", None)
                                    st.rerun()
                else:
                    st.warning("Enter a Trip Name and add at least one stop.")


# =========================================================
# TAB: ALL TRIPS
# =========================================================

with tab_all_trips:
    col_title, col_from_tpl = st.columns([5, 1.4], vertical_alignment="bottom")
    with col_from_tpl:
        if st.button(
            "📋 From Template",
            use_container_width=True,
            key="trips_from_template_btn",
            help="Jump to Library → 📋 Trip Templates to create a trip "
                 "from an existing template.",
        ):
            st.session_state["lib_hint_templates"] = True
            st.rerun()

    search_trip = st.text_input("🔍 Search Trips",
        placeholder="Destination, purpose, or executive name...", key="dash_search")

    col_dash1, col_dash2, col_dash3 = st.columns([2, 2, 1])
    with col_dash1:
        exec_filter_options = ["All"] + [
            f"{e['name']} (ID: {e['id']})"
            for e in db.get_all_executives(active_only=False)
        ]
        exec_filter = st.selectbox(
            "Filter by Executive", exec_filter_options, key="dash_filter_tab"
        )
        exec_id_filter = (
            None
            if exec_filter == "All"
            else int(exec_filter.split("(ID: ")[1].rstrip(")"))
        )
    with col_dash2:
        date_range = st.date_input(
            "Date Range (optional)", value=[], key="dash_date_tab"
        )
    with col_dash3:
        st.write("")
        st.write("")
        show_past = st.checkbox("Include past trips", value=True, key="dash_show_past")
    start_filter = date_range[0].isoformat() if len(date_range) > 0 else None
    end_filter = date_range[1].isoformat() if len(date_range) > 1 else None

    summary_data = db.get_spending_summary(exec_id=exec_id_filter,
        start_date=start_filter, end_date=end_filter)

    # ---- Filter out past trips when the toggle is off
    if not show_past:
        today_iso = datetime.now().date().isoformat()
        summary_data = [
            trip
            for trip in summary_data
            if not (trip.get("end_date") and trip["end_date"][:10] < today_iso)
        ]

    if search_trip:
        sl = search_trip.lower()
        summary_data = [trip for trip in summary_data
            if sl in trip.get("destination", "").lower()
            or sl in trip.get("purpose", "").lower()
            or sl in trip.get("executive_name", "").lower()]

    if "selected_trip_ids" not in st.session_state:
        st.session_state.selected_trip_ids = set()

    # Defensive: dedupe trips by trip_id. Coerce the key to a stable
    # string so `5` and `"5"` collapse, and drop rows with no ID.
    if summary_data:
        seen = set()
        deduped = []
        for t in summary_data:
            raw_id = t.get("trip_id")
            if raw_id is None:
                continue
            tid_key = str(raw_id)
            if tid_key in seen:
                continue
            seen.add(tid_key)
            deduped.append(t)
        summary_data = deduped

    if summary_data:
        st.subheader("Trip-Level Breakdown")

        col_select_all, col_delete_selected = st.columns([1, 3])
        with col_select_all:
            all_ids = [trip["trip_id"] for trip in summary_data]
            all_selected = all(id in st.session_state.selected_trip_ids for id in all_ids)
            if st.checkbox("Select All", value=all_selected, key="select_all_checkbox"):
                if not all_selected:
                    st.session_state.selected_trip_ids = set(all_ids)
                    st.rerun()
            else:
                if all_selected:
                    st.session_state.selected_trip_ids = set()
                    st.rerun()
        with col_delete_selected:
            if st.session_state.selected_trip_ids:
                st.write(f"**{len(st.session_state.selected_trip_ids)}** trip(s) selected.")
                if st.button("🗑️ Delete Selected", type="primary"):
                    st.session_state["confirm_mass_delete"] = True
            else:
                st.write("No trips selected.")

        if st.session_state.get("confirm_mass_delete", False):
            render_delete_confirmation(
                flag_key="confirm_mass_delete",
                title=f"Permanently delete {len(st.session_state.selected_trip_ids)} selected trip(s)?",
                on_confirm=lambda: db.delete_trips(list(st.session_state.selected_trip_ids)),
            )

        cols_widths = [0.4, 5, 0.8, 0.8]
        header_cols = st.columns(cols_widths)
        for label, idx in [
            ("", 0),
            ("Trip", 1),
            ("Open", 2),
            ("Delete", 3),
        ]:
            with header_cols[idx]:
                if label:
                    st.write(f"**{label}**")

        _rendered_trip_ids = set()
        for trip in summary_data:
            trip_base_currency = trip.get("base_currency", "USD")
            trip_id = trip["trip_id"]
            # Final guard: if upstream produced a duplicate, skip it
            # here rather than emitting a colliding widget key.
            _tid_key = str(trip_id)
            if _tid_key in _rendered_trip_ids:
                continue
            _rendered_trip_ids.add(_tid_key)
            trip_end = trip.get("end_date")
            is_past = bool(
                trip_end and trip_end[:10] < datetime.now().date().isoformat()
            )

            status = trip.get("status", "draft")
            if is_past and status == "draft":
                status_label = "🕒 Past (draft)"
            elif is_past:
                base_label = (
                    "✅ Approved" if status == "approved"
                    else "📄 Final" if status == "final"
                    else status
                )
                status_label = f"🕒 {base_label}"
            else:
                status_label = (
                    "📝 Draft" if status == "draft"
                    else "✅ Approved" if status == "approved"
                    else "📄 Final" if status == "final"
                    else status
                )

            with st.container():
                cols = st.columns(cols_widths)

                with cols[0]:
                    is_checked = trip_id in st.session_state.selected_trip_ids
                    if st.checkbox("", value=is_checked, key=f"sel_{trip_id}"):
                        st.session_state.selected_trip_ids.add(trip_id)
                    else:
                        st.session_state.selected_trip_ids.discard(trip_id)

                with cols[1]:
                    # Trip Name (Purpose) as the primary line
                    st.markdown(f"### {trip.get('purpose') or 'Untitled Trip'}")
                    # Executive and Company as secondary caption
                    st.caption(f"👤 {trip['executive_name']} · {trip['company_name']}")
                    # Destination and Status
                    st.write(f"📍 {trip['destination']}  ·  {status_label}")
                    # Financials
                    st.caption(
                        f"Budget: {fmt_money(trip['budget'], trip_base_currency)}  ·  "
                        f"Spent: {fmt_money(trip['total_spent'], trip_base_currency)}  ·  "
                        f"Confirmed: {fmt_money(trip['confirmed_spent'], trip_base_currency)}  ·  "
                        f"Estimated: {fmt_money(trip['estimated_spent'], trip_base_currency)}"
                    )

                with cols[2]:
                    if st.button(
                        "📂 Edit",
                        key=f"open_trip_{trip_id}",
                        use_container_width=True,
                    ):
                        st.session_state["dialog_trip_id"] = trip_id

                with cols[3]:
                    if st.button("🗑️", key=f"del_trip_dash_{trip_id}"):
                        st.session_state[f"confirm_del_trip_{trip_id}"] = True

                if st.session_state.get(f"confirm_del_trip_{trip_id}", False):

                    def _del_trip():
                        db.delete_trip(trip_id)
                        st.session_state.selected_trip_ids.discard(trip_id)

                    render_delete_confirmation(
                        flag_key=f"confirm_del_trip_{trip_id}",
                        title=f"Permanently delete trip to {trip['destination']}?",
                        on_confirm=_del_trip,
                    )
                st.divider()

        st.subheader("📊 Export Data")
        col_exp1, col_exp2, col_exp3, col_exp4 = st.columns(4)
        with col_exp1:
            headers = [
                "Executive",
                "Company",
                "Destination",
                "Budget",
                "Total Spent",
                "Confirmed",
                "Pending",
                "Status",
                "Base Currency",
            ]
            rows_with_currency = []
            for trip in summary_data:
                rows_with_currency.append(
                    [
                        trip["executive_name"],
                        trip["company_name"],
                        trip["destination"],
                        f"{trip['budget']:.2f}",
                        f"{trip['total_spent']:.2f}",
                        f"{trip['confirmed_spent']:.2f}",
                        f"{trip['estimated_spent']:.2f}",
                        trip["status"],
                        trip.get("base_currency", "USD"),
                    ]
                )
            output = io.StringIO()
            writer = csv.writer(output)
            writer.writerow(headers)
            for row in rows_with_currency:
                writer.writerow(row)
            st.download_button(
                "📊 Export Trips CSV",
                data=output.getvalue().encode("utf-8"),
                file_name=f"trips_summary_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                key="dash_csv_tab",
            )
        with col_exp2:
            if st.button("📄 Export Spending Report (Word)", key="dash_report_tab"):
                doc_stream = doc_generator.generate_spending_report_doc(
                    exec_filter if exec_filter != "All" else "All Executives",
                    summary_data,
                    start_filter,
                    end_filter,
                    get_currency_symbol("USD"),
                    base_currency="USD",
                )
                st.download_button(
                    "⬇️ Download Word Report",
                    data=doc_stream,
                    file_name=f"spending_report_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml",
                    key="dash_report_download_tab",
                )
        with col_exp3:
            if st.button("📊 Export to Excel", key="dash_excel_tab"):
                excel_stream = export_spending_to_excel(summary_data, "USD")
                if excel_stream:
                    st.download_button(
                        "⬇️ Download .xlsx",
                        data=excel_stream,
                        file_name=f"trips_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dash_excel_download_tab",
                    )
        with col_exp4:
            if st.button("📅 Export All to Calendar", key="dash_ics_tab"):
                # Build the payload: one entry per trip in the filtered set,
                # carrying that trip's items and executive timezone.
                trips_with_items = []
                for trip_row in summary_data:
                    trip_id = trip_row["trip_id"]
                    full_trip = db.get_trip(trip_id)
                    if not full_trip:
                        continue
                    items = db.get_items_for_trip(trip_id)
                    exec_profile = db.get_executive_profile(full_trip["exec_id"])
                    exec_tz = (
                        exec_profile.get("timezone", "UTC") if exec_profile else "UTC"
                    )
                    trips_with_items.append(
                        {
                            "trip": full_trip,
                            "items": items,
                            "exec_timezone": exec_tz,
                        }
                    )

                if not trips_with_items:
                    st.warning("No trips to export.")
                else:
                    ics_bytes = utils.generate_ics_for_trips(trips_with_items)
                    if ics_bytes:
                        st.download_button(
                            "⬇️ Download .ics",
                            data=ics_bytes,
                            file_name=(
                                f"All_Trips_Calendar_"
                                f"{datetime.now().strftime('%Y%m%d')}.ics"
                            ),
                            mime="text/calendar",
                            key="dash_ics_download_tab",
                        )
                    else:
                        st.error("Failed to generate calendar file.")
    else:
        st.info("No trips match the filters.")

    # ---- Open the trip edit dialog (if a trip is selected) ----
    # This runs once per script, regardless of how many trips are listed.
    # The dialog only appears when "dialog_trip_id" is set, which happens
    # when the user clicks a row's 📂 Edit button above.
    if st.session_state.get("dialog_trip_id"):
        _render_trip_edit_modal(
            st.session_state["dialog_trip_id"], country_list
        )


# =========================================================
# TAB: COMPANIES
# =========================================================

with tab_companies:
    with st.expander("➕ Add New Company", expanded=False):
        with st.form("add_company_manager_form"):
            new_name = st.text_input("Company Name*", key="mgr_comp_name")
            new_cc = st.text_input("Default Cost Center (optional)", key="mgr_comp_cc")
            new_policy = st.text_area("Policy Notes (optional)", key="mgr_comp_policy")
            if st.form_submit_button("Add Company"):
                if new_name:
                    db.add_company(new_name, new_cc, new_policy)
                    st.success(f"Company '{new_name}' added!")
                    st.rerun()
                else:
                    st.warning("Company Name is required.")

    search_company = st.text_input(
        "🔍 Search Companies", placeholder="Company name...", key="company_search"
    )

    companies = db.get_all_companies(active_only=False)
    if search_company:
        sl = search_company.lower()
        companies = [c for c in companies if sl in c["name"].lower()]

    if not companies:
        st.info("No companies match your search.")
    else:
        st.write(f"**{len(companies)} company(ies) found**")
        for comp in companies:
            comp_id = comp["id"]
            full_comp = db.get_company(comp_id)
            with st.container():
                col1, col2, col3, col4, col5, col6 = st.columns(
                    [2, 1.5, 2, 0.8, 0.8, 1]
                )
                with col1:
                    st.write(f"**{full_comp['name']}**")
                    if not full_comp.get("is_active", 1):
                        st.caption("⚠️ Inactive")
                with col2:
                    st.write(full_comp.get("default_cost_center", "") or "—")
                with col3:
                    st.write(full_comp.get("policy_notes", "") or "—")
                with col4:
                    if st.button("✏️", key=f"edit_comp_{comp_id}"):
                        st.session_state[f"edit_company_{comp_id}"] = True
                with col5:
                    if st.button("🗑️", key=f"del_comp_{comp_id}"):
                        st.session_state[f"confirm_del_comp_{comp_id}"] = True
                with col6:
                    if st.button("📄", key=f"export_comp_{comp_id}"):
                        st.session_state[f"export_company_{comp_id}"] = True

                if st.session_state.get(f"edit_company_{comp_id}", False):
                    with st.popover("Edit Company", use_container_width=True):
                        with st.form(key=f"edit_comp_form_{comp_id}"):
                            edit_name = st.text_input(
                                "Company Name*",
                                value=full_comp["name"],
                                key=f"edit_comp_name_{comp_id}",
                            )
                            edit_cc = st.text_input(
                                "Default Cost Center (optional)",
                                value=full_comp.get("default_cost_center", ""),
                                key=f"edit_comp_cc_{comp_id}",
                            )
                            edit_policy = st.text_area(
                                "Policy Notes (optional)",
                                value=full_comp.get("policy_notes", ""),
                                key=f"edit_comp_policy_{comp_id}",
                            )
                            edit_active = st.checkbox(
                                "Active",
                                value=bool(full_comp.get("is_active", 1)),
                                key=f"edit_comp_active_{comp_id}",
                                help="Uncheck to hide from pickers without deleting.",
                            )
                            if st.form_submit_button("💾 Save Changes"):
                                if edit_name:
                                    db.update_company(
                                        comp_id,
                                        edit_name,
                                        edit_cc,
                                        edit_policy,
                                        is_active=1 if edit_active else 0,
                                    )
                                    st.success(f"Company '{edit_name}' updated!")
                                    st.session_state.pop(
                                        f"edit_company_{comp_id}", None
                                    )
                                    st.rerun()
                                else:
                                    st.warning("Company Name is required.")
                            if st.form_submit_button("❌ Cancel"):
                                st.session_state.pop(f"edit_company_{comp_id}", None)
                                st.rerun()

                if st.session_state.get(f"confirm_del_comp_{comp_id}", False):

                    def _del_company(cid=comp_id):
                        # Return (ok, msg) so render_delete_confirmation can
                        # display a refusal without it being wiped by rerun.
                        return db.delete_company(cid)

                    render_delete_confirmation(
                        flag_key=f"confirm_del_comp_{comp_id}",
                        title=f"Permanently delete company '{full_comp['name']}'?",
                        on_confirm=_del_company,
                    )

                if st.session_state.get(f"export_company_{comp_id}", False):
                    with st.popover("Export Company Profile", use_container_width=True):
                        st.write(f"Export: **{full_comp['name']}**")
                        col_html, col_word, col_excel = st.columns(3)
                        with col_html:
                            if st.button("🌐 HTML", key=f"export_html_{comp_id}"):
                                html = doc_generator.generate_company_profile_html(
                                    comp_id
                                )
                                if html:
                                    st.download_button(
                                        "⬇️ Download HTML",
                                        data=html,
                                        file_name=f"{full_comp['name']}_profile.html",
                                        mime="text/html",
                                        key=f"download_html_{comp_id}",
                                    )
                                else:
                                    st.error("Failed to generate HTML.")
                        with col_word:
                            if st.button("📄 Word", key=f"export_word_{comp_id}"):
                                docx = doc_generator.generate_company_profile_docx(
                                    comp_id
                                )
                                if docx:
                                    st.download_button(
                                        "⬇️ Download Word",
                                        data=docx,
                                        file_name=f"{full_comp['name']}_profile.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml",
                                        key=f"download_word_{comp_id}",
                                    )
                                else:
                                    st.error("Failed to generate Word.")
                        with col_excel:
                            if st.button("📊 Excel", key=f"export_excel_{comp_id}"):
                                excel = export_company_profile_to_excel(comp_id)
                                if excel:
                                    st.download_button(
                                        "⬇️ Download Excel",
                                        data=excel,
                                        file_name=f"{full_comp['name']}_profile.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"download_excel_{comp_id}",
                                    )
                                else:
                                    st.error("Failed to generate Excel.")
                        if st.button("Close", key=f"close_export_comp_{comp_id}"):
                            st.session_state.pop(f"export_company_{comp_id}", None)
                            st.rerun()

                st.divider()

# =========================================================
# TAB: CONTACTS
# =========================================================

with tab_contacts:
    st.caption(
        "Manage contacts: local support, staff, partners, and other. "
        "These appear in the travel pack and can be assigned to specific "
        "itinerary items or the trip delegation."
    )

    companies = db.get_all_companies(active_only=True)
    company_options = {c["name"]: c["id"] for c in companies}
    company_names = list(company_options.keys())

    selected_company_label = st.selectbox(
        "Filter by Company",
        options=["All Companies"] + company_names,
        index=0,
        key="contact_filter_company",
    )
    if selected_company_label == "All Companies":
        filter_company_id = None
    else:
        filter_company_id = company_options[selected_company_label]

    search_term = st.text_input(
        "🔍 Search Contacts",
        placeholder="Name, role, phone, email, country, city...",
        key="contact_search",
    )

    type_options = ["All", "Emergency", "Local Support", "Staff", "Partner", "Other"]
    selected_type = st.radio(
        "Filter by Type",
        options=type_options,
        horizontal=True,
        index=0,
        key="contact_type_filter",
    )

    all_contacts_for_tags = (
        db.get_contacts(filter_company_id, active_only=False)
        if filter_company_id
        else db.get_contacts(active_only=False)
    )
    all_tags = set()
    for c in all_contacts_for_tags:
        if c.get("tags"):
            for tag in [t.strip() for t in c["tags"].split(",") if t.strip()]:
                all_tags.add(tag)
    tag_options = sorted(list(all_tags))
    selected_tags = st.multiselect(
        "🏷️ Filter by Tags", options=tag_options, key="contact_tag_filter"
    )

    with st.expander("➕ Add New Contact", expanded=False):
        with st.form("add_contact_form"):
            col1, col2 = st.columns(2)
            with col1:
                company_options_with_none = ["(No Company)"] + company_names
                add_company_label = st.selectbox(
                    "Company (optional)",
                    options=company_options_with_none,
                    index=0,
                    key="add_contact_company",
                )
                add_company_id = (
                    None
                    if add_company_label == "(No Company)"
                    else company_options[add_company_label]
                )
                add_name = st.text_input("Name*", key="add_contact_name")
                add_role = st.text_input("Role", key="add_contact_role")
                add_phone = st.text_input("Phone", key="add_contact_phone")
            with col2:
                add_email = st.text_input("Email", key="add_contact_email")
                country_list = sorted([c.name for c in pycountry.countries])
                add_country = st.selectbox(
                    "Country", options=[""] + country_list, key="add_contact_country"
                )
                add_city = st.text_input("City", key="add_contact_city")
                add_type = st.selectbox(
                    "Type",
                    options=["Emergency", "Local Support", "Staff", "Partner", "Other"],
                    key="add_contact_type",
                )
                add_notes = st.text_area("Notes", key="add_contact_notes")
                add_tags = st.text_input(
                    "Tags (comma-separated)", key="add_contact_tags"
                )

            if st.form_submit_button("➕ Add Contact"):
                if not add_name:
                    st.warning("Name is required.")
                else:
                    proceed = True
                    dupes = db.find_duplicate_contacts(
                        add_company_id, name=add_name, email=add_email, phone=add_phone
                    )
                    if dupes:
                        st.warning(
                            "⚠️ A contact with the same name, email, or phone "
                            "already exists in this company:"
                        )
                        for d in dupes:
                            st.write(f"- {d['name']} ({d.get('role','')})")
                        if not st.checkbox("Add anyway?", key="force_add_contact"):
                            proceed = False
                            st.info(
                                "Tick **Add anyway?** above and click **➕ Add Contact** "
                                "again to save the duplicate."
                            )
                    if proceed:
                        db.add_contact(
                            company_id=add_company_id,
                            name=add_name,
                            role=add_role,
                            phone=add_phone,
                            email=add_email,
                            country=add_country,
                            city=add_city,
                            notes=add_notes,
                            tags=add_tags,
                            type=add_type,
                        )
                        st.success(f"✅ Contact '{add_name}' added!")
                        st.rerun()

    # ---- Relationship Health ----
    with st.expander("💬 Relationship Health", expanded=False):
        gone_quiet_30 = db.get_contacts_gone_quiet(days=30)
        gone_quiet_60 = db.get_contacts_gone_quiet(days=60)
        gone_quiet_90 = db.get_contacts_gone_quiet(days=90)
        next_actions = db.get_next_actions_due()

        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            st.metric("Quiet 30+ days", len(gone_quiet_30))
        with col_b:
            st.metric("Quiet 60+ days", len(gone_quiet_60))
        with col_c:
            st.metric("Quiet 90+ days", len(gone_quiet_90))
        with col_d:
            st.metric("Next Actions", len(next_actions))

        st.markdown("**⏰ Next Actions Due**")
        if not next_actions:
            st.caption("No next actions set.")
        else:
            today_iso = datetime.now().date().isoformat()
            for c in next_actions[:10]:
                overdue = (
                    c.get("next_action_date")
                    and c["next_action_date"] < today_iso
                )
                flag = "🔴" if overdue else "🟢"
                col1, col2 = st.columns([5, 1])
                with col1:
                    st.markdown(
                        f"{flag} **{c['name']}** — {c['next_action']}"
                    )
                    st.caption(
                        f"Due: {c.get('next_action_date') or '—'}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"rel_open_{c['id']}",
                    ):
                        st.session_state[
                            "dialog_comm_contact_id"
                        ] = c["id"]
                        st.rerun()

        st.divider()

        st.markdown("**💤 Gone Quiet (90+ days)**")
        if not gone_quiet_90:
            st.caption("Everyone has been contacted in the last 90 days.")
        else:
            for c in gone_quiet_90[:10]:
                last = c.get("last_contacted_at") or "Never"
                col1, col2 = st.columns([5, 1])
                with col1:
                    st.markdown(f"**{c['name']}**")
                    st.caption(
                        f"Last: {last}  ·  "
                        f"{c.get('role') or '—'}"
                    )
                with col2:
                    if st.button(
                        "Open",
                        key=f"rel_quiet_{c['id']}",
                    ):
                        st.session_state[
                            "dialog_comm_contact_id"
                        ] = c["id"]
                        st.rerun()

    st.divider()

    all_contacts = (
        db.get_contacts(filter_company_id, active_only=False)
        if filter_company_id
        else db.get_contacts(active_only=False)
    )

    if search_term:
        s = search_term.lower()
        all_contacts = [
            c
            for c in all_contacts
            if s in c["name"].lower()
            or s in (c.get("role") or "").lower()
            or s in (c.get("phone") or "").lower()
            or s in (c.get("email") or "").lower()
            or s in (c.get("country") or "").lower()
            or s in (c.get("city") or "").lower()
        ]

    if selected_type != "All":
        all_contacts = [c for c in all_contacts if c.get("type") == selected_type]

    if selected_tags:
        filtered = []
        for c in all_contacts:
            if not c.get("tags"):
                continue
            contact_tags = [t.strip() for t in c["tags"].split(",") if t.strip()]
            if any(tag in contact_tags for tag in selected_tags):
                filtered.append(c)
        all_contacts = filtered

    if not all_contacts:
        st.info("No contacts match the filters.")
    else:
        st.write(f"**{len(all_contacts)} contact(s)**")

        col_export1, col_export2 = st.columns(2)
        with col_export1:
            if st.button("📥 Export CSV", key="export_contacts_csv"):
                output = io.StringIO()
                writer = csv.writer(output)
                writer.writerow(
                    [
                        "Company",
                        "Name",
                        "Role",
                        "Phone",
                        "Email",
                        "Country",
                        "City",
                        "Type",
                        "Notes",
                        "Tags",
                        "Active",
                    ]
                )
                for c in all_contacts:
                    company_name = (
                        db.get_company(c["company_id"])["name"]
                        if c.get("company_id")
                        else ""
                    )
                    writer.writerow(
                        [
                            company_name,
                            c["name"],
                            c.get("role", ""),
                            c.get("phone", ""),
                            c.get("email", ""),
                            c.get("country", ""),
                            c.get("city", ""),
                            c.get("type", "Local Support"),
                            c.get("notes", ""),
                            c.get("tags", ""),
                            "Yes" if c.get("is_active", 1) else "No",
                        ]
                    )
                st.download_button(
                    "⬇️ Download CSV",
                    data=output.getvalue().encode("utf-8"),
                    file_name=f"contacts_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="contact_csv_download",
                )

        with col_export2:
            uploaded_file = st.file_uploader(
                "📤 Import CSV",
                type=["csv"],
                key="contact_csv_upload",
                help="Columns: Company, Name, Role, Phone, Email, Country, City, Type, Notes, Tags",
            )
            if uploaded_file is not None:
                try:
                    content = uploaded_file.getvalue().decode("utf-8").splitlines()
                    reader = csv.DictReader(content)
                    expected = [
                        "Company",
                        "Name",
                        "Role",
                        "Phone",
                        "Email",
                        "Country",
                        "City",
                        "Type",
                        "Notes",
                        "Tags",
                    ]
                    if all(h in (reader.fieldnames or []) for h in expected):
                        if st.button("Start Import", key="contact_import_btn"):
                            added = skipped = 0
                            for row in reader:
                                company_name = row.get("Company", "").strip()
                                if company_name and company_name.lower() != "none":
                                    company_id = db._find_or_create_company(
                                        company_name
                                    )
                                else:
                                    company_id = None
                                name = row.get("Name", "").strip()
                                if not name:
                                    continue
                                dupes = db.find_duplicate_contacts(
                                    company_id,
                                    name=name,
                                    email=row.get("Email", "").strip() or None,
                                    phone=row.get("Phone", "").strip() or None,
                                )
                                if dupes:
                                    skipped += 1
                                    continue
                                db.add_contact(
                                    company_id=company_id,
                                    name=name,
                                    role=row.get("Role", "").strip() or None,
                                    phone=row.get("Phone", "").strip() or None,
                                    email=row.get("Email", "").strip() or None,
                                    country=row.get("Country", "").strip() or None,
                                    city=row.get("City", "").strip() or None,
                                    notes=row.get("Notes", "").strip() or None,
                                    tags=row.get("Tags", "").strip() or None,
                                    type=row.get("Type", "Local Support").strip()
                                    or "Local Support",
                                )
                                added += 1
                            # Stash the result so it survives st.rerun()
                            st.session_state["_contact_import_result"] = (
                                f"✅ Imported {added} contact(s). "
                                f"Skipped {skipped} duplicate(s)."
                            )
                            st.rerun()
                    else:
                        st.warning(f"CSV must have columns: {', '.join(expected)}")
                except Exception as e:
                    st.error(f"Import failed: {e}")

            # Show the import result from the previous run (survives rerun)
            _imp_msg = st.session_state.pop("_contact_import_result", None)
            if _imp_msg:
                st.success(_imp_msg)

        # Precompute default-contact IDs per company (one query per company)
        default_ids_by_company = {}
        for c in all_contacts:
            cid_company = c.get("company_id")
            if cid_company and cid_company not in default_ids_by_company:
                default_ids_by_company[cid_company] = set(
                    db.get_company_default_contacts(cid_company)
                )

        TYPE_ORDER = ["Emergency", "Local Support", "Staff", "Partner", "Other"]
        TYPE_ICONS = {
            "Emergency": "🚨",
            "Local Support": "📍",
            "Staff": "👤",
            "Partner": "🤝",
            "Other": "📌",
        }
        grouped = {}
        for c in all_contacts:
            t = c.get("type") or "Other"
            grouped.setdefault(t, []).append(c)

        # --- Compact grouped list -------------------------------------
        for grp_type in TYPE_ORDER:
            group = grouped.get(grp_type, [])
            if not group:
                continue
            icon = TYPE_ICONS.get(grp_type, "📌")
            st.markdown(f"##### {icon} {grp_type} · {len(group)}")

            # Column header row (once per group)
            hdr = st.columns([2.2, 1.5, 2.2, 2.2, 0.5, 0.5])
            for col, label in zip(
                hdr, ["Name", "Company", "Contact", "Location / Tags", "", ""]
            ):
                with col:
                    st.caption(label)

            for contact in group:
                cid = contact["id"]
                comp_name = (
                    db.get_company(contact["company_id"])["name"]
                    if contact.get("company_id")
                    else "—"
                )
                is_default = contact.get(
                    "company_id"
                ) and cid in default_ids_by_company.get(contact["company_id"], set())

                cols = st.columns([2.2, 1.5, 2.2, 2.2, 0.5, 0.5, 0.5])

                # ---- Name (+ markers, + role inline) ----
                with cols[0]:
                    line = f"**{contact['name']}**"
                    if is_default:
                        line += " ⭐"
                    if not contact.get("is_active", 1):
                        line += " ⚠️"
                    if contact.get("role"):
                        line += (
                            f" <span style='color:#8a8a8a;font-size:0.85em;'>"
                            f"· {contact['role']}</span>"
                        )
                    st.markdown(line, unsafe_allow_html=True)

                # ---- Company ----
                with cols[1]:
                    st.markdown(comp_name)

                # ---- Phone / email (stacked) ----
                with cols[2]:
                    bits = []
                    if contact.get("phone"):
                        bits.append(f"📞 {contact['phone']}")
                    if contact.get("email"):
                        bits.append(f"✉️ {contact['email']}")
                    st.markdown(
                        "  \n".join(bits) if bits else "—",
                        unsafe_allow_html=True,
                    )

                # ---- Location + tags ----
                with cols[3]:
                    bits = []
                    if contact.get("country") or contact.get("city"):
                        loc = (
                            f"{contact.get('city', '')}"
                            f"{', ' if contact.get('city') and contact.get('country') else ''}"
                            f"{contact.get('country', '')}"
                        )
                        bits.append(f"🌍 {loc}")
                    if contact.get("tags"):
                        tags_list = [
                            t.strip() for t in contact["tags"].split(",") if t.strip()
                        ]
                        if tags_list:
                            bits.append(
                                " ".join(_render_tag_pill(t) for t in tags_list)
                            )
                    st.markdown(
                        "  \n".join(bits) if bits else "—",
                        unsafe_allow_html=True,
                    )

                # ---- Edit / Delete ----
                with cols[4]:
                    if st.button("✏️", key=f"edit_contact_{cid}"):
                        st.session_state[f"editing_contact_{cid}"] = True
                with cols[5]:
                    if st.button("🗑️", key=f"del_contact_{cid}"):
                        st.session_state[f"delete_contact_{cid}"] = True
                with cols[6]:
                    if st.button(
                        "📞",
                        key=f"comm_contact_{cid}",
                        help="Log communication / view history",
                    ):
                        st.session_state["dialog_comm_contact_id"] = cid

                # ---- Inline edit form (unchanged) ----
                if st.session_state.get(f"editing_contact_{cid}", False):
                    with st.expander(f"Edit {contact['name']}", expanded=True):
                        with st.form(key=f"edit_contact_form_{cid}"):
                            edit_company_options_with_none = [
                                "(No Company)"
                            ] + company_names
                            current_company = (
                                comp_name if comp_name in company_names else None
                            )
                            edit_company_label = st.selectbox(
                                "Company (optional)",
                                options=edit_company_options_with_none,
                                index=(
                                    edit_company_options_with_none.index(
                                        current_company
                                    )
                                    if current_company in edit_company_options_with_none
                                    else 0
                                ),
                                key=f"edit_contact_company_{cid}",
                            )
                            edit_company_id = (
                                None
                                if edit_company_label == "(No Company)"
                                else company_options[edit_company_label]
                            )
                            edit_name = st.text_input(
                                "Name*",
                                value=contact["name"],
                                key=f"edit_name_{cid}",
                            )
                            edit_role = st.text_input(
                                "Role",
                                value=contact.get("role", ""),
                                key=f"edit_role_{cid}",
                            )
                            edit_phone = st.text_input(
                                "Phone",
                                value=contact.get("phone", ""),
                                key=f"edit_phone_{cid}",
                            )
                            edit_email = st.text_input(
                                "Email",
                                value=contact.get("email", ""),
                                key=f"edit_email_{cid}",
                            )
                            country_list = sorted([c.name for c in pycountry.countries])
                            edit_country = st.selectbox(
                                "Country",
                                options=[""] + country_list,
                                index=(
                                    ([""] + country_list).index(
                                        contact.get("country", "")
                                    )
                                    if contact.get("country", "") in country_list
                                    else 0
                                ),
                                key=f"edit_country_{cid}",
                            )
                            edit_city = st.text_input(
                                "City",
                                value=contact.get("city", ""),
                                key=f"edit_city_{cid}",
                            )
                            edit_type = st.selectbox(
                                "Type",
                                options=[
                                    "Emergency",
                                    "Local Support",
                                    "Staff",
                                    "Partner",
                                    "Other",
                                ],
                                index=[
                                    "Emergency",
                                    "Local Support",
                                    "Staff",
                                    "Partner",
                                    "Other",
                                ].index(contact.get("type", "Local Support")),
                                key=f"edit_contact_type_{cid}",
                            )
                            edit_notes = st.text_area(
                                "Notes",
                                value=contact.get("notes", ""),
                                key=f"edit_notes_{cid}",
                            )
                            edit_tags = st.text_input(
                                "Tags",
                                value=contact.get("tags", ""),
                                key=f"edit_tags_{cid}",
                            )
                            edit_active = st.checkbox(
                                "Active",
                                value=bool(contact.get("is_active", 1)),
                                key=f"edit_active_{cid}",
                            )

                            col_save, col_cancel = st.columns(2)
                            with col_save:
                                if st.form_submit_button("💾 Save"):
                                    if edit_name:
                                        db.update_contact(
                                            cid,
                                            name=edit_name,
                                            role=edit_role,
                                            phone=edit_phone,
                                            email=edit_email,
                                            country=edit_country,
                                            city=edit_city,
                                            notes=edit_notes,
                                            tags=edit_tags,
                                            is_active=1 if edit_active else 0,
                                            type=edit_type,
                                        )
                                        if edit_company_id != contact.get("company_id"):
                                            conn = sqlite3.connect(
                                                db.DB_PATH, timeout=30
                                            )
                                            try:
                                                cur = conn.cursor()
                                                cur.execute(
                                                    "UPDATE contacts SET company_id = ? WHERE id = ?",
                                                    (edit_company_id, cid),
                                                )
                                                conn.commit()
                                            finally:
                                                conn.close()
                                        st.session_state.pop(
                                            f"editing_contact_{cid}", None
                                        )
                                        st.success("Contact updated!")
                                        st.rerun()
                                    else:
                                        st.warning("Name is required.")
                            with col_cancel:
                                if st.form_submit_button("❌ Cancel"):
                                    st.session_state.pop(f"editing_contact_{cid}", None)
                                    st.rerun()

                if st.session_state.get(f"delete_contact_{cid}", False):
                    render_delete_confirmation(
                        flag_key=f"delete_contact_{cid}",
                        title=f"Permanently delete contact '{contact['name']}'?",
                        on_confirm=lambda cid=cid: db.delete_contact(cid),
                    )

            st.divider()

    if filter_company_id:
        st.subheader("⭐ Default Contacts")
        company = db.get_company(filter_company_id)
        if company:
            default_ids = db.get_company_default_contacts(filter_company_id)
            all_company_contacts = db.get_contacts(filter_company_id, active_only=True)
            if all_company_contacts:
                contact_options = {
                    f"{c['name']} ({c.get('role','')})": c["id"]
                    for c in all_company_contacts
                }
                selected_defaults = [
                    f"{c['name']} ({c.get('role','')})"
                    for c in all_company_contacts
                    if c["id"] in default_ids
                ]
                new_defaults = st.multiselect(
                    "Select contacts to auto‑include in new trips for this company",
                    options=list(contact_options.keys()),
                    default=selected_defaults,
                    key="default_contact_selector",
                )
                if st.button("💾 Save Default Contacts", key="save_default_contacts"):
                    db.set_company_default_contacts(
                        filter_company_id,
                        [contact_options[opt] for opt in new_defaults],
                    )
                    st.success("Default contacts updated!")
                    st.rerun()
            else:
                st.caption("No active contacts for this company yet.")
    else:
        st.info("Select a company above to set default contacts.")


# =========================================================
# TAB: VENUES
# =========================================================

with tab_venues:
    st.caption(
        "Manage reusable venues. Each venue can be attached to specific "
        "itinerary items."
    )

    col_search, col_country = st.columns(2)
    with col_search:
        search_term = st.text_input(
            "🔍 Search Venues", placeholder="Name, address, city...", key="venue_search"
        )
    with col_country:
        conn = sqlite3.connect(db.DB_PATH, timeout=30)
        try:
            c = conn.cursor()
            c.execute(
                "SELECT DISTINCT country FROM venues "
                "WHERE country IS NOT NULL AND country != '' ORDER BY country"
            )
            venue_countries = [row[0] for row in c.fetchall()]
        finally:
            conn.close()
        selected_country_filter = st.selectbox(
            "Filter by Country",
            options=["All Countries"] + venue_countries,
            index=0,
            key="venue_country_filter",
        )
        filter_country = (
            None
            if selected_country_filter == "All Countries"
            else selected_country_filter
        )

    with st.expander("➕ Add New Venue", expanded=False):
        with st.form("add_venue_form"):
            col1, col2 = st.columns(2)
            with col1:
                add_name = st.text_input("Venue Name*", key="add_venue_name")
                add_address = st.text_input("Address", key="add_venue_address")
                add_city = st.text_input("City", key="add_venue_city")
                country_list = sorted([c.name for c in pycountry.countries])
                add_country = st.selectbox(
                    "Country", options=[""] + country_list, key="add_venue_country"
                )
            with col2:
                add_wifi_ssid = st.text_input("WiFi SSID", key="add_venue_wifi_ssid")
                add_wifi_password = st.text_input(
                    "WiFi Password", key="add_venue_wifi_password"
                )
                add_badge_info = st.text_input("Badge Info", key="add_venue_badge_info")
                add_dress_code = st.selectbox(
                    "Dress Code",
                    options=DRESS_CODE_OPTIONS,
                    index=0,
                    key="add_venue_dress_code",
                )
            add_notes = st.text_area("Note", key="add_venue_notes")

            if st.form_submit_button("➕ Add Venue"):
                if not add_name:
                    st.warning("Venue Name is required.")
                else:
                    proceed = True
                    dupes = db.find_duplicate_venues(name=add_name)
                    if dupes:
                        st.warning(f"⚠️ A venue named '{add_name}' already exists:")
                        for d in dupes:
                            st.write(
                                f"- {d['name']} ({d.get('city','')}, {d.get('country','')})"
                            )
                        if not st.checkbox("Add anyway?", key="force_add_venue"):
                            proceed = False
                            st.info(
                                "Tick **Add anyway?** above and click **➕ Add Venue** "
                                "again to save the duplicate."
                            )
                    if proceed:
                        dress_code_value = (
                            add_dress_code if add_dress_code != "No Dress Code" else ""
                        )
                        db.add_venue(
                            name=add_name,
                            address=add_address,
                            city=add_city,
                            country=add_country,
                            wifi_ssid=add_wifi_ssid,
                            wifi_password=add_wifi_password,
                            badge_info=add_badge_info,
                            dress_code_notes=dress_code_value,
                            notes=add_notes,
                        )
                        st.success(f"✅ Venue '{add_name}' added!")
                        st.rerun()

    all_venues = db.get_venues(country=filter_country, active_only=False)
    if search_term:
        s = search_term.lower()
        all_venues = [
            v
            for v in all_venues
            if s in (v.get("name") or "").lower()
            or s in (v.get("address") or "").lower()
            or s in (v.get("city") or "").lower()
            or s in (v.get("country") or "").lower()
        ]

    if not all_venues:
        st.info("No venues match the filters.")
    else:
        st.write(f"**{len(all_venues)} venue(s)**")
        for v in all_venues:
            vid = v["id"]

            # --- Row: info, contact, edit, delete ---
            col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 0.6, 0.6])
            with col1:
                st.write(f"**{v['name']}**")
                if not v.get("is_active", 1):
                    st.caption("⚠️ Inactive")
                location_parts = [p for p in [v.get("city"), v.get("country")] if p]
                if location_parts:
                    st.caption("📍 " + ", ".join(location_parts))
            with col2:
                if v.get("address"):
                    st.write(f"🏠 {v['address']}")
                if v.get("wifi_ssid"):
                    st.write(f"📶 {v['wifi_ssid']}")
            with col3:
                if v.get("badge_info"):
                    st.write(f"🎫 {v['badge_info']}")
                if v.get("dress_code_notes"):
                    st.write(f"👔 {v['dress_code_notes']}")
            with col4:
                if st.button("✏️", key=f"edit_venue_{vid}"):
                    st.session_state[f"editing_venue_{vid}"] = True
            with col5:
                if st.button("🗑️", key=f"del_venue_{vid}"):
                    st.session_state[f"confirm_del_venue_{vid}"] = True

            # --- Inline edit form (Save / Cancel only) ---
            if st.session_state.get(f"editing_venue_{vid}", False):
                with st.expander(f"Edit {v['name']}", expanded=True):
                    with st.form(key=f"edit_venue_form_{vid}"):
                        col1, col2 = st.columns(2)
                        with col1:
                            edit_name = st.text_input(
                                "Venue Name*",
                                value=v["name"],
                                key=f"edit_venue_name_{vid}",
                            )
                            edit_address = st.text_input(
                                "Address",
                                value=v.get("address") or "",
                                key=f"edit_venue_address_{vid}",
                            )
                            edit_city = st.text_input(
                                "City",
                                value=v.get("city") or "",
                                key=f"edit_venue_city_{vid}",
                            )
                            country_list = sorted([c.name for c in pycountry.countries])
                            edit_country = st.selectbox(
                                "Country",
                                options=[""] + country_list,
                                index=(
                                    ([""] + country_list).index(v.get("country") or "")
                                    if (v.get("country") or "") in country_list
                                    else 0
                                ),
                                key=f"edit_venue_country_{vid}",
                            )
                        with col2:
                            edit_wifi_ssid = st.text_input(
                                "WiFi SSID",
                                value=v.get("wifi_ssid") or "",
                                key=f"edit_venue_wifi_ssid_{vid}",
                            )
                            edit_wifi_password = st.text_input(
                                "WiFi Password",
                                value=v.get("wifi_password") or "",
                                key=f"edit_venue_wifi_password_{vid}",
                            )
                            edit_badge_info = st.text_input(
                                "Badge Info",
                                value=v.get("badge_info") or "",
                                key=f"edit_venue_badge_{vid}",
                            )
                            current_dress_code = (
                                v.get("dress_code_notes") or "No Dress Code"
                            )
                            dress_code_options_edit = list(DRESS_CODE_OPTIONS)
                            if current_dress_code not in dress_code_options_edit:
                                dress_code_options_edit.append(current_dress_code)
                            edit_dress_code = st.selectbox(
                                "Dress Code",
                                options=dress_code_options_edit,
                                index=dress_code_options_edit.index(current_dress_code),
                                key=f"edit_venue_dress_{vid}",
                            )
                        edit_notes = st.text_area(
                            "Note",
                            value=v.get("notes") or "",
                            key=f"edit_venue_notes_{vid}",
                        )
                        edit_active = st.checkbox(
                            "Active",
                            value=bool(v.get("is_active", 1)),
                            key=f"edit_venue_active_{vid}",
                        )

                        col_save, col_cancel = st.columns(2)
                        with col_save:
                            if st.form_submit_button("💾 Save"):
                                if edit_name:
                                    dress_code_value = (
                                        edit_dress_code
                                        if edit_dress_code != "No Dress Code"
                                        else ""
                                    )
                                    db.update_venue(
                                        vid,
                                        name=edit_name,
                                        address=edit_address,
                                        city=edit_city,
                                        country=edit_country,
                                        wifi_ssid=edit_wifi_ssid,
                                        wifi_password=edit_wifi_password,
                                        badge_info=edit_badge_info,
                                        dress_code_notes=dress_code_value,
                                        notes=edit_notes,
                                        is_active=edit_active,
                                    )
                                    st.session_state.pop(f"editing_venue_{vid}", None)
                                    st.success("Venue updated!")
                                    st.rerun()
                                else:
                                    st.warning("Venue Name is required.")
                        with col_cancel:
                            if st.form_submit_button("❌ Cancel"):
                                st.session_state.pop(f"editing_venue_{vid}", None)
                                st.rerun()

            # --- Delete confirmation ---
            if st.session_state.get(f"confirm_del_venue_{vid}", False):

                def _del_venue(vid=vid):
                    db.delete_venue(vid)
                    st.session_state.pop(f"editing_venue_{vid}", None)

                render_delete_confirmation(
                    flag_key=f"confirm_del_venue_{vid}",
                    title=f"Permanently delete venue '{v['name']}'? Items referencing it will be unlinked.",
                    on_confirm=_del_venue,
                )
            st.divider()


# =========================================================
# TAB: LIBRARY
# =========================================================

with tab_library:
    st.caption("Reusable content: destination guides, visa rules, checklists, "
               "packing templates, and hospitals.")

    # One-shot hint shown when arriving from ✈️ Trips → 📋 From Template
    if st.session_state.pop("lib_hint_templates", False):
        st.info(
            "📋 You're here to create a trip from a template. "
            "Open the **📋 Trip Templates** sub-tab (first one on the left) → "
            "scroll to **🚀 Create Trip from Template**."
        )

    lib_templates, lib1, lib2, lib3, lib4, lib_emergency = st.tabs(
        [
            "📋 Trip Templates",
            "🌍 Destination Guides",
            "🛂 Visa Rules",
            "✅ Checklist Templates",
            "🎒 Packing Templates",
            "🚨 Emergency Directory",
        ]
    )

    # =========================================================
    # LIBRARY SUB-TAB: Trip Templates
    # =========================================================
    with lib_templates:
        st.subheader("📋 Trip Templates")
        st.caption("Reusable trip skeletons. Save any trip as a template from "
                   "its edit modal, then create new trips from here.")

        templates = db.get_trip_templates(active_only=False)
        if templates:
            st.write("**Saved Templates:**")
            for t in templates:
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.write(f"**{t['name']}**")
                    if not t.get("is_active", 1):
                        st.caption("⚠️ Inactive")
                    st.caption(f"Created: {fmt_date(t['created_at'])}")
                    if st.button("👁️ Preview", key=f"preview_{t['id']}"):
                        st.session_state[f"preview_template_{t['id']}"] = True
                    if st.session_state.get(f"preview_template_{t['id']}", False):
                        template_data = db.get_trip_template(t["id"])
                        if template_data:
                            st.write("**Departure:**")
                            st.write(f"City: {template_data.get('departure_city', 'N/A')}")
                            st.write(f"Region: {template_data.get('departure_region', 'N/A')}")
                            st.write(f"Country: {template_data.get('departure_country', 'N/A')}")
                            st.write("**Stops:**")
                            for stop in template_data.get("stops", []):
                                st.write(f"- {stop.get('city', '')}")
                            st.write("**Items:**")
                            for item in template_data.get("items", []):
                                st.write(f"- {item.get('item_type', '')}: "
                                         f"{item.get('description', '')}")
                            if st.button("Close Preview", key=f"close_preview_{t['id']}"):
                                st.session_state[f"preview_template_{t['id']}"] = False
                                st.rerun()
                with col2:
                    if t.get("is_active", 1):
                        if st.button("🚫", key=f"deactivate_tpl_{t['id']}",
                                     help="Deactivate — hides it from the template picker."):
                            db.update_trip_template(t["id"], is_active=0)
                            st.rerun()
                    else:
                        if st.button("✅", key=f"reactivate_tpl_{t['id']}",
                                     help="Reactivate — makes it selectable again."):
                            db.update_trip_template(t["id"], is_active=1)
                            st.rerun()
                    if st.button("🗑️", key=f"del_template_tab_{t['id']}"):
                        db.delete_trip_template(t["id"])
                        st.rerun()
        else:
            st.info("No templates yet. Save a trip as a template from its edit modal.")

        st.divider()
        st.subheader("🚀 Create Trip from Template")

        if not exec_id:
            st.warning("Select an executive in the sidebar first.")
        else:
            templates_active = db.get_trip_templates(active_only=True)
            if not templates_active:
                st.caption("No active templates available.")
            else:
                template_options = {t["name"]: t["id"] for t in templates_active}
                selected_template_name = st.selectbox(
                    "Select Template", list(template_options.keys()),
                    key="template_selector_tab")
                selected_template_id = template_options[selected_template_name]
                template_data = db.get_trip_template(selected_template_id)
                if template_data:
                    with st.form("apply_template_form_tab"):
                        col1, col2 = st.columns(2)
                        with col1:
                            new_trip_name = st.text_input("Trip Name*",
                                value=f"{selected_template_name} - "
                                      f"{datetime.now().strftime('%Y-%m-%d')}")
                            new_start = st.date_input("Start Date*",
                                value=datetime.now() + timedelta(days=7))
                        with col2:
                            new_budget = st.number_input("Budget",
                                min_value=0.0, step=100.0, value=1000.0)
                            new_end = st.date_input("End Date*",
                                value=datetime.now() + timedelta(days=10))
                        submitted = st.form_submit_button(
                            "🚀 Create Trip from Template")
                        if submitted:
                            if new_trip_name and new_start and new_end:
                                new_trip_id = db.apply_trip_template(
                                    selected_template_id, exec_id,
                                    new_trip_name, new_start, new_end, new_budget)
                                if new_trip_id:
                                    db.update_trip_status(new_trip_id, "draft")
                                    st.session_state["current_trip_id"] = new_trip_id
                                    st.success(f"✅ Trip '{new_trip_name}' "
                                               "created from template!")
                                    st.rerun()
                                else:
                                    st.error("Failed to create trip from template.")
                            else:
                                st.warning("Please fill in all required fields.")

    # ---- Guides ----
    with lib1:
        with st.expander("➕ Add New Destination Guide", expanded=False):
            with st.form("add_destination_guide_form"):
                country_list = sorted([c.name for c in pycountry.countries])
                col1, col2, col3 = st.columns(3)
                with col1:
                    dg_country = st.selectbox(
                        "Country*", options=country_list, key="dg_country"
                    )
                    dg_language = st.text_input("Primary Language", key="dg_language")
                with col2:
                    dg_currency = st.text_input(
                        "Local Currency", key="dg_currency", placeholder="e.g., JPY"
                    )
                    dg_apps = st.text_input(
                        "Recommended Apps (comma-separated)", key="dg_apps"
                    )
                with col3:
                    dg_police = st.text_input("Police Number", key="dg_police")
                    dg_ambulance = st.text_input("Ambulance Number", key="dg_ambulance")
                    dg_fire = st.text_input("Fire Number", key="dg_fire")
                dg_etiquette = st.text_area(
                    "Etiquette Notes (one per line)", key="dg_etiquette", height=100
                )
                dg_phrases = st.text_area(
                    "Common Phrases (one per line)", key="dg_phrases", height=100
                )
                dg_packing = st.text_area("Packing Tips", key="dg_packing", height=100)
                dg_conn = st.text_area(
                    "Connectivity Notes", key="dg_connectivity", height=80
                )
                dg_notes = st.text_area("Additional Note", key="dg_notes", height=80)
                if st.form_submit_button("➕ Add Guide"):
                    if not dg_country:
                        st.warning("Country is required.")
                    elif db.get_destination_guide_by_country(dg_country):
                        st.warning(f"A guide for {dg_country} already exists.")
                    else:
                        db.add_destination_guide(
                            country=dg_country,
                            language=dg_language or None,
                            currency=dg_currency or None,
                            emergency_police=dg_police or None,
                            emergency_ambulance=dg_ambulance or None,
                            emergency_fire=dg_fire or None,
                            etiquette_notes=dg_etiquette or None,
                            phrases=dg_phrases or None,
                            packing_tips=dg_packing or None,
                            connectivity_notes=dg_conn or None,
                            recommended_apps=dg_apps or None,
                            notes=dg_notes or None,
                        )
                        st.success(f"✅ Guide for '{dg_country}' added!")
                        st.rerun()

        guides = db.get_destination_guides(active_only=False)
        if not guides:
            st.info("No destination guides yet.")
        else:
            st.write(f"**{len(guides)} guide(s)**")
            for g in guides:
                gid = g["id"]

                # --- Row: info, emergency/apps, edit, delete ---
                col1, col2, col3, col4 = st.columns([3, 2, 0.6, 0.6])
                with col1:
                    st.write(f"**{g['country']}**")
                    if not g.get("is_active", 1):
                        st.caption("⚠️ Inactive")
                    if g.get("language"):
                        st.caption(
                            f"🗣️ {g['language']} · 💰 {g.get('currency') or '—'}"
                        )
                with col2:
                    if g.get("emergency_police") or g.get("emergency_ambulance"):
                        st.caption(
                            f"🚨 Police: {g.get('emergency_police') or '—'} · "
                            f"Ambulance: {g.get('emergency_ambulance') or '—'}"
                        )
                    if g.get("recommended_apps"):
                        st.caption(f"📱 {g['recommended_apps']}")
                with col3:
                    if st.button("✏️", key=f"edit_dg_{gid}"):
                        st.session_state[f"editing_dg_{gid}"] = True
                with col4:
                    if st.button("🗑️", key=f"del_dg_{gid}"):
                        st.session_state[f"confirm_del_dg_{gid}"] = True

                # --- Inline edit form (Save / Cancel only) ---
                if st.session_state.get(f"editing_dg_{gid}", False):
                    with st.expander(f"Edit Guide: {g['country']}", expanded=True):
                        with st.form(key=f"edit_dg_form_{gid}"):
                            country_list = sorted([c.name for c in pycountry.countries])
                            e_country = st.selectbox(
                                "Country*",
                                options=country_list,
                                index=(
                                    country_list.index(g["country"])
                                    if g["country"] in country_list
                                    else 0
                                ),
                                key=f"e_dg_country_{gid}",
                            )
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                e_language = st.text_input(
                                    "Language",
                                    value=g.get("language") or "",
                                    key=f"e_dg_lang_{gid}",
                                )
                                e_police = st.text_input(
                                    "Police",
                                    value=g.get("emergency_police") or "",
                                    key=f"e_dg_pol_{gid}",
                                )
                            with col2:
                                e_currency = st.text_input(
                                    "Currency",
                                    value=g.get("currency") or "",
                                    key=f"e_dg_cur_{gid}",
                                )
                                e_ambulance = st.text_input(
                                    "Ambulance",
                                    value=g.get("emergency_ambulance") or "",
                                    key=f"e_dg_amb_{gid}",
                                )
                            with col3:
                                e_apps = st.text_input(
                                    "Recommended Apps",
                                    value=g.get("recommended_apps") or "",
                                    key=f"e_dg_apps_{gid}",
                                )
                                e_fire = st.text_input(
                                    "Fire",
                                    value=g.get("emergency_fire") or "",
                                    key=f"e_dg_fire_{gid}",
                                )
                            e_etiquette = st.text_area(
                                "Etiquette Notes",
                                value=g.get("etiquette_notes") or "",
                                key=f"e_dg_etiq_{gid}",
                                height=100,
                            )
                            e_phrases = st.text_area(
                                "Common Phrases",
                                value=g.get("phrases") or "",
                                key=f"e_dg_phr_{gid}",
                                height=100,
                            )
                            e_packing = st.text_area(
                                "Packing Tips",
                                value=g.get("packing_tips") or "",
                                key=f"e_dg_pack_{gid}",
                                height=100,
                            )
                            e_conn = st.text_area(
                                "Connectivity Notes",
                                value=g.get("connectivity_notes") or "",
                                key=f"e_dg_conn_{gid}",
                                height=80,
                            )
                            e_notes = st.text_area(
                                "Note",
                                value=g.get("notes") or "",
                                key=f"e_dg_notes_{gid}",
                                height=80,
                            )
                            e_active = st.checkbox(
                                "Active",
                                value=bool(g.get("is_active", 1)),
                                key=f"e_dg_act_{gid}",
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.form_submit_button("💾 Save"):
                                    db.update_destination_guide(
                                        gid,
                                        country=e_country,
                                        language=e_language,
                                        currency=e_currency,
                                        emergency_police=e_police,
                                        emergency_ambulance=e_ambulance,
                                        emergency_fire=e_fire,
                                        etiquette_notes=e_etiquette,
                                        phrases=e_phrases,
                                        packing_tips=e_packing,
                                        connectivity_notes=e_conn,
                                        recommended_apps=e_apps,
                                        notes=e_notes,
                                        is_active=1 if e_active else 0,
                                    )
                                    st.session_state.pop(f"editing_dg_{gid}", None)
                                    st.success("Guide updated!")
                                    st.rerun()
                            with c2:
                                if st.form_submit_button("❌ Cancel"):
                                    st.session_state.pop(f"editing_dg_{gid}", None)
                                    st.rerun()

                # --- Delete confirmation ---
                if st.session_state.get(f"confirm_del_dg_{gid}", False):

                    def _del_dg(gid=gid):
                        db.delete_destination_guide(gid)
                        st.session_state.pop(f"editing_dg_{gid}", None)

                    render_delete_confirmation(
                        flag_key=f"confirm_del_dg_{gid}",
                        title=f"Permanently delete guide for '{g['country']}'?",
                        on_confirm=_del_dg,
                    )
                st.divider()

    # ---- Visa Rules ----
    # ---- Visa Rules ----
    with lib2:
        with st.expander("➕ Add New Visa Rule", expanded=False):
            with st.form("add_visa_rule_form"):
                country_list = sorted([c.name for c in pycountry.countries])
                col1, col2 = st.columns(2)
                with col1:
                    vr_passport = st.selectbox(
                        "Passport Nationality*", options=country_list, key="vr_passport"
                    )
                    vr_max_stay = st.number_input(
                        "Max Stay (days)", min_value=0, value=0, key="vr_max_stay"
                    )
                with col2:
                    vr_dest = st.selectbox(
                        "Destination Country*", options=country_list, key="vr_dest"
                    )
                    vr_processing = st.text_input(
                        "Processing Time", key="vr_processing"
                    )
                vr_required = st.checkbox(
                    "Visa Required", value=True, key="vr_required"
                )
                vr_fee = st.text_input("Fee", key="vr_fee")
                vr_notes = st.text_area("Notes", key="vr_notes", height=100)
                if st.form_submit_button("➕ Add Visa Rule"):
                    if db.get_visa_rule_for_pair(vr_passport, vr_dest):
                        st.warning(
                            f"A rule for {vr_passport} → {vr_dest} already exists."
                        )
                    else:
                        db.add_visa_rule(
                            passport_nationality=vr_passport,
                            destination_country=vr_dest,
                            visa_required=1 if vr_required else 0,
                            max_stay_days=vr_max_stay or None,
                            processing_time=vr_processing or None,
                            fee=vr_fee or None,
                            notes=vr_notes or None,
                        )
                        st.success("✅ Visa rule added!")
                        st.rerun()

        rules = db.get_visa_rules(active_only=False)
        if not rules:
            st.info("No visa rules yet.")
        else:
            st.write(f"**{len(rules)} rule(s)**")
            for r in rules:
                rid = r["id"]

                # --- Row: pair, requirement, details, edit, delete ---
                col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 0.6, 0.6])
                with col1:
                    st.write(
                        f"**{r['passport_nationality']}** → "
                        f"**{r['destination_country']}**"
                    )
                    if not r.get("is_active", 1):
                        st.caption("⚠️ Inactive")
                with col2:
                    st.write(
                        "✅ Visa Required"
                        if r["visa_required"]
                        else "❌ Visa Not Required"
                    )
                with col3:
                    parts = []
                    if r.get("max_stay_days"):
                        parts.append(f"Max stay: {r['max_stay_days']} days")
                    if r.get("processing_time"):
                        parts.append(f"Processing: {r['processing_time']}")
                    if r.get("fee"):
                        parts.append(f"Fee: {r['fee']}")
                    st.caption(" · ".join(parts) if parts else "—")
                with col4:
                    if st.button("✏️", key=f"edit_vr_{rid}"):
                        st.session_state[f"editing_vr_{rid}"] = True
                with col5:
                    if st.button("🗑️", key=f"del_vr_{rid}"):
                        st.session_state[f"confirm_del_vr_{rid}"] = True

                # --- Inline edit form (Save / Cancel only) ---
                if st.session_state.get(f"editing_vr_{rid}", False):
                    with st.expander("Edit Visa Rule", expanded=True):
                        with st.form(key=f"edit_vr_form_{rid}"):
                            country_list = sorted([c.name for c in pycountry.countries])
                            col1, col2 = st.columns(2)
                            with col1:
                                e_passport = st.selectbox(
                                    "Passport Nationality",
                                    options=country_list,
                                    index=(
                                        country_list.index(r["passport_nationality"])
                                        if r["passport_nationality"] in country_list
                                        else 0
                                    ),
                                    key=f"e_vr_pass_{rid}",
                                )
                                e_max_stay = st.number_input(
                                    "Max Stay (days)",
                                    min_value=0,
                                    value=int(r.get("max_stay_days") or 0),
                                    key=f"e_vr_ms_{rid}",
                                )
                            with col2:
                                e_dest = st.selectbox(
                                    "Destination Country",
                                    options=country_list,
                                    index=(
                                        country_list.index(r["destination_country"])
                                        if r["destination_country"] in country_list
                                        else 0
                                    ),
                                    key=f"e_vr_dest_{rid}",
                                )
                                e_processing = st.text_input(
                                    "Processing Time",
                                    value=r.get("processing_time") or "",
                                    key=f"e_vr_proc_{rid}",
                                )
                            e_required = st.checkbox(
                                "Visa Required",
                                value=bool(r["visa_required"]),
                                key=f"e_vr_req_{rid}",
                            )
                            e_fee = st.text_input(
                                "Fee", value=r.get("fee") or "", key=f"e_vr_fee_{rid}"
                            )
                            e_notes = st.text_area(
                                "Notes",
                                value=r.get("notes") or "",
                                key=f"e_vr_notes_{rid}",
                                height=100,
                            )
                            e_active = st.checkbox(
                                "Active",
                                value=bool(r.get("is_active", 1)),
                                key=f"e_vr_act_{rid}",
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.form_submit_button("💾 Save"):
                                    db.update_visa_rule(
                                        rid,
                                        passport_nationality=e_passport,
                                        destination_country=e_dest,
                                        visa_required=1 if e_required else 0,
                                        max_stay_days=e_max_stay or None,
                                        processing_time=e_processing or None,
                                        fee=e_fee or None,
                                        notes=e_notes or None,
                                        is_active=1 if e_active else 0,
                                    )
                                    st.session_state.pop(f"editing_vr_{rid}", None)
                                    st.success("Visa rule updated!")
                                    st.rerun()
                            with c2:
                                if st.form_submit_button("❌ Cancel"):
                                    st.session_state.pop(f"editing_vr_{rid}", None)
                                    st.rerun()

                # --- Delete confirmation ---
                if st.session_state.get(f"confirm_del_vr_{rid}", False):

                    def _del_vr(rid=rid):
                        db.delete_visa_rule(rid)
                        st.session_state.pop(f"editing_vr_{rid}", None)

                    render_delete_confirmation(
                        flag_key=f"confirm_del_vr_{rid}",
                        title="Permanently delete this visa rule?",
                        on_confirm=_del_vr,
                    )
                st.divider()



    # ---- Checklist Templates ----
    # ---- Checklist Templates ----
    with lib3:
        with st.expander("➕ Add New Checklist Template", expanded=False):
            with st.form("add_checklist_tpl_form"):
                ct_name = st.text_input("Template Name*", key="ct_name")
                ct_desc = st.text_input("Description (optional)", key="ct_desc")
                ct_items = st.text_area(
                    "Items (one per line)*", key="ct_items", height=150
                )
                if st.form_submit_button("➕ Add Template"):
                    if not ct_name or not ct_items.strip():
                        st.warning("Name and at least one item are required.")
                    else:
                        items_list = [
                            line.strip()
                            for line in ct_items.splitlines()
                            if line.strip()
                        ]
                        db.add_checklist_template(
                            name=ct_name,
                            description=ct_desc or None,
                            items_json=json.dumps(items_list),
                        )
                        st.success(f"✅ Checklist template '{ct_name}' added!")
                        st.rerun()

        templates = db.get_checklist_templates(active_only=False)
        if not templates:
            st.info("No checklist templates yet.")
        else:
            st.write(f"**{len(templates)} template(s)**")
            for t in templates:
                tid = t["id"]
                try:
                    items_list = json.loads(t.get("items_json") or "[]")
                except Exception:
                    items_list = []

                # --- Row: info, count, edit, delete ---
                col1, col2, col3, col4 = st.columns([3, 2, 0.6, 0.6])
                with col1:
                    st.write(f"**{t['name']}**")
                    if not t.get("is_active", 1):
                        st.caption("⚠️ Inactive")
                    if t.get("description"):
                        st.caption(t["description"])
                with col2:
                    st.caption(f"📋 {len(items_list)} item(s)")
                with col3:
                    if st.button("✏️", key=f"edit_ct_{tid}"):
                        st.session_state[f"editing_ct_{tid}"] = True
                with col4:
                    if st.button("🗑️", key=f"del_ct_{tid}"):
                        st.session_state[f"confirm_del_ct_{tid}"] = True

                # --- Inline edit form (Save / Cancel only) ---
                if st.session_state.get(f"editing_ct_{tid}", False):
                    with st.expander(f"Edit Template: {t['name']}", expanded=True):
                        with st.form(key=f"edit_ct_form_{tid}"):
                            e_name = st.text_input(
                                "Name*", value=t["name"], key=f"e_ct_name_{tid}"
                            )
                            e_desc = st.text_input(
                                "Description",
                                value=t.get("description") or "",
                                key=f"e_ct_desc_{tid}",
                            )
                            e_items = st.text_area(
                                "Items (one per line)",
                                value="\n".join(items_list),
                                height=200,
                                key=f"e_ct_items_{tid}",
                            )
                            e_active = st.checkbox(
                                "Active",
                                value=bool(t.get("is_active", 1)),
                                key=f"e_ct_act_{tid}",
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.form_submit_button("💾 Save"):
                                    new_items = [
                                        line.strip()
                                        for line in e_items.splitlines()
                                        if line.strip()
                                    ]
                                    db.update_checklist_template(
                                        tid,
                                        name=e_name,
                                        description=e_desc or None,
                                        items_json=json.dumps(new_items),
                                        is_active=1 if e_active else 0,
                                    )
                                    st.session_state.pop(f"editing_ct_{tid}", None)
                                    st.success("Template updated!")
                                    st.rerun()
                            with c2:
                                if st.form_submit_button("❌ Cancel"):
                                    st.session_state.pop(f"editing_ct_{tid}", None)
                                    st.rerun()

                # --- Delete confirmation ---
                if st.session_state.get(f"confirm_del_ct_{tid}", False):

                    def _del_ct(tid=tid):
                        db.delete_checklist_template(tid)
                        st.session_state.pop(f"editing_ct_{tid}", None)

                    render_delete_confirmation(
                        flag_key=f"confirm_del_ct_{tid}",
                        title="Permanently delete this checklist template?",
                        on_confirm=_del_ct,
                    )
                st.divider()




    # ---- Packing Templates ----
    # ---- Packing Templates ----
    with lib4:
        PACKING_CATEGORIES = [
            "Documents & Badges",
            "Devices & Chargers",
            "Clothing",
            "Formalwear",
            "Presentation Materials",
            "Toiletries",
            "Personal",
            "Other",
        ]
        with st.expander("➕ Add New Packing Template", expanded=False):
            with st.form("add_packing_tpl_form"):
                pt_name = st.text_input("Template Name*", key="pt_name")
                pt_cat = st.selectbox(
                    "Category", options=PACKING_CATEGORIES, key="pt_cat"
                )
                pt_items = st.text_area(
                    "Items (one per line)*", key="pt_items", height=150
                )
                if st.form_submit_button("➕ Add Template"):
                    if not pt_name or not pt_items.strip():
                        st.warning("Name and at least one item are required.")
                    else:
                        items_list = [
                            line.strip()
                            for line in pt_items.splitlines()
                            if line.strip()
                        ]
                        db.add_packing_template(
                            name=pt_name,
                            category=pt_cat,
                            items_json=json.dumps(items_list),
                        )
                        st.success(f"✅ Packing template '{pt_name}' added!")
                        st.rerun()

        pack_templates = db.get_packing_templates(active_only=False)
        if not pack_templates:
            st.info("No packing templates yet.")
        else:
            st.write(f"**{len(pack_templates)} template(s)**")
            for t in pack_templates:
                tid = t["id"]
                try:
                    items_list = json.loads(t.get("items_json") or "[]")
                except Exception:
                    items_list = []

                # --- Row: info, count, edit, delete ---
                col1, col2, col3, col4 = st.columns([3, 2, 0.6, 0.6])
                with col1:
                    st.write(f"**{t['name']}**")
                    if not t.get("is_active", 1):
                        st.caption("⚠️ Inactive")
                    if t.get("category"):
                        st.caption(f"📁 {t['category']}")
                with col2:
                    st.caption(f"📦 {len(items_list)} item(s)")
                with col3:
                    if st.button("✏️", key=f"edit_pt_{tid}"):
                        st.session_state[f"editing_pt_{tid}"] = True
                with col4:
                    if st.button("🗑️", key=f"del_pt_{tid}"):
                        st.session_state[f"confirm_del_pt_{tid}"] = True

                # --- Inline edit form (Save / Cancel only) ---
                if st.session_state.get(f"editing_pt_{tid}", False):
                    with st.expander(f"Edit Template: {t['name']}", expanded=True):
                        with st.form(key=f"edit_pt_form_{tid}"):
                            e_name = st.text_input(
                                "Name*", value=t["name"], key=f"e_pt_name_{tid}"
                            )
                            e_cat = st.selectbox(
                                "Category",
                                options=PACKING_CATEGORIES,
                                index=(
                                    PACKING_CATEGORIES.index(t["category"])
                                    if t.get("category") in PACKING_CATEGORIES
                                    else 0
                                ),
                                key=f"e_pt_cat_{tid}",
                            )
                            e_items = st.text_area(
                                "Items (one per line)",
                                value="\n".join(items_list),
                                height=200,
                                key=f"e_pt_items_{tid}",
                            )
                            e_active = st.checkbox(
                                "Active",
                                value=bool(t.get("is_active", 1)),
                                key=f"e_pt_act_{tid}",
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.form_submit_button("💾 Save"):
                                    new_items = [
                                        line.strip()
                                        for line in e_items.splitlines()
                                        if line.strip()
                                    ]
                                    db.update_packing_template(
                                        tid,
                                        name=e_name,
                                        category=e_cat,
                                        items_json=json.dumps(new_items),
                                        is_active=1 if e_active else 0,
                                    )
                                    st.session_state.pop(f"editing_pt_{tid}", None)
                                    st.success("Packing template updated!")
                                    st.rerun()
                            with c2:
                                if st.form_submit_button("❌ Cancel"):
                                    st.session_state.pop(f"editing_pt_{tid}", None)
                                    st.rerun()

                # --- Delete confirmation ---
                if st.session_state.get(f"confirm_del_pt_{tid}", False):

                    def _del_pt(tid=tid):
                        db.delete_packing_template(tid)
                        st.session_state.pop(f"editing_pt_{tid}", None)

                    render_delete_confirmation(
                        flag_key=f"confirm_del_pt_{tid}",
                        title="Permanently delete this packing template?",
                        on_confirm=_del_pt,
                    )
                st.divider()


    # =========================================================
    # LIBRARY SUB-TAB: Emergency Directory
    # =========================================================
    with lib_emergency:
        st.caption(
            "Hospitals, emergency numbers, and embassies / consulates — "
            "everything you'd need in a crisis."
        )

        e_hosp, e_numbers, e_emb = st.tabs(
            ["🏥 Hospitals", "☎️ Emergency Numbers", "🏛️ Embassies / Consulates"]
        )

        # ==================== Hospitals ====================
        with e_hosp:
            st.markdown("**🏥 Hospitals**")
            with st.expander("➕ Add New Hospital", expanded=False):
                with st.form("add_hospital_form"):
                    country_list = sorted([c.name for c in pycountry.countries])
                    col1, col2 = st.columns(2)
                    with col1:
                        h_city = st.text_input("City*", key="h_city")
                        h_name = st.text_input("Hospital Name*", key="h_name")
                        h_phone = st.text_input("Phone", key="h_phone")
                    with col2:
                        h_country = st.selectbox(
                            "Country", options=[""] + country_list, key="h_country"
                        )
                        h_address = st.text_input("Address", key="h_address")
                        h_maps = st.text_input("Google Maps URL", key="h_maps")
                    h_notes = st.text_area("Note", key="h_notes", height=80)
                    if st.form_submit_button("➕ Add Hospital"):
                        if not h_city or not h_name:
                            st.warning("City and Hospital Name are required.")
                        else:
                            db.add_hospital(
                                city=h_city,
                                country=h_country,
                                name=h_name,
                                address=h_address or None,
                                phone=h_phone or None,
                                maps_url=h_maps or None,
                                notes=h_notes or None,
                            )
                            st.success(f"✅ Hospital '{h_name}' added!")
                            st.rerun()

            col_f1, col_f2 = st.columns(2)
            with col_f1:
                hosp_search = st.text_input(
                    "🔍 Search", key="hosp_search", placeholder="City, country, name..."
                )
            with col_f2:
                conn = sqlite3.connect(db.DB_PATH, timeout=30)
                try:
                    c = conn.cursor()
                    c.execute(
                        "SELECT DISTINCT country FROM hospitals "
                        "WHERE country IS NOT NULL AND country != '' "
                        "ORDER BY country"
                    )
                    hosp_countries = [row[0] for row in c.fetchall()]
                finally:
                    conn.close()
                hosp_country_filter = st.selectbox(
                    "Filter by Country",
                    options=["All Countries"] + hosp_countries,
                    index=0,
                    key="hosp_country_filter",
                )
                hosp_filter_country = (
                    None
                    if hosp_country_filter == "All Countries"
                    else hosp_country_filter
                )

            all_hospitals = db.get_hospitals(
                country=hosp_filter_country, active_only=False
            )
            if hosp_search:
                s = hosp_search.lower()
                all_hospitals = [
                    h
                    for h in all_hospitals
                    if s in (h.get("name") or "").lower()
                    or s in (h.get("city") or "").lower()
                    or s in (h.get("country") or "").lower()
                    or s in (h.get("address") or "").lower()
                ]

            if not all_hospitals:
                st.info("No hospitals match the filters.")
            else:
                st.write(f"**{len(all_hospitals)} hospital(s)**")
                for h in all_hospitals:
                    hid = h["id"]

                    # --- Row: info, address, edit, delete ---
                    col1, col2, col3, col4 = st.columns([3, 3, 0.6, 0.6])
                    with col1:
                        st.write(f"**{h['name']}**")
                        if not h.get("is_active", 1):
                            st.caption("⚠️ Inactive")
                        loc = ", ".join(
                            p for p in [h.get("city"), h.get("country")] if p
                        )
                        if loc:
                            st.caption(f"📍 {loc}")
                    with col2:
                        if h.get("address"):
                            st.caption(f"🏠 {h['address']}")
                        if h.get("phone"):
                            st.write(f"📞 {h['phone']}")
                        if h.get("maps_url"):
                            st.markdown(f"[Open in Maps]({h['maps_url']})")
                    with col3:
                        if st.button("✏️", key=f"edit_hosp_{hid}"):
                            st.session_state[f"editing_hosp_{hid}"] = True
                    with col4:
                        if st.button("🗑️", key=f"del_hosp_{hid}"):
                            st.session_state[f"confirm_del_hosp_{hid}"] = True

                    # --- Inline edit form (Save / Cancel only) ---
                    if st.session_state.get(f"editing_hosp_{hid}", False):
                        with st.expander(f"Edit {h['name']}", expanded=True):
                            with st.form(key=f"edit_hosp_form_{hid}"):
                                country_list = sorted(
                                    [c.name for c in pycountry.countries]
                                )
                                col1, col2 = st.columns(2)
                                with col1:
                                    e_city = st.text_input(
                                        "City*", value=h["city"], key=f"e_h_city_{hid}"
                                    )
                                    e_name = st.text_input(
                                        "Hospital Name*",
                                        value=h["name"],
                                        key=f"e_h_name_{hid}",
                                    )
                                    e_phone = st.text_input(
                                        "Phone",
                                        value=h.get("phone") or "",
                                        key=f"e_h_phone_{hid}",
                                    )
                                with col2:
                                    e_country = st.selectbox(
                                        "Country",
                                        options=[""] + country_list,
                                        index=(
                                            ([""] + country_list).index(
                                                h.get("country") or ""
                                            )
                                            if (h.get("country") or "") in country_list
                                            else 0
                                        ),
                                        key=f"e_h_country_{hid}",
                                    )
                                    e_address = st.text_input(
                                        "Address",
                                        value=h.get("address") or "",
                                        key=f"e_h_addr_{hid}",
                                    )
                                    e_maps = st.text_input(
                                        "Google Maps URL",
                                        value=h.get("maps_url") or "",
                                        key=f"e_h_maps_{hid}",
                                    )
                                e_notes = st.text_area(
                                    "Note",
                                    value=h.get("notes") or "",
                                    key=f"e_h_notes_{hid}",
                                    height=80,
                                )
                                e_active = st.checkbox(
                                    "Active",
                                    value=bool(h.get("is_active", 1)),
                                    key=f"e_h_act_{hid}",
                                )
                                c1, c2 = st.columns(2)
                                with c1:
                                    if st.form_submit_button("💾 Save"):
                                        db.update_hospital(
                                            hid,
                                            city=e_city,
                                            country=e_country,
                                            name=e_name,
                                            address=e_address,
                                            phone=e_phone,
                                            maps_url=e_maps,
                                            notes=e_notes,
                                            is_active=1 if e_active else 0,
                                        )
                                        st.session_state.pop(
                                            f"editing_hosp_{hid}", None
                                        )
                                        st.success("Hospital updated!")
                                        st.rerun()
                                with c2:
                                    if st.form_submit_button("❌ Cancel"):
                                        st.session_state.pop(
                                            f"editing_hosp_{hid}", None
                                        )
                                        st.rerun()

                    # --- Delete confirmation ---
                    if st.session_state.get(f"confirm_del_hosp_{hid}", False):

                        def _del_hosp(hid=hid):
                            db.delete_hospital(hid)
                            st.session_state.pop(f"editing_hosp_{hid}", None)

                        render_delete_confirmation(
                            flag_key=f"confirm_del_hosp_{hid}",
                            title=f"Permanently delete '{h['name']}'?",
                            on_confirm=_del_hosp,
                        )
                    st.divider()

        # ==================== Emergency Numbers ====================
        with e_numbers:
            st.markdown("**☎️ Emergency Numbers**")
            st.caption(
                "Emergency service numbers per country. "
                "These come from the Destination Guides — edit them there."
            )
            numbers = []
            for dg in db.get_destination_guides(active_only=False):
                if (
                    dg.get("emergency_police")
                    or dg.get("emergency_ambulance")
                    or dg.get("emergency_fire")
                ):
                    numbers.append(dg)
            if not numbers:
                st.info(
                    "No emergency numbers yet. Add them via "
                    "**Library → 🌍 Destination Guides**."
                )
            else:
                for dg in numbers:
                    col1, col2 = st.columns([2, 4])
                    with col1:
                        st.write(f"**{dg['country']}**")
                        if not dg.get("is_active", 1):
                            st.caption("⚠️ Inactive")
                    with col2:
                        bits = []
                        if dg.get("emergency_police"):
                            bits.append(f"🚓 Police: {dg['emergency_police']}")
                        if dg.get("emergency_ambulance"):
                            bits.append(f"🚑 Ambulance: {dg['emergency_ambulance']}")
                        if dg.get("emergency_fire"):
                            bits.append(f"🚒 Fire: {dg['emergency_fire']}")
                        st.write("  ·  ".join(bits))
                    st.divider()

        # ==================== Embassies ====================
        with e_emb:
            st.markdown("**🏛️ Embassies / Consulates**")
            with st.expander("➕ Add New Embassy", expanded=False):
                with st.form("add_embassy_form"):
                    country_list = sorted([c.name for c in pycountry.countries])
                    col1, col2 = st.columns(2)
                    with col1:
                        e_host = st.selectbox(
                            "Host Country (where the " "executive is)*",
                            options=country_list,
                            key="emb_host",
                        )
                        e_city = st.text_input(
                            "City", placeholder="e.g. Tokyo", key="emb_city"
                        )
                        e_hours = st.text_input(
                            "Hours",
                            placeholder="e.g. Mon-Fri 9:00-16:00",
                            key="emb_hours",
                        )
                    with col2:
                        e_rep = st.selectbox(
                            "Representing Country (whose " "embassy it is)*",
                            options=country_list,
                            key="emb_rep",
                        )
                        e_name = st.text_input(
                            "Embassy Name",
                            placeholder="e.g. Embassy of Nigeria",
                            key="emb_name",
                        )
                        e_phone = st.text_input("Phone", key="emb_phone")
                    e_address = st.text_input("Address", key="emb_address")
                    e_email = st.text_input("Email", key="emb_email")
                    e_website = st.text_input("Website", key="emb_website")
                    e_maps = st.text_input("Google Maps URL", key="emb_maps")
                    e_notes = st.text_area("Notes", key="emb_notes", height=80)

                    if st.form_submit_button("➕ Add Embassy"):
                        if e_host and e_rep:
                            db.add_embassy(
                                host_country=e_host,
                                representing_country=e_rep,
                                city=e_city or None,
                                name=e_name or None,
                                address=e_address or None,
                                phone=e_phone or None,
                                email=e_email or None,
                                website=e_website or None,
                                maps_url=e_maps or None,
                                hours=e_hours or None,
                                notes=e_notes or None,
                            )
                            st.success(f"✅ Embassy added: {e_rep} in {e_host}.")
                            st.rerun()
                        else:
                            st.warning(
                                "Host Country and Representing " "Country are required."
                            )

            col_f1, col_f2 = st.columns(2)
            with col_f1:
                emb_search = st.text_input(
                    "🔍 Search", key="emb_search", placeholder="Country, city, name..."
                )
            with col_f2:
                emb_host_filter = st.selectbox(
                    "Filter by Host Country",
                    options=["All Countries"]
                    + sorted([c.name for c in pycountry.countries]),
                    index=0,
                    key="emb_host_filter",
                )
                host_filter = (
                    None if emb_host_filter == "All Countries" else emb_host_filter
                )

            all_embassies = db.get_embassies(
                host_country=host_filter, active_only=False
            )
            if emb_search:
                s = emb_search.lower()
                all_embassies = [
                    e
                    for e in all_embassies
                    if s in (e.get("name") or "").lower()
                    or s in (e.get("city") or "").lower()
                    or s in (e.get("host_country") or "").lower()
                    or s in (e.get("representing_country") or "").lower()
                    or s in (e.get("address") or "").lower()
                ]

            if not all_embassies:
                st.info("No embassies match the filters.")
            else:
                st.write(f"**{len(all_embassies)} embassy / consulate entries**")
                for emb in all_embassies:
                    eid = emb["id"]

                    # --- Row: info, contact, edit, delete ---
                    col1, col2, col3, col4 = st.columns([3, 3, 0.6, 0.6])
                    with col1:
                        title = emb.get("name") or (
                            f"{emb['representing_country']} Embassy"
                        )
                        st.write(f"**{title}**")
                        if not emb.get("is_active", 1):
                            st.caption("⚠️ Inactive")
                        st.caption(
                            f"{emb['representing_country']} → "
                            f"{emb['host_country']}"
                            + (f" · {emb['city']}" if emb.get("city") else "")
                        )
                    with col2:
                        if emb.get("address"):
                            st.caption(f"🏠 {emb['address']}")
                        if emb.get("phone"):
                            st.write(f"📞 {emb['phone']}")
                        if emb.get("hours"):
                            st.caption(f"🕐 {emb['hours']}")
                        if emb.get("maps_url"):
                            st.markdown(f"[Open in Maps]({emb['maps_url']})")
                    with col3:
                        if st.button("✏️", key=f"edit_emb_{eid}"):
                            st.session_state[f"editing_emb_{eid}"] = True
                    with col4:
                        if st.button("🗑️", key=f"del_emb_{eid}"):
                            st.session_state[f"confirm_del_emb_{eid}"] = True

                    # --- Inline edit form (Save / Cancel only) ---
                    if st.session_state.get(f"editing_emb_{eid}", False):
                        with st.expander("Edit Embassy", expanded=True):
                            with st.form(key=f"edit_emb_form_{eid}"):
                                country_list = sorted(
                                    [c.name for c in pycountry.countries]
                                )
                                col1, col2 = st.columns(2)
                                with col1:
                                    ee_host = st.selectbox(
                                        "Host Country*",
                                        options=country_list,
                                        index=(
                                            country_list.index(emb["host_country"])
                                            if emb["host_country"] in country_list
                                            else 0
                                        ),
                                        key=f"ee_host_{eid}",
                                    )
                                    ee_city = st.text_input(
                                        "City",
                                        value=emb.get("city") or "",
                                        key=f"ee_city_{eid}",
                                    )
                                    ee_hours = st.text_input(
                                        "Hours",
                                        value=emb.get("hours") or "",
                                        key=f"ee_hours_{eid}",
                                    )
                                with col2:
                                    ee_rep = st.selectbox(
                                        "Representing Country*",
                                        options=country_list,
                                        index=(
                                            country_list.index(
                                                emb["representing_country"]
                                            )
                                            if emb["representing_country"]
                                            in country_list
                                            else 0
                                        ),
                                        key=f"ee_rep_{eid}",
                                    )
                                    ee_name = st.text_input(
                                        "Embassy Name",
                                        value=emb.get("name") or "",
                                        key=f"ee_name_{eid}",
                                    )
                                    ee_phone = st.text_input(
                                        "Phone",
                                        value=emb.get("phone") or "",
                                        key=f"ee_phone_{eid}",
                                    )
                                ee_address = st.text_input(
                                    "Address",
                                    value=emb.get("address") or "",
                                    key=f"ee_addr_{eid}",
                                )
                                ee_email = st.text_input(
                                    "Email",
                                    value=emb.get("email") or "",
                                    key=f"ee_email_{eid}",
                                )
                                ee_website = st.text_input(
                                    "Website",
                                    value=emb.get("website") or "",
                                    key=f"ee_website_{eid}",
                                )
                                ee_maps = st.text_input(
                                    "Google Maps URL",
                                    value=emb.get("maps_url") or "",
                                    key=f"ee_maps_{eid}",
                                )
                                ee_notes = st.text_area(
                                    "Notes",
                                    value=emb.get("notes") or "",
                                    key=f"ee_notes_{eid}",
                                    height=80,
                                )
                                ee_active = st.checkbox(
                                    "Active",
                                    value=bool(emb.get("is_active", 1)),
                                    key=f"ee_act_{eid}",
                                )
                                c1, c2 = st.columns(2)
                                with c1:
                                    if st.form_submit_button("💾 Save"):
                                        db.update_embassy(
                                            eid,
                                            host_country=ee_host,
                                            representing_country=ee_rep,
                                            city=ee_city or None,
                                            name=ee_name or None,
                                            address=ee_address or None,
                                            phone=ee_phone or None,
                                            email=ee_email or None,
                                            website=ee_website or None,
                                            maps_url=ee_maps or None,
                                            hours=ee_hours or None,
                                            notes=ee_notes or None,
                                            is_active=1 if ee_active else 0,
                                        )
                                        st.session_state.pop(f"editing_emb_{eid}", None)
                                        st.success("Embassy updated!")
                                        st.rerun()
                                with c2:
                                    if st.form_submit_button("❌ Cancel"):
                                        st.session_state.pop(f"editing_emb_{eid}", None)
                                        st.rerun()

                    # --- Delete confirmation ---
                    if st.session_state.get(f"confirm_del_emb_{eid}", False):

                        def _del_emb(eid=eid):
                            db.delete_embassy(eid)
                            st.session_state.pop(f"editing_emb_{eid}", None)

                        render_delete_confirmation(
                            flag_key=f"confirm_del_emb_{eid}",
                            title=f"Permanently delete this embassy record?",
                            on_confirm=_del_emb,
                        )
                    st.divider()

# =========================================================
# GLOBAL DIALOGS
# =========================================================
# Rendered outside tab blocks so they can be triggered from
# any tab and still appear.

if st.session_state.get("dialog_meeting_id"):
    _render_meeting_detail_dialog(
        st.session_state["dialog_meeting_id"]
    )

if st.session_state.get("dialog_comm_contact_id"):
    _render_communications_dialog(
        st.session_state["dialog_comm_contact_id"]
    )

if st.session_state.get("dialog_event_id"):
    _render_event_detail_dialog(
        st.session_state["dialog_event_id"]
    )

if st.session_state.get("dialog_rsvp_import_event_id"):
    _render_rsvp_import_dialog(
        st.session_state["dialog_rsvp_import_event_id"]
    )
    

if st.session_state.get("dialog_speaking_id"):
    _render_speaking_detail_dialog(
        st.session_state["dialog_speaking_id"]
    )