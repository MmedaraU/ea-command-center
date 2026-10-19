from datetime import datetime, timedelta, date
from icalendar import Calendar, Event
import pytz


def detect_conflicts(items, exec_timezone_str="UTC"):
    """
    Check for overlapping events (flights, meetings, transport).
    Hotels are excluded.
    Returns a list of warning strings with item IDs and time details.
    All times are converted to UTC for comparison.
    """
    conflicts = []
    time_events = []

    # Fallback timezone
    fallback_tz = pytz.timezone(exec_timezone_str)

    for item in items:
        # Skip hotels – they represent stays, not timed events
        if item.get("item_type", "").lower() == "hotel":
            continue
        try:
            # Parse naive datetime strings
            start_naive = datetime.fromisoformat(item["datetime_start"])
            end_naive = (
                datetime.fromisoformat(item["datetime_end"])
                if item["datetime_end"]
                else start_naive + timedelta(hours=1)
            )

            # Get timezone from item, fallback to executive's
            tz_str = item.get("timezone") or exec_timezone_str
            tz = pytz.timezone(tz_str)

            # Localize and convert to UTC
            start_aware = tz.localize(start_naive)
            end_aware = tz.localize(end_naive)
            start_utc = start_aware.astimezone(pytz.UTC)
            end_utc = end_aware.astimezone(pytz.UTC)

            time_events.append(
                {
                    "id": item["id"],
                    "desc": item["description"],
                    "start": start_utc,
                    "end": end_utc,
                    "tz_str": tz_str,
                }
            )
        except (TypeError, ValueError):
            continue

    for i, ev1 in enumerate(time_events):
        for j, ev2 in enumerate(time_events):
            if i < j:
                if ev1["start"] < ev2["end"] and ev2["start"] < ev1["end"]:
                    conflicts.append(
                        f"⏰ '{ev1['desc']}' (ID: {ev1['id']}) overlaps with "
                        f"'{ev2['desc']}' (ID: {ev2['id']}) "
                        f"from {ev1['start'].strftime('%H:%M')} to {ev1['end'].strftime('%H:%M')} UTC "
                        f"vs {ev2['start'].strftime('%H:%M')}–{ev2['end'].strftime('%H:%M')} UTC"
                    )

    return conflicts


def generate_ics(items, exec_timezone_str, destination):
    """
    Generate an .ics calendar file (as bytes) from itinerary items.
    Hotels become all-day/multi-day events.
    Each event uses its own timezone (fallback to executive's).
    """
    cal = Calendar()
    cal.add("prodid", "-//EA Command Center//local//")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")

    fallback_tz = pytz.timezone(exec_timezone_str)

    for item in items:
        event = Event()
        item_type = item.get("item_type", "").lower()
        item_tz_str = item.get("timezone") or exec_timezone_str
        tz = pytz.timezone(item_tz_str)

        if item_type == "hotel":
            start_dt = datetime.fromisoformat(item["datetime_start"])
            end_dt = (
                datetime.fromisoformat(item["datetime_end"])
                if item["datetime_end"]
                else start_dt + timedelta(days=1)
            )
            start_date = start_dt.date()
            end_date = end_dt.date()

            event.add("summary", f"🏨 Hotel: {item['description']}")
            event.add("dtstart", start_date)
            event.add("dtend", end_date)
            event.add("location", item.get("location", ""))
            desc = f"Confirmation: {item.get('confirmation_code', 'N/A')}"
            if item.get("notes"):
                desc += f"\nNotes: {item['notes']}"
            event.add("description", desc)

        else:
            start_naive = datetime.fromisoformat(item["datetime_start"])
            end_naive = (
                datetime.fromisoformat(item["datetime_end"])
                if item["datetime_end"]
                else start_naive + timedelta(hours=1)
            )
            # Localize to item timezone, then convert to UTC for ICS (many clients handle UTC)
            start_aware = tz.localize(start_naive)
            end_aware = tz.localize(end_naive)
            start_utc = start_aware.astimezone(pytz.UTC)
            end_utc = end_aware.astimezone(pytz.UTC)

            emoji_map = {"flight": "✈️", "meeting": "🤝", "transport": "🚗"}
            emoji = emoji_map.get(item_type, "📌")

            event.add("summary", f"{emoji} {item['description']}")
            event.add("dtstart", start_utc)
            event.add("dtend", end_utc)
            event.add("location", item.get("location", ""))
            desc = f"Type: {item['item_type']}\nConf: {item.get('confirmation_code', 'N/A')}"
            if item.get("cost"):
                desc += f"\nCost: ${item['cost']:.2f}"
            if item.get("notes"):
                desc += f"\nNotes: {item['notes']}"
            desc += f"\nOriginal timezone: {item_tz_str}"
            event.add("description", desc)

        cal.add_component(event)

    return cal.to_ical()


def generate_ics_for_trips(trips_with_items, exec_timezone_str="UTC"):
    """
    Generate a single .ics calendar file containing items from multiple trips.

    Parameters
    ----------
    trips_with_items : list of dict
        Each dict has:
          - "trip": the trip row (dict) — used for the purpose/name prefix
          - "items": the trip's itinerary items (list of dicts)
          - "exec_timezone": IANA timezone string (optional; falls back to
            exec_timezone_str if missing)
    exec_timezone_str : str
        Fallback timezone when a trip doesn't specify one.

    Each event's summary is prefixed with the trip name, e.g.
    "Q1 Asia Roadshow · LOS → NRT (BA 075)" so multi-trip exports
    stay readable.

    Returns bytes (the .ics content).
    """
    cal = Calendar()
    cal.add("prodid", "-//EA Command Center//multi-trip//")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")

    for entry in trips_with_items:
        trip = entry.get("trip") or {}
        items = entry.get("items") or []
        trip_tz_str = entry.get("exec_timezone") or exec_timezone_str
        trip_purpose = (trip.get("purpose") or "").strip()
        prefix = f"{trip_purpose} · " if trip_purpose else ""

        for item in items:
            event = Event()
            item_type = (item.get("item_type") or "").lower()
            item_tz_str = item.get("timezone") or trip_tz_str

            try:
                tz = pytz.timezone(item_tz_str)
            except Exception:
                tz = pytz.timezone(exec_timezone_str)

            if item_type == "hotel":
                start_dt = datetime.fromisoformat(item["datetime_start"])
                end_dt = (
                    datetime.fromisoformat(item["datetime_end"])
                    if item.get("datetime_end")
                    else start_dt + timedelta(days=1)
                )
                event.add("summary", f"🏨 {prefix}Hotel: {item.get('description', '')}")
                event.add("dtstart", start_dt.date())
                event.add("dtend", end_dt.date())
                event.add("location", item.get("location", ""))
                desc = f"Trip: {trip_purpose}\n"
                desc += f"Confirmation: {item.get('confirmation_code', 'N/A')}"
                if item.get("notes"):
                    desc += f"\nNotes: {item['notes']}"
                event.add("description", desc)
            else:
                start_naive = datetime.fromisoformat(item["datetime_start"])
                end_naive = (
                    datetime.fromisoformat(item["datetime_end"])
                    if item.get("datetime_end")
                    else start_naive + timedelta(hours=1)
                )
                try:
                    start_utc = tz.localize(start_naive).astimezone(pytz.UTC)
                    end_utc = tz.localize(end_naive).astimezone(pytz.UTC)
                except Exception:
                    # Fallback if the datetime is already tz-aware
                    start_utc = start_naive.astimezone(pytz.UTC)
                    end_utc = end_naive.astimezone(pytz.UTC)

                emoji_map = {"flight": "✈️", "meeting": "🤝", "transport": "🚗"}
                emoji = emoji_map.get(item_type, "📌")

                event.add("summary", f"{emoji} {prefix}{item.get('description', '')}")
                event.add("dtstart", start_utc)
                event.add("dtend", end_utc)
                event.add("location", item.get("location", ""))
                desc = f"Trip: {trip_purpose}\n"
                desc += f"Type: {item.get('item_type', '')}\n"
                desc += f"Conf: {item.get('confirmation_code', 'N/A')}"
                if item.get("cost"):
                    desc += (
                        f"\nCost: {item['cost']:.2f} {item.get('cost_currency', 'USD')}"
                    )
                if item.get("notes"):
                    desc += f"\nNotes: {item['notes']}"
                desc += f"\nOriginal timezone: {item_tz_str}"
                event.add("description", desc)

            cal.add_component(event)

    return cal.to_ical()


def format_datetime_with_timezone(
    dt_str, item_tz_str, display_tz_str, format_str="%d-%m-%Y %H:%M"
):
    """
    Convert a naive datetime string to a formatted string in a given display timezone.
    If item_tz_str is None, use display_tz_str as both source and target (no conversion).
    """
    if not dt_str:
        return ""
    try:
        dt_naive = datetime.fromisoformat(dt_str)
        # If no item timezone, assume it's already in display timezone
        if not item_tz_str:
            display_tz = pytz.timezone(display_tz_str)
            # Treat naive as if it's in display timezone
            dt_aware = display_tz.localize(dt_naive)
            return dt_aware.strftime(format_str)
        else:
            item_tz = pytz.timezone(item_tz_str)
            display_tz = pytz.timezone(display_tz_str)
            dt_aware = item_tz.localize(dt_naive)
            dt_display = dt_aware.astimezone(display_tz)
            return dt_display.strftime(format_str)
    except Exception:
        return dt_str


import csv
import io

# Header aliases — a header row is detected if every non-empty cell maps
# to one of these. Case- and space-insensitive.
_BULK_COLUMN_ALIASES = {
    "type": "item_type",
    "item_type": "item_type",
    "description": "description",
    "title": "description",
    "summary": "description",
    "start": "datetime_start",
    "start_time": "datetime_start",
    "start_datetime": "datetime_start",
    "begin": "datetime_start",
    "end": "datetime_end",
    "end_time": "datetime_end",
    "end_datetime": "datetime_end",
    "finish": "datetime_end",
    "location": "location",
    "where": "location",
    "venue": "location",
    "cost": "cost",
    "amount": "cost",
    "price": "cost",
    "currency": "cost_currency",
    "ccy": "cost_currency",
    "confirmation": "confirmation_code",
    "conf": "confirmation_code",
    "confirmation_code": "confirmation_code",
    "code": "confirmation_code",
    "notes": "notes",
    "note": "notes",
    "comment": "notes",
}

_BULK_DATETIME_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
    "%d-%m-%Y %H:%M",
    "%d-%m-%Y",
    "%d %b %Y %H:%M",
    "%d %b %Y",
    "%b %d %Y %H:%M",
    "%b %d %Y",
    "%d %B %Y %H:%M",
    "%d %B %Y",
]


def _normalise_header(h):
    return (h or "").strip().lower().replace(" ", "_")


def _parse_bulk_datetime(value):
    """Try each known datetime format. Returns a datetime or None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in _BULK_DATETIME_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def parse_itinerary_paste(raw_text, default_timezone=None):
    """
    Parse a pasted TSV or CSV blob into a list of item dicts.

    Column order is flexible when a header row is present. Without a header,
    positional columns are assumed to be:
      type, description, start, end, location, cost, currency, confirmation

    Returns (items, errors):
      items  — list of dicts ready for db.add_itinerary_item(**)
      errors — list of human-readable strings describing skipped/fixed rows
    """
    items = []
    errors = []

    if not raw_text or not raw_text.strip():
        return items, ["No input provided."]

    lines = [ln for ln in raw_text.splitlines() if ln.strip()]
    if not lines:
        return items, ["No non-empty lines found."]

    sample = lines[0]
    if "\t" in sample:
        delimiter = "\t"
    elif ";" in sample and "," not in sample:
        delimiter = ";"
    else:
        delimiter = ","

    reader = csv.reader(lines, delimiter=delimiter)
    rows = [row for row in reader]
    if not rows:
        return items, ["No rows parsed."]

    first_row_norm = [_normalise_header(c) for c in rows[0]]
    nonempty = [h for h in first_row_norm if h]
    is_header = bool(nonempty) and all(h in _BULK_COLUMN_ALIASES for h in nonempty)

    if is_header:
        header_map = {}
        for pos, h in enumerate(first_row_norm):
            if h in _BULK_COLUMN_ALIASES:
                header_map[pos] = _BULK_COLUMN_ALIASES[h]
        data_rows = rows[1:]
    else:
        default_order = [
            "item_type",
            "description",
            "datetime_start",
            "datetime_end",
            "location",
            "cost",
            "cost_currency",
            "confirmation_code",
        ]
        header_map = {pos: name for pos, name in enumerate(default_order)}
        data_rows = rows

    for row_num, row in enumerate(data_rows, start=1):
        if not any(cell.strip() for cell in row):
            continue

        record = {}
        for pos, cell in enumerate(row):
            if pos not in header_map:
                continue
            record[header_map[pos]] = cell.strip()

        desc = record.get("description", "").strip()
        start_raw = record.get("datetime_start", "").strip()
        if not desc:
            errors.append(f"Row {row_num}: missing description — skipped.")
            continue
        if not start_raw:
            errors.append(f"Row {row_num}: missing start time — skipped.")
            continue

        start_dt = _parse_bulk_datetime(start_raw)
        if start_dt is None:
            errors.append(
                f"Row {row_num}: could not parse start time "
                f"'{start_raw}' — skipped."
            )
            continue

        end_raw = record.get("datetime_end", "").strip()
        if end_raw:
            end_dt = _parse_bulk_datetime(end_raw)
            if end_dt is None:
                end_dt = start_dt + timedelta(hours=1)
                errors.append(
                    f"Row {row_num}: could not parse end time "
                    f"'{end_raw}' — defaulted to +1 hour."
                )
        else:
            end_dt = start_dt + timedelta(hours=1)

        cost_raw = record.get("cost", "").strip()
        try:
            cost_val = float(cost_raw.replace(",", "")) if cost_raw else 0.0
        except ValueError:
            cost_val = 0.0
            errors.append(f"Row {row_num}: cost '{cost_raw}' not numeric — set to 0.")

        currency_val = record.get("cost_currency", "").strip().upper() or "USD"
        item_type = record.get("item_type", "").strip() or "Other"

        items.append(
            {
                "item_type": item_type,
                "description": desc,
                "datetime_start": start_dt.isoformat(),
                "datetime_end": end_dt.isoformat(),
                "location": record.get("location", "").strip(),
                "cost": cost_val,
                "cost_currency": currency_val,
                "confirmation_code": record.get("confirmation_code", "").strip(),
                "notes": record.get("notes", "").strip(),
                "timezone": default_timezone,
                "is_confirmed": 0,
                "delegation_ids": [],
                "contact_ids": [],
                "venue_id": None,
                "cost_date": start_dt.date().isoformat(),
            }
        )

    return items, errors


# Header aliases for expense bulk import
_EXPENSE_COLUMN_ALIASES = {
    "traveler": "traveler",
    "traveller": "traveler",
    "name": "traveler",
    "person": "traveler",
    "member": "traveler",
    "date": "expense_date",
    "expense_date": "expense_date",
    "when": "expense_date",
    "category": "category",
    "type": "category",
    "description": "description",
    "desc": "description",
    "detail": "description",
    "amount": "amount",
    "cost": "amount",
    "total": "amount",
    "currency": "currency",
    "ccy": "currency",
    "reimbursable": "is_reimbursable",
    "reim": "is_reimbursable",
    "notes": "notes",
    "note": "notes",
    "comment": "notes",
    "receipt": "receipt_path",
    "receipt_path": "receipt_path",
}

_EXPENSE_CATEGORIES = [
    "Meals", "Transport", "Lodging", "Incidentals",
    "Entertainment", "Communication", "Other",
]


def parse_expense_paste(raw_text, default_currency="USD"):
    """
    Parse a pasted TSV or CSV blob into a list of expense dicts.

    Expected columns (header row strongly recommended):
      traveler, date, category, description, amount,
      currency, reimbursable, notes, receipt_path

    Returns (expenses, errors):
      expenses — list of dicts; `traveler` is a raw string the caller
                 must resolve to a contact_id
      errors   — list of human-readable warnings
    """
    expenses = []
    errors = []

    if not raw_text or not raw_text.strip():
        return expenses, ["No input provided."]

    lines = [ln for ln in raw_text.splitlines() if ln.strip()]
    if not lines:
        return expenses, ["No non-empty lines found."]

    sample = lines[0]
    if "\t" in sample:
        delimiter = "\t"
    elif ";" in sample and "," not in sample:
        delimiter = ";"
    else:
        delimiter = ","

    reader = csv.reader(lines, delimiter=delimiter)
    rows = [row for row in reader]
    if not rows:
        return expenses, ["No rows parsed."]

    first_norm = [_normalise_header(c) for c in rows[0]]
    nonempty = [h for h in first_norm if h]
    is_header = bool(nonempty) and all(
        h in _EXPENSE_COLUMN_ALIASES for h in nonempty
    )
    if not is_header:
        return expenses, [
            "A header row is required for expense import. "
            "Expected columns: traveler, date, category, description, "
            "amount, currency, reimbursable, notes."
        ]

    header_map = {}
    for pos, h in enumerate(first_norm):
        if h in _EXPENSE_COLUMN_ALIASES:
            header_map[pos] = _EXPENSE_COLUMN_ALIASES[h]

    for row_num, row in enumerate(rows[1:], start=2):
        if not any(cell.strip() for cell in row):
            continue

        record = {}
        for pos, cell in enumerate(row):
            if pos in header_map:
                record[header_map[pos]] = cell.strip()

        traveler = record.get("traveler", "").strip()
        if not traveler:
            errors.append(f"Row {row_num}: missing traveler — skipped.")
            continue

        date_raw = record.get("expense_date", "").strip()
        date_dt = _parse_bulk_datetime(date_raw) if date_raw else None
        if date_dt is None:
            errors.append(
                f"Row {row_num}: could not parse date '{date_raw}' — skipped."
            )
            continue

        amount_raw = record.get("amount", "").strip()
        try:
            amount_val = float(amount_raw.replace(",", "")) if amount_raw else 0.0
        except ValueError:
            errors.append(
                f"Row {row_num}: amount '{amount_raw}' not numeric — skipped."
            )
            continue
        if amount_val <= 0:
            errors.append(
                f"Row {row_num}: amount must be > 0 — skipped."
            )
            continue

        category = record.get("category", "").strip() or "Other"
        if category not in _EXPENSE_CATEGORIES:
            # Coerce unknown categories rather than reject the row
            errors.append(
                f"Row {row_num}: category '{category}' unknown — "
                f"set to 'Other'."
            )
            category = "Other"

        reim_raw = record.get("is_reimbursable", "").strip().lower()
        if reim_raw in ("", "yes", "y", "true", "1"):
            is_reim = 1
        elif reim_raw in ("no", "n", "false", "0"):
            is_reim = 0
        else:
            is_reim = 1
            errors.append(
                f"Row {row_num}: reimbursable '{reim_raw}' not understood — "
                f"defaulted to Yes."
            )

        expenses.append({
            "traveler": traveler,
            "expense_date": date_dt.date().isoformat(),
            "category": category,
            "description": record.get("description", "").strip(),
            "amount": amount_val,
            "currency": (
                record.get("currency", "").strip().upper() or default_currency
            ),
            "is_reimbursable": is_reim,
            "notes": record.get("notes", "").strip(),
            "receipt_path": record.get("receipt_path", "").strip() or None,
        })

    return expenses, errors

    # =========================================================
# ICS EXPORTS — Meetings, Events, Speaking
# =========================================================


def _parse_time(time_str, fallback_hour=9, fallback_min=0):
    """
    Parse a user-entered time like '14:30', '2:30 PM', or '1430'.
    Returns (hour, minute). Falls back to defaults if unparseable.
    """
    if not time_str:
        return fallback_hour, fallback_min
    s = str(time_str).strip().lower()
    try:
        if "am" in s or "pm" in s:
            # 12-hour format
            t = s.replace("am", "").replace("pm", "").strip()
            if ":" in t:
                h, m = t.split(":")
            else:
                h, m = t, "0"
            h = int(h)
            m = int(m)
            if "pm" in s and h != 12:
                h += 12
            if "am" in s and h == 12:
                h = 0
            return h, m
        if ":" in s:
            h, m = s.split(":")
            return int(h), int(m)
        if len(s) == 4 and s.isdigit():
            return int(s[:2]), int(s[2:])
        if s.isdigit():
            return int(s), 0
    except Exception:
        pass
    return fallback_hour, fallback_min


def _ics_header():
    cal = Calendar()
    cal.add("prodid", "-//EA Command Center//local//")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")
    return cal


def generate_ics_for_meetings(meetings, exec_timezones=None):
    """
    Build an .ics file from a list of meeting records.

    meetings: list of dicts (from db.get_meetings)
    exec_timezones: dict {exec_id: tz_string} for per-executive timezone.
                    Falls back to UTC if missing.
    """
    cal = _ics_header()
    tz_map = exec_timezones or {}

    for m in meetings:
        if not m.get("meeting_date"):
            continue
        exec_id = m.get("exec_id")
        tz_str = tz_map.get(exec_id) or "UTC"
        try:
            tz = pytz.timezone(tz_str)
        except Exception:
            tz = pytz.UTC

        try:
            base_date = datetime.fromisoformat(m["meeting_date"]).date()
        except Exception:
            continue

        hour, minute = _parse_time(m.get("meeting_time"))
        start_naive = datetime(
            base_date.year, base_date.month, base_date.day, hour, minute
        )
        duration = m.get("duration_minutes") or 60
        end_naive = start_naive + timedelta(minutes=duration)

        try:
            start_utc = tz.localize(start_naive).astimezone(pytz.UTC)
            end_utc = tz.localize(end_naive).astimezone(pytz.UTC)
        except Exception:
            start_utc = start_naive.replace(tzinfo=pytz.UTC)
            end_utc = end_naive.replace(tzinfo=pytz.UTC)

        event = Event()
        event.add("uid", f"meeting-{m['id']}@ea-command-center")
        event.add("summary", f"📅 {m.get('title') or 'Meeting'}")
        event.add("dtstart", start_utc)
        event.add("dtend", end_utc)

        loc = m.get("location") or m.get("video_link") or ""
        if loc:
            event.add("location", loc)

        desc_parts = []
        if m.get("executive_name"):
            desc_parts.append(f"Executive: {m['executive_name']}")
        if m.get("meeting_type"):
            desc_parts.append(f"Type: {m['meeting_type']}")
        if m.get("status"):
            desc_parts.append(f"Status: {m['status']}")
        if m.get("agenda"):
            desc_parts.append(f"\nAgenda:\n{m['agenda']}")
        if m.get("notes_native"):
            desc_parts.append(f"\nNotes:\n{m['notes_native']}")
        if m.get("calendar_event_url"):
            desc_parts.append(f"\nCalendar: {m['calendar_event_url']}")
        if desc_parts:
            event.add("description", "\n".join(desc_parts))

        cal.add_component(event)

    return cal.to_ical()


def generate_ics_for_events(events, exec_timezones=None):
    """
    Build an .ics file from a list of event records.
    Events render as one event per day of the event (single day is most
    common). Multi-day events use DTEND = end date + 1 day.
    """
    cal = _ics_header()
    tz_map = exec_timezones or {}

    for ev in events:
        if not ev.get("event_date"):
            continue
        exec_id = ev.get("exec_id")
        tz_str = tz_map.get(exec_id) or "UTC"
        try:
            tz = pytz.timezone(tz_str)
        except Exception:
            tz = pytz.UTC

        try:
            base_date = datetime.fromisoformat(ev["event_date"]).date()
        except Exception:
            continue

        # Parse start/end times
        start_h, start_m = _parse_time(ev.get("start_time"), 9, 0)
        end_h, end_m = _parse_time(ev.get("end_time"), start_h + 2, start_m)

        start_naive = datetime(
            base_date.year, base_date.month, base_date.day, start_h, start_m
        )
        end_naive = datetime(
            base_date.year, base_date.month, base_date.day, end_h, end_m
        )
        # If end time is before start, assume +12 hours
        if end_naive <= start_naive:
            end_naive = start_naive + timedelta(hours=2)

        try:
            start_utc = tz.localize(start_naive).astimezone(pytz.UTC)
            end_utc = tz.localize(end_naive).astimezone(pytz.UTC)
        except Exception:
            start_utc = start_naive.replace(tzinfo=pytz.UTC)
            end_utc = end_naive.replace(tzinfo=pytz.UTC)

        event = Event()
        event.add("uid", f"event-{ev['id']}@ea-command-center")
        event.add("summary", f"🎉 {ev.get('name') or 'Event'}")
        event.add("dtstart", start_utc)
        event.add("dtend", end_utc)

        loc = ev.get("venue_name") or ev.get("location_override") or ""
        if ev.get("venue_city"):
            loc = f"{loc}, {ev['venue_city']}".strip(", ")
        if loc:
            event.add("location", loc)

        desc_parts = []
        if ev.get("executive_name"):
            desc_parts.append(f"Executive: {ev['executive_name']}")
        if ev.get("event_type"):
            desc_parts.append(f"Type: {ev['event_type']}")
        if ev.get("status"):
            desc_parts.append(f"Status: {ev['status']}")
        if ev.get("objective"):
            desc_parts.append(f"\nObjective:\n{ev['objective']}")
        if ev.get("calendar_event_url"):
            desc_parts.append(f"\nCalendar: {ev['calendar_event_url']}")
        if desc_parts:
            event.add("description", "\n".join(desc_parts))

        cal.add_component(event)

    return cal.to_ical()


def generate_ics_for_speaking(engagements, exec_timezones=None):
    """
    Build an .ics file from a list of speaking engagement records.
    """
    cal = _ics_header()
    tz_map = exec_timezones or {}

    for eng in engagements:
        if not eng.get("engagement_date"):
            continue
        exec_id = eng.get("exec_id")
        tz_str = tz_map.get(exec_id) or eng.get("timezone") or "UTC"
        try:
            tz = pytz.timezone(tz_str)
        except Exception:
            tz = pytz.UTC

        try:
            base_date = datetime.fromisoformat(eng["engagement_date"]).date()
        except Exception:
            continue

        start_h, start_m = _parse_time(eng.get("engagement_time"), 9, 0)
        # Duration is free-form ("45 min"), try to parse
        duration_min = 60
        dur_str = (eng.get("duration") or "").lower()
        try:
            if "min" in dur_str:
                duration_min = int(
                    "".join(c for c in dur_str.split("min")[0] if c.isdigit())
                )
            elif "hour" in dur_str or "hr" in dur_str:
                hours = int(
                    "".join(
                        c
                        for c in dur_str.split("hour")[0].split("hr")[0]
                        if c.isdigit()
                    )
                )
                duration_min = hours * 60
        except Exception:
            pass

        start_naive = datetime(
            base_date.year, base_date.month, base_date.day, start_h, start_m
        )
        end_naive = start_naive + timedelta(minutes=duration_min)

        try:
            start_utc = tz.localize(start_naive).astimezone(pytz.UTC)
            end_utc = tz.localize(end_naive).astimezone(pytz.UTC)
        except Exception:
            start_utc = start_naive.replace(tzinfo=pytz.UTC)
            end_utc = end_naive.replace(tzinfo=pytz.UTC)

        event = Event()
        event.add("uid", f"speaking-{eng['id']}@ea-command-center")
        event.add(
            "summary",
            f"🎤 {eng.get('title') or 'Speaking Engagement'}",
        )
        event.add("dtstart", start_utc)
        event.add("dtend", end_utc)

        if eng.get("location"):
            event.add("location", eng["location"])

        desc_parts = []
        if eng.get("executive_name"):
            desc_parts.append(f"Executive: {eng['executive_name']}")
        if eng.get("organizer_name"):
            desc_parts.append(f"Organizer: {eng['organizer_name']}")
        if eng.get("engagement_type"):
            desc_parts.append(f"Type: {eng['engagement_type']}")
        if eng.get("format"):
            desc_parts.append(f"Format: {eng['format']}")
        if eng.get("topic"):
            desc_parts.append(f"\nTopic: {eng['topic']}")
        if eng.get("audience"):
            desc_parts.append(f"Audience: {eng['audience']}")
        if desc_parts:
            event.add("description", "\n".join(desc_parts))

        cal.add_component(event)

    return cal.to_ical()
