from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
from datetime import datetime
import json
import os
import base64
from jinja2 import Environment, FileSystemLoader
import pytz
import database as db
from weasyprint import HTML
import currency


# ---------------------------------------------------------
# Stage 6 – currency conversion helper for export totals
# ---------------------------------------------------------
def _convert_item_cost(item, trip_base_currency):
    """
    Convert a single itinerary item's cost to the trip's base currency.

    Uses the item's `cost_date` (falling back to the date part of
    `datetime_start`) so historical rates are honoured. If the new
    conversion pipeline raises for any reason, falls back to the legacy
    `exchange_rate_snapshot` multiplier so exports never crash.
    """
    cost = item.get("cost", 0) or 0
    cost_cur = (item.get("cost_currency") or "USD").upper()
    base_cur = (trip_base_currency or "USD").upper()

    on_date = item.get("cost_date")
    if not on_date and item.get("datetime_start"):
        on_date = item["datetime_start"][:10]  # YYYY-MM-DD

    try:
        return currency.convert_amount(cost, cost_cur, base_cur, on_date)
    except Exception:
        # No rate available anywhere — treat the amount as if it were already
        # in the base currency so exports never crash on missing data.
        return float(cost)


def _sum_items_in_base(items, trip_base_currency):
    """
    Return (total, confirmed, estimated) for a list of items,
    converted to `trip_base_currency`.
    """
    total = 0.0
    confirmed = 0.0
    estimated = 0.0
    for item in items:
        converted = _convert_item_cost(item, trip_base_currency)
        total += converted
        if item.get("is_confirmed", 0):
            confirmed += converted
        else:
            estimated += converted
    return total, confirmed, estimated

def _build_weather_by_stop(stops):
    """
    Fetch weather per stop, one entry per stop that has a city and dates.

    Each entry carries:
      city, country, region, start_date, end_date,
      location, fetched_at, is_current, is_past, available,
      current (only populated for the stop containing today),
      daily  (filtered to the stop's own date window)

    Falls back to the base get_weather() snapshot when the range function
    is missing or returns nothing. Stops outside the forecast horizon or
    in the past come back with available=False and empty daily.
    """
    result = []
    if not stops:
        return result

    try:
        import weather as weather_module
    except ImportError:
        return result

    today = datetime.now().date()

    for stop in stops:
        city = (stop.get("city") or "").strip()
        start_raw = stop.get("start_date")
        end_raw = stop.get("end_date")
        if not (city and start_raw and end_raw):
            continue

        start_iso = str(start_raw)[:10]
        end_iso = str(end_raw)[:10]

        try:
            start_d = datetime.fromisoformat(start_iso).date()
            end_d = datetime.fromisoformat(end_iso).date()
        except Exception:
            continue

        is_current = start_d <= today <= end_d
        is_past = end_d < today

        # Try range-aware fetch first
        data = None
        if hasattr(weather_module, "get_weather_for_range"):
            try:
                data = weather_module.get_weather_for_range(
                    city, start_iso, end_iso
                )
            except Exception:
                data = None

        # Fall back to base snapshot
        if not data and hasattr(weather_module, "get_weather"):
            try:
                data = weather_module.get_weather(city)
            except Exception:
                data = None

        # Filter the daily array to the stop's own window. Some weather
        # sources return a week regardless of what was asked for.
        raw_daily = (data or {}).get("daily") or []
        in_window = [
            d for d in raw_daily
            if start_iso <= (d.get("date") or "") <= end_iso
        ]

        # "Now" is only meaningful for the stop we're currently at.
        current = (data or {}).get("current") if is_current else None

        available = bool(in_window or current)
        if is_past:
            available = False

        result.append({
            "city": city,
            "country": stop.get("country") or "",
            "region": stop.get("region") or "",
            "start_date": start_iso,
            "end_date": end_iso,
            "location": (data or {}).get("location") or city,
            "fetched_at": (data or {}).get("fetched_at") or "",
            "is_current": is_current,
            "is_past": is_past,
            "available": available,
            "current": current,
            "daily": in_window,
        })

    return result

def generate_executive_profile_doc(profile_data, exec_id, currency_symbol="$"):
    """
    Generate a Word document with the executive profile in a bulleted list.
    """
    doc = Document()
    doc.add_heading("Executive Profile", 0)

    doc.add_heading("Personal Information", level=1)
    fields = [
        ("Name", profile_data.get("Name") or ""),
        ("Email", profile_data.get("Email") or ""),
        ("Timezone", profile_data.get("Timezone") or ""),
        ("Seat Preference", profile_data.get("Seat Preference") or ""),
        ("Dietary", profile_data.get("Dietary") or ""),
        ("Company", profile_data.get("Company") or ""),
        ("Cost Center", profile_data.get("Cost Center") or ""),
        ("Policy Notes", profile_data.get("Policy Notes") or ""),
        ("Preferred Airline", profile_data.get("Preferred Airline") or ""),
        ("TSA PreCheck", profile_data.get("TSA PreCheck") or ""),
        ("Meal Preference", profile_data.get("Meal Preference") or ""),
    ]
    for field, value in fields:
        if value:
            doc.add_paragraph(f"{field}: {value}", style="List Bullet")

    doc.add_heading("Memberships", level=1)
    mems = profile_data.get("Memberships", "")
    if mems:
        doc.add_paragraph(mems, style="List Bullet")
    else:
        doc.add_paragraph("No memberships recorded.")

    doc.add_page_break()
    doc.add_paragraph(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph("Executive Travel Planner")

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


def generate_itinerary_doc(
    exec_data,
    items,
    stops,
    dep_city,
    dep_region,
    dep_country,
    trip_id,
    trip_budget,
    display_symbol,
    display_currency,
    base_currency="USD",
    convert_to_base=False,
):
    """
    Generate a Word document with trip itinerary.
    Includes contacts column.
    """
    doc = Document()
    doc.add_heading(f"Itinerary: {exec_data['name']}", 0)

    # Trip overview
    doc.add_heading("Trip Overview", level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Detail"
    hdr_cells[1].text = "Value"

    dep_parts = [p for p in [dep_city, dep_region, dep_country] if p]
    departure_display = ", ".join(dep_parts) if dep_parts else "Not specified"
    rows_data = [
        ("Executive", exec_data["name"]),
        ("Departure", departure_display),
        ("Budget", f"{display_symbol}{trip_budget:.2f} {display_currency}"),
        ("Base Currency", base_currency),
    ]
    for label, value in rows_data:
        row_cells = table.add_row().cells
        row_cells[0].text = label
        row_cells[1].text = str(value)

    # Stops
    if stops:
        doc.add_heading("Stops", level=1)
        for stop in stops:
            loc = stop["city"]
            loc_parts = []
            if stop.get("region"):
                loc_parts.append(stop["region"])
            if stop.get("country"):
                loc_parts.append(stop["country"])
            if loc_parts:
                loc += f" ({', '.join(loc_parts)})"
            doc.add_paragraph(
                f"• {loc}: {stop['start_date']} to {stop['end_date']}",
                style="List Bullet",
            )

    # Itinerary items
    if items:
        doc.add_heading("Itinerary", level=1)
        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        headers = ["Date/Time", "Type", "Description", "Location", "Contacts", "Cost"]
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            hdr_cells[i].paragraphs[0].runs[0].bold = True

        for item in items:
            row_cells = table.add_row().cells
            start_dt = datetime.fromisoformat(item["datetime_start"])
            row_cells[0].text = start_dt.strftime("%Y-%m-%d %H:%M")
            row_cells[1].text = item.get("item_type", "")
            row_cells[2].text = item.get("description", "")
            row_cells[3].text = item.get("location", "")
            contacts = db.get_item_contacts(item["id"])
            contact_str = ", ".join([c["name"] for c in contacts]) if contacts else ""
            row_cells[4].text = contact_str
            cost = item.get("cost", 0)
            currency = item.get("cost_currency", "USD")
            row_cells[5].text = f"{cost:.2f} {currency}"

    doc.add_page_break()
    doc.add_paragraph("Generated by Executive Travel Planner")
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream, f"{exec_data['name']}_itinerary.docx"


def generate_expense_report_doc(
    exec_data,
    items,
    stops,
    dep_city,
    dep_region,
    dep_country,
    trip_id,
    trip_budget,
    trip_purpose,
    display_symbol,
    base_currency="USD",
):
    """
    Generate a Word document with expense report.
    Includes contacts column.
    """
    doc = Document()
    doc.add_heading(f"Expense Report: {trip_purpose}", 0)

        # Summary (converted to trip base currency – Stage 6)
    trip_base_cur = exec_data.get("base_currency", base_currency) or base_currency
    total_spent, confirmed_spent, estimated_spent = _sum_items_in_base(
        items, trip_base_cur
    )



    doc.add_heading("Summary", level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Metric"
    hdr_cells[1].text = "Value"

    rows_data = [
        ("Total Budget", f"{display_symbol}{trip_budget:.2f} {base_currency}"),
        ("Total Spent", f"{display_symbol}{total_spent:.2f} {base_currency}"),
        (
            "Confirmed (Booked)",
            f"{display_symbol}{confirmed_spent:.2f} {base_currency}",
        ),
        (
            "Pending (Quoted)",
            f"{display_symbol}{estimated_spent:.2f} {base_currency}",
        ),
        (
            "Remaining",
            f"{display_symbol}{trip_budget - total_spent:.2f} {base_currency}",
        ),
    ]
    for label, value in rows_data:
        row_cells = table.add_row().cells
        row_cells[0].text = label
        row_cells[1].text = value

    # Itemized list
    doc.add_heading("Expense Details", level=1)
    if items:
        table = doc.add_table(rows=1, cols=7)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        headers = [
            "Date/Time",
            "Type",
            "Description",
            "Location",
            "Contacts",
            "Cost",
            "Confirmed",
        ]
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            hdr_cells[i].paragraphs[0].runs[0].bold = True

        for item in items:
            row_cells = table.add_row().cells
            start_dt = datetime.fromisoformat(item["datetime_start"])
            row_cells[0].text = start_dt.strftime("%Y-%m-%d %H:%M")
            row_cells[1].text = item.get("item_type", "")
            row_cells[2].text = item.get("description", "")
            row_cells[3].text = item.get("location", "")
            contacts = db.get_item_contacts(item["id"])
            contact_str = ", ".join([c["name"] for c in contacts]) if contacts else ""
            row_cells[4].text = contact_str
            cost = item.get("cost", 0)
            currency = item.get("cost_currency", "USD")
            row_cells[5].text = f"{cost:.2f} {currency}"
            row_cells[6].text = "✅" if item.get("is_confirmed", 0) else "❌"
    else:
        doc.add_paragraph("No expenses recorded.")

    doc.add_page_break()
    doc.add_paragraph("Generated by Executive Travel Planner")
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream, f"{exec_data['name']}_ExpenseReport.docx"


def generate_spending_report_doc(
    filter_label,
    summary_data,
    start_filter,
    end_filter,
    currency_symbol,
    base_currency="USD",
):
    """
    Generate a Word document with spending summary report.
    """
    doc = Document()
    doc.add_heading("Spending Report", 0)

    doc.add_heading("Filters", level=1)
    doc.add_paragraph(f"Executive: {filter_label}")
    if start_filter:
        doc.add_paragraph(f"Start Date: {start_filter}")
    if end_filter:
        doc.add_paragraph(f"End Date: {end_filter}")
    doc.add_paragraph(f"Base Currency (for numeric columns): {base_currency}")

    doc.add_heading("Trip Breakdown", level=1)
    if summary_data:
        table = doc.add_table(rows=1, cols=7)
        table.style = "Table Grid"
        hdr_cells = table.rows[0].cells
        headers = [
            "Executive",
            "Destination",
            "Budget",
            "Total Spent",
            "Confirmed",
            "Status",
            "Currency",
        ]
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            hdr_cells[i].paragraphs[0].runs[0].bold = True

        for trip in summary_data:
            row_cells = table.add_row().cells
            row_cells[0].text = trip.get("executive_name", "")
            row_cells[1].text = trip.get("destination", "")
            row_cells[2].text = (
                f"{trip.get('budget', 0):.2f} {trip.get('base_currency', 'USD')}"
            )
            row_cells[3].text = (
                f"{trip.get('total_spent', 0):.2f} {trip.get('base_currency', 'USD')}"
            )
            row_cells[4].text = (
                f"{trip.get('confirmed_spent', 0):.2f} {trip.get('base_currency', 'USD')}"
            )
            row_cells[5].text = trip.get("status", "").title()
            row_cells[6].text = trip.get("base_currency", "USD")
    else:
        doc.add_paragraph("No trips found.")

    doc.add_page_break()
    doc.add_paragraph("Generated by Executive Travel Planner")
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


def generate_all_executive_profiles_doc(profiles):
    """
    Generate a Word document with a list of all executive profiles.
    """
    doc = Document()
    doc.add_heading("All Executive Profiles", 0)

    for p in profiles:
        name = p.get("Name", "Unknown Executive")
        doc.add_heading(name, level=1)

        fields = [
            ("Email", p.get("Email")),
            ("Company", p.get("Company")),
            ("Timezone", p.get("Timezone")),
            ("Seat Preference", p.get("Seat Preference")),
            ("Memberships", p.get("Memberships")),
            ("Meal Preference", p.get("Meal Preference")),
        ]
        for label, value in fields:
            if value:
                doc.add_paragraph(f"{label}: {value}", style="List Bullet")

        doc.add_paragraph()

    doc.add_page_break()
    doc.add_paragraph("Generated by Executive Travel Planner")
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# ---- Travel Pack HTML ----
def generate_travel_pack_html(trip_id, exec_timezone, display_mode="Home"):
    """
    Generate a self-contained HTML travel pack.
    Includes contacts per item, delegation, venue, and receipts.
    """
    trip = db.get_trip(trip_id)
    if not trip:
        return None
    executive = db.get_executive_profile(trip["exec_id"])
    if not executive:
        executive = {}

    trip_venues = db.get_venues_for_trip(trip_id)
    expense_summary = db.get_expense_summary(trip_id)
    stops = db.get_trip_stops(trip_id)

    # ---- Emergency info ----
    emergency_hospitals = db.get_hospitals_for_trip(trip_id)

    # Country-level emergency numbers from Destination Guides
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

    # ---- Weather per stop ----
    weather_by_stop = _build_weather_by_stop(stops)

    # ---- Packing lists (per member) ----
    packing_lists = []
    for member in db.get_trip_delegation_members(trip_id):
        plist = db.get_packing_list(trip_id, member["id"])
        if not plist:
            continue
        packing_items_local = db.get_packing_items(plist["id"])
        if not packing_items_local:
            continue
        packing_lists.append({
            "member_name": member["name"],
            "role": member.get("role", ""),
            "total_count": len(packing_items_local),
            "packed_count": sum(
                1 for it in packing_items_local if it["packed"]
            ),
            "items": packing_items_local,
        })

    # ---- Trip checklists ----
    trip_checklists = []
    for cl in db.get_trip_checklists(trip_id):
        cl_items_local = db.get_checklist_items(cl["id"])
        if not cl_items_local:
            continue
        trip_checklists.append({
            "name": cl["name"],
            "description": cl.get("description", ""),
            "total_count": len(cl_items_local),
            "done_count": sum(1 for it in cl_items_local if it["is_done"]),
            "items": cl_items_local,
        })

    total_allowance = sum(s["allowance"] for s in expense_summary)
    total_spent_expenses = sum(s["spent"] for s in expense_summary)
    items = db.get_items_for_trip(trip_id)
    contacts = db.get_trip_contacts(trip_id)
    memberships = db.get_memberships(trip["exec_id"])

    # ---- Delegation: only members assigned to items on this trip ----
    company_id_for_delegation = (
        executive.get("company_id") if executive else None
    )
    delegation = []
    if company_id_for_delegation:
        assigned_ids = set()
        for item in items:
            for m in db.get_item_delegation_members(item["id"]):
                assigned_ids.add(m["id"])
        if assigned_ids:
            all_members = db.get_delegation_members(
                company_id_for_delegation, active_only=True
            )
            delegation = [m for m in all_members if m["id"] in assigned_ids]

    # ---- Budget totals (converted to trip base currency – Stage 6) ----
    base_cur = trip.get("base_currency", "USD")
    total_spent, confirmed_spent, estimated_spent = _sum_items_in_base(
        items, base_cur
    )

    # ---- Format items with timezone + participants + contacts ----
    formatted_items = []
    receipts = []
    fallback_tz = pytz.timezone(exec_timezone)

    for item in items:
        item_tz_str = item.get("timezone") or exec_timezone
        try:
            item_tz = pytz.timezone(item_tz_str)
        except Exception:
            item_tz = fallback_tz

        dt_naive = datetime.fromisoformat(item["datetime_start"])
        dt_aware = item_tz.localize(dt_naive)

        if display_mode == "Home":
            display_tz = fallback_tz
        else:
            display_tz = item_tz

        dt_display = dt_aware.astimezone(display_tz)
        formatted_time = dt_display.strftime("%d-%m-%Y %H:%M %Z")

        # Delegation members assigned to this item
        delegation_members = db.get_item_delegation_members(item["id"])
        delegation_names = (
            [m["name"] for m in delegation_members]
            if delegation_members else []
        )

        # Contacts assigned to this item
        item_contacts = db.get_item_contacts(item["id"])
        contact_names = (
            [c["name"] for c in item_contacts] if item_contacts else []
        )

        # ---- Venue lookup for this item ----
        venue_name_for_item = None
        if item.get("venue_id"):
            v = db.get_venue(item["venue_id"])
            if v:
                venue_name_for_item = v["name"]

        item_dict = {
            "id": item["id"],
            "item_type": item["item_type"],
            "description": item["description"],
            "location": item.get("location", ""),
            "cost": item.get("cost", 0),
            "cost_currency": item.get("cost_currency", "USD"),
            "confirmation_code": item.get("confirmation_code", ""),
            "notes": item.get("notes", ""),
            "is_confirmed": item.get("is_confirmed", 0),
            "formatted_time": formatted_time,
            "participants": delegation_names,  # used for Agenda attendees
            "contacts": contact_names,          # for future use
            "venue_id": item.get("venue_id"),
            "venue_name": venue_name_for_item,
        }
        formatted_items.append(item_dict)

        # Collect receipts
        receipt_path = item.get("receipt_path")
        if receipt_path and os.path.exists(receipt_path):
            with open(receipt_path, "rb") as f:
                image_data = f.read()
                b64 = base64.b64encode(image_data).decode("utf-8")
                ext = os.path.splitext(receipt_path)[1].lower()
                mime = (
                    "image/png"
                    if ext == ".png"
                    else "image/jpeg"
                    if ext in [".jpg", ".jpeg"]
                    else "application/pdf"
                )
                data_uri = f"data:{mime};base64,{b64}"
                receipts.append(
                    {"description": item["description"], "data_uri": data_uri}
                )

    # ---- Build context for Jinja2 ----
    context = {
        "trip": trip,
        "executive": executive,
        "delegation": delegation,  # <-- ADDED (fixes missing Delegation section)
        "stops": stops,
        "items": formatted_items,
        "contacts": contacts,
        "memberships": memberships,
        "total_spent": total_spent,
        "confirmed_spent": confirmed_spent,
        "estimated_spent": estimated_spent,
        "receipts": receipts,
        "trip_venues": trip_venues,
        "now": datetime.now(),
        "expense_summary": expense_summary,
        "total_allowance": total_allowance,
        "total_spent_expenses": total_spent_expenses,
        "packing_lists": packing_lists,
        "trip_checklists": trip_checklists,
        "emergency_hospitals": emergency_hospitals,
        "trip_emergency_numbers": trip_emergency_numbers,
        "weather_by_stop": weather_by_stop,
    }

    env = Environment(loader=FileSystemLoader("templates"))
    template = env.get_template("travel_pack.html")
    html = template.render(**context)
    return html


# ---- Company Profile HTML ----
def generate_company_profile_html(company_id):
    """Generate a self‑contained HTML company profile."""
    company = db.get_company(company_id)
    if not company:
        return None
    executives = db.get_executives_by_company(company_id)
    contacts = db.get_contacts(company_id, active_only=True)
    delegation = db.get_delegation_members(company_id, active_only=True)

    context = {
        "company": company,
        "executives": executives,
        "contacts": contacts,
        "participants": delegation,  # template uses 'participants'
        "now": datetime.now(),
    }
    env = Environment(loader=FileSystemLoader("templates"))
    template = env.get_template("company_profile.html")
    return template.render(**context)


# ---- Company Profile Word ----
def generate_company_profile_docx(company_id):
    """Generate a Word document with company profile."""
    company = db.get_company(company_id)
    if not company:
        return None
    executives = db.get_executives_by_company(company_id)
    contacts = db.get_contacts(company_id, active_only=True)
    delegation = db.get_delegation_members(company_id, active_only=True)

    doc = Document()
    doc.add_heading(f"Company Profile: {company['name']}", 0)
    doc.add_paragraph(f"Cost Center: {company.get('default_cost_center') or 'N/A'}")
    if company.get("policy_notes"):
        doc.add_paragraph(f"Policy Notes: {company['policy_notes']}")

    doc.add_heading("Executives", level=1)
    if executives:
        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "Name"
        hdr[1].text = "Email"
        hdr[2].text = "Timezone"
        hdr[3].text = "Seat Preference"
        hdr[4].text = "Meal Preference"
        for e in executives:
            row = table.add_row().cells
            row[0].text = e.get("name", "")
            row[1].text = e.get("email", "")
            row[2].text = e.get("timezone", "")
            row[3].text = e.get("seat_preference", "")
            row[4].text = e.get("meal_preference", "")
    else:
        doc.add_paragraph("No executives assigned.")

    doc.add_heading("Contacts", level=1)
    if contacts:
        for c in contacts:
            p = doc.add_paragraph()
            p.add_run(f"{c['name']}").bold = True
            if c.get("role"):
                p.add_run(f" ({c['role']})")
            if c.get("type"):
                p.add_run(f" [{c['type']}]")
            location = c.get("city", "")
            if c.get("country"):
                location += f", {c['country']}" if location else c["country"]
            p.add_run(f"\n📞 {c.get('phone', '')}  ✉️ {c.get('email', '')}")
            if location:
                p.add_run(f"  🌍 {location}")
            if c.get("tags"):
                p.add_run(f"\n🏷️ {c['tags']}")
    else:
        doc.add_paragraph("No contacts.")

    doc.add_paragraph(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


def generate_travel_pack_pdf(trip_id, exec_timezone, display_mode="Home"):
    """
    Generate a PDF version of the travel pack using WeasyPrint.
    Returns a BytesIO stream.
    """
    html_str = generate_travel_pack_html(trip_id, exec_timezone, display_mode)
    if not html_str:
        return None
    pdf_bytes = HTML(string=html_str).write_pdf()
    return io.BytesIO(pdf_bytes)


def generate_travel_pack_docx(trip_id, exec_timezone, display_mode="Home"):
    """
    Generate a Word (.docx) version of the travel pack.
    Mirrors the HTML travel pack structure with the same sections.
    """
    trip = db.get_trip(trip_id)
    if not trip:
        return None

    executive = db.get_executive_profile(trip["exec_id"])
    if not executive:
        executive = {}

    venue = db.get_venue(trip_id)
    stops = db.get_trip_stops(trip_id)
    items = db.get_items_for_trip(trip_id)
    contacts = db.get_trip_contacts(trip_id)
    trip_venues = db.get_venues_for_trip(trip_id)

    # ---- Delegation: only members assigned to items on this trip ----
    company_id_for_delegation = executive.get("company_id") if executive else None
    delegation = []
    if company_id_for_delegation:
        assigned_ids = set()
        for item in items:
            for m in db.get_item_delegation_members(item["id"]):
                assigned_ids.add(m["id"])
        if assigned_ids:
            all_members = db.get_delegation_members(
                company_id_for_delegation, active_only=True
            )
            delegation = [m for m in all_members if m["id"] in assigned_ids]

        # ---- Budget totals (converted to trip base currency – Stage 6) ----
    base_cur = trip.get("base_currency", "USD")
    total_spent, confirmed_spent, estimated_spent = _sum_items_in_base(items, base_cur)

    # =========================================================
    # BUILD THE DOCUMENT
    # =========================================================
    doc = Document()

    # ---- Cover / Trip Overview ----
    doc.add_heading(trip.get("purpose", "Travel Pack"), 0)

    doc.add_heading("Trip Overview", level=1)
    dep_parts = [
        p for p in [
            trip.get("departure_city"),
            trip.get("departure_region"),
            trip.get("departure_country"),
        ] if p
    ]
    departure_display = ", ".join(dep_parts) if dep_parts else "Not specified"
    dest_display = stops[0]["city"] if stops else "N/A"

    overview_table = doc.add_table(rows=1, cols=2)
    overview_table.style = "Table Grid"
    hdr = overview_table.rows[0].cells
    hdr[0].text = "Detail"
    hdr[1].text = "Value"

    overview_rows = [
        ("From", departure_display),
        ("To", dest_display),
        ("Dates", f"{trip['start_date'][:10]} – {trip['end_date'][:10]}"),
        ("Status", trip.get("status", "").title()),
        ("Budget", f"{trip.get('budget', 0):.2f} {trip.get('base_currency', 'USD')}"),
        ("Total Spent", f"{total_spent:.2f} {trip.get('base_currency', 'USD')}"),
        ("Confirmed", f"{confirmed_spent:.2f} {trip.get('base_currency', 'USD')}"),
        ("Pending", f"{estimated_spent:.2f} {trip.get('base_currency', 'USD')}"),
    ]
    for label, value in overview_rows:
        row = overview_table.add_row().cells
        row[0].text = label
        row[1].text = str(value)

    # ---- Delegation ----
    if delegation:
        doc.add_heading("Delegation", level=1)
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "Name"
        hdr[1].text = "Role"
        hdr[2].text = "Email"
        hdr[3].text = "Phone"
        for member in delegation:
            row = table.add_row().cells
            row[0].text = member.get("name", "")
            row[1].text = member.get("role", "")
            row[2].text = member.get("email", "")
            row[3].text = member.get("phone", "")

    # ---- Route (Stops) ----
    if stops:
        doc.add_heading("Route", level=1)
        for stop in stops:
            loc = stop["city"]
            loc_parts = []
            if stop.get("region"):
                loc_parts.append(stop["region"])
            if stop.get("country"):
                loc_parts.append(stop["country"])
            if loc_parts:
                loc += f" ({', '.join(loc_parts)})"
            doc.add_paragraph(
                f"{loc}: {stop['start_date'][:10]} to {stop['end_date'][:10]}",
                style="List Bullet",
            )

    # ---- Itinerary (Flights & Transport only) ----
    itinerary_items = [
        item for item in items if item["item_type"] in ["Flight", "Transport"]
    ]
    if itinerary_items:
        doc.add_heading("Itinerary", level=1)
        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        headers = ["Date/Time", "Type", "Description", "Location", "Confirmation", "Cost"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            hdr[i].paragraphs[0].runs[0].bold = True

        for item in itinerary_items:
            row = table.add_row().cells
            row[0].text = item["datetime_start"][:16]
            row[1].text = item["item_type"]
            row[2].text = item["description"]
            row[3].text = item.get("location", "")
            row[4].text = item.get("confirmation_code", "") or "—"
            row[5].text = (
                f"{item.get('cost', 0):.2f} {item.get('cost_currency', 'USD')}"
            )

    # ---- Hotels (separate section) ----
    hotels = [item for item in items if item["item_type"] == "Hotel"]
    if hotels:
        doc.add_heading("Hotels", level=1)
        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        headers = ["Check-in / Check-out", "Hotel", "Location", "Confirmation", "Cost"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            hdr[i].paragraphs[0].runs[0].bold = True

        for hotel in hotels:
            row = table.add_row().cells
            end_str = hotel["datetime_end"][:16] if hotel.get("datetime_end") else ""
            row[0].text = f"{hotel['datetime_start'][:16]} – {end_str}" if end_str else hotel["datetime_start"][:16]
            row[1].text = hotel["description"]
            row[2].text = hotel.get("location", "")
            row[3].text = hotel.get("confirmation_code", "") or "—"
            row[4].text = (
                f"{hotel.get('cost', 0):.2f} {hotel.get('cost_currency', 'USD')}"
            )

    # ---- Agenda (Meetings, Conferences, Dinners, Site Visits) ----
    agenda_items = [
        item
        for item in items
        if item["item_type"] in ["Meeting", "Conference", "Dinner", "Site Visit"]
    ]
    if agenda_items:
        doc.add_heading("Agenda", level=1)
        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        headers = ["Date/Time", "Type", "Description", "Location", "Attendees"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            hdr[i].paragraphs[0].runs[0].bold = True

        for item in agenda_items:
            row = table.add_row().cells
            row[0].text = item["datetime_start"][:16]
            row[1].text = item["item_type"]
            row[2].text = item["description"]
            row[3].text = item.get("location", "")
            # Venue name
            venue_name = ""
            if item.get("venue_id"):
                v = db.get_venue(item["venue_id"])
                if v:
                    venue_name = v["name"]
            row[4].text = venue_name or "—"
            attendees = db.get_item_delegation_members(item["id"])
            attendee_names = ", ".join([a["name"] for a in attendees])
            row[4].text = attendee_names or "—"

    # ---- Venues (unique venues referenced by this trip's items) ----
    if trip_venues:
        doc.add_heading("Venues", level=1)
        for v in trip_venues:
            p = doc.add_paragraph()
            p.add_run(v["name"]).bold = True
            if v.get("address"):
                p.add_run(f"\nAddress: {v['address']}")
            location_parts = [x for x in [v.get("city"), v.get("country")] if x]
            if location_parts:
                p.add_run(f"\nLocation: {', '.join(location_parts)}")
            if v.get("wifi_ssid") or v.get("wifi_password"):
                p.add_run(
                    f"\nWiFi: {v.get('wifi_ssid', '')} / {v.get('wifi_password', '')}"
                )
            if v.get("badge_info"):
                p.add_run(f"\nBadge Info: {v['badge_info']}")
            if v.get("dress_code_notes"):
                p.add_run(f"\nDress Code: {v['dress_code_notes']}")
            if v.get("notes"):
                p.add_run(f"\nNotes: {v['notes']}")
            doc.add_paragraph()  # spacing

    # ---- Local Support Contacts ----
    if contacts:
        doc.add_heading("Local Support", level=1)
        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        headers = ["Name", "Role", "Phone", "Email", "Country"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            hdr[i].paragraphs[0].runs[0].bold = True

        for c in contacts:
            row = table.add_row().cells
            row[0].text = c.get("name", "")
            row[1].text = c.get("role", "")
            row[2].text = c.get("phone", "")
            row[3].text = c.get("email", "")
            row[4].text = c.get("country", "")

    # ---- Expenses & Per Diem ----
    expense_summary = db.get_expense_summary(trip_id)
    if expense_summary:
        doc.add_heading("Expenses & Per Diem", level=1)
        total_allowance = sum(s["allowance"] for s in expense_summary)
        total_spent_expenses = sum(s["spent"] for s in expense_summary)
        doc.add_paragraph(
            f"Delegation Allowance: {total_allowance:,.2f} {trip.get('base_currency', 'USD')}  ·  "
            f"Total Spent: {total_spent_expenses:,.2f} {trip.get('base_currency', 'USD')}  ·  "
            f"Remaining: {(total_allowance - total_spent_expenses):,.2f} {trip.get('base_currency', 'USD')}"
        )
        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        headers = ["Traveler", "Daily Rate", "Days", "Allowance", "Spent", "Remaining"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            hdr[i].paragraphs[0].runs[0].bold = True
        for s in expense_summary:
            row = table.add_row().cells
            row[0].text = s["name"]
            row[1].text = f"{s['daily_rate']:.2f}"
            row[2].text = str(s["days"])
            row[3].text = f"{s['allowance']:.2f}"
            row[4].text = f"{s['spent']:.2f}"
            row[5].text = f"{s['remaining']:.2f}"

    # ---- Packing Lists (per traveler) ----
    packing_lists_docx = []
    for member in db.get_trip_delegation_members(trip_id):
        plist = db.get_packing_list(trip_id, member["id"])
        if not plist:
            continue
        items = db.get_packing_items(plist["id"])
        if not items:
            continue
        packing_lists_docx.append((member, items))

    if packing_lists_docx:
        doc.add_heading("Packing Lists", level=1)
        for member, items in packing_lists_docx:
            packed_count = sum(1 for it in items if it["packed"])
            doc.add_heading(f"{member['name']} ({packed_count} of {len(items)} packed)", level=2)
            current_cat = None
            for it in items:
                if it.get("category") and it["category"] != current_cat:
                    current_cat = it["category"]
                    p = doc.add_paragraph()
                    r = p.add_run(current_cat)
                    r.bold = True
                prefix = "☑ " if it["packed"] else "☐ "
                doc.add_paragraph(f"{prefix}{it['item_name']}", style="List Bullet")

    # ---- Trip Checklists ----
    checklists_docx = []
    for cl in db.get_trip_checklists(trip_id):
        items = db.get_checklist_items(cl["id"])
        if not items:
            continue
        checklists_docx.append((cl, items))

    if checklists_docx:
        doc.add_heading("Checklists", level=1)
        for cl, items in checklists_docx:
            done_count = sum(1 for it in items if it["is_done"])
            doc.add_heading(f"{cl['name']} ({done_count} of {len(items)} done)", level=2)
            if cl.get("description"):
                doc.add_paragraph(cl["description"])
            for it in items:
                prefix = "☑ " if it["is_done"] else "☐ "
                doc.add_paragraph(f"{prefix}{it['item_text']}", style="List Bullet")

    # ---- Emergency Info ----
    emergency_hospitals = db.get_hospitals_for_trip(trip_id)
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

    if emergency_hospitals or trip_emergency_numbers:
        doc.add_heading("Emergency Info", level=1)

        if trip_emergency_numbers:
            doc.add_heading("Emergency Services", level=2)
            for loc, nums in trip_emergency_numbers.items():
                p = doc.add_paragraph()
                p.add_run(loc).bold = True
                for label, val in [
                    ("Police", nums.get("police")),
                    ("Ambulance", nums.get("ambulance")),
                    ("Fire", nums.get("fire")),
                ]:
                    if val:
                        p.add_run(f"\n{label}: {val}")

        if emergency_hospitals:
            doc.add_heading("Nearest Hospitals", level=2)
            for h in emergency_hospitals:
                p = doc.add_paragraph()
                p.add_run(h["name"]).bold = True
                if h.get("city"):
                    loc_str = h["city"] + (
                        f", {h['country']}" if h.get("country") else ""
                    )
                    p.add_run(f" — {loc_str}")
                if h.get("address"):
                    p.add_run(f"\nAddress: {h['address']}")
                if h.get("phone"):
                    p.add_run(f"\nPhone: {h['phone']}")
                if h.get("maps_url"):
                    p.add_run(f"\nMaps: {h['maps_url']}")
                if h.get("notes"):
                    p.add_run(f"\nNotes: {h['notes']}")

    # ---- Weather per stop ----
    weather_by_stop = _build_weather_by_stop(stops)
    if weather_by_stop:
        doc.add_heading("Weather", level=1)

        for entry in weather_by_stop:
            # Sub-heading: "London, United Kingdom  (2026-10-12 → 2026-10-14)"
            header = entry["city"]
            if entry["country"]:
                header += f", {entry['country']}"
            header += f"  ({entry['start_date']} → {entry['end_date']})"
            doc.add_heading(header, level=2)

            if entry["is_past"]:
                doc.add_paragraph("This stop has already passed.")
                continue

            if not entry["available"]:
                doc.add_paragraph(
                    "Forecast not yet available for these dates — "
                    "check back closer to departure."
                )
                continue

            cur = entry.get("current")
            if cur:
                doc.add_paragraph(
                    f"Now: {cur['icon']} {cur['temp']}°C — {cur['desc']}  ·  "
                    f"Humidity: {cur['humidity']}%  ·  "
                    f"Wind: {cur['wind']} km/h"
                )

            daily = entry.get("daily") or []
            if daily:
                table = doc.add_table(rows=1, cols=5)
                table.style = "Table Grid"
                hdr = table.rows[0].cells
                headers = [
                    "Date",
                    "Condition",
                    "High / Low",
                    "Precipitation",
                    "UV Index",
                ]
                for i, h in enumerate(headers):
                    hdr[i].text = h
                    hdr[i].paragraphs[0].runs[0].bold = True
                for d in daily:
                    row = table.add_row().cells
                    row[0].text = d["date"]
                    row[1].text = f"{d['icon']} {d['desc']}"
                    row[2].text = f"{d['max']}° / {d['min']}°"
                    row[3].text = f"{d['precip']} mm"
                    row[4].text = f"{d['uv']}"
            else:
                doc.add_paragraph("No daily forecast available.")

    # ---- Receipts ----
    receipt_items = [
        item for item in items
        if item.get("receipt_path") and os.path.exists(item["receipt_path"])
    ]
    if receipt_items:
        doc.add_heading("Receipts", level=1)
        for item in receipt_items:
            try:
                doc.add_picture(item["receipt_path"], width=Inches(2))
                doc.add_paragraph(item["description"])
                doc.add_paragraph()  # spacing
            except Exception:
                pass

    # ---- Footer ----
    doc.add_page_break()
    doc.add_paragraph(
        f"Generated by Executive Travel Planner on {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream
