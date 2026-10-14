import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import database as db
from datetime import datetime
import currency  # Stage 6
from doc_generator import _sum_items_in_base

def export_profile_to_excel(exec_id, currency_symbol="$"):
    """
    Export a single executive profile to an Excel file.
    One sheet with executive details and memberships.
    """
    profile = db.get_full_executive_profile(exec_id)
    if not profile:
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "Executive Profile"

    # Title
    ws["A1"] = "Executive Profile"
    ws["A1"].font = Font(size=14, bold=True)
    ws.merge_cells("A1:D1")

    # Headers
    headers = ["Field", "Value"]
    ws.append(headers)
    for cell in ws[2]:
        cell.font = Font(bold=True)

    # Data rows
    fields = [
        ("Name", profile.get("Name", "")),
        ("Email", profile.get("Email", "")),
        ("Timezone", profile.get("Timezone", "")),
        ("Seat Preference", profile.get("Seat Preference", "")),
        ("Dietary", profile.get("Dietary", "")),
        ("Company", profile.get("Company", "")),
        ("Cost Center", profile.get("Cost Center", "")),
        ("Policy Notes", profile.get("Policy Notes", "")),
        ("Preferred Airline", profile.get("Preferred Airline", "")),
        ("TSA PreCheck", profile.get("TSA PreCheck", "")),
        ("Meal Preference", profile.get("Meal Preference", "")),
    ]
    for field, value in fields:
        ws.append([field, value])

    # Memberships
    ws.append([])
    ws.append(["Memberships"])
    mems = db.get_memberships(exec_id)
    if mems:
        ws.append(
            ["Category", "Program", "Number", "Tier", "Alliance", "Airport", "Notes"]
        )
        for m in mems:
            ws.append(
                [
                    m.get("category", ""),
                    m.get("program_name", ""),
                    m.get("membership_number", ""),
                    m.get("tier", ""),
                    m.get("alliance", ""),
                    m.get("airport_code", ""),
                    m.get("notes", ""),
                ]
            )
    else:
        ws.append(["No memberships recorded."])

    # Auto-fit columns
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream


def export_itinerary_to_excel(items, trip_data, currency_symbol, base_currency="USD"):
    """
    Export itinerary items to an Excel file.
    Includes contacts column.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Itinerary"

    # Title
    ws["A1"] = f"Itinerary: {trip_data.get('purpose', 'Trip')}"
    ws["A1"].font = Font(size=14, bold=True)
    ws.merge_cells("A1:G1")

    # Headers
    headers = [
        "Date/Time",
        "Type",
        "Description",
        "Location",
        "Contacts",
        "Cost (Original)",
        "Confirmed",
    ]
    ws.append(headers)
    for cell in ws[2]:
        cell.font = Font(bold=True)

    for item in items:
        start_dt = (
            datetime.fromisoformat(item["datetime_start"])
            if item.get("datetime_start")
            else None
        )
        date_str = start_dt.strftime("%Y-%m-%d %H:%M") if start_dt else ""
        contacts = db.get_item_contacts(item["id"])
        contact_str = ", ".join([c["name"] for c in contacts]) if contacts else ""
        ws.append(
            [
                date_str,
                item.get("item_type", ""),
                item.get("description", ""),
                item.get("location", ""),
                contact_str,
                f"{item.get('cost', 0):.2f} {item.get('cost_currency', 'USD')}",
                "Yes" if item.get("is_confirmed", 0) else "No",
            ]
        )

    # Auto-fit columns
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream


def export_expense_to_excel(items, trip_data, currency_symbol, base_currency="USD"):
    """
    Export expense report to Excel.
    Includes contacts column.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Expense Report"

    # Title
    ws["A1"] = f"Expense Report: {trip_data.get('purpose', 'Trip')}"
    ws["A1"].font = Font(size=14, bold=True)
    ws.merge_cells("A1:H1")

    # Summary (converted to trip base currency – Stage 6)
    trip_base_cur = trip_data.get("base_currency", base_currency) or base_currency
    total_spent, confirmed_spent, estimated_spent = _sum_items_in_base(
        items, trip_base_cur
    )
    trip_budget = trip_data.get("budget", 0)
    

    summary_data = [
        ("Total Budget", f"{trip_budget:.2f} {base_currency}"),
        ("Total Spent", f"{total_spent:.2f} {base_currency}"),
        ("Confirmed (Booked)", f"{confirmed_spent:.2f} {base_currency}"),
        ("Pending (Quoted)", f"{estimated_spent:.2f} {base_currency}"),
        ("Remaining", f"{trip_budget - total_spent:.2f} {base_currency}"),
    ]
    ws.append([])
    for label, value in summary_data:
        ws.append([label, value])
    ws.append([])

    # Details
    headers = [
        "Date/Time",
        "Type",
        "Description",
        "Location",
        "Contacts",
        "Cost (Original)",
        "Confirmed",
    ]
    ws.append(headers)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    for item in items:
        start_dt = (
            datetime.fromisoformat(item["datetime_start"])
            if item.get("datetime_start")
            else None
        )
        date_str = start_dt.strftime("%Y-%m-%d %H:%M") if start_dt else ""
        contacts = db.get_item_contacts(item["id"])
        contact_str = ", ".join([c["name"] for c in contacts]) if contacts else ""
        ws.append(
            [
                date_str,
                item.get("item_type", ""),
                item.get("description", ""),
                item.get("location", ""),
                contact_str,
                f"{item.get('cost', 0):.2f} {item.get('cost_currency', 'USD')}",
                "Yes" if item.get("is_confirmed", 0) else "No",
            ]
        )

    # Auto-fit columns
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream


def export_spending_to_excel(summary_data, base_currency="USD"):
    """
    Export spending summary to Excel.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Spending Summary"

    ws["A1"] = "Spending Summary"
    ws["A1"].font = Font(size=14, bold=True)
    ws.merge_cells("A1:I1")

    headers = [
        "Executive",
        "Company",
        "Destination",
        "Budget (Base Currency)",
        "Total Spent (Base Currency)",
        "Confirmed (Base Currency)",
        "Pending (Base Currency)",
        "Status",
        "Currency",
    ]
    ws.append(headers)
    for cell in ws[2]:
        cell.font = Font(bold=True)

    def safe_float(value):
        return float(value) if value is not None else 0.0

    for trip in summary_data:
        ws.append(
            [
                trip.get("executive_name", ""),
                trip.get("company_name", ""),
                trip.get("destination", ""),
                f"{safe_float(trip.get('budget')):.2f}",
                f"{safe_float(trip.get('total_spent')):.2f}",
                f"{safe_float(trip.get('confirmed_spent')):.2f}",
                f"{safe_float(trip.get('estimated_spent')):.2f}",
                trip.get("status", "").title(),
                trip.get("base_currency", base_currency),
            ]
        )

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream


def export_company_profile_to_excel(company_id):
    """
    Export company profile with sheets: Company Info, Executives, Contacts, Delegation.
    Contacts include city and type.
    """
    company = db.get_company(company_id)
    if not company:
        return None
    executives = db.get_executives_by_company(company_id)
    contacts = db.get_contacts(company_id, active_only=True)
    delegation = db.get_delegation_members(company_id, active_only=True)

    wb = Workbook()

    # Sheet 1: Company Info
    ws1 = wb.active
    ws1.title = "Company Info"
    ws1["A1"] = "Field"
    ws1["B1"] = "Value"
    ws1["A1"].font = Font(bold=True)
    ws1["B1"].font = Font(bold=True)
    ws1.append(["Name", company["name"]])
    ws1.append(["Cost Center", company.get("default_cost_center", "")])
    ws1.append(["Policy Notes", company.get("policy_notes", "")])
    ws1.column_dimensions["A"].width = 20
    ws1.column_dimensions["B"].width = 40

    # Sheet 2: Executives
    ws2 = wb.create_sheet("Executives")
    headers = [
        "Name",
        "Email",
        "Timezone",
        "Seat Preference",
        "Dietary",
        "Passport",
        "Preferred Airline",
        "TSA PreCheck",
        "Meal Preference",
    ]
    for col_idx, header in enumerate(headers, 1):
        cell = ws2.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
    for row_idx, e in enumerate(executives, 2):
        ws2.cell(row=row_idx, column=1, value=e.get("name", ""))
        ws2.cell(row=row_idx, column=2, value=e.get("email", ""))
        ws2.cell(row=row_idx, column=3, value=e.get("timezone", ""))
        ws2.cell(row=row_idx, column=4, value=e.get("seat_preference", ""))
        ws2.cell(row=row_idx, column=7, value=e.get("dietary_restrictions", ""))
        ws2.cell(row=row_idx, column=8, value=e.get("passport_number", ""))
        ws2.cell(row=row_idx, column=9, value=e.get("preferred_airline", ""))
        ws2.cell(row=row_idx, column=10, value=e.get("tsa_precheck", ""))
        ws2.cell(row=row_idx, column=11, value=e.get("meal_preference", ""))
    for col in range(1, len(headers) + 1):
        ws2.column_dimensions[get_column_letter(col)].width = 20

    # Sheet 3: Contacts (with city and type)
    ws3 = wb.create_sheet("Contacts")
    headers = [
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
    for col_idx, header in enumerate(headers, 1):
        cell = ws3.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
    for row_idx, c in enumerate(contacts, 2):
        ws3.cell(row=row_idx, column=1, value=c.get("name", ""))
        ws3.cell(row=row_idx, column=2, value=c.get("role", ""))
        ws3.cell(row=row_idx, column=3, value=c.get("phone", ""))
        ws3.cell(row=row_idx, column=4, value=c.get("email", ""))
        ws3.cell(row=row_idx, column=5, value=c.get("country", ""))
        ws3.cell(row=row_idx, column=6, value=c.get("city", ""))
        ws3.cell(row=row_idx, column=7, value=c.get("type", "Local Support"))
        ws3.cell(row=row_idx, column=8, value=c.get("notes", ""))
        ws3.cell(row=row_idx, column=9, value=c.get("tags", ""))
    for col in range(1, len(headers) + 1):
        ws3.column_dimensions[get_column_letter(col)].width = 20

    # Sheet 4: Delegation
    ws4 = wb.create_sheet("Delegation")
    headers = ["Name", "Email", "Role", "Phone"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws4.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
    for row_idx, m in enumerate(delegation, 2):
        ws4.cell(row=row_idx, column=1, value=m.get("name", ""))
        ws4.cell(row=row_idx, column=2, value=m.get("email", ""))
        ws4.cell(row=row_idx, column=3, value=m.get("role", ""))
        ws4.cell(row=row_idx, column=4, value=m.get("phone", ""))
    for col in range(1, len(headers) + 1):
        ws4.column_dimensions[get_column_letter(col)].width = 20

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream
