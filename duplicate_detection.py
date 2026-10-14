# duplicate_detection.py
import sqlite3
import os
from datetime import datetime, timedelta
import database as db
from utils import detect_conflicts


# ------------------------------------------------
# 1. TRIP DUPLICATES
# ------------------------------------------------
def find_duplicate_trips(exec_id, purpose, start_date, end_date, exclude_trip_id=None):
    """
    Returns a list of trips for the same executive that have:
    - same purpose (exact match)
    - overlapping date ranges
    """
    conn = sqlite3.connect(db.DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    query = """
        SELECT id, purpose, start_date, end_date, destination
        FROM trips
        WHERE exec_id = ?
          AND purpose = ?
          AND start_date <= ? AND end_date >= ?  -- overlap
    """
    params = [exec_id, purpose, end_date, start_date]
    if exclude_trip_id:
        query += " AND id != ?"
        params.append(exclude_trip_id)

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()
    return [dict(row) for row in rows]


# ------------------------------------------------
# 2. EXECUTIVE DUPLICATES
# ------------------------------------------------
def find_duplicate_executive(email, name, company_id, exclude_exec_id=None):
    """
    Returns a list of executives with the same email OR same name+company.
    """
    conn = sqlite3.connect(db.DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    query = """
        SELECT id, name, email, company_id
        FROM executives
        WHERE (email = ? OR (name = ? AND company_id = ?))
    """
    params = [email, name, company_id]
    if exclude_exec_id:
        query += " AND id != ?"
        params.append(exclude_exec_id)

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()
    return [dict(row) for row in rows]


# ------------------------------------------------
# 3. ITINERARY ITEM DUPLICATES (exact match)
# ------------------------------------------------
def find_duplicate_item(
    trip_id,
    description,
    datetime_start,
    datetime_end,
    cost,
    item_type,
    exclude_item_id=None,
):
    """
    Finds items in the same trip with identical description, type,
    start/end times, and cost (within a small epsilon).
    """
    conn = sqlite3.connect(db.DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    tolerance = 0.01
    query = """
        SELECT id, description, datetime_start, datetime_end, cost
        FROM itinerary_items
        WHERE trip_id = ?
          AND description = ?
          AND item_type = ?
          AND datetime_start = ?
          AND datetime_end = ?
          AND ABS(cost - ?) < ?
    """
    params = [
        trip_id,
        description,
        item_type,
        datetime_start,
        datetime_end,
        cost,
        tolerance,
    ]
    if exclude_item_id:
        query += " AND id != ?"
        params.append(exclude_item_id)

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()
    return [dict(row) for row in rows]


# ------------------------------------------------
# 4. EXPENSE DUPLICATES (fuzzy, for the scan)
# ------------------------------------------------
def find_similar_expenses(trip_id, days_window=1, amount_tolerance=5.0):
    """
    Flags items in the same trip that have similar description, same type,
    occur on the same day (or within days_window), and have amounts within
    tolerance. Useful for catching duplicate hotel stays, meals, etc.
    """
    items = db.get_items_for_trip(trip_id)
    if len(items) < 2:
        return []

    duplicates = []
    for i, item_a in enumerate(items):
        for item_b in items[i + 1 :]:
            # Same type
            if item_a["item_type"] != item_b["item_type"]:
                continue
            # Similar description (case-insensitive, exact match for simplicity)
            if (
                item_a["description"].lower().strip()
                != item_b["description"].lower().strip()
            ):
                continue
            # Dates within window
            dt_a = datetime.fromisoformat(item_a["datetime_start"])
            dt_b = datetime.fromisoformat(item_b["datetime_start"])
            if abs((dt_a - dt_b).days) > days_window:
                continue
            # Amount within tolerance
            if abs(item_a["cost"] - item_b["cost"]) > amount_tolerance:
                continue
            duplicates.append((item_a, item_b))
    return duplicates


# ------------------------------------------------
# 5. RECEIPT DUPLICATES
# ------------------------------------------------
def find_duplicate_receipts(trip_id):
    """
    Checks all items in a trip for receipt files with the same name.
    Returns a dict {filename: [item_ids]}.
    """
    items = db.get_items_for_trip(trip_id)
    filename_map = {}
    for item in items:
        path = item.get("receipt_path")
        if path and os.path.exists(path):
            fname = os.path.basename(path)
            filename_map.setdefault(fname, []).append(item["id"])
    # Keep only those with more than one item
    return {fname: ids for fname, ids in filename_map.items() if len(ids) > 1}


# ------------------------------------------------
# 6. COMPREHENSIVE SCAN (for the dashboard)
# ------------------------------------------------
def scan_trip_for_duplicates(trip_id):
    """
    Runs all checks for the current trip and returns a structured report.
    """
    trip = db.get_trip(trip_id)
    if not trip:
        return None

    report = {
        "trip_id": trip_id,
        "conflicts": [],  # scheduling overlaps
        "duplicate_items": [],  # exact matches
        "similar_expenses": [],  # fuzzy matches
        "duplicate_receipts": {},  # filename collisions
    }

    # 1. Scheduling conflicts (using existing utils)
    items = db.get_items_for_trip(trip_id)
    report["conflicts"] = detect_conflicts(items)

    # 2. Exact duplicate items
    for item in items:
        dupes = find_duplicate_item(
            trip_id,
            item["description"],
            item["datetime_start"],
            item["datetime_end"],
            item["cost"],
            item["item_type"],
            exclude_item_id=item["id"],
        )
        if dupes:
            report["duplicate_items"].append({"original": item, "duplicates": dupes})

    # 3. Similar expenses (fuzzy)
    report["similar_expenses"] = find_similar_expenses(trip_id)

    # 4. Duplicate receipts
    report["duplicate_receipts"] = find_duplicate_receipts(trip_id)

    return report
