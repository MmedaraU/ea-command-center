import sqlite3
import os
import json
from datetime import date, datetime, timedelta

DB_PATH = "ea_command_center.db"

def _table_exists(cur, table_name):
    """True if `table_name` exists in the current SQLite database."""
    cur.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    )
    return cur.fetchone() is not None

# =========================================================
# MIGRATION
# =========================================================


def migrate_db():
    """Add new columns/tables if they don't exist. Handles schema upgrades.

    Every ALTER TABLE block is gated on the target table actually existing,
    so this function is safe to run against a fresh/empty database, a
    partially-populated one, or a fully-migrated one. The CREATE TABLE IF
    NOT EXISTS statements below then ensure every table exists for the
    next run.
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()

        # --- trips table ---
        if _table_exists(c, "trips"):
            c.execute("PRAGMA table_info(trips)")
            existing_trips = [row[1] for row in c.fetchall()]
            for col, ddl in [
                ("budget", "REAL DEFAULT 0"),
                ("departure_city", "TEXT"),
                ("departure_region", "TEXT"),
                ("departure_country", "TEXT"),
                ("base_currency", "TEXT DEFAULT 'USD'"),
                ("display_currency", "TEXT DEFAULT 'USD'"),
                ("trip_contacts", "TEXT"),
                ("receipt_upload_folder_url", "TEXT"),
            ]:
                if col not in existing_trips:
                    c.execute(f"ALTER TABLE trips ADD COLUMN {col} {ddl}")

        # --- itinerary_items table ---
        if _table_exists(c, "itinerary_items"):
            c.execute("PRAGMA table_info(itinerary_items)")
            existing_items = [row[1] for row in c.fetchall()]
            for col, ddl in [
                ("is_confirmed", "INTEGER DEFAULT 0"),
                ("receipt_path", "TEXT"),
                ("cost_currency", "TEXT DEFAULT 'USD'"),
                ("timezone", "TEXT"),
                ("venue_id", "INTEGER"),
                ("cost_date", "TEXT"),
            ]:
                if col not in existing_items:
                    c.execute(
                        f"ALTER TABLE itinerary_items ADD COLUMN {col} {ddl}"
                    )
            # Backfill cost_date
            c.execute(
                """UPDATE itinerary_items
                   SET cost_date = SUBSTR(datetime_start, 1, 10)
                   WHERE cost_date IS NULL AND datetime_start IS NOT NULL"""
            )

        # --- executives table ---
        if _table_exists(c, "executives"):
            c.execute("PRAGMA table_info(executives)")
            existing_execs = [row[1] for row in c.fetchall()]
            for col, ddl in [
                ("passport_number", "TEXT"),
                ("preferred_airline", "TEXT"),
                ("tsa_precheck", "TEXT"),
                ("meal_preference", "TEXT"),
                ("is_active", "INTEGER DEFAULT 1"),
            ]:
                if col not in existing_execs:
                    c.execute(
                        f"ALTER TABLE executives ADD COLUMN {col} {ddl}"
                    )

        # --- companies table ---
        if _table_exists(c, "companies"):
            c.execute("PRAGMA table_info(companies)")
            existing_company_cols = [row[1] for row in c.fetchall()]
            if "default_contact_ids" not in existing_company_cols:
                c.execute(
                    "ALTER TABLE companies ADD COLUMN default_contact_ids TEXT"
                )
            if "is_active" not in existing_company_cols:
                c.execute(
                    "ALTER TABLE companies ADD COLUMN is_active INTEGER DEFAULT 1"
                )

        # --- executive_memberships ---
        if _table_exists(c, "executive_memberships"):
            c.execute("PRAGMA table_info(executive_memberships)")
            existing_membership_cols = [row[1] for row in c.fetchall()]
            for col, ddl in [
                ("tier", "TEXT"),
                ("alliance", "TEXT"),
                ("airport_code", "TEXT"),
                ("notes", "TEXT"),
            ]:
                if col not in existing_membership_cols:
                    c.execute(
                        f"ALTER TABLE executive_memberships ADD COLUMN {col} {ddl}"
                    )

        # --- contacts table ---
        if _table_exists(c, "contacts"):
            c.execute("PRAGMA table_info(contacts)")
            existing_contact_cols = [row[1] for row in c.fetchall()]
            for col, ddl in [
                ("type", "TEXT DEFAULT 'Local Support'"),
                ("city", "TEXT"),
                ("last_contacted_at", "TEXT"),
                ("last_channel", "TEXT"),
                ("next_action", "TEXT"),
                ("next_action_date", "TEXT"),
            ]:
                if col not in existing_contact_cols:
                    c.execute(f"ALTER TABLE contacts ADD COLUMN {col} {ddl}")

        # --- Support tables that may or may not exist yet ---
        c.execute("""CREATE TABLE IF NOT EXISTS executive_passports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exec_id INTEGER NOT NULL,
            country TEXT NOT NULL,
            passport_number TEXT NOT NULL,
            expiry_date TEXT,
            issued_date TEXT,
            notes TEXT,
            FOREIGN KEY (exec_id) REFERENCES executives(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS executive_memberships (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exec_id INTEGER NOT NULL,
            category TEXT NOT NULL,
            program_name TEXT NOT NULL,
            membership_number TEXT NOT NULL,
            tier TEXT,
            alliance TEXT,
            airport_code TEXT,
            notes TEXT,
            FOREIGN KEY (exec_id) REFERENCES executives(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS trip_stops (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id INTEGER NOT NULL,
            stop_order INTEGER NOT NULL,
            city TEXT NOT NULL,
            country TEXT,
            region TEXT,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            notes TEXT,
            FOREIGN KEY (trip_id) REFERENCES trips(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            is_active INTEGER DEFAULT 1
        )""")
        c.execute("SELECT COUNT(*) FROM categories")
        if c.fetchone()[0] == 0:
            for cat in [
                "Flight", "Hotel", "Meeting", "Transport",
                "Dinner", "Car Rental", "Conference", "Site Visit",
                "Transfer", "Tour", "Wellness", "Activity",
            ]:
                c.execute("INSERT INTO categories (name) VALUES (?)", (cat,))

        c.execute("""CREATE TABLE IF NOT EXISTS trip_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            departure_city TEXT,
            departure_region TEXT,
            departure_country TEXT,
            display_currency TEXT DEFAULT 'USD',
            base_currency TEXT DEFAULT 'USD',
            stops_json TEXT,
            items_json TEXT,
            is_active INTEGER DEFAULT 1
        )""")

        # --- Trip delegation (who is on this trip) ---
        c.execute("""CREATE TABLE IF NOT EXISTS trip_delegation (
            trip_id INTEGER NOT NULL,
            contact_id INTEGER NOT NULL,
            PRIMARY KEY (trip_id, contact_id),
            FOREIGN KEY (trip_id) REFERENCES trips(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE
        )""")

        # --- Item delegation (who attends this item) ---
        c.execute("""CREATE TABLE IF NOT EXISTS item_delegation (
            item_id INTEGER NOT NULL,
            contact_id INTEGER NOT NULL,
            PRIMARY KEY (item_id, contact_id),
            FOREIGN KEY (item_id) REFERENCES itinerary_items(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS item_contacts (
            item_id INTEGER NOT NULL,
            contact_id INTEGER NOT NULL,
            PRIMARY KEY (item_id, contact_id),
            FOREIGN KEY (item_id) REFERENCES itinerary_items(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE
        )""")

        # --- Venues ---
        c.execute("""CREATE TABLE IF NOT EXISTS venues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            address TEXT,
            city TEXT,
            country TEXT,
            wifi_ssid TEXT,
            wifi_password TEXT,
            badge_info TEXT,
            dress_code_notes TEXT,
            notes TEXT,
            is_active INTEGER DEFAULT 1
        )""")

        # --- Exchange rates ---
        c.execute("""CREATE TABLE IF NOT EXISTS exchange_rates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rate_date TEXT NOT NULL,
            base_currency TEXT NOT NULL,
            target_currency TEXT NOT NULL,
            rate REAL NOT NULL,
            source TEXT,
            fetched_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(rate_date, base_currency, target_currency)
        )""")

        # --- Destination Guides ---
        c.execute("""CREATE TABLE IF NOT EXISTS destination_guides (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            country TEXT NOT NULL,
            language TEXT,
            currency TEXT,
            emergency_police TEXT,
            emergency_ambulance TEXT,
            emergency_fire TEXT,
            etiquette_notes TEXT,
            phrases TEXT,
            packing_tips TEXT,
            connectivity_notes TEXT,
            recommended_apps TEXT,
            notes TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # --- Visa Rules ---
        c.execute("""CREATE TABLE IF NOT EXISTS visa_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            passport_nationality TEXT NOT NULL,
            destination_country TEXT NOT NULL,
            visa_required INTEGER DEFAULT 1,
            max_stay_days INTEGER,
            processing_time TEXT,
            fee TEXT,
            notes TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # --- Checklist Templates ---
        c.execute("""CREATE TABLE IF NOT EXISTS checklist_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT,
            items_json TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # --- Packing Templates ---
        c.execute("""CREATE TABLE IF NOT EXISTS packing_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT,
            items_json TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # --- Indexes ---
        for idx in [
            "CREATE INDEX IF NOT EXISTS idx_contacts_company ON contacts(company_id)",
            "CREATE INDEX IF NOT EXISTS idx_contacts_country ON contacts(country)",
            "CREATE INDEX IF NOT EXISTS idx_contacts_city ON contacts(city)",
            "CREATE INDEX IF NOT EXISTS idx_contacts_active ON contacts(is_active)",
            "CREATE INDEX IF NOT EXISTS idx_trip_delegation_trip ON trip_delegation(trip_id)",
            "CREATE INDEX IF NOT EXISTS idx_trip_delegation_contact ON trip_delegation(contact_id)",
            "CREATE INDEX IF NOT EXISTS idx_item_delegation_item ON item_delegation(item_id)",
            "CREATE INDEX IF NOT EXISTS idx_item_delegation_contact ON item_delegation(contact_id)",
            "CREATE INDEX IF NOT EXISTS idx_item_contacts_item ON item_contacts(item_id)",
            "CREATE INDEX IF NOT EXISTS idx_trip_contacts ON trips(trip_contacts)",
            "CREATE INDEX IF NOT EXISTS idx_venues_name ON venues(name)",
            "CREATE INDEX IF NOT EXISTS idx_venues_country ON venues(country)",
            "CREATE INDEX IF NOT EXISTS idx_venues_active ON venues(is_active)",
            "CREATE INDEX IF NOT EXISTS idx_items_venue ON itinerary_items(venue_id)",
            "CREATE INDEX IF NOT EXISTS idx_rates_pair_date ON exchange_rates(base_currency, target_currency, rate_date)",
            "CREATE INDEX IF NOT EXISTS idx_dest_country ON destination_guides(country)",
            "CREATE INDEX IF NOT EXISTS idx_visa_pair ON visa_rules(passport_nationality, destination_country)",
            "CREATE INDEX IF NOT EXISTS idx_hospitals_city ON hospitals(city)",
            "CREATE INDEX IF NOT EXISTS idx_hospitals_country ON hospitals(country)",
            "CREATE INDEX IF NOT EXISTS idx_hospitals_active ON hospitals(is_active)",
            "CREATE INDEX IF NOT EXISTS idx_categories_active ON categories(is_active)",
            "CREATE INDEX IF NOT EXISTS idx_executives_active ON executives(is_active)",
            "CREATE INDEX IF NOT EXISTS idx_companies_active ON companies(is_active)",
        ]:
            try:
                c.execute(idx)
            except sqlite3.OperationalError:
                # Table may not exist yet — skip index creation for it
                pass

        # --- Contacts: allow NULL company_id ---
        if _table_exists(c, "contacts"):
            c.execute("PRAGMA table_info(contacts)")
            columns = c.fetchall()
            for col in columns:
                if col[1] == "company_id" and col[3] == 1:
                    conn.commit()
                    c.execute("PRAGMA foreign_keys=OFF")
                    c.execute("CREATE TABLE contacts_new AS SELECT * FROM contacts")
                    c.execute("DROP TABLE contacts")
                    c.execute("""
                        CREATE TABLE contacts (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            company_id INTEGER,
                            name TEXT NOT NULL,
                            role TEXT,
                            phone TEXT,
                            email TEXT,
                            country TEXT,
                            city TEXT,
                            notes TEXT,
                            is_active INTEGER DEFAULT 1,
                            tags TEXT,
                            type TEXT DEFAULT 'Local Support',
                            FOREIGN KEY (company_id) REFERENCES companies(id)
                                ON DELETE SET NULL
                        )
                    """)
                    c.execute("INSERT INTO contacts SELECT * FROM contacts_new")
                    c.execute("DROP TABLE contacts_new")
                    c.execute("PRAGMA foreign_keys=ON")
                    conn.commit()
                    break

        conn.commit()
    finally:
        conn.close()


def init_db():
    """Create all tables and run migrations."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("PRAGMA busy_timeout=30000")

        c.execute("""CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            default_cost_center TEXT,
            policy_notes TEXT,
            default_contact_ids TEXT,
            is_active INTEGER DEFAULT 1
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS executives (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            name TEXT NOT NULL,
            email TEXT,
            timezone TEXT DEFAULT 'America/New_York',
            seat_preference TEXT,
            hotel_loyalty TEXT,
            frequent_flyer_number TEXT,
            dietary_restrictions TEXT,
            passport_number TEXT,
            preferred_airline TEXT,
            tsa_precheck TEXT,
            meal_preference TEXT,
            is_active INTEGER DEFAULT 1,
            FOREIGN KEY (company_id) REFERENCES companies(id)
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS trips (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exec_id INTEGER,
            destination TEXT NOT NULL,
            start_date TEXT,
            end_date TEXT,
            purpose TEXT,
            status TEXT DEFAULT 'draft',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            budget REAL DEFAULT 0,
            departure_city TEXT,
            departure_region TEXT,
            departure_country TEXT,
            base_currency TEXT DEFAULT 'USD',
            display_currency TEXT DEFAULT 'USD',
            trip_contacts TEXT,
            receipt_upload_folder_url TEXT,
            FOREIGN KEY (exec_id) REFERENCES executives(id)
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS itinerary_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id INTEGER,
            item_type TEXT,
            datetime_start TEXT,
            datetime_end TEXT,
            description TEXT,
            location TEXT,
            cost REAL,
            confirmation_code TEXT,
            notes TEXT,
            is_confirmed INTEGER DEFAULT 0,
            receipt_path TEXT,
            cost_currency TEXT DEFAULT 'USD',
            timezone TEXT,
            venue_id INTEGER,
            cost_date TEXT,
            FOREIGN KEY (trip_id) REFERENCES trips(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            name TEXT NOT NULL,
            role TEXT,
            phone TEXT,
            email TEXT,
            country TEXT,
            city TEXT,
            notes TEXT,
            is_active INTEGER DEFAULT 1,
            tags TEXT,
            type TEXT DEFAULT 'Local Support',
            FOREIGN KEY (company_id) REFERENCES companies(id) ON DELETE SET NULL
        )""")

        conn.commit()
    finally:
        conn.close()

    migrate_db()

    # Phase 3–5 tables need a fresh connection
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("PRAGMA busy_timeout=30000")

        c.execute("""CREATE TABLE IF NOT EXISTS per_diem (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id INTEGER NOT NULL,
            contact_id INTEGER NOT NULL,
            daily_rate REAL DEFAULT 0,
            days INTEGER DEFAULT 0,
            currency TEXT DEFAULT 'USD',
            notes TEXT,
            UNIQUE(trip_id, contact_id),
            FOREIGN KEY (trip_id) REFERENCES trips(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id INTEGER NOT NULL,
            contact_id INTEGER,
            expense_date TEXT NOT NULL,
            category TEXT,
            description TEXT,
            amount REAL DEFAULT 0,
            currency TEXT DEFAULT 'USD',
            receipt_path TEXT,
            notes TEXT,
            is_reimbursable INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (trip_id) REFERENCES trips(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS packing_lists (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id INTEGER NOT NULL,
            contact_id INTEGER NOT NULL,
            template_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(trip_id, contact_id),
            FOREIGN KEY (trip_id) REFERENCES trips(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE,
            FOREIGN KEY (template_id) REFERENCES packing_templates(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS packing_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            list_id INTEGER NOT NULL,
            category TEXT,
            item_name TEXT NOT NULL,
            packed INTEGER DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (list_id) REFERENCES packing_lists(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS trip_checklists (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            description TEXT,
            template_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (trip_id) REFERENCES trips(id) ON DELETE CASCADE,
            FOREIGN KEY (template_id) REFERENCES checklist_templates(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS trip_checklist_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            checklist_id INTEGER NOT NULL,
            item_text TEXT NOT NULL,
            is_done INTEGER DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (checklist_id) REFERENCES trip_checklists(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS hospitals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            city TEXT NOT NULL,
            country TEXT,
            name TEXT NOT NULL,
            address TEXT,
            phone TEXT,
            maps_url TEXT,
            notes TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        # --- Embassies / Consulates ---
        c.execute("""CREATE TABLE IF NOT EXISTS embassies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            host_country TEXT NOT NULL,
            city TEXT,
            representing_country TEXT NOT NULL,
            name TEXT,
            address TEXT,
            phone TEXT,
            email TEXT,
            website TEXT,
            maps_url TEXT,
            hours TEXT,
            notes TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")
        c.execute(
            "CREATE INDEX IF NOT EXISTS idx_embassies_host "
            "ON embassies(host_country)"
        )
        c.execute(
            "CREATE INDEX IF NOT EXISTS idx_embassies_rep "
            "ON embassies(representing_country)"
        )

        for idx in [
            "CREATE INDEX IF NOT EXISTS idx_per_diem_trip ON per_diem(trip_id)",
            "CREATE INDEX IF NOT EXISTS idx_expenses_trip ON expenses(trip_id)",
            "CREATE INDEX IF NOT EXISTS idx_expenses_contact ON expenses(contact_id)",
            "CREATE INDEX IF NOT EXISTS idx_packing_lists_trip ON packing_lists(trip_id)",
            "CREATE INDEX IF NOT EXISTS idx_packing_items_list ON packing_items(list_id)",
            "CREATE INDEX IF NOT EXISTS idx_trip_checklists_trip ON trip_checklists(trip_id)",
            "CREATE INDEX IF NOT EXISTS idx_trip_checklist_items ON trip_checklist_items(checklist_id)",
        ]:
            try:
                c.execute(idx)
            except sqlite3.OperationalError:
                # Column doesn't exist yet — the migration script will fix it.
                pass

        # =========================================================
        # EA COMMAND CENTER — Phase 1 new tables
        # =========================================================

        # --- Meetings ---
        c.execute("""CREATE TABLE IF NOT EXISTS meetings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exec_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            meeting_type TEXT,
            status TEXT DEFAULT 'draft',
            meeting_date TEXT,
            meeting_time TEXT,
            duration_minutes INTEGER,
            location TEXT,
            video_link TEXT,
            calendar_event_url TEXT,
            agenda TEXT,
            notes_native TEXT,
            notes_url TEXT,
            notes_label TEXT,
            follow_up_date TEXT,
            linked_trip_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (exec_id) REFERENCES executives(id) ON DELETE CASCADE,
            FOREIGN KEY (linked_trip_id) REFERENCES trips(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS meeting_attendees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meeting_id INTEGER NOT NULL,
            contact_id INTEGER NOT NULL,
            is_internal INTEGER DEFAULT 0,
            is_vip INTEGER DEFAULT 0,
            role TEXT,
            assistant_contact_id INTEGER,
            status TEXT DEFAULT 'pending',
            response_notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (meeting_id) REFERENCES meetings(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE,
            FOREIGN KEY (assistant_contact_id) REFERENCES contacts(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS meeting_action_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meeting_id INTEGER NOT NULL,
            description TEXT NOT NULL,
            owner_contact_id INTEGER,
            owner_exec_id INTEGER,
            due_date TEXT,
            status TEXT DEFAULT 'not_started',
            pushed_to_ticktick INTEGER DEFAULT 0,
            ticktick_id TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (meeting_id) REFERENCES meetings(id) ON DELETE CASCADE,
            FOREIGN KEY (owner_contact_id) REFERENCES contacts(id) ON DELETE SET NULL,
            FOREIGN KEY (owner_exec_id) REFERENCES executives(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS meeting_decisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meeting_id INTEGER NOT NULL,
            decision TEXT NOT NULL,
            made_by TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (meeting_id) REFERENCES meetings(id) ON DELETE CASCADE
        )""")

        # --- Events ---
        c.execute("""CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exec_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            objective TEXT,
            event_type TEXT,
            status TEXT DEFAULT 'planning',
            event_date TEXT,
            start_time TEXT,
            end_time TEXT,
            venue_id INTEGER,
            location_override TEXT,
            headcount INTEGER,
            budget_total REAL DEFAULT 0,
            currency TEXT DEFAULT 'USD',
            calendar_event_url TEXT,
            run_of_show_url TEXT,
            logistics_url TEXT,
            notes_native TEXT,
            debrief_notes TEXT,
            rsvp_form_url TEXT,
            linked_trip_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (exec_id) REFERENCES executives(id) ON DELETE CASCADE,
            FOREIGN KEY (venue_id) REFERENCES venues(id) ON DELETE SET NULL,
            FOREIGN KEY (linked_trip_id) REFERENCES trips(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS event_guests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER NOT NULL,
            contact_id INTEGER,
            guest_name TEXT,
            guest_email TEXT,
            invite_status TEXT DEFAULT 'not_invited',
            plus_ones INTEGER DEFAULT 0,
            plus_one_names TEXT,
            dietary TEXT,
            allergy_details TEXT,
            accessibility TEXT,
            badge_name TEXT,
            seating_preference TEXT,
            rsvp_timestamp TEXT,
            rsvp_source TEXT,
            checked_in_at TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (event_id) REFERENCES events(id) ON DELETE CASCADE,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS event_budget_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER NOT NULL,
            category TEXT,
            description TEXT,
            estimated_amount REAL DEFAULT 0,
            actual_amount REAL DEFAULT 0,
            paid_amount REAL DEFAULT 0,
            vendor_engagement_id INTEGER,
            status TEXT DEFAULT 'estimated',
            receipt_urls TEXT,
            invoice_urls TEXT,
            notes TEXT,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (event_id) REFERENCES events(id) ON DELETE CASCADE
        )""")

        # --- Vendors ---
        c.execute("""CREATE TABLE IF NOT EXISTS vendors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT,
            contact_id INTEGER,
            email TEXT,
            phone TEXT,
            website TEXT,
            address TEXT,
            city TEXT,
            country TEXT,
            notes TEXT,
            preferred INTEGER DEFAULT 0,
            rating INTEGER,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            last_used_at TEXT,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS vendor_engagements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_id INTEGER NOT NULL,
            vendor_id INTEGER NOT NULL,
            category TEXT,
            status TEXT DEFAULT 'shortlisted',
            quote_amount REAL,
            agreed_amount REAL,
            currency TEXT DEFAULT 'USD',
            contract_status TEXT DEFAULT 'not_sent',
            contract_url TEXT,
            deposit_amount REAL,
            deposit_due_date TEXT,
            deposit_paid_date TEXT,
            balance_amount REAL,
            balance_due_date TEXT,
            balance_paid_date TEXT,
            payment_terms TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (event_id) REFERENCES events(id) ON DELETE CASCADE,
            FOREIGN KEY (vendor_id) REFERENCES vendors(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS vendor_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vendor_engagement_id INTEGER NOT NULL,
            amount REAL NOT NULL,
            currency TEXT DEFAULT 'USD',
            due_date TEXT,
            paid_date TEXT,
            method TEXT,
            reference TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (vendor_engagement_id) REFERENCES vendor_engagements(id) ON DELETE CASCADE
        )""")

        # --- Speaking Engagements ---
        c.execute("""CREATE TABLE IF NOT EXISTS speaking_engagements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exec_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            organizer_contact_id INTEGER,
            organizer_name TEXT,
            engagement_type TEXT,
            format TEXT,
            status TEXT DEFAULT 'invited',
            engagement_date TEXT,
            engagement_time TEXT,
            duration TEXT,
            location TEXT,
            timezone TEXT,
            audience TEXT,
            topic TEXT,
            fee_offered REAL,
            fee_agreed REAL,
            currency TEXT DEFAULT 'USD',
            travel_covered INTEGER DEFAULT 0,
            accommodation_covered INTEGER DEFAULT 0,
            recording_rights TEXT,
            exclusivity TEXT,
            calendar_event_url TEXT,
            notes_url TEXT,
            notes_native TEXT,
            linked_trip_id INTEGER,
            promoted INTEGER DEFAULT 0,
            thank_you_sent INTEGER DEFAULT 0,
            thank_you_date TEXT,
            feedback_received TEXT,
            feedback_rating INTEGER,
            testimonial TEXT,
            testimonial_permission INTEGER DEFAULT 0,
            recording_received INTEGER DEFAULT 0,
            recording_url TEXT,
            would_repeat INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (exec_id) REFERENCES executives(id) ON DELETE CASCADE,
            FOREIGN KEY (organizer_contact_id) REFERENCES contacts(id) ON DELETE SET NULL,
            FOREIGN KEY (linked_trip_id) REFERENCES trips(id) ON DELETE SET NULL
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS speaking_deliverables (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            engagement_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            owner TEXT,
            due_date TEXT,
            status TEXT DEFAULT 'not_started',
            file_url TEXT,
            submitted_to_organizer INTEGER DEFAULT 0,
            organizer_feedback TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (engagement_id) REFERENCES speaking_engagements(id) ON DELETE CASCADE
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS speaking_cfps (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            engagement_id INTEGER NOT NULL,
            cfp_url TEXT,
            submission_deadline TEXT,
            submission_date TEXT,
            proposal_title TEXT,
            abstract TEXT,
            proposal_doc_url TEXT,
            status TEXT DEFAULT 'draft',
            decision_date TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (engagement_id) REFERENCES speaking_engagements(id) ON DELETE CASCADE
        )""")

        # --- Communications ---
        c.execute("""CREATE TABLE IF NOT EXISTS communications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            contact_id INTEGER NOT NULL,
            exec_id INTEGER,
            comm_date TEXT NOT NULL,
            channel TEXT NOT NULL,
            direction TEXT DEFAULT 'outbound',
            summary TEXT,
            next_action TEXT,
            next_action_date TEXT,
            linked_to_type TEXT,
            linked_to_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (contact_id) REFERENCES contacts(id) ON DELETE CASCADE,
            FOREIGN KEY (exec_id) REFERENCES executives(id) ON DELETE SET NULL
        )""")

        # --- Budget templates ---
        c.execute("""CREATE TABLE IF NOT EXISTS budget_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT,
            event_type TEXT,
            is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS budget_template_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            template_id INTEGER NOT NULL,
            category TEXT,
            description TEXT,
            default_amount REAL DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (template_id) REFERENCES budget_templates(id) ON DELETE CASCADE
        )""")

        # --- Restore log ---
        c.execute("""CREATE TABLE IF NOT EXISTS restore_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            performed_at TEXT DEFAULT CURRENT_TIMESTAMP,
            mode TEXT,
            source_filename TEXT,
            notes TEXT,
            success INTEGER DEFAULT 1,
            error_message TEXT
        )""")

        # --- Indexes for new tables ---
        for idx in [
            "CREATE INDEX IF NOT EXISTS idx_meetings_exec ON meetings(exec_id)",
            "CREATE INDEX IF NOT EXISTS idx_meetings_status ON meetings(status)",
            "CREATE INDEX IF NOT EXISTS idx_meetings_date ON meetings(meeting_date)",
            "CREATE INDEX IF NOT EXISTS idx_meeting_attendees_meeting ON meeting_attendees(meeting_id)",
            "CREATE INDEX IF NOT EXISTS idx_meeting_attendees_contact ON meeting_attendees(contact_id)",
            "CREATE INDEX IF NOT EXISTS idx_meeting_actions_meeting ON meeting_action_items(meeting_id)",
            "CREATE INDEX IF NOT EXISTS idx_meeting_actions_due ON meeting_action_items(due_date)",
            "CREATE INDEX IF NOT EXISTS idx_meeting_decisions_meeting ON meeting_decisions(meeting_id)",
            "CREATE INDEX IF NOT EXISTS idx_events_exec ON events(exec_id)",
            "CREATE INDEX IF NOT EXISTS idx_events_status ON events(status)",
            "CREATE INDEX IF NOT EXISTS idx_events_date ON events(event_date)",
            "CREATE INDEX IF NOT EXISTS idx_event_guests_event ON event_guests(event_id)",
            "CREATE INDEX IF NOT EXISTS idx_event_guests_contact ON event_guests(contact_id)",
            "CREATE INDEX IF NOT EXISTS idx_event_guests_status ON event_guests(invite_status)",
            "CREATE INDEX IF NOT EXISTS idx_event_budget_event ON event_budget_items(event_id)",
            "CREATE INDEX IF NOT EXISTS idx_vendors_name ON vendors(name)",
            "CREATE INDEX IF NOT EXISTS idx_vendors_category ON vendors(category)",
            "CREATE INDEX IF NOT EXISTS idx_vendors_active ON vendors(is_active)",
            "CREATE INDEX IF NOT EXISTS idx_vendor_engagements_event ON vendor_engagements(event_id)",
            "CREATE INDEX IF NOT EXISTS idx_vendor_engagements_vendor ON vendor_engagements(vendor_id)",
            "CREATE INDEX IF NOT EXISTS idx_vendor_engagements_status ON vendor_engagements(status)",
            "CREATE INDEX IF NOT EXISTS idx_vendor_payments_engagement ON vendor_payments(vendor_engagement_id)",
            "CREATE INDEX IF NOT EXISTS idx_vendor_payments_due ON vendor_payments(due_date)",
            "CREATE INDEX IF NOT EXISTS idx_speaking_exec ON speaking_engagements(exec_id)",
            "CREATE INDEX IF NOT EXISTS idx_speaking_status ON speaking_engagements(status)",
            "CREATE INDEX IF NOT EXISTS idx_speaking_date ON speaking_engagements(engagement_date)",
            "CREATE INDEX IF NOT EXISTS idx_speaking_deliv_engagement ON speaking_deliverables(engagement_id)",
            "CREATE INDEX IF NOT EXISTS idx_speaking_deliv_due ON speaking_deliverables(due_date)",
            "CREATE INDEX IF NOT EXISTS idx_speaking_cfps_engagement ON speaking_cfps(engagement_id)",
            "CREATE INDEX IF NOT EXISTS idx_speaking_cfps_deadline ON speaking_cfps(submission_deadline)",
            "CREATE INDEX IF NOT EXISTS idx_comms_contact ON communications(contact_id)",
            "CREATE INDEX IF NOT EXISTS idx_comms_date ON communications(comm_date)",
            "CREATE INDEX IF NOT EXISTS idx_comms_next_action ON communications(next_action_date)",
            "CREATE INDEX IF NOT EXISTS idx_budget_templates_active ON budget_templates(is_active)",
            "CREATE INDEX IF NOT EXISTS idx_budget_template_items_template ON budget_template_items(template_id)",
            "CREATE INDEX IF NOT EXISTS idx_restore_log_date ON restore_log(performed_at)",
        ]:
            try:
                c.execute(idx)
            except sqlite3.OperationalError:
                pass

        conn.commit()
    finally:
        conn.close()


# =========================================================
# COMPANIES
# =========================================================


def add_company(name, default_cost_center=None, policy_notes=None, is_active=1):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO companies
               (name, default_cost_center, policy_notes, is_active)
               VALUES (?, ?, ?, ?)""",
            (name, default_cost_center, policy_notes, 1 if is_active else 0),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_all_companies(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT id, name, is_active FROM companies"
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY name"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_company(company_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM companies WHERE id = ?", (company_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_company(
    company_id,
    name,
    default_cost_center=None,
    policy_notes=None,
    default_contact_ids=None,
    is_active=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields = [
            "name = ?",
            "default_cost_center = ?",
            "policy_notes = ?",
            "default_contact_ids = ?",
        ]
        params = [name, default_cost_center, policy_notes, default_contact_ids]
        if is_active is not None:
            fields.append("is_active = ?")
            params.append(1 if is_active else 0)
        params.append(company_id)
        c.execute(
            f"UPDATE companies SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def set_company_active(company_id, is_active):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE companies SET is_active = ? WHERE id = ?",
            (1 if is_active else 0, company_id),
        )
        conn.commit()
    finally:
        conn.close()


def delete_company(company_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "SELECT COUNT(*) FROM executives WHERE company_id = ?",
            (company_id,),
        )
        count = c.fetchone()[0]
        if count > 0:
            return (
                False,
                f"Cannot delete: {count} executive(s) are still assigned "
                "to this company.",
            )
        c.execute("DELETE FROM companies WHERE id = ?", (company_id,))
        conn.commit()
        return True, "Company deleted successfully."
    finally:
        conn.close()


def get_executives_by_company(company_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM executives WHERE company_id = ?", (company_id,)
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def _find_or_create_company(name):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("SELECT id FROM companies WHERE name = ?", (name,))
        row = c.fetchone()
        if row:
            return row[0]
        c.execute("INSERT INTO companies (name) VALUES (?)", (name,))
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


# =========================================================
# CONTACTS
# =========================================================


def add_contact(
    company_id,
    name,
    role=None,
    phone=None,
    email=None,
    country=None,
    city=None,
    notes=None,
    tags=None,
    type="Local Support",
    is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO contacts
               (company_id, name, role, phone, email, country, city,
                notes, tags, type, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                company_id, name, role, phone, email, country, city,
                notes, tags, type, 1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_contacts(company_id=None, active_only=True, country=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions = []
        params = []
        if active_only:
            conditions.append("is_active = 1")
        if company_id is not None:
            conditions.append("company_id = ?")
            params.append(company_id)
        if country is not None:
            conditions.append("country = ?")
            params.append(country)
        q = "SELECT * FROM contacts"
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY company_id, name"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_contact(contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM contacts WHERE id = ?", (contact_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_contact(
    contact_id,
    name=None,
    role=None,
    phone=None,
    email=None,
    country=None,
    city=None,
    notes=None,
    tags=None,
    is_active=None,
    type=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for col, val in [
            ("name", name),
            ("role", role),
            ("phone", phone),
            ("email", email),
            ("country", country),
            ("city", city),
            ("notes", notes),
            ("tags", tags),
            ("type", type),
        ]:
            if val is not None:
                fields.append(f"{col} = ?")
                params.append(val)
        if is_active is not None:
            fields.append("is_active = ?")
            params.append(1 if is_active else 0)
        if not fields:
            return
        params.append(contact_id)
        c.execute(
            f"UPDATE contacts SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def set_contact_active(contact_id, is_active):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE contacts SET is_active = ? WHERE id = ?",
            (1 if is_active else 0, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def delete_contact(contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM contacts WHERE id = ?", (contact_id,))
        conn.commit()
    finally:
        conn.close()


def find_duplicate_contacts(company_id, name=None, email=None, phone=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions = []
        params = []
        if company_id is not None:
            conditions.append("company_id = ?")
            params.append(company_id)
        else:
            conditions.append("company_id IS NULL")
        sub = []
        if name:
            sub.append("name = ?")
            params.append(name)
        if email:
            sub.append("email = ?")
            params.append(email)
        if phone:
            sub.append("phone = ?")
            params.append(phone)
        if not sub:
            return []
        q = "SELECT * FROM contacts WHERE is_active = 1 AND "
        if conditions:
            q += conditions[0] + " AND "
        q += "(" + " OR ".join(sub) + ")"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()

def get_contacts_for_delegation(company_id, active_only=True):
    """
    Contacts available for trip delegation:
    - all active contacts of the given company, OR
    - active contacts with no company (universal / freelance)
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if active_only:
            conditions.append("is_active = 1")
        if company_id:
            conditions.append("(company_id = ? OR company_id IS NULL)")
            params.append(company_id)
        else:
            conditions.append("company_id IS NULL")
        q = "SELECT * FROM contacts"
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY (company_id IS NULL), name"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()

def get_delegation_members(company_id, active_only=True):
    """
    Alias for get_contacts_for_delegation().

    Kept because doc_generator.py calls this name in four places:
      - generate_travel_pack_html
      - generate_travel_pack_docx
      - generate_company_profile_html
      - generate_company_profile_docx

    Same semantics as the function it delegates to — contacts of the
    given company plus unaffiliated (freelance) contacts.
    """
    return get_contacts_for_delegation(company_id, active_only=active_only)


def get_contacts_for_local_support(company_id, active_only=True):
    """
    Contacts available as local support for a trip:
    - all active contacts of the given company, OR
    - active contacts with no company (universal / freelance)
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if active_only:
            conditions.append("is_active = 1")
        if company_id:
            conditions.append("(company_id = ? OR company_id IS NULL)")
            params.append(company_id)
        else:
            conditions.append("company_id IS NULL")
        q = "SELECT * FROM contacts"
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY (company_id IS NULL), name"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()

# =========================================================
# TRIP DELEGATION
# =========================================================


def add_trip_delegation(trip_id, contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "INSERT OR IGNORE INTO trip_delegation (trip_id, contact_id) "
            "VALUES (?, ?)",
            (trip_id, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def remove_trip_delegation(trip_id, contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM trip_delegation WHERE trip_id = ? AND contact_id = ?",
            (trip_id, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def set_trip_delegation(trip_id, contact_ids):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM trip_delegation WHERE trip_id = ?", (trip_id,))
        for cid in contact_ids or []:
            c.execute(
                "INSERT OR IGNORE INTO trip_delegation (trip_id, contact_id) "
                "VALUES (?, ?)",
                (trip_id, cid),
            )
        conn.commit()
    finally:
        conn.close()


def get_trip_delegation_members(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT c.* FROM contacts c
               JOIN trip_delegation td ON c.id = td.contact_id
               WHERE td.trip_id = ?
               ORDER BY c.name""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_trip_delegation_ids(trip_id):
    return [m["id"] for m in get_trip_delegation_members(trip_id)]


# =========================================================
# ITEM DELEGATION
# =========================================================


def add_item_delegation_member(item_id, contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "INSERT OR IGNORE INTO item_delegation (item_id, contact_id) "
            "VALUES (?, ?)",
            (item_id, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def remove_item_delegation_member(item_id, contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM item_delegation WHERE item_id = ? AND contact_id = ?",
            (item_id, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_item_delegation_members(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT c.* FROM contacts c
               JOIN item_delegation id ON c.id = id.contact_id
               WHERE id.item_id = ?
               ORDER BY c.name""",
            (item_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def set_item_delegation_members(item_id, contact_ids):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM item_delegation WHERE item_id = ?", (item_id,))
        for cid in contact_ids or []:
            c.execute(
                "INSERT OR IGNORE INTO item_delegation (item_id, contact_id) "
                "VALUES (?, ?)",
                (item_id, cid),
            )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# ITEM CONTACTS
# =========================================================


def add_item_contact(item_id, contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "INSERT OR IGNORE INTO item_contacts (item_id, contact_id) "
            "VALUES (?, ?)",
            (item_id, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def remove_item_contact(item_id, contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM item_contacts WHERE item_id = ? AND contact_id = ?",
            (item_id, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_item_contacts(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT c.* FROM contacts c
               JOIN item_contacts ic ON c.id = ic.contact_id
               WHERE ic.item_id = ?
               ORDER BY c.name""",
            (item_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def set_item_contacts(item_id, contact_ids):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM item_contacts WHERE item_id = ?", (item_id,))
        for cid in contact_ids or []:
            c.execute(
                "INSERT OR IGNORE INTO item_contacts (item_id, contact_id) "
                "VALUES (?, ?)",
                (item_id, cid),
            )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# TRIP CONTACTS
# =========================================================


def get_trip_contacts(trip_id):
    trip = get_trip(trip_id)
    if not trip or not trip.get("trip_contacts"):
        return []
    try:
        contact_ids = json.loads(trip["trip_contacts"])
    except Exception:
        return []
    if not contact_ids:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        placeholders = ",".join(["?"] * len(contact_ids))
        c.execute(
            f"SELECT * FROM contacts WHERE id IN ({placeholders}) "
            "AND is_active = 1",
            contact_ids,
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_trip_contacts(trip_id, contact_ids):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trips SET trip_contacts = ? WHERE id = ?",
            (json.dumps(contact_ids), trip_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_company_default_contacts(company_id):
    comp = get_company(company_id)
    if not comp or not comp.get("default_contact_ids"):
        return []
    try:
        return json.loads(comp["default_contact_ids"])
    except Exception:
        return []


def set_company_default_contacts(company_id, contact_ids):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE companies SET default_contact_ids = ? WHERE id = ?",
            (json.dumps(contact_ids), company_id),
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# EXECUTIVES
# =========================================================


def add_executive(
    company_id,
    name,
    email,
    timezone,
    seat_preference,
    dietary_restrictions,
    preferred_airline,
    tsa_precheck,
    meal_preference,
    is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO executives
               (company_id, name, email, timezone, seat_preference,
                dietary_restrictions, preferred_airline, tsa_precheck,
                meal_preference, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                company_id, name, email, timezone, seat_preference,
                dietary_restrictions, preferred_airline, tsa_precheck,
                meal_preference, 1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_all_executives(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = """
            SELECT e.id, e.name, e.is_active, c.name AS company_name
            FROM executives e
            JOIN companies c ON e.company_id = c.id
        """
        if active_only:
            q += " WHERE e.is_active = 1"
        q += " ORDER BY e.name"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_executive_profile(exec_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """
            SELECT e.*, c.name as company_name, c.default_cost_center,
                   c.policy_notes
            FROM executives e
            JOIN companies c ON e.company_id = c.id
            WHERE e.id = ?
            """,
            (exec_id,),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_full_executive_profile(exec_id):
    raw = get_executive_profile(exec_id)
    if not raw:
        return None
    return {
        "Name": raw.get("name", ""),
        "Email": raw.get("email", ""),
        "Timezone": raw.get("timezone", ""),
        "Seat Preference": raw.get("seat_preference", ""),
        "Dietary": raw.get("dietary_restrictions", ""),
        "Company": raw.get("company_name", ""),
        "Cost Center": raw.get("default_cost_center", ""),
        "Policy Notes": raw.get("policy_notes", ""),
        "Preferred Airline": raw.get("preferred_airline", ""),
        "TSA PreCheck": raw.get("tsa_precheck", ""),
        "Meal Preference": raw.get("meal_preference", ""),
        "Active": "Yes" if raw.get("is_active", 1) else "No",
    }


def get_all_executive_profiles(active_only=False):
    all_execs = get_all_executives(active_only=active_only)
    profiles = []
    for e in all_execs:
        profile = get_full_executive_profile(e["id"])
        if profile:
            mems = get_memberships(e["id"])
            profile["Memberships"] = "; ".join(
                f"{m['program_name']}: {m['membership_number']}" for m in mems
            )
            profile["ID"] = e["id"]
            profiles.append(profile)
    return profiles


def update_executive(
    exec_id,
    company_id,
    name,
    email,
    timezone,
    seat_preference,
    dietary_restrictions,
    preferred_airline,
    tsa_precheck,
    meal_preference,
    is_active=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields = [
            "company_id = ?", "name = ?", "email = ?", "timezone = ?",
            "seat_preference = ?", "dietary_restrictions = ?",
            "preferred_airline = ?", "tsa_precheck = ?",
            "meal_preference = ?",
        ]
        params = [
            company_id, name, email, timezone, seat_preference,
            dietary_restrictions, preferred_airline, tsa_precheck,
            meal_preference,
        ]
        if is_active is not None:
            fields.append("is_active = ?")
            params.append(1 if is_active else 0)
        params.append(exec_id)
        c.execute(
            f"UPDATE executives SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def set_executive_active(exec_id, is_active):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE executives SET is_active = ? WHERE id = ?",
            (1 if is_active else 0, exec_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_executive_trip_count(exec_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM trips WHERE exec_id = ?", (exec_id,))
        return c.fetchone()[0]
    finally:
        conn.close()


def delete_executive(exec_id, force=False):
    import shutil

    trip_count = get_executive_trip_count(exec_id)
    if trip_count > 0 and not force:
        return (
            False,
            f"Cannot delete: Executive has {trip_count} trip(s). "
            "Delete trips first or use force delete.",
        )

    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        if force:
            c.execute("SELECT id FROM trips WHERE exec_id = ?", (exec_id,))
            trip_ids = [row[0] for row in c.fetchall()]
            for trip_id in trip_ids:
                folder = f"ea_receipts/trip_{trip_id}"
                if os.path.exists(folder):
                    shutil.rmtree(folder)
            c.execute(
                "DELETE FROM itinerary_items WHERE trip_id IN "
                "(SELECT id FROM trips WHERE exec_id = ?)",
                (exec_id,),
            )
            c.execute(
                "DELETE FROM trip_stops WHERE trip_id IN "
                "(SELECT id FROM trips WHERE exec_id = ?)",
                (exec_id,),
            )
            c.execute("DELETE FROM trips WHERE exec_id = ?", (exec_id,))
        c.execute(
            "DELETE FROM executive_memberships WHERE exec_id = ?", (exec_id,)
        )
        c.execute(
            "DELETE FROM executive_passports WHERE exec_id = ?", (exec_id,)
        )
        c.execute("DELETE FROM executives WHERE id = ?", (exec_id,))
        conn.commit()
    finally:
        conn.close()
    return (
        True,
        f"Executive and {trip_count} trip(s) deleted successfully.",
    )


# =========================================================
# PASSPORTS
# =========================================================


def add_passport(
    exec_id, country, passport_number, expiry_date=None,
    issued_date=None, notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO executive_passports
               (exec_id, country, passport_number, expiry_date,
                issued_date, notes)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (exec_id, country, passport_number, expiry_date,
             issued_date, notes),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_passports(exec_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM executive_passports WHERE exec_id = ? "
            "ORDER BY country",
            (exec_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()
def find_duplicate_passports(exec_id, country, passport_number=None):
    """
    Return existing passports for this exec in this country. If
    `passport_number` is given, matches are narrowed further (a row with
    the same number is a stronger signal than a same-country row).

    A person can legitimately have two passports for the same country
    (diplomatic + regular, or an expired + new one during renewal), so
    this is a warning, not a hard refusal.
    """
    if not exec_id or not country:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM executive_passports "
            "WHERE exec_id = ? AND country = ? "
            "ORDER BY id",
            (exec_id, country),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def find_duplicate_memberships(exec_id, program_name):
    """
    Return existing memberships for this exec with the same program name.
    Program-name matching is case-insensitive so 'British Airways' and
    'british airways' collapse.

    A person can legitimately have two accounts for the same program
    (personal + corporate), so this is a warning, not a hard refusal.
    """
    if not exec_id or not program_name:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM executive_memberships "
            "WHERE exec_id = ? AND LOWER(program_name) = LOWER(?) "
            "ORDER BY id",
            (exec_id, program_name.strip()),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()

def delete_passport(passport_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM executive_passports WHERE id = ?", (passport_id,)
        )
        conn.commit()
    finally:
        conn.close()


def update_passport(
    passport_id, country, passport_number,
    expiry_date=None, issued_date=None, notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """UPDATE executive_passports
               SET country = ?, passport_number = ?, expiry_date = ?,
                   issued_date = ?, notes = ?
               WHERE id = ?""",
            (country, passport_number, expiry_date, issued_date,
             notes, passport_id),
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# MEMBERSHIPS
# =========================================================


def add_membership(
    exec_id, category, program_name, membership_number,
    tier=None, alliance=None, airport_code=None, notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO executive_memberships
               (exec_id, category, program_name, membership_number,
                tier, alliance, airport_code, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (exec_id, category, program_name, membership_number,
             tier, alliance, airport_code, notes),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_memberships(exec_id, category=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        if category:
            c.execute(
                """SELECT * FROM executive_memberships
                   WHERE exec_id = ? AND category = ?
                   ORDER BY program_name""",
                (exec_id, category),
            )
        else:
            c.execute(
                """SELECT * FROM executive_memberships
                   WHERE exec_id = ?
                   ORDER BY category, program_name""",
                (exec_id,),
            )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def delete_membership(membership_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM executive_memberships WHERE id = ?",
            (membership_id,),
        )
        conn.commit()
    finally:
        conn.close()


def update_membership(
    membership_id, category, program_name, membership_number,
    tier=None, alliance=None, airport_code=None, notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """UPDATE executive_memberships
               SET category = ?, program_name = ?, membership_number = ?,
                   tier = ?, alliance = ?, airport_code = ?, notes = ?
               WHERE id = ?""",
            (category, program_name, membership_number,
             tier, alliance, airport_code, notes, membership_id),
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# CATEGORIES
# =========================================================


def add_category(name, is_active=1):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        try:
            c.execute(
                "INSERT INTO categories (name, is_active) VALUES (?, ?)",
                (name, 1 if is_active else 0),
            )
            conn.commit()
            return c.lastrowid
        except sqlite3.IntegrityError:
            return None
    finally:
        conn.close()


def get_all_categories(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT id, name, is_active FROM categories"
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY name"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def set_category_active(category_id, is_active):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE categories SET is_active = ? WHERE id = ?",
            (1 if is_active else 0, category_id),
        )
        conn.commit()
    finally:
        conn.close()


def delete_category(category_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM categories WHERE id = ?", (category_id,))
        conn.commit()
    finally:
        conn.close()


# =========================================================
# TRIPS
# =========================================================


def create_or_get_trip(
    exec_id, destination_summary, start_date, end_date, purpose,
    display_currency="USD", base_currency="USD", status="draft",
    trip_contacts=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """SELECT id FROM trips
               WHERE exec_id = ? AND destination = ? AND start_date = ?
                 AND status = ?
               ORDER BY created_at DESC LIMIT 1""",
            (exec_id, destination_summary, start_date, status),
        )
        row = c.fetchone()
        if row:
            return row[0]
        contacts_json = json.dumps(trip_contacts) if trip_contacts else None
        c.execute(
            """INSERT INTO trips (exec_id, destination, start_date, end_date,
                                  purpose, status, display_currency,
                                  base_currency, trip_contacts)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                exec_id, destination_summary, start_date, end_date,
                purpose, status, display_currency, base_currency,
                contacts_json,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM trips WHERE id = ?", (trip_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_trip_budget(trip_id, budget):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trips SET budget = ? WHERE id = ?", (budget, trip_id)
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_status(trip_id, status):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trips SET status = ? WHERE id = ?", (status, trip_id)
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_departure_details(trip_id, city, region, country):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """UPDATE trips
               SET departure_city = ?, departure_region = ?,
                   departure_country = ?
               WHERE id = ?""",
            (city, region, country, trip_id),
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_purpose(trip_id, purpose):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trips SET purpose = ? WHERE id = ?", (purpose, trip_id)
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_dates(trip_id, start_date, end_date, destination):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """UPDATE trips
               SET start_date = ?, end_date = ?, destination = ?
               WHERE id = ?""",
            (start_date, end_date, destination, trip_id),
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_base_currency(trip_id, base_currency):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trips SET base_currency = ? WHERE id = ?",
            (base_currency, trip_id),
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_display_currency(trip_id, display_currency):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trips SET display_currency = ? WHERE id = ?",
            (display_currency, trip_id),
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_currencies(trip_id, base_currency, display_currency):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """UPDATE trips SET base_currency = ?, display_currency = ?
               WHERE id = ?""",
            (base_currency, display_currency, trip_id),
        )
        conn.commit()
    finally:
        conn.close()


def update_trip_receipt_folder(trip_id, url):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trips SET receipt_upload_folder_url = ? WHERE id = ?",
            (url or None, trip_id),
        )
        conn.commit()
    finally:
        conn.close()


def delete_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM trips WHERE id = ?", (trip_id,))
        conn.commit()
    finally:
        conn.close()


def delete_trips(trip_ids):
    if not trip_ids:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        placeholders = ",".join("?" * len(trip_ids))
        c.execute(
            f"DELETE FROM trips WHERE id IN ({placeholders})", trip_ids
        )
        conn.commit()
    finally:
        conn.close()


def duplicate_trip(trip_id, exec_id):
    original = get_trip(trip_id)
    if not original:
        return None
    new_purpose = original["purpose"]
    new_purpose += " (Copy)"

    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO trips
               (exec_id, destination, start_date, end_date, purpose, status,
                created_at, budget, departure_city, departure_region,
                departure_country, base_currency, display_currency,
                trip_contacts, receipt_upload_folder_url)
               VALUES (?, ?, ?, ?, ?, 'draft', CURRENT_TIMESTAMP,
                       ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                exec_id,
                original["destination"],
                original["start_date"],
                original["end_date"],
                new_purpose,
                original.get("budget", 0),
                original.get("departure_city", ""),
                original.get("departure_region", ""),
                original.get("departure_country", ""),
                original.get("base_currency", "USD"),
                original.get("display_currency", "USD"),
                original.get("trip_contacts"),
                original.get("receipt_upload_folder_url"),
            ),
        )
        new_trip_id = c.lastrowid
        conn.commit()
    finally:
        conn.close()

    # Duplicate stops
    for stop in get_trip_stops(trip_id):
        add_trip_stop(
            new_trip_id,
            stop["stop_order"],
            stop["city"],
            stop["country"],
            stop["region"],
            stop["start_date"],
            stop["end_date"],
            stop.get("notes", ""),
        )

    # Duplicate delegation
    set_trip_delegation(new_trip_id, get_trip_delegation_ids(trip_id))

    # Duplicate items
    for item in get_items_for_trip(trip_id):
        new_item_id = add_itinerary_item(
            new_trip_id,
            item["item_type"],
            item["description"],
            item["datetime_start"],
            item["datetime_end"],
            item["location"],
            item["cost"],
            item["confirmation_code"],
            item["notes"],
            item.get("is_confirmed", 0),
            item.get("cost_currency", "USD"),
            item.get("timezone"),
            item.get("venue_id"),
            item.get("cost_date"),
        )
        deleg_ids = [m["id"] for m in get_item_delegation_members(item["id"])]
        if deleg_ids:
            set_item_delegation_members(new_item_id, deleg_ids)
        contact_ids = [c["id"] for c in get_item_contacts(item["id"])]
        if contact_ids:
            set_item_contacts(new_item_id, contact_ids)

    return new_trip_id


# =========================================================
# TRIP STOPS
# =========================================================


def add_trip_stop(
    trip_id, stop_order, city, country, region, start_date, end_date,
    notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO trip_stops
               (trip_id, stop_order, city, country, region, start_date,
                end_date, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (trip_id, stop_order, city, country, region, start_date,
             end_date, notes),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_trip_stops(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM trip_stops WHERE trip_id = ? ORDER BY stop_order",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def find_duplicate_trip_stops(trip_id, city, start_date, end_date):
    """
    Return trip stops on this trip with the same city (case-insensitive)
    whose date range overlaps [start_date, end_date].

    A legitimate itinerary can revisit a city on separate legs (London
    → Paris → London), so the date-range overlap is what distinguishes
    a genuine duplicate from an intentional revisit.
    """
    if not trip_id or not city or not start_date or not end_date:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM trip_stops
               WHERE trip_id = ?
                 AND LOWER(city) = LOWER(?)
                 AND start_date <= ?
                 AND end_date >= ?
               ORDER BY stop_order""",
            (trip_id, city.strip(), end_date, start_date),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_trip_stop(
    stop_id, city, country, region, start_date, end_date, notes
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """UPDATE trip_stops
               SET city = ?, country = ?, region = ?, start_date = ?,
                   end_date = ?, notes = ?
               WHERE id = ?""",
            (city, country, region, start_date, end_date, notes, stop_id),
        )
        conn.commit()
    finally:
        conn.close()


def delete_trip_stop(stop_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM trip_stops WHERE id = ?", (stop_id,))
        conn.commit()
    finally:
        conn.close()


def reorder_trip_stops(trip_id, new_order_ids):
    """
    Persist a new ordering for a trip's stops. `new_order_ids` is a list
    of stop IDs in the desired top-to-bottom order. Uses a two-pass
    negate-then-restore strategy so a UNIQUE(trip_id, stop_order) index
    never collides mid-update.
    """
    if not new_order_ids:
        return False
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        for new_pos, stop_id in enumerate(new_order_ids, start=1):
            c.execute(
                "UPDATE trip_stops SET stop_order = ? WHERE id = ? AND trip_id = ?",
                (-new_pos, stop_id, trip_id),
            )
        for new_pos, stop_id in enumerate(new_order_ids, start=1):
            c.execute(
                "UPDATE trip_stops SET stop_order = ? WHERE id = ? AND trip_id = ?",
                (new_pos, stop_id, trip_id),
            )
        conn.commit()
        return True
    finally:
        conn.close()


def delete_all_trip_stops(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM trip_stops WHERE trip_id = ?", (trip_id,))
        conn.commit()
    finally:
        conn.close()


# =========================================================
# ITINERARY ITEMS
# =========================================================


def add_itinerary_item(
    trip_id, item_type, description, datetime_start, datetime_end,
    location, cost, confirmation_code, notes, is_confirmed=0,
    cost_currency="USD", timezone=None, venue_id=None, cost_date=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO itinerary_items
               (trip_id, item_type, description, datetime_start,
                datetime_end, location, cost, confirmation_code, notes,
                is_confirmed, cost_currency, timezone, venue_id, cost_date)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                trip_id, item_type, description, datetime_start,
                datetime_end, location, cost, confirmation_code, notes,
                is_confirmed, cost_currency, timezone, venue_id, cost_date,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_items_for_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM itinerary_items
               WHERE trip_id = ?
               ORDER BY datetime_start""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_itinerary_item(
    item_id, item_type, description, datetime_start, datetime_end,
    location, cost, confirmation_code, notes, is_confirmed,
    cost_currency="USD", timezone=None, venue_id=None, cost_date=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """UPDATE itinerary_items SET
                item_type = ?, description = ?, datetime_start = ?,
                datetime_end = ?, location = ?, cost = ?,
                confirmation_code = ?, notes = ?, is_confirmed = ?,
                cost_currency = ?, timezone = ?, venue_id = ?,
                cost_date = ?
               WHERE id = ?""",
            (
                item_type, description, datetime_start, datetime_end,
                location, cost, confirmation_code, notes, is_confirmed,
                cost_currency, timezone, venue_id, cost_date, item_id,
            ),
        )
        conn.commit()
    finally:
        conn.close()


def delete_itinerary_item(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM itinerary_items WHERE id = ?", (item_id,))
        conn.commit()
    finally:
        conn.close()


def update_receipt_path(item_id, file_path):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE itinerary_items SET receipt_path = ? WHERE id = ?",
            (file_path, item_id),
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# BUDGET & SPENDING
# =========================================================


def get_trip_spending(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "SELECT COALESCE(SUM(cost), 0) FROM itinerary_items "
            "WHERE trip_id = ?",
            (trip_id,),
        )
        total_all = c.fetchone()[0]
        c.execute(
            "SELECT COALESCE(SUM(cost), 0) FROM itinerary_items "
            "WHERE trip_id = ? AND is_confirmed = 1",
            (trip_id,),
        )
        total_confirmed = c.fetchone()[0]
        c.execute(
            "SELECT COALESCE(SUM(cost), 0) FROM itinerary_items "
            "WHERE trip_id = ? AND is_confirmed = 0",
            (trip_id,),
        )
        total_estimated = c.fetchone()[0]
    finally:
        conn.close()
    return {
        "total_all": total_all,
        "total_confirmed": total_confirmed,
        "total_estimated": total_estimated,
    }


def get_spending_summary(
    exec_id=None, company_id=None, start_date=None, end_date=None
):
    import currency

    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        trip_query = """
            SELECT t.*, e.name AS executive_name, c.name AS company_name
            FROM trips t
            JOIN executives e ON t.exec_id = e.id
            JOIN companies c ON e.company_id = c.id
            WHERE 1=1
        """
        params = []
        if exec_id:
            trip_query += " AND e.id = ?"
            params.append(exec_id)
        if company_id:
            trip_query += " AND c.id = ?"
            params.append(company_id)
        if start_date:
            trip_query += " AND t.start_date >= ?"
            params.append(start_date)
        if end_date:
            trip_query += " AND t.end_date <= ?"
            params.append(end_date)
        trip_query += " ORDER BY t.start_date DESC"
        c.execute(trip_query, params)
        trips = c.fetchall()

        summary = []
        for trip in trips:
            base_cur = (trip["base_currency"] or "USD").upper()
            c.execute(
                "SELECT * FROM itinerary_items WHERE trip_id = ?",
                (trip["id"],),
            )
            items = c.fetchall()
            total = confirmed = estimated = 0.0
            for item in items:
                cost = item["cost"] or 0
                cost_cur = (item["cost_currency"] or "USD").upper()
                on_date = item["cost_date"]
                if not on_date and item["datetime_start"]:
                    on_date = item["datetime_start"][:10]
                try:
                    conv = currency.convert_amount(
                        cost, cost_cur, base_cur, on_date
                    )
                except Exception:
                    conv = float(cost)
                total += conv
                if item["is_confirmed"]:
                    confirmed += conv
                else:
                    estimated += conv
            summary.append(
                {
                    "trip_id": trip["id"],
                    "purpose": trip["purpose"],  # ← add this line
                    "executive_name": trip["executive_name"],
                    "company_name": trip["company_name"],
                    "destination": trip["destination"],
                    "start_date": trip["start_date"],
                    "end_date": trip["end_date"],
                    "budget": trip["budget"] or 0,
                    "status": trip["status"],
                    "base_currency": base_cur,
                    "display_currency": trip["display_currency"],
                    "total_spent": total,
                    "confirmed_spent": confirmed,
                    "estimated_spent": estimated,
                }
            )
        return summary
    finally:
        conn.close()


# =========================================================
# IMPORT / MERGE
# =========================================================


def merge_database_data(data):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        added_execs = added_trips = added_stops = added_items = 0

        for exec_data in data.get("executives", []):
            email = exec_data.get("email")
            if not email:
                continue
            c.execute("SELECT id FROM executives WHERE email = ?", (email,))
            if c.fetchone():
                continue
            company_name = exec_data.get("company_name")
            if not company_name:
                continue
            company_id = _find_or_create_company(company_name)
            c.execute(
                """INSERT INTO executives
                   (company_id, name, email, timezone, seat_preference,
                    hotel_loyalty, frequent_flyer_number,
                    dietary_restrictions, passport_number,
                    preferred_airline, tsa_precheck, meal_preference,
                    is_active)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                (
                    company_id,
                    exec_data.get("name"),
                    email,
                    exec_data.get("timezone", "America/New_York"),
                    exec_data.get("seat_preference"),
                    exec_data.get("hotel_loyalty"),
                    exec_data.get("frequent_flyer_number"),
                    exec_data.get("dietary_restrictions"),
                    exec_data.get("passport_number"),
                    exec_data.get("preferred_airline"),
                    exec_data.get("tsa_precheck"),
                    exec_data.get("meal_preference"),
                ),
            )
            added_execs += 1
            conn.commit()

        for trip_data in data.get("trips", []):
            exec_email = trip_data.get("executive_email")
            if not exec_email:
                continue
            c.execute(
                "SELECT id FROM executives WHERE email = ?", (exec_email,)
            )
            row = c.fetchone()
            if not row:
                continue
            exec_id = row[0]
            dest = trip_data.get("destination")
            start = trip_data.get("start_date")
            if not dest or not start:
                continue
            c.execute(
                """SELECT id FROM trips
                   WHERE exec_id = ? AND destination = ? AND start_date = ?""",
                (exec_id, dest, start),
            )
            if c.fetchone():
                continue
            c.execute(
                """INSERT INTO trips
                   (exec_id, destination, start_date, end_date, purpose,
                    status, budget, departure_city, departure_region,
                    departure_country, display_currency, base_currency,
                    trip_contacts)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL)""",
                (
                    exec_id, dest, start, trip_data.get("end_date"),
                    trip_data.get("purpose"),
                    trip_data.get("status", "draft"),
                    trip_data.get("budget", 0),
                    trip_data.get("departure_city"),
                    trip_data.get("departure_region"),
                    trip_data.get("departure_country"),
                    trip_data.get("display_currency", "USD"),
                    trip_data.get("base_currency", "USD"),
                ),
            )
            trip_id = c.lastrowid
            added_trips += 1
            conn.commit()

            for stop in trip_data.get("stops", []):
                c.execute(
                    """INSERT INTO trip_stops
                       (trip_id, stop_order, city, country, region,
                        start_date, end_date, notes)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        trip_id, stop.get("stop_order", 0),
                        stop.get("city"), stop.get("country"),
                        stop.get("region"), stop.get("start_date"),
                        stop.get("end_date"), stop.get("notes"),
                    ),
                )
                added_stops += 1
                conn.commit()

            for item in trip_data.get("items", []):
                c.execute(
                    """INSERT INTO itinerary_items
                       (trip_id, item_type, description, datetime_start,
                        datetime_end, location, cost, cost_currency,
                        is_confirmed, confirmation_code, notes, timezone,
                        cost_date)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        trip_id, item.get("type"),
                        item.get("description"),
                        item.get("datetime_start"),
                        item.get("datetime_end"),
                        item.get("location"),
                        item.get("cost", 0),
                        item.get("cost_currency", "USD"),
                        1 if item.get("is_confirmed") else 0,
                        item.get("confirmation_code"),
                        item.get("notes"),
                        None,
                        item.get("cost_date") or (
                            item.get("datetime_start", "")[:10]
                            if item.get("datetime_start") else None
                        ),
                    ),
                )
                added_items += 1
                conn.commit()

        return (
            f"✅ Imported {added_execs} executives, {added_trips} trips, "
            f"{added_stops} stops, {added_items} items."
        )
    finally:
        conn.close()


def import_executives_from_csv(reader):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        added = skipped = 0
        for row in reader:
            email = row.get("email")
            if not email:
                continue
            c.execute(
                "SELECT id FROM executives WHERE email = ?", (email,)
            )
            if c.fetchone():
                skipped += 1
                continue
            company_name = row.get("company_name")
            if not company_name:
                continue
            company_id = _find_or_create_company(company_name)
            c.execute(
                """INSERT INTO executives
                   (company_id, name, email, timezone, seat_preference,
                    hotel_loyalty, frequent_flyer_number,
                    dietary_restrictions, passport_number,
                    preferred_airline, tsa_precheck, meal_preference,
                    is_active)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
                (
                    company_id, row.get("name"), email,
                    row.get("timezone", "America/New_York"),
                    row.get("seat_preference"),
                    row.get("hotel_loyalty"),
                    row.get("frequent_flyer_number"),
                    row.get("dietary_restrictions"),
                    row.get("passport_number"),
                    row.get("preferred_airline"),
                    row.get("tsa_precheck"),
                    row.get("meal_preference"),
                ),
            )
            added += 1
            conn.commit()
        return f"✅ Added {added} executives. Skipped {skipped} duplicates."
    finally:
        conn.close()


# =========================================================
# TRIP TEMPLATES
# =========================================================


def save_trip_as_template(trip_id, template_name, description=None):
    trip = get_trip(trip_id)
    if not trip:
        return None
    stops = get_trip_stops(trip_id)
    items = get_items_for_trip(trip_id)

    stops_data = [
        {
            "city": s.get("city", ""),
            "country": s.get("country", ""),
            "region": s.get("region", ""),
            "notes": s.get("notes", ""),
        }
        for s in stops
    ]
    items_data = [
        {
            "item_type": i.get("item_type", ""),
            "description": i.get("description", ""),
            "location": i.get("location", ""),
            "cost_currency": i.get("cost_currency", "USD"),
            "is_confirmed": i.get("is_confirmed", 0),
            "confirmation_code": i.get("confirmation_code", ""),
            "notes": i.get("notes", ""),
        }
        for i in items
    ]

    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO trip_templates
               (name, description, departure_city, departure_region,
                departure_country, display_currency, base_currency,
                stops_json, items_json, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)""",
            (
                template_name, description,
                trip.get("departure_city", ""),
                trip.get("departure_region", ""),
                trip.get("departure_country", ""),
                trip.get("display_currency", "USD"),
                trip.get("base_currency", "USD"),
                json.dumps(stops_data, indent=2),
                json.dumps(items_data, indent=2),
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_trip_templates(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = """
            SELECT id, name, description, created_at, departure_city,
                   departure_region, departure_country, display_currency,
                   base_currency, is_active
            FROM trip_templates
        """
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY name"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_trip_template(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM trip_templates WHERE id = ?", (template_id,)
        )
        row = c.fetchone()
        if not row:
            return None
        template = dict(row)
        template["stops"] = json.loads(template.get("stops_json") or "[]")
        template["items"] = json.loads(template.get("items_json") or "[]")
        return template
    finally:
        conn.close()


def update_trip_template(
    template_id, name=None, description=None, is_active=None
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        if name is not None:
            fields.append("name = ?")
            params.append(name)
        if description is not None:
            fields.append("description = ?")
            params.append(description)
        if is_active is not None:
            fields.append("is_active = ?")
            params.append(1 if is_active else 0)
        if not fields:
            return
        params.append(template_id)
        c.execute(
            f"UPDATE trip_templates SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def set_trip_template_active(template_id, is_active):
    update_trip_template(template_id, is_active=is_active)


def delete_trip_template(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM trip_templates WHERE id = ?", (template_id,)
        )
        conn.commit()
        return True
    finally:
        conn.close()


def apply_trip_template(
    template_id, exec_id, new_purpose, start_date, end_date, budget=0
):
    template = get_trip_template(template_id)
    if not template:
        return None
    stops = template.get("stops", [])
    stop_names = [s["city"] for s in stops if s.get("city")]
    dest_summary = " → ".join(stop_names) if stop_names else "Template Trip"

    trip_id = create_or_get_trip(
        exec_id, dest_summary,
        start_date.isoformat(), end_date.isoformat(), new_purpose,
        template.get("display_currency", "USD"),
        template.get("base_currency", "USD"),
    )
    update_trip_budget(trip_id, budget)
    update_trip_departure_details(
        trip_id,
        template.get("departure_city", ""),
        template.get("departure_region", ""),
        template.get("departure_country", ""),
    )

    stop_order = 0
    for stop in stops:
        stop_order += 1
        add_trip_stop(
            trip_id, stop_order,
            stop.get("city", ""), stop.get("country", ""),
            stop.get("region", ""),
            start_date.isoformat(), end_date.isoformat(),
            stop.get("notes", ""),
        )
    for item in template.get("items", []):
        add_itinerary_item(
            trip_id,
            item.get("item_type", ""),
            item.get("description", ""),
            start_date.isoformat() + "T08:00:00",
            start_date.isoformat() + "T09:00:00",
            item.get("location", ""),
            0,
            item.get("confirmation_code", ""),
            item.get("notes", ""),
            0,
            item.get("cost_currency", "USD"),
            None, None,
            start_date.isoformat(),
        )
    return trip_id


# =========================================================
# EXPORT
# =========================================================


def get_all_rows(table_name):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(f"SELECT * FROM {table_name}")
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def export_all_data():
    tables = [
        # Original travel planner tables
        "companies", "executives", "contacts", "trip_delegation",
        "trips", "itinerary_items", "trip_stops", "trip_templates",
        "categories", "executive_memberships", "executive_passports",
        "item_delegation", "item_contacts", "packing_lists",
        "packing_items", "trip_checklists", "trip_checklist_items",
        "hospitals", "exchange_rates",
        # EA Command Center tables
        "meetings", "meeting_attendees", "meeting_action_items",
        "meeting_decisions",
        "events", "event_guests", "event_budget_items",
        "vendors", "vendor_engagements", "vendor_payments",
        "speaking_engagements", "speaking_deliverables", "speaking_cfps",
        "communications",
        "budget_templates", "budget_template_items",
        "restore_log",
    ]
    return {t: get_all_rows(t) for t in tables}


# =========================================================
# VENUES
# =========================================================


def add_venue(
    name, address=None, city=None, country=None,
    wifi_ssid=None, wifi_password=None, badge_info=None,
    dress_code_notes=None, notes=None, is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO venues
               (name, address, city, country, wifi_ssid, wifi_password,
                badge_info, dress_code_notes, notes, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                name, address, city, country, wifi_ssid, wifi_password,
                badge_info, dress_code_notes, notes, 1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_venues(country=None, active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if active_only:
            conditions.append("is_active = 1")
        if country:
            conditions.append("country = ?")
            params.append(country)
        q = "SELECT * FROM venues"
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY name"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_venue(venue_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM venues WHERE id = ?", (venue_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_venue(
    venue_id, name=None, address=None, city=None, country=None,
    wifi_ssid=None, wifi_password=None, badge_info=None,
    dress_code_notes=None, notes=None, is_active=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for col, val in [
            ("name", name), ("address", address), ("city", city),
            ("country", country), ("wifi_ssid", wifi_ssid),
            ("wifi_password", wifi_password), ("badge_info", badge_info),
            ("dress_code_notes", dress_code_notes), ("notes", notes),
        ]:
            if val is not None:
                fields.append(f"{col} = ?")
                params.append(val)
        if is_active is not None:
            fields.append("is_active = ?")
            params.append(1 if is_active else 0)
        if not fields:
            return
        params.append(venue_id)
        c.execute(
            f"UPDATE venues SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def set_venue_active(venue_id, is_active):
    update_venue(venue_id, is_active=is_active)


def delete_venue(venue_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE itinerary_items SET venue_id = NULL WHERE venue_id = ?",
            (venue_id,),
        )
        c.execute("DELETE FROM venues WHERE id = ?", (venue_id,))
        conn.commit()
    finally:
        conn.close()


def get_venues_for_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT DISTINCT v.* FROM venues v
               JOIN itinerary_items i ON i.venue_id = v.id
               WHERE i.trip_id = ? AND v.is_active = 1
               ORDER BY v.name""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def find_duplicate_venues(name=None, address=None):
    if not name and not address:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if name:
            conditions.append("name = ?")
            params.append(name)
        if address:
            conditions.append("address = ?")
            params.append(address)
        q = (
            "SELECT * FROM venues WHERE is_active = 1 AND "
            f"({' OR '.join(conditions)})"
        )
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# =========================================================
# DESTINATION GUIDES
# =========================================================


def add_destination_guide(
    country, language=None, currency=None,
    emergency_police=None, emergency_ambulance=None, emergency_fire=None,
    etiquette_notes=None, phrases=None, packing_tips=None,
    connectivity_notes=None, recommended_apps=None, notes=None, is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO destination_guides
               (country, language, currency, emergency_police,
                emergency_ambulance, emergency_fire, etiquette_notes,
                phrases, packing_tips, connectivity_notes,
                recommended_apps, notes, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                country, language, currency, emergency_police,
                emergency_ambulance, emergency_fire, etiquette_notes,
                phrases, packing_tips, connectivity_notes,
                recommended_apps, notes, 1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_destination_guides(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT * FROM destination_guides"
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY country"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_destination_guide(guide_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM destination_guides WHERE id = ?", (guide_id,)
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_destination_guide_by_country(country):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM destination_guides
               WHERE country = ? AND is_active = 1""",
            (country,),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_destination_guide(guide_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(guide_id)
        c.execute(
            f"UPDATE destination_guides SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def set_destination_guide_active(guide_id, is_active):
    update_destination_guide(guide_id, is_active=1 if is_active else 0)


def delete_destination_guide(guide_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM destination_guides WHERE id = ?", (guide_id,)
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# VISA RULES
# =========================================================


def add_visa_rule(
    passport_nationality, destination_country, visa_required=1,
    max_stay_days=None, processing_time=None, fee=None, notes=None,
    is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO visa_rules
               (passport_nationality, destination_country, visa_required,
                max_stay_days, processing_time, fee, notes, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                passport_nationality, destination_country, visa_required,
                max_stay_days, processing_time, fee, notes,
                1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_visa_rules(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT * FROM visa_rules"
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY passport_nationality, destination_country"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_visa_rule(rule_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM visa_rules WHERE id = ?", (rule_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_visa_rule_for_pair(passport_nationality, destination_country):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM visa_rules
               WHERE passport_nationality = ? AND destination_country = ?
                 AND is_active = 1""",
            (passport_nationality, destination_country),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_visa_rule(rule_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(rule_id)
        c.execute(
            f"UPDATE visa_rules SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def set_visa_rule_active(rule_id, is_active):
    update_visa_rule(rule_id, is_active=1 if is_active else 0)


def delete_visa_rule(rule_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM visa_rules WHERE id = ?", (rule_id,))
        conn.commit()
    finally:
        conn.close()


# =========================================================
# CHECKLIST TEMPLATES
# =========================================================


def add_checklist_template(
    name, description=None, items_json="[]", is_active=1
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO checklist_templates
               (name, description, items_json, is_active)
               VALUES (?, ?, ?, ?)""",
            (name, description, items_json, 1 if is_active else 0),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_checklist_templates(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT * FROM checklist_templates"
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY name"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_checklist_template(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM checklist_templates WHERE id = ?",
            (template_id,),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_checklist_template(template_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(template_id)
        c.execute(
            f"UPDATE checklist_templates SET {', '.join(fields)} "
            "WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def set_checklist_template_active(template_id, is_active):
    update_checklist_template(template_id, is_active=1 if is_active else 0)


def delete_checklist_template(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM checklist_templates WHERE id = ?", (template_id,)
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# PACKING TEMPLATES
# =========================================================


def add_packing_template(
    name, category=None, items_json="[]", is_active=1
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO packing_templates
               (name, category, items_json, is_active)
               VALUES (?, ?, ?, ?)""",
            (name, category, items_json, 1 if is_active else 0),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_packing_templates(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT * FROM packing_templates"
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY name"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_packing_template(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM packing_templates WHERE id = ?", (template_id,)
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_packing_template(template_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(template_id)
        c.execute(
            f"UPDATE packing_templates SET {', '.join(fields)} "
            "WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def set_packing_template_active(template_id, is_active):
    update_packing_template(template_id, is_active=1 if is_active else 0)


def delete_packing_template(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM packing_templates WHERE id = ?", (template_id,)
        )
        conn.commit()
    finally:
        conn.close()


# =========================================================
# PER DIEM
# =========================================================


def set_per_diem(
    trip_id, contact_id, daily_rate, days, currency="USD", notes=None
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "SELECT id FROM per_diem WHERE trip_id = ? AND contact_id = ?",
            (trip_id, contact_id),
        )
        row = c.fetchone()
        if row:
            c.execute(
                """UPDATE per_diem
                   SET daily_rate = ?, days = ?, currency = ?, notes = ?
                   WHERE id = ?""",
                (daily_rate, days, currency, notes, row[0]),
            )
        else:
            c.execute(
                """INSERT INTO per_diem
                   (trip_id, contact_id, daily_rate, days, currency, notes)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (trip_id, contact_id, daily_rate, days, currency, notes),
            )
        conn.commit()
    finally:
        conn.close()


def get_per_diem(trip_id, contact_id=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        if contact_id:
            c.execute(
                "SELECT * FROM per_diem WHERE trip_id = ? AND contact_id = ?",
                (trip_id, contact_id),
            )
            row = c.fetchone()
            return dict(row) if row else None
        c.execute(
            "SELECT * FROM per_diem WHERE trip_id = ? ORDER BY contact_id",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def delete_per_diem(per_diem_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM per_diem WHERE id = ?", (per_diem_id,))
        conn.commit()
    finally:
        conn.close()


# =========================================================
# EXPENSES
# =========================================================


def add_expense(
    trip_id, contact_id, expense_date, category=None, description=None,
    amount=0, currency="USD", receipt_path=None, notes=None,
    is_reimbursable=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO expenses
               (trip_id, contact_id, expense_date, category, description,
                amount, currency, receipt_path, notes, is_reimbursable)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                trip_id, contact_id, expense_date, category, description,
                amount, currency, receipt_path, notes, is_reimbursable,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_expenses(trip_id, contact_id=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        if contact_id:
            c.execute(
                """SELECT * FROM expenses
                   WHERE trip_id = ? AND contact_id = ?
                   ORDER BY expense_date DESC, id DESC""",
                (trip_id, contact_id),
            )
        else:
            c.execute(
                """SELECT * FROM expenses
                   WHERE trip_id = ?
                   ORDER BY expense_date DESC, id DESC""",
                (trip_id,),
            )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_expense(expense_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM expenses WHERE id = ?", (expense_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_expense(expense_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(expense_id)
        c.execute(
            f"UPDATE expenses SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def delete_expense(expense_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM expenses WHERE id = ?", (expense_id,))
        conn.commit()
    finally:
        conn.close()


def get_expense_summary(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT DISTINCT c.id, c.name, c.role
               FROM contacts c
               WHERE c.id IN (
                   SELECT contact_id FROM per_diem WHERE trip_id = ?
                   UNION
                   SELECT contact_id FROM expenses
                   WHERE trip_id = ? AND contact_id IS NOT NULL
               )""",
            (trip_id, trip_id),
        )
        members = c.fetchall()
        summary = []
        for m in members:
            c.execute(
                """SELECT daily_rate, days, currency FROM per_diem
                   WHERE trip_id = ? AND contact_id = ?""",
                (trip_id, m["id"]),
            )
            pd_row = c.fetchone()
            daily_rate = pd_row["daily_rate"] if pd_row else 0
            days = pd_row["days"] if pd_row else 0
            currency = pd_row["currency"] if pd_row else "USD"
            allowance = (daily_rate or 0) * (days or 0)
            c.execute(
                """SELECT COALESCE(SUM(amount), 0) as total,
                          COUNT(*) as cnt
                   FROM expenses WHERE trip_id = ? AND contact_id = ?""",
                (trip_id, m["id"]),
            )
            exp_row = c.fetchone()
            spent = exp_row["total"] or 0
            entry_count = exp_row["cnt"] or 0
            summary.append({
                "member_id": m["id"],
                "name": m["name"],
                "role": m["role"],
                "daily_rate": daily_rate,
                "days": days,
                "currency": currency,
                "allowance": allowance,
                "spent": spent,
                "remaining": allowance - spent,
                "entry_count": entry_count,
            })
        return summary
    finally:
        conn.close()


def get_trip_delegation_members_for_expenses(trip_id):
    """Alias — kept for backward compatibility with app.py call sites."""
    return get_trip_delegation_members(trip_id)


# =========================================================
# PACKING LISTS
# =========================================================


def create_packing_list(trip_id, contact_id, template_id=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """SELECT id FROM packing_lists
               WHERE trip_id = ? AND contact_id = ?""",
            (trip_id, contact_id),
        )
        row = c.fetchone()
        if row:
            return row[0]
        c.execute(
            """INSERT INTO packing_lists (trip_id, contact_id, template_id)
               VALUES (?, ?, ?)""",
            (trip_id, contact_id, template_id),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_packing_list(trip_id, contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM packing_lists
               WHERE trip_id = ? AND contact_id = ?""",
            (trip_id, contact_id),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_packing_lists_for_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM packing_lists WHERE trip_id = ?", (trip_id,)
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def delete_packing_list(list_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM packing_lists WHERE id = ?", (list_id,))
        conn.commit()
    finally:
        conn.close()


def add_packing_item(list_id, item_name, category=None, sort_order=0):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO packing_items
               (list_id, category, item_name, sort_order)
               VALUES (?, ?, ?, ?)""",
            (list_id, category, item_name, sort_order),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_packing_items(list_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM packing_items
               WHERE list_id = ?
               ORDER BY sort_order ASC, id ASC""",
            (list_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def toggle_packing_item(item_id, packed):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE packing_items SET packed = ? WHERE id = ?",
            (1 if packed else 0, item_id),
        )
        conn.commit()
    finally:
        conn.close()


def delete_packing_item(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM packing_items WHERE id = ?", (item_id,))
        conn.commit()
    finally:
        conn.close()


def apply_packing_template(trip_id, contact_id, template_id):
    tpl = get_packing_template(template_id)
    if not tpl:
        return 0
    list_id = create_packing_list(trip_id, contact_id, template_id)
    try:
        items_to_add = json.loads(tpl.get("items_json") or "[]")
    except Exception:
        items_to_add = []
    existing = {
        item["item_name"].lower() for item in get_packing_items(list_id)
    }
    added = 0
    for i, name in enumerate(items_to_add):
        if name.lower() in existing:
            continue
        add_packing_item(
            list_id, item_name=name,
            category=tpl.get("category"), sort_order=i,
        )
        added += 1
    return added


# =========================================================
# TRIP CHECKLISTS
# =========================================================


def create_trip_checklist(
    trip_id, name, description=None, template_id=None
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO trip_checklists
               (trip_id, name, description, template_id)
               VALUES (?, ?, ?, ?)""",
            (trip_id, name, description, template_id),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_trip_checklists(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM trip_checklists
               WHERE trip_id = ? ORDER BY created_at ASC""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_trip_checklist(checklist_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM trip_checklists WHERE id = ?", (checklist_id,)
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_trip_checklist(checklist_id, name=None, description=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        if name is not None:
            fields.append("name = ?")
            params.append(name)
        if description is not None:
            fields.append("description = ?")
            params.append(description)
        if not fields:
            return
        params.append(checklist_id)
        c.execute(
            f"UPDATE trip_checklists SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_trip_checklist(checklist_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM trip_checklists WHERE id = ?", (checklist_id,)
        )
        conn.commit()
    finally:
        conn.close()


def add_checklist_item(checklist_id, item_text, sort_order=0):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO trip_checklist_items
               (checklist_id, item_text, sort_order)
               VALUES (?, ?, ?)""",
            (checklist_id, item_text, sort_order),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_checklist_items(checklist_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM trip_checklist_items
               WHERE checklist_id = ?
               ORDER BY sort_order ASC, id ASC""",
            (checklist_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def toggle_checklist_item(item_id, is_done):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "UPDATE trip_checklist_items SET is_done = ? WHERE id = ?",
            (1 if is_done else 0, item_id),
        )
        conn.commit()
    finally:
        conn.close()


def delete_checklist_item(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM trip_checklist_items WHERE id = ?", (item_id,)
        )
        conn.commit()
    finally:
        conn.close()


def apply_checklist_template(trip_id, template_id):
    tpl = get_checklist_template(template_id)
    if not tpl:
        return None
    checklist_id = create_trip_checklist(
        trip_id, name=tpl["name"],
        description=tpl.get("description"), template_id=template_id,
    )
    try:
        items = json.loads(tpl.get("items_json") or "[]")
    except Exception:
        items = []
    for i, text in enumerate(items):
        add_checklist_item(checklist_id, text, sort_order=i)
    return checklist_id


# =========================================================
# HOSPITALS
# =========================================================


def add_hospital(
    city, country, name, address=None, phone=None, maps_url=None,
    notes=None, is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO hospitals
               (city, country, name, address, phone, maps_url, notes,
                is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                city, country, name, address, phone, maps_url, notes,
                1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_hospitals(city=None, country=None, active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if active_only:
            conditions.append("is_active = 1")
        if city:
            conditions.append("city = ?")
            params.append(city)
        if country:
            conditions.append("country = ?")
            params.append(country)
        q = "SELECT * FROM hospitals"
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY country, city, name"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_hospital(hospital_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM hospitals WHERE id = ?", (hospital_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_hospital(hospital_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(hospital_id)
        c.execute(
            f"UPDATE hospitals SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def set_hospital_active(hospital_id, is_active):
    update_hospital(hospital_id, is_active=1 if is_active else 0)


def delete_hospital(hospital_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM hospitals WHERE id = ?", (hospital_id,))
        conn.commit()
    finally:
        conn.close()


def get_hospitals_for_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT DISTINCT h.* FROM hospitals h
               JOIN trip_stops s ON LOWER(s.city) = LOWER(h.city)
               WHERE s.trip_id = ? AND h.is_active = 1
               ORDER BY s.stop_order, h.name""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()

# =========================================================
# EMBASSIES / CONSULATES
# =========================================================


def add_embassy(
    host_country,
    representing_country,
    city=None,
    name=None,
    address=None,
    phone=None,
    email=None,
    website=None,
    maps_url=None,
    hours=None,
    notes=None,
    is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO embassies
               (host_country, city, representing_country, name, address,
                phone, email, website, maps_url, hours, notes, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                host_country,
                city,
                representing_country,
                name,
                address,
                phone,
                email,
                website,
                maps_url,
                hours,
                notes,
                1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_embassies(host_country=None, representing_country=None, active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if active_only:
            conditions.append("is_active = 1")
        if host_country:
            conditions.append("host_country = ?")
            params.append(host_country)
        if representing_country:
            conditions.append("representing_country = ?")
            params.append(representing_country)
        q = "SELECT * FROM embassies"
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY host_country, representing_country, name"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_embassy(embassy_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM embassies WHERE id = ?", (embassy_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_embassy(embassy_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(embassy_id)
        c.execute(
            f"UPDATE embassies SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def set_embassy_active(embassy_id, is_active):
    update_embassy(embassy_id, is_active=1 if is_active else 0)


def delete_embassy(embassy_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM embassies WHERE id = ?", (embassy_id,))
        conn.commit()
    finally:
        conn.close()


def get_embassies_for_trip(trip_id, representing_countries=None):
    """
    Return embassies relevant to a trip:
    - host_country matches one of the trip's stop countries
    - representing_country matches one of the executive's passport countries
    """
    if not representing_countries:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT DISTINCT country FROM trip_stops
               WHERE trip_id = ? AND country IS NOT NULL AND country != ''""",
            (trip_id,),
        )
        host_countries = [row[0] for row in c.fetchall()]
        if not host_countries:
            return []
        h_ph = ",".join(["?"] * len(host_countries))
        r_ph = ",".join(["?"] * len(representing_countries))
        c.execute(
            f"""SELECT * FROM embassies
                WHERE is_active = 1
                  AND host_country IN ({h_ph})
                  AND representing_country IN ({r_ph})
                ORDER BY host_country, representing_country, name""",
            host_countries + representing_countries,
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_executive_passport_countries(exec_id):
    """Return a list of country names from an executive's passports."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "SELECT DISTINCT country FROM executive_passports "
            "WHERE exec_id = ? AND country IS NOT NULL AND country != ''",
            (exec_id,),
        )
        return [row[0] for row in c.fetchall()]
    finally:
        conn.close()


# =========================================================
# EXCHANGE RATES
# =========================================================


def save_exchange_rate(
    rate_date, base_currency, target_currency, rate, source=None
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT OR REPLACE INTO exchange_rates
               (rate_date, base_currency, target_currency, rate, source)
               VALUES (?, ?, ?, ?, ?)""",
            (rate_date, base_currency, target_currency, rate, source),
        )
        conn.commit()
    finally:
        conn.close()


def get_exchange_rate(rate_date, base_currency, target_currency):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM exchange_rates
               WHERE rate_date = ? AND base_currency = ?
                 AND target_currency = ?""",
            (rate_date, base_currency, target_currency),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_nearest_exchange_rate(
    rate_date, base_currency, target_currency, days=7
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM exchange_rates
               WHERE base_currency = ? AND target_currency = ?
                 AND rate_date BETWEEN date(?, ?) AND date(?, ?)
               ORDER BY ABS(julianday(rate_date) - julianday(?))
               LIMIT 1""",
            (
                base_currency, target_currency,
                rate_date, f"-{days} days",
                rate_date, f"+{days} days",
                rate_date,
            ),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


# =========================================================
# MEETINGS
# =========================================================


def add_meeting(
    exec_id,
    title,
    meeting_type=None,
    status="draft",
    meeting_date=None,
    meeting_time=None,
    duration_minutes=None,
    location=None,
    video_link=None,
    calendar_event_url=None,
    agenda=None,
    notes_native=None,
    notes_url=None,
    notes_label=None,
    follow_up_date=None,
    linked_trip_id=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO meetings
               (exec_id, title, meeting_type, status, meeting_date,
                meeting_time, duration_minutes, location, video_link,
                calendar_event_url, agenda, notes_native, notes_url,
                notes_label, follow_up_date, linked_trip_id)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                exec_id,
                title,
                meeting_type,
                status,
                meeting_date,
                meeting_time,
                duration_minutes,
                location,
                video_link,
                calendar_event_url,
                agenda,
                notes_native,
                notes_url,
                notes_label,
                follow_up_date,
                linked_trip_id,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_meeting(meeting_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT m.*, e.name AS executive_name
               FROM meetings m
               LEFT JOIN executives e ON m.exec_id = e.id
               WHERE m.id = ?""",
            (meeting_id,),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_meetings(
    exec_id=None,
    status=None,
    start_date=None,
    end_date=None,
    without_notes=False,
    awaiting_response=False,
    search=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if exec_id:
            conditions.append("m.exec_id = ?")
            params.append(exec_id)
        if status:
            if isinstance(status, (list, tuple)):
                placeholders = ",".join("?" * len(status))
                conditions.append(f"m.status IN ({placeholders})")
                params.extend(status)
            else:
                conditions.append("m.status = ?")
                params.append(status)
        if start_date:
            conditions.append("m.meeting_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("m.meeting_date <= ?")
            params.append(end_date)
        if without_notes:
            conditions.append(
                "((m.notes_url IS NULL OR m.notes_url = '') "
                "AND (m.notes_native IS NULL OR m.notes_native = ''))"
            )
        if awaiting_response:
            conditions.append(
                "m.status IN ('request_sent', 'pending', 'tentative', 'no_response')"
            )
        if search:
            like = f"%{search.lower()}%"
            conditions.append(
                "(LOWER(m.title) LIKE ? OR LOWER(COALESCE(m.agenda,'')) LIKE ?)"
            )
            params.extend([like, like])

        q = """SELECT m.*, e.name AS executive_name
               FROM meetings m
               LEFT JOIN executives e ON m.exec_id = e.id"""
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY m.meeting_date DESC, m.meeting_time DESC, m.id DESC"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_meeting(meeting_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        fields.append("updated_at = CURRENT_TIMESTAMP")
        params.append(meeting_id)
        c.execute(f"UPDATE meetings SET {', '.join(fields)} WHERE id = ?", params)
        conn.commit()
    finally:
        conn.close()


def delete_meeting(meeting_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM meetings WHERE id = ?", (meeting_id,))
        conn.commit()
    finally:
        conn.close()


def find_duplicate_meetings(exec_id, title, meeting_date, exclude_id=None):
    if not exec_id or not title or not meeting_date:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = """SELECT * FROM meetings
               WHERE exec_id = ? AND LOWER(title) = LOWER(?)
                 AND meeting_date = ?"""
        params = [exec_id, title.strip(), meeting_date]
        if exclude_id:
            q += " AND id != ?"
            params.append(exclude_id)
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# --- Attendees ---


def add_meeting_attendee(
    meeting_id,
    contact_id,
    is_internal=0,
    is_vip=0,
    role=None,
    assistant_contact_id=None,
    status="pending",
    response_notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO meeting_attendees
               (meeting_id, contact_id, is_internal, is_vip, role,
                assistant_contact_id, status, response_notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                meeting_id,
                contact_id,
                is_internal,
                is_vip,
                role,
                assistant_contact_id,
                status,
                response_notes,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_meeting_attendees(meeting_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ma.*, c.name AS contact_name, c.email AS contact_email,
                      c.phone AS contact_phone, c.role AS contact_role
               FROM meeting_attendees ma
               LEFT JOIN contacts c ON ma.contact_id = c.id
               WHERE ma.meeting_id = ?
               ORDER BY ma.is_vip DESC, c.name""",
            (meeting_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_meeting_attendee(attendee_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(attendee_id)
        c.execute(
            f"UPDATE meeting_attendees SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def remove_meeting_attendee(attendee_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM meeting_attendees WHERE id = ?", (attendee_id,))
        conn.commit()
    finally:
        conn.close()


# --- Action items ---


def add_meeting_action_item(
    meeting_id,
    description,
    owner_contact_id=None,
    owner_exec_id=None,
    due_date=None,
    status="not_started",
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO meeting_action_items
               (meeting_id, description, owner_contact_id, owner_exec_id,
                due_date, status)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                meeting_id,
                description,
                owner_contact_id,
                owner_exec_id,
                due_date,
                status,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_meeting_action_items(meeting_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ai.*,
                      c.name AS owner_contact_name,
                      e.name AS owner_exec_name
               FROM meeting_action_items ai
               LEFT JOIN contacts c ON ai.owner_contact_id = c.id
               LEFT JOIN executives e ON ai.owner_exec_id = e.id
               WHERE ai.meeting_id = ?
               ORDER BY ai.due_date IS NULL, ai.due_date ASC, ai.id""",
            (meeting_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_meeting_action_item(item_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(item_id)
        c.execute(
            f"UPDATE meeting_action_items SET {', '.join(fields)} " "WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_meeting_action_item(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM meeting_action_items WHERE id = ?", (item_id,))
        conn.commit()
    finally:
        conn.close()


def get_action_items_due(exec_id=None, start_date=None, end_date=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = ["ai.status != 'done'"], []
        if exec_id:
            conditions.append("m.exec_id = ?")
            params.append(exec_id)
        if start_date:
            conditions.append("ai.due_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("ai.due_date <= ?")
            params.append(end_date)

        q = """SELECT ai.*, m.title AS meeting_title,
                      m.id AS meeting_id, m.exec_id,
                      c.name AS owner_contact_name,
                      e.name AS owner_exec_name
               FROM meeting_action_items ai
               JOIN meetings m ON ai.meeting_id = m.id
               LEFT JOIN contacts c ON ai.owner_contact_id = c.id
               LEFT JOIN executives e ON ai.owner_exec_id = e.id
               WHERE """
        q += " AND ".join(conditions)
        q += " ORDER BY ai.due_date IS NULL, ai.due_date ASC"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# --- Decisions ---


def add_meeting_decision(meeting_id, decision, made_by=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO meeting_decisions (meeting_id, decision, made_by)
               VALUES (?, ?, ?)""",
            (meeting_id, decision, made_by),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_meeting_decisions(meeting_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM meeting_decisions
               WHERE meeting_id = ?
               ORDER BY id""",
            (meeting_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def delete_meeting_decision(decision_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM meeting_decisions WHERE id = ?", (decision_id,))
        conn.commit()
    finally:
        conn.close()


# =========================================================
# COMMUNICATIONS
# =========================================================


def add_communication(
    contact_id,
    comm_date,
    channel,
    direction="outbound",
    summary=None,
    next_action=None,
    next_action_date=None,
    exec_id=None,
    linked_to_type=None,
    linked_to_id=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO communications
               (contact_id, exec_id, comm_date, channel, direction,
                summary, next_action, next_action_date,
                linked_to_type, linked_to_id)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                contact_id,
                exec_id,
                comm_date,
                channel,
                direction,
                summary,
                next_action,
                next_action_date,
                linked_to_type,
                linked_to_id,
            ),
        )
        conn.commit()
        comm_id = c.lastrowid
    finally:
        conn.close()

    # Recompute denormalized fields on the contact
    _refresh_contact_communication_summary(contact_id)

    return comm_id


def _refresh_contact_communication_summary(contact_id):
    """
    Recompute last_contacted_at, last_channel, next_action, next_action_date
    on the contacts table from the communications history.
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        # Latest communication
        c.execute(
            """SELECT comm_date, channel FROM communications
               WHERE contact_id = ?
               ORDER BY comm_date DESC, id DESC
               LIMIT 1""",
            (contact_id,),
        )
        last_row = c.fetchone()
        last_date = last_row[0] if last_row else None
        last_channel = last_row[1] if last_row else None

        # Next pending action — earliest next_action_date in the future
        c.execute(
            """SELECT next_action, next_action_date FROM communications
               WHERE contact_id = ?
                 AND next_action IS NOT NULL
                 AND TRIM(next_action) != ''
               ORDER BY next_action_date IS NULL,
                        next_action_date ASC,
                        id DESC
               LIMIT 1""",
            (contact_id,),
        )
        next_row = c.fetchone()
        next_action = next_row[0] if next_row else None
        next_date = next_row[1] if next_row else None

        c.execute(
            """UPDATE contacts
               SET last_contacted_at = ?,
                   last_channel = ?,
                   next_action = ?,
                   next_action_date = ?
               WHERE id = ?""",
            (last_date, last_channel, next_action, next_date, contact_id),
        )
        conn.commit()
    finally:
        conn.close()


def get_communications_for_contact(contact_id, limit=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = """SELECT cm.*, e.name AS executive_name
               FROM communications cm
               LEFT JOIN executives e ON cm.exec_id = e.id
               WHERE cm.contact_id = ?
               ORDER BY cm.comm_date DESC, cm.id DESC"""
        params = [contact_id]
        if limit:
            q += " LIMIT ?"
            params.append(limit)
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_communication(comm_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM communications WHERE id = ?", (comm_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_communication(comm_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(comm_id)
        c.execute(
            f"UPDATE communications SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()

    # Recompute summary if contact_id changed or content changed
    existing = get_communication(comm_id)
    if existing:
        _refresh_contact_communication_summary(existing["contact_id"])


def delete_communication(comm_id):
    existing = get_communication(comm_id)
    if not existing:
        return
    contact_id = existing["contact_id"]
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM communications WHERE id = ?", (comm_id,))
        conn.commit()
    finally:
        conn.close()
    _refresh_contact_communication_summary(contact_id)


def get_contacts_gone_quiet(days=30, exec_id=None, active_only=True):
    """
    Return contacts whose last_contacted_at is older than `days` days,
    or who have never been contacted. Ordered by oldest first.
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        cutoff = (datetime.now().date() - timedelta(days=days)).isoformat()
        conditions = ["(last_contacted_at IS NULL OR last_contacted_at <= ?)"]
        params = [cutoff]
        if active_only:
            conditions.append("is_active = 1")
        q = "SELECT * FROM contacts WHERE " + " AND ".join(conditions)
        q += """ ORDER BY
                    last_contacted_at IS NULL DESC,
                    last_contacted_at ASC,
                    name"""
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_next_actions_due(start_date=None, end_date=None, exec_id=None):
    """
    Return contacts with a next_action_date in the given window.
    If no window given, returns everything with a next_action_date set.
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions = ["next_action IS NOT NULL", "TRIM(next_action) != ''"]
        params = []
        if start_date:
            conditions.append("next_action_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("next_action_date <= ?")
            params.append(end_date)
        q = "SELECT * FROM contacts WHERE " + " AND ".join(conditions)
        q += """ ORDER BY
                    next_action_date IS NULL,
                    next_action_date ASC,
                    name"""
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def count_communications_for_contact(contact_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "SELECT COUNT(*) FROM communications WHERE contact_id = ?",
            (contact_id,),
        )
        return c.fetchone()[0]
    finally:
        conn.close()


# --- Auto-population helpers ---


def auto_log_meeting_completion(meeting_id, exec_id):
    """
    When a meeting is marked 'completed', create one inbound+outbound
    communication entry per attendee (excluding those already logged
    for this meeting via the linked_to fields).

    Returns the number of new entries created.
    """
    meeting = get_meeting(meeting_id)
    if not meeting:
        return 0
    attendees = get_meeting_attendees(meeting_id)
    if not attendees:
        return 0

    summary = f"Attended: {meeting.get('title') or 'Meeting'}"
    comm_date = meeting.get("meeting_date") or datetime.now().date().isoformat()

    created = 0
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        for att in attendees:
            # Skip if we already auto-logged for this meeting + contact
            c.execute(
                """SELECT 1 FROM communications
                   WHERE contact_id = ?
                     AND linked_to_type = 'meeting'
                     AND linked_to_id = ?""",
                (att["contact_id"], meeting_id),
            )
            if c.fetchone():
                continue
            c.execute(
                """INSERT INTO communications
                   (contact_id, exec_id, comm_date, channel, direction,
                    summary, linked_to_type, linked_to_id)
                   VALUES (?, ?, ?, 'in_person', 'outbound', ?, 'meeting', ?)""",
                (att["contact_id"], exec_id, comm_date, summary, meeting_id),
            )
            created += 1
            _refresh_contact_communication_summary(att["contact_id"])
        conn.commit()
    finally:
        conn.close()
    return created

# =========================================================
# EVENTS
# =========================================================


def add_event(
    exec_id,
    name,
    objective=None,
    event_type=None,
    status="planning",
    event_date=None,
    start_time=None,
    end_time=None,
    venue_id=None,
    location_override=None,
    headcount=None,
    budget_total=0,
    currency="USD",
    calendar_event_url=None,
    run_of_show_url=None,
    logistics_url=None,
    notes_native=None,
    debrief_notes=None,
    rsvp_form_url=None,
    linked_trip_id=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO events
               (exec_id, name, objective, event_type, status, event_date,
                start_time, end_time, venue_id, location_override,
                headcount, budget_total, currency, calendar_event_url,
                run_of_show_url, logistics_url, notes_native,
                debrief_notes, rsvp_form_url, linked_trip_id)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                       ?, ?, ?, ?)""",
            (
                exec_id,
                name,
                objective,
                event_type,
                status,
                event_date,
                start_time,
                end_time,
                venue_id,
                location_override,
                headcount,
                budget_total,
                currency,
                calendar_event_url,
                run_of_show_url,
                logistics_url,
                notes_native,
                debrief_notes,
                rsvp_form_url,
                linked_trip_id,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_event(event_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ev.*, e.name AS executive_name,
                      v.name AS venue_name, v.city AS venue_city,
                      v.country AS venue_country
               FROM events ev
               LEFT JOIN executives e ON ev.exec_id = e.id
               LEFT JOIN venues v ON ev.venue_id = v.id
               WHERE ev.id = ?""",
            (event_id,),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_events(
    exec_id=None,
    status=None,
    start_date=None,
    end_date=None,
    search=None,
    upcoming_only=False,
    overdue_rsvps=False,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if exec_id:
            conditions.append("ev.exec_id = ?")
            params.append(exec_id)
        if status:
            if isinstance(status, (list, tuple)):
                placeholders = ",".join("?" * len(status))
                conditions.append(f"ev.status IN ({placeholders})")
                params.extend(status)
            else:
                conditions.append("ev.status = ?")
                params.append(status)
        if start_date:
            conditions.append("ev.event_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("ev.event_date <= ?")
            params.append(end_date)
        if upcoming_only:
            today_iso = datetime.now().date().isoformat()
            conditions.append("ev.event_date >= ?")
            params.append(today_iso)
        if overdue_rsvps:
            # Events within 14 days that still have guests in pending/invited
            today_iso = datetime.now().date().isoformat()
            cutoff = (datetime.now().date() + timedelta(days=14)).isoformat()
            conditions.append("ev.event_date >= ?")
            params.append(today_iso)
            conditions.append("ev.event_date <= ?")
            params.append(cutoff)
            conditions.append("""EXISTS (
                    SELECT 1 FROM event_guests eg
                    WHERE eg.event_id = ev.id
                      AND eg.invite_status IN ('invited','pending')
                )""")
        if search:
            like = f"%{search.lower()}%"
            conditions.append(
                "(LOWER(ev.name) LIKE ? " "OR LOWER(COALESCE(ev.objective,'')) LIKE ?)"
            )
            params.extend([like, like])

        q = """SELECT ev.*, e.name AS executive_name,
                      v.name AS venue_name, v.city AS venue_city,
                      (SELECT COUNT(*) FROM event_guests eg
                       WHERE eg.event_id = ev.id) AS guest_count,
                      (SELECT COUNT(*) FROM event_guests eg
                       WHERE eg.event_id = ev.id
                         AND eg.invite_status = 'accepted') AS accepted_count
               FROM events ev
               LEFT JOIN executives e ON ev.exec_id = e.id
               LEFT JOIN venues v ON ev.venue_id = v.id"""
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY ev.event_date DESC, ev.start_time DESC, ev.id DESC"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_event(event_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        fields.append("updated_at = CURRENT_TIMESTAMP")
        params.append(event_id)
        c.execute(f"UPDATE events SET {', '.join(fields)} WHERE id = ?", params)
        conn.commit()
    finally:
        conn.close()


def delete_event(event_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM events WHERE id = ?", (event_id,))
        conn.commit()
    finally:
        conn.close()


def find_duplicate_events(exec_id, name, event_date, exclude_id=None):
    if not exec_id or not name or not event_date:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = """SELECT * FROM events
               WHERE exec_id = ? AND LOWER(name) = LOWER(?)
                 AND event_date = ?"""
        params = [exec_id, name.strip(), event_date]
        if exclude_id:
            q += " AND id != ?"
            params.append(exclude_id)
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# --- Guests ---


def add_event_guest(
    event_id,
    contact_id=None,
    guest_name=None,
    guest_email=None,
    invite_status="not_invited",
    plus_ones=0,
    plus_one_names=None,
    dietary=None,
    allergy_details=None,
    accessibility=None,
    badge_name=None,
    seating_preference=None,
    rsvp_timestamp=None,
    rsvp_source=None,
    notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO event_guests
               (event_id, contact_id, guest_name, guest_email,
                invite_status, plus_ones, plus_one_names, dietary,
                allergy_details, accessibility, badge_name,
                seating_preference, rsvp_timestamp, rsvp_source, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                event_id,
                contact_id,
                guest_name,
                guest_email,
                invite_status,
                plus_ones,
                plus_one_names,
                dietary,
                allergy_details,
                accessibility,
                badge_name,
                seating_preference,
                rsvp_timestamp,
                rsvp_source,
                notes,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_event_guests(event_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT eg.*, c.name AS contact_name,
                      c.email AS contact_email, c.phone AS contact_phone
               FROM event_guests eg
               LEFT JOIN contacts c ON eg.contact_id = c.id
               WHERE eg.event_id = ?
               ORDER BY eg.invite_status, c.name, eg.guest_name""",
            (event_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_event_guest(guest_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM event_guests WHERE id = ?", (guest_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def find_event_guest_by_email(event_id, email):
    if not event_id or not email:
        return None
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM event_guests
               WHERE event_id = ?
                 AND LOWER(COALESCE(guest_email,'')) = LOWER(?)""",
            (event_id, email.strip()),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def find_event_guest_by_name(event_id, name):
    if not event_id or not name:
        return None
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM event_guests
               WHERE event_id = ?
                 AND LOWER(COALESCE(guest_name,'')) = LOWER(?)""",
            (event_id, name.strip()),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_event_guest(guest_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(guest_id)
        c.execute(
            f"UPDATE event_guests SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_event_guest(guest_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM event_guests WHERE id = ?", (guest_id,))
        conn.commit()
    finally:
        conn.close()


def get_event_guest_status_counts(event_id):
    """Return {'status': count} for an event."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """SELECT invite_status, COUNT(*) FROM event_guests
               WHERE event_id = ?
               GROUP BY invite_status""",
            (event_id,),
        )
        return {row[0] or "not_invited": row[1] for row in c.fetchall()}
    finally:
        conn.close()


def get_events_with_overdue_rsvps(days_ahead=14):
    """
    Return events within `days_ahead` days that still have guests in
    'invited' or 'pending' status.
    """
    today = datetime.now().date()
    cutoff = (today + timedelta(days=days_ahead)).isoformat()
    today_iso = today.isoformat()
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT DISTINCT ev.*, e.name AS executive_name,
                      (SELECT COUNT(*) FROM event_guests eg
                       WHERE eg.event_id = ev.id
                         AND eg.invite_status IN ('invited','pending'))
                          AS pending_count
               FROM events ev
               LEFT JOIN executives e ON ev.exec_id = e.id
               LEFT JOIN event_guests eg ON eg.event_id = ev.id
               WHERE ev.event_date >= ?
                 AND ev.event_date <= ?
                 AND eg.invite_status IN ('invited','pending')
               ORDER BY ev.event_date ASC""",
            (today_iso, cutoff),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# =========================================================
# VENDORS
# =========================================================


def add_vendor(
    name, category=None, contact_id=None, email=None, phone=None,
    website=None, address=None, city=None, country=None, notes=None,
    preferred=0, rating=None, is_active=1,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO vendors
               (name, category, contact_id, email, phone, website,
                address, city, country, notes, preferred, rating, is_active)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                name, category, contact_id, email, phone, website,
                address, city, country, notes, preferred, rating,
                1 if is_active else 0,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_vendors(category=None, active_only=True, preferred_only=False,
                search=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if active_only:
            conditions.append("is_active = 1")
        if category:
            conditions.append("category = ?")
            params.append(category)
        if preferred_only:
            conditions.append("preferred = 1")
        if search:
            like = f"%{search.lower()}%"
            conditions.append(
                "(LOWER(name) LIKE ? OR LOWER(COALESCE(email,'')) LIKE ? "
                "OR LOWER(COALESCE(city,'')) LIKE ?)"
            )
            params.extend([like, like, like])
        q = "SELECT * FROM vendors"
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY preferred DESC, name"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_vendor(vendor_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute("SELECT * FROM vendors WHERE id = ?", (vendor_id,))
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_vendor(vendor_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(vendor_id)
        c.execute(
            f"UPDATE vendors SET {', '.join(fields)} WHERE id = ?", params
        )
        conn.commit()
    finally:
        conn.close()


def delete_vendor(vendor_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM vendors WHERE id = ?", (vendor_id,))
        conn.commit()
    finally:
        conn.close()


def find_duplicate_vendors(name, exclude_id=None):
    if not name:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT * FROM vendors WHERE LOWER(name) = LOWER(?)"
        params = [name.strip()]
        if exclude_id:
            q += " AND id != ?"
            params.append(exclude_id)
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# --- Vendor engagements ---


def add_vendor_engagement(
    event_id, vendor_id, category=None, status="shortlisted",
    quote_amount=None, agreed_amount=None, currency="USD",
    contract_status="not_sent", contract_url=None,
    deposit_amount=None, deposit_due_date=None, deposit_paid_date=None,
    balance_amount=None, balance_due_date=None, balance_paid_date=None,
    payment_terms=None, notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO vendor_engagements
               (event_id, vendor_id, category, status, quote_amount,
                agreed_amount, currency, contract_status, contract_url,
                deposit_amount, deposit_due_date, deposit_paid_date,
                balance_amount, balance_due_date, balance_paid_date,
                payment_terms, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                event_id, vendor_id, category, status, quote_amount,
                agreed_amount, currency, contract_status, contract_url,
                deposit_amount, deposit_due_date, deposit_paid_date,
                balance_amount, balance_due_date, balance_paid_date,
                payment_terms, notes,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_vendor_engagements(event_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ve.*, v.name AS vendor_name, v.email AS vendor_email,
                      v.phone AS vendor_phone, v.category AS vendor_category
               FROM vendor_engagements ve
               JOIN vendors v ON ve.vendor_id = v.id
               WHERE ve.event_id = ?
               ORDER BY ve.status, v.name""",
            (event_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_vendor_engagement(engagement_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ve.*, v.name AS vendor_name
               FROM vendor_engagements ve
               JOIN vendors v ON ve.vendor_id = v.id
               WHERE ve.id = ?""",
            (engagement_id,),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_vendor_engagement(engagement_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        fields.append("updated_at = CURRENT_TIMESTAMP")
        params.append(engagement_id)
        c.execute(
            f"UPDATE vendor_engagements SET {', '.join(fields)} "
            "WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_vendor_engagement(engagement_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM vendor_engagements WHERE id = ?", (engagement_id,)
        )
        conn.commit()
    finally:
        conn.close()


# --- Vendor payments ---


def add_vendor_payment(
    vendor_engagement_id, amount, currency="USD",
    due_date=None, paid_date=None, method=None, reference=None, notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO vendor_payments
               (vendor_engagement_id, amount, currency, due_date,
                paid_date, method, reference, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                vendor_engagement_id, amount, currency, due_date,
                paid_date, method, reference, notes,
            ),
        )
        conn.commit()
        new_id = c.lastrowid
    finally:
        conn.close()

    _recompute_budget_from_engagement(vendor_engagement_id)
    return new_id


def get_vendor_payments(engagement_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM vendor_payments
               WHERE vendor_engagement_id = ?
               ORDER BY due_date IS NULL, due_date ASC, id""",
            (engagement_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_vendor_payment(payment_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(payment_id)
        c.execute(
            f"UPDATE vendor_payments SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()

    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "SELECT vendor_engagement_id FROM vendor_payments WHERE id = ?",
            (payment_id,),
        )
        row = c.fetchone()
        if row:
            _recompute_budget_from_engagement(row[0])
    finally:
        conn.close()


def delete_vendor_payment(payment_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "SELECT vendor_engagement_id FROM vendor_payments WHERE id = ?",
            (payment_id,),
        )
        row = c.fetchone()
        engagement_id = row[0] if row else None
        c.execute("DELETE FROM vendor_payments WHERE id = ?", (payment_id,))
        conn.commit()
    finally:
        conn.close()
    if engagement_id:
        _recompute_budget_from_engagement(engagement_id)


def get_payments_due(start_date=None, end_date=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions = ["vp.paid_date IS NULL", "vp.due_date IS NOT NULL"]
        params = []
        if start_date:
            conditions.append("vp.due_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("vp.due_date <= ?")
            params.append(end_date)
        q = """SELECT vp.*, v.name AS vendor_name,
                      ve.event_id, ev.name AS event_name
               FROM vendor_payments vp
               JOIN vendor_engagements ve ON vp.vendor_engagement_id = ve.id
               JOIN vendors v ON ve.vendor_id = v.id
               JOIN events ev ON ve.event_id = ev.id
               WHERE """
        q += " AND ".join(conditions)
        q += " ORDER BY vp.due_date ASC"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# =========================================================
# EVENT BUDGET
# =========================================================


def add_budget_item(
    event_id, category=None, description=None,
    estimated_amount=0, actual_amount=0, paid_amount=0,
    vendor_engagement_id=None, status="estimated",
    receipt_urls=None, invoice_urls=None, notes=None, sort_order=0,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO event_budget_items
               (event_id, category, description, estimated_amount,
                actual_amount, paid_amount, vendor_engagement_id,
                status, receipt_urls, invoice_urls, notes, sort_order)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                event_id, category, description, estimated_amount,
                actual_amount, paid_amount, vendor_engagement_id,
                status, receipt_urls, invoice_urls, notes, sort_order,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_budget_items(event_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT bi.*, v.name AS vendor_name
               FROM event_budget_items bi
               LEFT JOIN vendor_engagements ve
                    ON bi.vendor_engagement_id = ve.id
               LEFT JOIN vendors v ON ve.vendor_id = v.id
               WHERE bi.event_id = ?
               ORDER BY bi.sort_order ASC, bi.id ASC""",
            (event_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_budget_item(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            "SELECT * FROM event_budget_items WHERE id = ?", (item_id,)
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def update_budget_item(item_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(item_id)
        c.execute(
            f"UPDATE event_budget_items SET {', '.join(fields)} "
            "WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_budget_item(item_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM event_budget_items WHERE id = ?", (item_id,)
        )
        conn.commit()
    finally:
        conn.close()


def get_budget_summary(event_id):
    """Return totals for an event's budget."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """SELECT
                 COALESCE(SUM(estimated_amount), 0),
                 COALESCE(SUM(actual_amount), 0),
                 COALESCE(SUM(paid_amount), 0),
                 COUNT(*)
               FROM event_budget_items
               WHERE event_id = ?""",
            (event_id,),
        )
        row = c.fetchone()
        return {
            "estimated": row[0] or 0,
            "actual": row[1] or 0,
            "paid": row[2] or 0,
            "item_count": row[3] or 0,
        }
    finally:
        conn.close()


def _recompute_budget_from_engagement(engagement_id):
    """
    When a payment is recorded, keep the linked budget line item's
    paid_amount in sync. Also ensure a line item exists if the
    engagement is confirmed with an agreed amount.
    """
    engagement = get_vendor_engagement(engagement_id)
    if not engagement:
        return

    event_id = engagement["event_id"]
    vendor_name = engagement.get("vendor_name") or "Vendor"
    category = engagement.get("category") or "Other"
    agreed = engagement.get("agreed_amount") or 0
    currency = engagement.get("currency") or "USD"

    # Sum payments
    payments = get_vendor_payments(engagement_id)
    paid_sum = sum(
        float(p.get("amount") or 0)
        for p in payments
        if p.get("paid_date")
    )

    # Find the linked line item, if any
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM event_budget_items
               WHERE vendor_engagement_id = ?""",
            (engagement_id,),
        )
        item = c.fetchone()
    finally:
        conn.close()

    if item:
        # Update existing
        new_status = item["status"]
        if paid_sum > 0 and paid_sum >= agreed and agreed > 0:
            new_status = "paid"
        elif agreed > 0:
            new_status = "committed"

        update_budget_item(
            item["id"],
            actual_amount=agreed,
            paid_amount=paid_sum,
            status=new_status,
            description=item["description"] or f"{vendor_name} — {category}",
        )
    else:
        # Auto-create a line item if this is a confirmed engagement
        if engagement.get("status") == "confirmed" and agreed > 0:
            new_status = "paid" if paid_sum >= agreed else "committed"
            add_budget_item(
                event_id=event_id,
                category=category,
                description=f"{vendor_name} — {category}",
                estimated_amount=agreed,
                actual_amount=agreed,
                paid_amount=paid_sum,
                vendor_engagement_id=engagement_id,
                status=new_status,
                sort_order=999,
            )


def recompute_budget_after_engagement_change(engagement_id):
    """Call after confirming a vendor or updating its agreed amount."""
    _recompute_budget_from_engagement(engagement_id)


def get_events_over_budget(threshold=0.9):
    """
    Return events where sum(actual) / budget_total > threshold.
    Events with budget_total = 0 are skipped.
    """
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ev.*, e.name AS executive_name,
                      (SELECT COALESCE(SUM(actual_amount), 0)
                       FROM event_budget_items bi
                       WHERE bi.event_id = ev.id) AS spent_actual,
                      (SELECT COALESCE(SUM(estimated_amount), 0)
                       FROM event_budget_items bi
                       WHERE bi.event_id = ev.id) AS spent_estimated
               FROM events ev
               LEFT JOIN executives e ON ev.exec_id = e.id
               WHERE ev.budget_total > 0
                 AND ev.status NOT IN ('completed', 'cancelled')
               ORDER BY ev.event_date ASC"""
        )
        rows = [dict(r) for r in c.fetchall()]
        return [
            r for r in rows
            if (r.get("spent_actual") or 0) / r["budget_total"]
            >= threshold
        ]
    finally:
        conn.close()


def get_unsigned_contracts_within_days(days=7):
    today = datetime.now().date()
    cutoff = (today + timedelta(days=days)).isoformat()
    today_iso = today.isoformat()
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ve.*, v.name AS vendor_name,
                      ev.name AS event_name, ev.event_date
               FROM vendor_engagements ve
               JOIN vendors v ON ve.vendor_id = v.id
               JOIN events ev ON ve.event_id = ev.id
               WHERE ve.contract_status IN ('not_sent','sent')
                 AND ev.event_date >= ?
                 AND ev.event_date <= ?
               ORDER BY ev.event_date ASC""",
            (today_iso, cutoff),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# =========================================================
# BUDGET TEMPLATES
# =========================================================


BUDGET_TEMPLATE_DEFAULTS = {
    "Client Dinner": [
        ("Venue", "Restaurant reservation", 0),
        ("Catering", "Food & beverage", 0),
        ("Transport", "Car service", 0),
        ("Decor", "Flowers, candles", 0),
        ("Staff", "Server, coordinator", 0),
        ("Misc", "Gratuity, incidentals", 0),
    ],
    "Offsite / Retreat": [
        ("Venue", "Meeting space rental", 0),
        ("Catering", "Breakfast, lunch, snacks", 0),
        ("AV", "Projector, mics, tech support", 0),
        ("Transport", "Shuttles, parking", 0),
        ("Accommodation", "Room block", 0),
        ("Materials", "Printed agendas, name badges", 0),
        ("Staff", "Coordinator, facilitator", 0),
        ("Misc", "Contingency", 0),
    ],
    "Conference / Speaking": [
        ("Travel", "Flights", 0),
        ("Accommodation", "Hotel", 0),
        ("Ground Transport", "Taxis, rental car", 0),
        ("Meals", "Per diem", 0),
        ("Materials", "Printed handouts", 0),
        ("Misc", "Registration fees, incidentals", 0),
    ],
    "Board Meeting": [
        ("Venue", "Meeting room", 0),
        ("Catering", "Coffee, lunch", 0),
        ("AV", "Video conference setup", 0),
        ("Materials", "Board packets (print/digital)", 0),
        ("Staff", "Coordinator", 0),
    ],
}


def add_budget_template(name, description=None, event_type=None,
                        items=None, is_active=1):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO budget_templates
               (name, description, event_type, is_active)
               VALUES (?, ?, ?, ?)""",
            (name, description, event_type, 1 if is_active else 0),
        )
        template_id = c.lastrowid
        for i, item in enumerate(items or []):
            category, description_, default_amount = (
                item[0], item[1], item[2] if len(item) > 2 else 0
            )
            c.execute(
                """INSERT INTO budget_template_items
                   (template_id, category, description, default_amount,
                    sort_order)
                   VALUES (?, ?, ?, ?, ?)""",
                (template_id, category, description_, default_amount, i),
            )
        conn.commit()
        return template_id
    finally:
        conn.close()


def get_budget_templates(active_only=True):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = "SELECT * FROM budget_templates"
        if active_only:
            q += " WHERE is_active = 1"
        q += " ORDER BY name"
        c.execute(q)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_budget_template_items(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM budget_template_items
               WHERE template_id = ?
               ORDER BY sort_order, id""",
            (template_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def delete_budget_template(template_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM budget_templates WHERE id = ?", (template_id,)
        )
        conn.commit()
    finally:
        conn.close()


def apply_budget_template_to_event(event_id, template_id):
    items = get_budget_template_items(template_id)
    created = 0
    for item in items:
        add_budget_item(
            event_id=event_id,
            category=item.get("category"),
            description=item.get("description"),
            estimated_amount=item.get("default_amount") or 0,
            status="estimated",
            sort_order=item.get("sort_order") or 0,
        )
        created += 1
    return created


def seed_default_budget_templates():
    """Create the built-in templates if they don't exist yet."""
    existing = {t["name"] for t in get_budget_templates(active_only=False)}
    for name, items in BUDGET_TEMPLATE_DEFAULTS.items():
        if name not in existing:
            add_budget_template(
                name=name,
                description=f"Default template: {name}",
                items=items,
            )


# =========================================================
# SPEAKING ENGAGEMENTS
# =========================================================


def add_speaking_engagement(
    exec_id, title, organizer_contact_id=None, organizer_name=None,
    engagement_type=None, format=None, status="invited",
    engagement_date=None, engagement_time=None, duration=None,
    location=None, timezone=None, audience=None, topic=None,
    fee_offered=None, fee_agreed=None, currency="USD",
    travel_covered=0, accommodation_covered=0,
    recording_rights=None, exclusivity=None,
    calendar_event_url=None, notes_url=None, notes_native=None,
    linked_trip_id=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO speaking_engagements
               (exec_id, title, organizer_contact_id, organizer_name,
                engagement_type, format, status, engagement_date,
                engagement_time, duration, location, timezone, audience,
                topic, fee_offered, fee_agreed, currency, travel_covered,
                accommodation_covered, recording_rights, exclusivity,
                calendar_event_url, notes_url, notes_native, linked_trip_id)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                       ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                exec_id, title, organizer_contact_id, organizer_name,
                engagement_type, format, status, engagement_date,
                engagement_time, duration, location, timezone, audience,
                topic, fee_offered, fee_agreed, currency,
                1 if travel_covered else 0,
                1 if accommodation_covered else 0,
                recording_rights, exclusivity, calendar_event_url,
                notes_url, notes_native, linked_trip_id,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_speaking_engagement(engagement_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT se.*, e.name AS executive_name,
                      c.name AS organizer_contact_name,
                      c.email AS organizer_contact_email
               FROM speaking_engagements se
               LEFT JOIN executives e ON se.exec_id = e.id
               LEFT JOIN contacts c ON se.organizer_contact_id = c.id
               WHERE se.id = ?""",
            (engagement_id,),
        )
        row = c.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_speaking_engagements(
    exec_id=None, status=None, start_date=None, end_date=None,
    search=None, upcoming_only=False,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions, params = [], []
        if exec_id:
            conditions.append("se.exec_id = ?")
            params.append(exec_id)
        if status:
            if isinstance(status, (list, tuple)):
                placeholders = ",".join("?" * len(status))
                conditions.append(f"se.status IN ({placeholders})")
                params.extend(status)
            else:
                conditions.append("se.status = ?")
                params.append(status)
        if start_date:
            conditions.append("se.engagement_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("se.engagement_date <= ?")
            params.append(end_date)
        if upcoming_only:
            today_iso = datetime.now().date().isoformat()
            conditions.append("se.engagement_date >= ?")
            params.append(today_iso)
            conditions.append(
                "se.status NOT IN ('closed', 'cancelled', 'declined')"
            )
        if search:
            like = f"%{search.lower()}%"
            conditions.append(
                "(LOWER(se.title) LIKE ? "
                "OR LOWER(COALESCE(se.organizer_name,'')) LIKE ? "
                "OR LOWER(COALESCE(se.topic,'')) LIKE ?)"
            )
            params.extend([like, like, like])

        q = """SELECT se.*, e.name AS executive_name
               FROM speaking_engagements se
               LEFT JOIN executives e ON se.exec_id = e.id"""
        if conditions:
            q += " WHERE " + " AND ".join(conditions)
        q += " ORDER BY se.engagement_date DESC, se.id DESC"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_speaking_engagement(engagement_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        fields.append("updated_at = CURRENT_TIMESTAMP")
        params.append(engagement_id)
        c.execute(
            f"UPDATE speaking_engagements SET {', '.join(fields)} "
            "WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_speaking_engagement(engagement_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM speaking_engagements WHERE id = ?",
            (engagement_id,),
        )
        conn.commit()
    finally:
        conn.close()


def find_duplicate_speaking_engagements(exec_id, title, engagement_date,
                                        exclude_id=None):
    if not exec_id or not title or not engagement_date:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = """SELECT * FROM speaking_engagements
               WHERE exec_id = ? AND LOWER(title) = LOWER(?)
                 AND engagement_date = ?"""
        params = [exec_id, title.strip(), engagement_date]
        if exclude_id:
            q += " AND id != ?"
            params.append(exclude_id)
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# --- Deliverables ---


def add_speaking_deliverable(
    engagement_id, name, owner=None, due_date=None,
    status="not_started", file_url=None,
    submitted_to_organizer=0, organizer_feedback=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO speaking_deliverables
               (engagement_id, name, owner, due_date, status, file_url,
                submitted_to_organizer, organizer_feedback)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                engagement_id, name, owner, due_date, status, file_url,
                1 if submitted_to_organizer else 0, organizer_feedback,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_speaking_deliverables(engagement_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM speaking_deliverables
               WHERE engagement_id = ?
               ORDER BY due_date IS NULL, due_date ASC, id""",
            (engagement_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_speaking_deliverable(deliverable_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(deliverable_id)
        c.execute(
            f"UPDATE speaking_deliverables SET {', '.join(fields)} "
            "WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_speaking_deliverable(deliverable_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            "DELETE FROM speaking_deliverables WHERE id = ?",
            (deliverable_id,),
        )
        conn.commit()
    finally:
        conn.close()


def seed_default_deliverables(engagement_id):
    """Create the 5 core deliverables if the engagement has none yet."""
    existing = get_speaking_deliverables(engagement_id)
    if existing:
        return 0
    defaults = [
        ("Speaker bio", None, "not_started"),
        ("Headshot (high-res)", None, "not_started"),
        ("Slides / deck", None, "not_started"),
        ("A/V requirements", None, "not_started"),
        ("Introduction script", None, "not_started"),
    ]
    for name, owner, status in defaults:
        add_speaking_deliverable(
            engagement_id, name=name, owner=owner, status=status
        )
    return len(defaults)


def get_deliverables_due(start_date=None, end_date=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions = ["sd.status != 'done'", "sd.due_date IS NOT NULL"]
        params = []
        if start_date:
            conditions.append("sd.due_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("sd.due_date <= ?")
            params.append(end_date)
        q = """SELECT sd.*, se.title AS engagement_title,
                      se.id AS engagement_id,
                      e.name AS executive_name
               FROM speaking_deliverables sd
               JOIN speaking_engagements se ON sd.engagement_id = se.id
               LEFT JOIN executives e ON se.exec_id = e.id
               WHERE """
        q += " AND ".join(conditions)
        q += " ORDER BY sd.due_date ASC"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# --- CFPs ---


def add_speaking_cfp(
    engagement_id, cfp_url=None, submission_deadline=None,
    submission_date=None, proposal_title=None, abstract=None,
    proposal_doc_url=None, status="draft", decision_date=None, notes=None,
):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO speaking_cfps
               (engagement_id, cfp_url, submission_deadline,
                submission_date, proposal_title, abstract,
                proposal_doc_url, status, decision_date, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                engagement_id, cfp_url, submission_deadline,
                submission_date, proposal_title, abstract,
                proposal_doc_url, status, decision_date, notes,
            ),
        )
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()


def get_speaking_cfps(engagement_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM speaking_cfps
               WHERE engagement_id = ?
               ORDER BY submission_deadline IS NULL,
                        submission_deadline ASC, id""",
            (engagement_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def update_speaking_cfp(cfp_id, **kwargs):
    if not kwargs:
        return
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        fields, params = [], []
        for k, v in kwargs.items():
            fields.append(f"{k} = ?")
            params.append(v)
        params.append(cfp_id)
        c.execute(
            f"UPDATE speaking_cfps SET {', '.join(fields)} WHERE id = ?",
            params,
        )
        conn.commit()
    finally:
        conn.close()


def delete_speaking_cfp(cfp_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute("DELETE FROM speaking_cfps WHERE id = ?", (cfp_id,))
        conn.commit()
    finally:
        conn.close()


def get_cfps_closing_within_days(days=14):
    today = datetime.now().date()
    cutoff = (today + timedelta(days=days)).isoformat()
    today_iso = today.isoformat()
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT cf.*, se.title AS engagement_title,
                      se.id AS engagement_id,
                      e.name AS executive_name
               FROM speaking_cfps cf
               JOIN speaking_engagements se ON cf.engagement_id = se.id
               LEFT JOIN executives e ON se.exec_id = e.id
               WHERE cf.status = 'draft'
                 AND cf.submission_deadline >= ?
                 AND cf.submission_deadline <= ?
               ORDER BY cf.submission_deadline ASC""",
            (today_iso, cutoff),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


# --- Pipeline helpers ---


def get_speaking_pipeline_counts():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """SELECT status, COUNT(*) FROM speaking_engagements
               GROUP BY status"""
        )
        return {row[0] or "invited": row[1] for row in c.fetchall()}
    finally:
        conn.close()


def get_speaking_pipeline_value():
    """Sum of agreed fees for confirmed + prep + ready + delivered."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """SELECT COALESCE(SUM(fee_agreed), 0)
               FROM speaking_engagements
               WHERE status IN ('confirmed', 'prep_in_progress',
                                'ready', 'delivered')
                 AND fee_agreed IS NOT NULL"""
        )
        return c.fetchone()[0] or 0
    finally:
        conn.close()


def get_thank_yous_overdue(days_since_delivery=7):
    """Engagements marked delivered more than N days ago without
    thank_you_sent."""
    today = datetime.now().date()
    cutoff = (today - timedelta(days=days_since_delivery)).isoformat()
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT se.*, e.name AS executive_name
               FROM speaking_engagements se
               LEFT JOIN executives e ON se.exec_id = e.id
               WHERE se.status IN ('delivered', 'follow_up')
                 AND se.thank_you_sent = 0
                 AND se.engagement_date IS NOT NULL
                 AND se.engagement_date <= ?
               ORDER BY se.engagement_date ASC""",
            (cutoff,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()

# =========================================================
# CROSS-MODULE QUERIES
# =========================================================


def get_meetings_for_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT m.*, e.name AS executive_name
               FROM meetings m
               LEFT JOIN executives e ON m.exec_id = e.id
               WHERE m.linked_trip_id = ?
               ORDER BY m.meeting_date ASC, m.meeting_time ASC""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_events_for_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT ev.*, e.name AS executive_name
               FROM events ev
               LEFT JOIN executives e ON ev.exec_id = e.id
               WHERE ev.linked_trip_id = ?
               ORDER BY ev.event_date ASC, ev.start_time ASC""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_speaking_for_trip(trip_id):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT se.*, e.name AS executive_name
               FROM speaking_engagements se
               LEFT JOIN executives e ON se.exec_id = e.id
               WHERE se.linked_trip_id = ?
               ORDER BY se.engagement_date ASC""",
            (trip_id,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_trips_overlapping_dates(start_date, end_date, exec_id=None):
    """
    Return trips whose [start_date, end_date] overlaps the given window.
    Used by auto-suggest in meeting/event/engagement forms.
    """
    if not start_date or not end_date:
        return []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        conditions = ["t.start_date <= ?", "t.end_date >= ?"]
        params = [end_date, start_date]
        if exec_id:
            conditions.append("t.exec_id = ?")
            params.append(exec_id)
        q = """SELECT t.*, e.name AS executive_name
               FROM trips t
               LEFT JOIN executives e ON t.exec_id = e.id
               WHERE """ + " AND ".join(conditions)
        q += " ORDER BY t.start_date ASC"
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()


def get_upcoming_activities(start_date, end_date, exec_id=None):
    """
    Return a unified list of activities (meetings, events, speaking
    engagements) in the given date window. Each item has a common shape:
      {kind, id, title, date, time, exec_id, executive_name, status, extra}
    """
    items = []
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()

        # Meetings
        q = """SELECT m.id, m.title, m.meeting_date AS date,
                      m.meeting_time AS time, m.status,
                      m.exec_id, e.name AS executive_name,
                      m.location, m.linked_trip_id
               FROM meetings m
               LEFT JOIN executives e ON m.exec_id = e.id
               WHERE m.meeting_date >= ? AND m.meeting_date <= ?"""
        params = [start_date, end_date]
        if exec_id:
            q += " AND m.exec_id = ?"
            params.append(exec_id)
        c.execute(q, params)
        for r in c.fetchall():
            row = dict(r)
            row["kind"] = "meeting"
            row["extra"] = row.get("location") or ""
            items.append(row)

        # Events
        q = """SELECT ev.id, ev.name AS title, ev.event_date AS date,
                      ev.start_time AS time, ev.status,
                      ev.exec_id, e.name AS executive_name,
                      ev.linked_trip_id
               FROM events ev
               LEFT JOIN executives e ON ev.exec_id = e.id
               WHERE ev.event_date >= ? AND ev.event_date <= ?"""
        params = [start_date, end_date]
        if exec_id:
            q += " AND ev.exec_id = ?"
            params.append(exec_id)
        c.execute(q, params)
        for r in c.fetchall():
            row = dict(r)
            row["kind"] = "event"
            row["extra"] = ""
            items.append(row)

        # Speaking engagements
        q = """SELECT se.id, se.title, se.engagement_date AS date,
                      se.engagement_time AS time, se.status,
                      se.exec_id, e.name AS executive_name,
                      se.location, se.linked_trip_id
               FROM speaking_engagements se
               LEFT JOIN executives e ON se.exec_id = e.id
               WHERE se.engagement_date >= ?
                 AND se.engagement_date <= ?"""
        params = [start_date, end_date]
        if exec_id:
            q += " AND se.exec_id = ?"
            params.append(exec_id)
        c.execute(q, params)
        for r in c.fetchall():
            row = dict(r)
            row["kind"] = "speaking"
            row["extra"] = row.get("location") or ""
            items.append(row)
    finally:
        conn.close()

    # Sort by date, then time
    def _sort_key(it):
        return (
            it.get("date") or "9999-12-31",
            it.get("time") or "99:99",
        )
    items.sort(key=_sort_key)
    return items


def get_upcoming_trips(start_date, end_date, exec_id=None):
    """Return trips starting in the given window, shaped like activities."""
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        q = """SELECT t.id, t.purpose AS title, t.start_date AS date,
                      t.destination, t.exec_id,
                      e.name AS executive_name, t.status
               FROM trips t
               LEFT JOIN executives e ON t.exec_id = e.id
               WHERE t.start_date >= ? AND t.start_date <= ?"""
        params = [start_date, end_date]
        if exec_id:
            q += " AND t.exec_id = ?"
            params.append(exec_id)
        q += " ORDER BY t.start_date ASC"
        c.execute(q, params)
        items = []
        for r in c.fetchall():
            row = dict(r)
            row["kind"] = "trip"
            row["time"] = None
            row["extra"] = row.get("destination") or ""
            items.append(row)
        return items
    finally:
        conn.close()


def global_search(query, limit_per_kind=10):
    """
    Search across meetings, events, speaking engagements, trips, contacts,
    and vendors. Returns dict of kind -> list of results.
    """
    if not query or not query.strip():
        return {}
    like = f"%{query.lower()}%"
    results = {}
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()

        c.execute(
            """SELECT id, title, meeting_date AS date, status
               FROM meetings
               WHERE LOWER(title) LIKE ?
                  OR LOWER(COALESCE(agenda,'')) LIKE ?
                  OR LOWER(COALESCE(notes_native,'')) LIKE ?
               ORDER BY meeting_date DESC
               LIMIT ?""",
            (like, like, like, limit_per_kind),
        )
        results["meetings"] = [dict(r) for r in c.fetchall()]

        c.execute(
            """SELECT id, name AS title, event_date AS date, status
               FROM events
               WHERE LOWER(name) LIKE ?
                  OR LOWER(COALESCE(objective,'')) LIKE ?
                  OR LOWER(COALESCE(notes_native,'')) LIKE ?
               ORDER BY event_date DESC
               LIMIT ?""",
            (like, like, like, limit_per_kind),
        )
        results["events"] = [dict(r) for r in c.fetchall()]

        c.execute(
            """SELECT id, title, engagement_date AS date, status
               FROM speaking_engagements
               WHERE LOWER(title) LIKE ?
                  OR LOWER(COALESCE(topic,'')) LIKE ?
                  OR LOWER(COALESCE(organizer_name,'')) LIKE ?
               ORDER BY engagement_date DESC
               LIMIT ?""",
            (like, like, like, limit_per_kind),
        )
        results["speaking"] = [dict(r) for r in c.fetchall()]

        c.execute(
            """SELECT id, purpose AS title, start_date AS date, status
               FROM trips
               WHERE LOWER(purpose) LIKE ?
                  OR LOWER(destination) LIKE ?
               ORDER BY start_date DESC
               LIMIT ?""",
            (like, like, limit_per_kind),
        )
        results["trips"] = [dict(r) for r in c.fetchall()]

        c.execute(
            """SELECT id, name AS title, email, role, company_id
               FROM contacts
               WHERE LOWER(name) LIKE ?
                  OR LOWER(COALESCE(email,'')) LIKE ?
                  OR LOWER(COALESCE(role,'')) LIKE ?
               ORDER BY name
               LIMIT ?""",
            (like, like, like, limit_per_kind),
        )
        results["contacts"] = [dict(r) for r in c.fetchall()]

        c.execute(
            """SELECT id, name AS title, category, email
               FROM vendors
               WHERE LOWER(name) LIKE ?
                  OR LOWER(COALESCE(category,'')) LIKE ?
                  OR LOWER(COALESCE(email,'')) LIKE ?
               ORDER BY name
               LIMIT ?""",
            (like, like, like, limit_per_kind),
        )
        results["vendors"] = [dict(r) for r in c.fetchall()]
    finally:
        conn.close()
    return results


# =========================================================
# BACKUP & RESTORE
# =========================================================

import shutil


BACKUP_DIR = ".backups"
BACKUP_KEEP_COUNT = 60


def ensure_backup_dir():
    os.makedirs(BACKUP_DIR, exist_ok=True)


def _latest_auto_backup_age_minutes():
    """Return minutes since the newest auto-backup, or None if none exist."""
    ensure_backup_dir()
    files = [
        os.path.join(BACKUP_DIR, f)
        for f in os.listdir(BACKUP_DIR)
        if f.startswith("ea_command_center_") and f.endswith(".db")
    ]
    if not files:
        return None
    newest = max(files, key=os.path.getmtime)
    age_seconds = datetime.now().timestamp() - os.path.getmtime(newest)
    return age_seconds / 60


def make_auto_backup():
    """
    Copy the current DB to .backups/ with a timestamp.
    Returns the path of the new backup, or None on failure.
    """
    if not os.path.exists(DB_PATH):
        return None
    ensure_backup_dir()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(
        BACKUP_DIR, f"ea_command_center_{timestamp}.db"
    )
    try:
        # Use the SQLite backup API for a consistent copy while
        # the DB may be in use.
        src_conn = sqlite3.connect(DB_PATH, timeout=30)
        dst_conn = sqlite3.connect(dest, timeout=30)
        try:
            src_conn.backup(dst_conn)
        finally:
            dst_conn.close()
            src_conn.close()
        return dest
    except Exception:
        try:
            shutil.copy2(DB_PATH, dest)
            return dest
        except Exception:
            return None


def prune_auto_backups(keep=BACKUP_KEEP_COUNT):
    """Delete the oldest backups beyond `keep`."""
    ensure_backup_dir()
    files = [
        os.path.join(BACKUP_DIR, f)
        for f in os.listdir(BACKUP_DIR)
        if f.startswith("ea_command_center_") and f.endswith(".db")
    ]
    if len(files) <= keep:
        return 0
    files.sort(key=os.path.getmtime)
    to_delete = files[: len(files) - keep]
    deleted = 0
    for f in to_delete:
        try:
            os.remove(f)
            deleted += 1
        except Exception:
            pass
    return deleted


def run_auto_backup_if_due(min_age_minutes=60):
    """
    Create an auto-backup if the newest is older than `min_age_minutes`.
    Returns a tuple (made: bool, path: str or None).
    """
    age = _latest_auto_backup_age_minutes()
    if age is not None and age < min_age_minutes:
        return False, None
    path = make_auto_backup()
    if path:
        prune_auto_backups()
        return True, path
    return False, None


def get_auto_backup_count():
    ensure_backup_dir()
    return len([
        f for f in os.listdir(BACKUP_DIR)
        if f.startswith("ea_command_center_") and f.endswith(".db")
    ])


def log_restore(mode, source_filename, notes=None,
                success=1, error_message=None):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    try:
        c = conn.cursor()
        c.execute(
            """INSERT INTO restore_log
               (mode, source_filename, notes, success, error_message)
               VALUES (?, ?, ?, ?, ?)""",
            (mode, source_filename, notes, success, error_message),
        )
        conn.commit()
    finally:
        conn.close()


def get_restore_log(limit=20):
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        c = conn.cursor()
        c.execute(
            """SELECT * FROM restore_log
               ORDER BY performed_at DESC
               LIMIT ?""",
            (limit,),
        )
        return [dict(r) for r in c.fetchall()]
    finally:
        conn.close()