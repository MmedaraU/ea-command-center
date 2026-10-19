"""
Seed fictional test data for EA Command Center.

Usage:
    python seed_test_data.py            # add test data to existing DB
    python seed_test_data.py --reset    # wipe DB first, then add

After running, restart the app or refresh the browser to see the data.
"""

import os
import sys
from datetime import datetime, timedelta

import database as db


DB_PATH = "ea_command_center.db"


def today():
    return datetime.now().date()


def d(offset):
    return (today() + timedelta(days=offset)).isoformat()


def dt(offset, hour=9, minute=0):
    return (datetime.now() + timedelta(days=offset)).replace(
        hour=hour, minute=minute, second=0, microsecond=0
    ).isoformat()


def reset_db():
    for suffix in ("", "-wal", "-shm"):
        p = DB_PATH + suffix
        if os.path.exists(p):
            os.remove(p)
    db.init_db()
    db.seed_default_budget_templates()
    print("  Database reset.\n")


def safe(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs)
    except Exception as e:
        print(f"  ⚠️ {fn.__name__}: {e}")
        return None


def seed():
    print("Seeding test data...\n")

    # =========================================================
    # 1. COMPANIES
    # =========================================================
    print("🏢 Companies")
    northwind = db.add_company(
        "Northwind Capital", "CC-1001",
        "Preferred vendors on file. No alcohol at client meetings.",
    )
    meridian = db.add_company(
        "Meridian Health", "CC-2001",
        "All international travel requires board pre-approval.",
    )
    atlas = db.add_company(
        "Atlas Ventures", "CC-3001",
        "Founders travel business class on trips over 6 hours.",
    )
    print(f"  → Northwind={northwind}, Meridian={meridian}, Atlas={atlas}")

    # =========================================================
    # 2. EXECUTIVES
    # =========================================================
    print("\n👤 Executives")
    sarah = db.add_executive(
        northwind, "Sarah Chen", "sarah.chen@northwind.example",
        "America/New_York", "Aisle", "Pescatarian", "Delta", "Known",
        "Vegetarian",
    )
    marcus = db.add_executive(
        meridian, "Marcus Webb", "marcus.webb@meridian.example",
        "Europe/London", "Window", "", "British Airways", "Known", "",
    )
    priya = db.add_executive(
        atlas, "Priya Raman", "priya.raman@atlas.example",
        "America/Los_Angeles", "Aisle", "Gluten-Free", "United", "Known",
        "Vegetarian",
    )
    print(f"  → Sarah={sarah}, Marcus={marcus}, Priya={priya}")

    db.add_passport(sarah, "United States", "US12345678",
                    expiry_date=d(365 * 3), issued_date=d(-365 * 7))
    db.add_passport(marcus, "United Kingdom", "GB87654321",
                    expiry_date=d(365 * 5), issued_date=d(-365 * 4))
    db.add_passport(priya, "United States", "US98765432",
                    expiry_date=d(180), issued_date=d(-365 * 9))
    db.add_passport(priya, "India", "IN11223344",
                    expiry_date=d(365 * 6), issued_date=d(-365 * 5))
    print("  → 4 passports")

    db.add_membership(sarah, "airline", "Delta SkyMiles", "DL-998877",
                      tier="Platinum", alliance="SkyTeam")
    db.add_membership(sarah, "hotel", "Marriott Bonvoy", "MB-112233",
                      tier="Titanium")
    db.add_membership(marcus, "airline", "British Airways Executive Club",
                      "BA-334455", tier="Gold", alliance="oneworld")
    db.add_membership(priya, "airline", "United MileagePlus", "UA-556677",
                      tier="1K", alliance="Star Alliance")
    print("  → 4 memberships")

    # =========================================================
    # 3. CONTACTS
    # =========================================================
    print("\n👥 Contacts")
    nw_board_chair = db.add_contact(
        northwind, "David Okafor", "Board Chair",
        "+1-212-555-0101", "david.okafor@northwind.example",
        "United States", "New York", tags="board,vip", type="Staff",
    )
    nw_cfo = db.add_contact(
        northwind, "Elena Volkov", "Chief Financial Officer",
        "+1-212-555-0102", "elena.volkov@northwind.example",
        "United States", "New York", type="Staff",
    )
    nw_client_acme = db.add_contact(
        northwind, "James Park", "CEO, Acme Corp",
        "+1-415-555-0201", "james.park@acme.example",
        "United States", "San Francisco", tags="client,acme", type="Partner",
    )
    nw_driver = db.add_contact(
        northwind, "Tomás Reyes", "Executive Driver",
        "+1-212-555-0301", "tomas@nycdrivers.example",
        "United States", "New York", tags="nyc,driver",
        type="Local Support",
    )
    mh_partner = db.add_contact(
        meridian, "Dr. Aisha Bello", "Partner, Meridian Health",
        "+44-20-555-0101", "aisha.bello@meridian.example",
        "United Kingdom", "London", type="Staff",
    )
    mh_singapore = db.add_contact(
        meridian, "Wei Lin", "Regional Director, Singapore",
        "+65-6555-0101", "wei.lin@meridian.example",
        "Singapore", "Singapore", tags="singapore,apac", type="Staff",
    )
    mh_tokyo = db.add_contact(
        meridian, "Yuki Tanaka", "Director, Tokyo",
        "+81-3-5555-0101", "yuki.tanaka@meridian.example",
        "Japan", "Tokyo", tags="tokyo,apac", type="Staff",
    )
    mh_podcast_host = db.add_contact(
        meridian, "Rebecca Hollis", "Host, The Health Signal Podcast",
        "", "rebecca@healthsignal.example",
        "United Kingdom", "London", tags="podcast", type="Partner",
    )
    av_partner = db.add_contact(
        atlas, "Daniel Cho", "General Partner, Atlas Ventures",
        "+1-415-555-0101", "daniel.cho@atlas.example",
        "United States", "San Francisco", type="Staff",
    )
    av_founder_1 = db.add_contact(
        atlas, "Maya Patel", "Founder, Brightline",
        "+1-415-555-0201", "maya@brightline.example",
        "United States", "San Francisco",
        tags="portfolio,brightline", type="Partner",
    )
    av_founder_2 = db.add_contact(
        atlas, "Chris Nakamura", "Founder, Cobalt AI",
        "+1-415-555-0202", "chris@cobaltai.example",
        "United States", "San Francisco",
        tags="portfolio,cobalt", type="Partner",
    )
    av_conference_org = db.add_contact(
        atlas, "Sofia Marchetti", "Program Director, TechCrunch Disrupt",
        "", "sofia@techcrunch.example",
        "United States", "San Francisco",
        tags="conference", type="Partner",
    )
    av_journalist = db.add_contact(
        atlas, "Rachel Kim", "Senior Reporter, The Information",
        "", "rachel@theinformation.example",
        "United States", "New York", tags="press", type="Partner",
    )
    freelance_driver_paris = db.add_contact(
        None, "Pierre Laurent", "Driver, Paris",
        "+33-1-5555-0101", "pierre@parisdrivers.example",
        "France", "Paris", tags="paris,driver", type="Local Support",
    )
    print("  → 14 contacts")

    # =========================================================
    # 4. VENUES
    # =========================================================
    print("\n🏢 Venues")
    v_gramercy = db.add_venue(
        "Gramercy Tavern", "42 E 20th St", "New York", "United States",
        wifi_ssid="GramTav-Guest", wifi_password="welcome123",
        dress_code_notes="Smart Casual",
    )
    v_claridges = db.add_venue(
        "Claridge's", "Brook Street, Mayfair", "London", "United Kingdom",
        wifi_ssid="Claridges", dress_code_notes="Business Formal",
    )
    v_marina_bay = db.add_venue(
        "Marina Bay Sands Expo", "10 Bayfront Ave", "Singapore",
        "Singapore", wifi_ssid="MBS-Exhibitor",
        badge_info="Pick up badge at Level 1",
    )
    v_tokyo_imperial = db.add_venue(
        "Imperial Hotel Tokyo", "1-1-1 Uchisaiwaicho", "Tokyo", "Japan",
        wifi_ssid="Imperial-Guest",
    )
    v_hudson_yards = db.add_venue(
        "The Shed, Hudson Yards", "545 W 30th St", "New York",
        "United States", wifi_ssid="TheShed-Guest",
        dress_code_notes="Cocktail",
    )
    print("  → 5 venues")

    # =========================================================
    # 5. VENDORS
    # =========================================================
    print("\n🍽️ Vendors")
    vendor_caterer = db.add_vendor(
        "Sunset Catering", "Catering",
        email="hello@sunsetcatering.example",
        phone="+1-212-555-0401", city="New York",
        country="United States", rating=5, preferred=1,
    )
    vendor_av = db.add_vendor(
        "Clear Audio Visual", "AV",
        email="bookings@clearaudio.example",
        phone="+1-212-555-0402", city="New York",
        country="United States", rating=4,
    )
    vendor_florist = db.add_vendor(
        "Wild Stem Studio", "Florist",
        email="orders@wildstem.example",
        city="New York", country="United States", rating=5,
        preferred=1,
    )
    vendor_photographer = db.add_vendor(
        "Amir Hassan Photography", "Photographer",
        email="amir@amirhassan.example",
        city="London", country="United Kingdom", rating=5,
    )
    vendor_transport = db.add_vendor(
        "Blacklane London", "Transport",
        email="london@blacklane.example",
        city="London", country="United Kingdom", rating=4,
    )
    print("  → 5 vendors")

    # =========================================================
    # 6. MEETINGS
    # =========================================================
    print("\n📅 Meetings")

    m1 = db.add_meeting(
        exec_id=sarah, title="Q4 Board Prep",
        meeting_type="Board", status="confirmed",
        meeting_date=d(0), meeting_time="14:00", duration_minutes=90,
        location="Northwind HQ — Board Room",
        calendar_event_url="https://calendar.google.com/event?eid=boardprep1",
        agenda="1. Q4 financials\n2. Portfolio review\n3. Succession planning",
        notes_native="David wants a tighter deck. Trim to 20 slides.",
        notes_url="https://docs.google.com/document/d/boardprep_q4",
        notes_label="Board prep notes", follow_up_date=d(3),
    )
    db.add_meeting_attendee(m1, nw_board_chair, is_internal=1, is_vip=1,
                            status="accepted")
    db.add_meeting_attendee(m1, nw_cfo, is_internal=1, status="accepted")
    db.add_meeting_action_item(
        m1, "Send final board deck to David",
        owner_contact_id=nw_cfo, due_date=d(2), status="in_progress",
    )
    db.add_meeting_action_item(
        m1, "Book board dinner at Gramercy Tavern",
        owner_exec_id=sarah, due_date=d(5), status="not_started",
    )
    db.add_meeting_decision(m1, "Approve Q4 dividend payout",
                            made_by="Board")

    m2 = db.add_meeting(
        exec_id=sarah, title="Acme Corp — Q1 Pipeline Review",
        meeting_type="Client", status="request_sent",
        meeting_date=d(1), meeting_time="10:00", duration_minutes=60,
        video_link="https://meet.google.com/abc-defg-hij",
        agenda="Review Q1 deal flow and pricing model.",
    )
    db.add_meeting_attendee(m2, nw_client_acme, is_internal=0, is_vip=1,
                            status="pending")
    db.add_meeting_attendee(m2, nw_cfo, is_internal=1, status="accepted")

    m3 = db.add_meeting(
        exec_id=sarah, title="Weekly 1:1 with Elena",
        meeting_type="Internal", status="completed",
        meeting_date=d(-3), meeting_time="09:00", duration_minutes=30,
        location="Northwind HQ",
        notes_native="Elena flagging cash flow timing for Q1.",
    )
    db.add_meeting_attendee(m3, nw_cfo, is_internal=1, status="accepted")
    db.add_meeting_action_item(
        m3, "Review cash flow projections",
        owner_exec_id=sarah, due_date=d(2), status="not_started",
    )

    m4 = db.add_meeting(
        exec_id=marcus, title="Meridian Partner Sync",
        meeting_type="Internal", status="confirmed",
        meeting_date=d(0), meeting_time="11:00", duration_minutes=60,
        location="Meridian HQ, London",
    )
    db.add_meeting_attendee(m4, mh_partner, is_internal=1, status="accepted")

    m5 = db.add_meeting(
        exec_id=priya, title="Investor Update — Q4",
        meeting_type="Client", status="pending",
        meeting_date=d(7), meeting_time="08:00", duration_minutes=45,
        video_link="https://zoom.us/j/9876543210",
        notes_url="https://docs.google.com/document/d/investor_q4",
    )
    db.add_meeting_attendee(m5, av_partner, is_internal=1, status="accepted")
    db.add_meeting_attendee(m5, av_founder_1, is_internal=0, status="pending")
    db.add_meeting_action_item(
        m5, "Draft investor update narrative",
        owner_exec_id=priya, due_date=d(4), status="not_started",
    )

    m6 = db.add_meeting(
        exec_id=marcus, title="Podcast Recording — The Health Signal",
        meeting_type="VIP", status="no_response",
        meeting_date=d(-1), meeting_time="15:00", duration_minutes=45,
    )
    db.add_meeting_attendee(m6, mh_podcast_host, is_internal=0, is_vip=1,
                            status="no_response")

    # --- NEW: Recurring weekly meeting (parent only) ---
    m7 = db.add_meeting(
        exec_id=sarah,
        title="Weekly Leadership Sync",
        meeting_type="Internal",
        status="confirmed",
        meeting_date=d(2),
        meeting_time="09:00",
        duration_minutes=30,
        video_link="https://meet.google.com/weekly-leadership",
        agenda="Weekly priorities, blockers, and cross-team updates.",
    )
    db.add_meeting_attendee(m7, nw_cfo, is_internal=1, status="accepted")
    db.add_meeting_attendee(m7, nw_board_chair, is_internal=1,
                            status="accepted")
    db.update_meeting(
        m7,
        recurrence="weekly",
        recurrence_end_date=d(60),
    )

    print("  → 7 meetings (1 recurring parent)")

    # =========================================================
    # 7. EVENTS
    # =========================================================
    print("\n🎉 Events")

    e1 = db.add_event(
        exec_id=sarah, name="Acme Client Dinner",
        objective="Close the Q1 deal over dinner",
        event_type="Dinner", status="confirmed",
        event_date=d(14), start_time="19:00", end_time="22:00",
        venue_id=v_gramercy, headcount=12, budget_total=3500,
        currency="USD", rsvp_form_url="https://forms.gle/acmedinner2026",
        calendar_event_url="https://calendar.google.com/event?eid=acmedinner",
        run_of_show_url="https://docs.google.com/document/d/runofshow_acme",
        notes_native="David confirmed. Need dietary for two vegetarians.",
    )
    db.add_event_guest(e1, contact_id=nw_client_acme,
                       invite_status="accepted", plus_ones=1,
                       dietary="Pescatarian", badge_name="James Park")
    db.add_event_guest(e1, contact_id=nw_board_chair,
                       invite_status="accepted", dietary="None")
    db.add_event_guest(e1, contact_id=nw_cfo, invite_status="pending")
    db.add_event_guest(e1, contact_id=av_founder_1,
                       invite_status="waitlist")

    ve1 = db.add_vendor_engagement(
        event_id=e1, vendor_id=vendor_caterer, category="Catering",
        status="confirmed", quote_amount=3200, agreed_amount=3050,
        currency="USD", contract_status="signed",
        contract_url="https://docs.google.com/document/d/catering_contract",
        deposit_amount=1000, deposit_due_date=d(7),
        deposit_paid_date=d(-1), balance_amount=2050,
        balance_due_date=d(14), payment_terms="Net 15",
    )
    if ve1:
        db.recompute_budget_after_engagement_change(ve1)
        db.add_vendor_payment(
            ve1, amount=1000, currency="USD", due_date=d(7),
            paid_date=d(-1), method="bank_transfer",
            reference="SUN-INV-001",
        )
        db.add_vendor_payment(
            ve1, amount=2050, currency="USD", due_date=d(14),
            method="bank_transfer",
        )

    db.add_budget_item(
        e1, category="Venue", description="Gramercy private room",
        estimated_amount=500, actual_amount=500, paid_amount=500,
        status="paid",
    )
    db.add_budget_item(
        e1, category="Decor", description="Floral arrangements",
        estimated_amount=300, actual_amount=0, paid_amount=0,
        status="estimated",
    )

    e2 = db.add_event(
        exec_id=marcus, name="Meridian Board Retreat",
        objective="Annual strategy retreat", event_type="Retreat",
        status="planning", event_date=d(21), venue_id=v_claridges,
        headcount=8, budget_total=12000, currency="GBP",
        notes_native="Dr. Bello handling agenda. Need AV for hybrid "
                     "board members.",
    )
    db.add_event_guest(e2, contact_id=mh_partner,
                       invite_status="accepted", badge_name="Dr. Bello")
    db.add_event_guest(e2, guest_name="Sir Alan Pemberton",
                       guest_email="alan@pemberton.example",
                       invite_status="invited")
    db.add_event_guest(e2, guest_name="Dame Catherine Wells",
                       guest_email="catherine@wells.example",
                       invite_status="pending")

    ve2 = db.add_vendor_engagement(
        event_id=e2, vendor_id=vendor_photographer,
        category="Photographer", status="quoted",
        quote_amount=800, currency="GBP",
    )
    ve3 = db.add_vendor_engagement(
        event_id=e2, vendor_id=vendor_av, category="AV",
        status="confirmed", agreed_amount=1500, currency="GBP",
        contract_status="sent",
    )
    if ve3:
        db.recompute_budget_after_engagement_change(ve3)

    e3 = db.add_event(
        exec_id=priya, name="Brightline Series A Launch",
        objective="Announce the round", event_type="Launch",
        status="confirmed", event_date=d(6), start_time="18:00",
        end_time="21:00", venue_id=v_hudson_yards, headcount=60,
        budget_total=15000, currency="USD",
        rsvp_form_url="https://forms.gle/brightlinelaunch",
    )
    db.add_event_guest(e3, contact_id=av_founder_1,
                       invite_status="accepted", plus_ones=2,
                       badge_name="Maya Patel")
    db.add_event_guest(e3, contact_id=av_partner,
                       invite_status="accepted")
    db.add_event_guest(e3, contact_id=av_journalist,
                       invite_status="pending")
    db.add_event_guest(e3, guest_name="Sam Torres",
                       guest_email="sam@tieronevc.example",
                       invite_status="invited")
    db.add_event_guest(e3, guest_name="Nina Rosales",
                       guest_email="nina@techwatch.example",
                       invite_status="declined")

    ve4 = db.add_vendor_engagement(
        event_id=e3, vendor_id=vendor_caterer, category="Catering",
        status="confirmed", agreed_amount=8500, currency="USD",
        contract_status="signed", deposit_amount=3000,
        deposit_due_date=d(2),
        payment_terms="50% deposit, 50% on day of event",
    )
    if ve4:
        db.recompute_budget_after_engagement_change(ve4)

    db.add_budget_item(
        e3, category="Venue", description="The Shed — main hall",
        estimated_amount=4000, actual_amount=4200, paid_amount=4200,
        status="paid",
    )
    db.add_budget_item(
        e3, category="AV", description="Stage, lighting, sound",
        estimated_amount=2500, actual_amount=2800, paid_amount=1400,
        status="committed",
    )

    # --- NEW: Recurring monthly event (parent only) ---
    e4 = db.add_event(
        exec_id=priya,
        name="Monthly Board Dinner",
        objective="Informal board relationship building",
        event_type="Dinner",
        status="planning",
        event_date=d(30),
        start_time="19:00",
        end_time="22:00",
        headcount=8,
        budget_total=2400,
        currency="USD",
        notes_native="Rotating venue. First dinner at The Shed.",
    )
    db.update_event(
        e4,
        recurrence="monthly",
        recurrence_end_date=d(180),
    )

    print("  → 4 events (1 recurring parent)")

    # =========================================================
    # 8. SPEAKING ENGAGEMENTS
    # =========================================================
    print("\n🎤 Speaking Engagements")

    s1 = db.add_speaking_engagement(
        exec_id=priya, title="TechCrunch Disrupt — Opening Keynote",
        organizer_contact_id=av_conference_org,
        organizer_name="TechCrunch Events",
        engagement_type="Keynote", format="In-Person",
        status="confirmed", engagement_date=d(45),
        engagement_time="09:30", duration="45 min",
        location="Moscone Center, San Francisco",
        audience="2,000+ attendees",
        topic="AI in Enterprise: The Next Wave",
        fee_offered=15000, fee_agreed=15000, currency="USD",
        travel_covered=1, accommodation_covered=1,
        recording_rights="Shared", exclusivity="No competing SF keynotes",
        calendar_event_url="https://calendar.google.com/event?eid=tcdisrupt",
        notes_url="https://docs.google.com/document/d/tcdisrupt_brief",
    )
    db.seed_default_deliverables(s1)
    deliverables = db.get_speaking_deliverables(s1)
    if deliverables and len(deliverables) >= 5:
        db.update_speaking_deliverable(
            deliverables[0]["id"], status="done",
            file_url="https://docs.google.com/document/d/bio_priya",
            submitted_to_organizer=1,
        )
        db.update_speaking_deliverable(
            deliverables[1]["id"], status="done",
            file_url="https://drive.google.com/file/d/headshot_priya",
            submitted_to_organizer=1,
        )
        db.update_speaking_deliverable(
            deliverables[2]["id"], status="in_progress",
            due_date=d(20), owner="Priya",
        )
        db.update_speaking_deliverable(
            deliverables[3]["id"], status="done", submitted_to_organizer=1,
        )
        db.update_speaking_deliverable(
            deliverables[4]["id"], status="not_started",
            due_date=d(30), owner="EA",
        )

    db.add_speaking_cfp(
        s1, cfp_url="https://techcrunch.com/cfp",
        submission_deadline=d(-45), submission_date=d(-50),
        proposal_title="AI in Enterprise: The Next Wave",
        abstract="A 45-minute keynote on how AI adoption is reshaping "
                 "enterprise software procurement.",
        proposal_doc_url="https://docs.google.com/document/d/cfp_ai",
        status="accepted", decision_date=d(-30),
    )

    s2 = db.add_speaking_engagement(
        exec_id=marcus, title="The Health Signal Podcast — Episode 142",
        organizer_contact_id=mh_podcast_host,
        organizer_name="The Health Signal",
        engagement_type="Podcast", format="Virtual",
        status="follow_up", engagement_date=d(-5),
        engagement_time="15:00", duration="45 min", location="Zoom",
        audience="Health industry listeners",
        topic="Value-Based Care in 2026",
        fee_offered=0, fee_agreed=0, currency="GBP",
        recording_rights="Shared", travel_covered=0,
        accommodation_covered=0,
    )
    db.seed_default_deliverables(s2)
    db.update_speaking_engagement(
        s2, thank_you_sent=0, recording_received=0, would_repeat=1,
    )

    s3 = db.add_speaking_engagement(
        exec_id=sarah, title="Davos Women in Finance Panel",
        organizer_name="World Economic Forum",
        engagement_type="Panel", format="In-Person",
        status="under_review", engagement_date=d(120),
        location="Davos, Switzerland",
        audience="Global finance leaders",
        topic="Capital Allocation in Uncertain Times",
        fee_offered=0, fee_agreed=0, currency="USD",
        travel_covered=1, accommodation_covered=1,
    )
    db.add_speaking_cfp(
        s3, cfp_url="https://weforum.org/cfp",
        submission_deadline=d(15), submission_date=d(-10),
        proposal_title="Capital Allocation in Uncertain Times",
        abstract="Panel discussion with four global allocators.",
        status="under_review",
    )

    s4 = db.add_speaking_engagement(
        exec_id=priya, title="SXSW 2027 Fireside Chat",
        organizer_name="SXSW", engagement_type="Fireside",
        format="In-Person", status="invited",
        engagement_date=d(150), location="Austin, TX",
        topic="The Founder-Investor Relationship",
        fee_offered=5000, currency="USD",
    )
    print("  → 4 engagements")

    # =========================================================
    # 9. TRIPS
    # =========================================================
    print("\n✈️ Trips")

    t1 = db.create_or_get_trip(
        sarah, "London → Paris", d(25), d(35),
        "Q1 European Roadshow", base_currency="USD",
    )
    db.update_trip_budget(t1, 25000)
    db.update_trip_departure_details(
        t1, "New York", "New York", "United States",
    )
    db.update_trip_status(t1, "draft")
    db.set_trip_delegation(t1, [nw_cfo])
    db.add_trip_stop(t1, 1, "London", "United Kingdom", "",
                     d(25), d(30), "Client meetings and board calls")
    db.add_trip_stop(t1, 2, "Paris", "France", "",
                     d(30), d(35), "Investor meetings")

    db.add_itinerary_item(
        trip_id=t1, item_type="Flight",
        description="JFK → LHR (BA 178)",
        datetime_start=dt(25, 21, 30), datetime_end=dt(26, 9, 30),
        location="JFK T7 → Heathrow T5", cost=0,
        confirmation_code="BA-178-XXX", notes="", is_confirmed=1,
        cost_currency="USD", timezone="America/New_York",
        cost_date=d(25),
    )
    db.add_itinerary_item(
        trip_id=t1, item_type="Hotel",
        description="The Connaught, London",
        datetime_start=dt(26, 14, 0), datetime_end=dt(30, 11, 0),
        location="Carlos Place, Mayfair", cost=0,
        confirmation_code="CONN-556677", notes="", is_confirmed=1,
        cost_currency="USD", timezone="Europe/London", cost_date=d(26),
    )
    db.add_itinerary_item(
        trip_id=t1, item_type="Meeting",
        description="Client meeting — Barings",
        datetime_start=dt(27, 10, 0), datetime_end=dt(27, 12, 0),
        location="Barings London office", cost=0,
        confirmation_code="", notes="Bring NDA", is_confirmed=1,
        cost_currency="USD", timezone="Europe/London", cost_date=d(27),
    )
    db.add_itinerary_item(
        trip_id=t1, item_type="Transport",
        description="Eurostar LHR → Paris",
        datetime_start=dt(30, 14, 0), datetime_end=dt(30, 17, 30),
        location="St Pancras → Gare du Nord", cost=0,
        confirmation_code="ES-998877", notes="", is_confirmed=1,
        cost_currency="USD", timezone="Europe/London", cost_date=d(30),
    )
    db.add_itinerary_item(
        trip_id=t1, item_type="Hotel",
        description="Le Bristol Paris",
        datetime_start=dt(30, 18, 0), datetime_end=dt(35, 11, 0),
        location="112 Rue du Faubourg Saint-Honoré", cost=0,
        confirmation_code="BRISTOL-112", notes="", is_confirmed=1,
        cost_currency="USD", timezone="Europe/Paris", cost_date=d(30),
    )

    t2 = db.create_or_get_trip(
        marcus, "Singapore → Tokyo", d(10), d(20),
        "APAC Partner Meetings", base_currency="GBP",
    )
    db.update_trip_budget(t2, 18000)
    db.update_trip_departure_details(t2, "London", "", "United Kingdom")
    db.update_trip_status(t2, "approved")
    db.set_trip_delegation(t2, [mh_partner])
    db.add_trip_stop(t2, 1, "Singapore", "Singapore", "",
                     d(10), d(15), "")
    db.add_trip_stop(t2, 2, "Tokyo", "Japan", "", d(15), d(20), "")

    db.add_itinerary_item(
        trip_id=t2, item_type="Flight",
        description="LHR → SIN (SQ 317)",
        datetime_start=dt(10, 22, 5), datetime_end=dt(11, 18, 30),
        location="Heathrow → Changi", cost=0,
        confirmation_code="SQ-317-XXX", notes="", is_confirmed=1,
        cost_currency="GBP", timezone="Europe/London", cost_date=d(10),
    )
    db.add_itinerary_item(
        trip_id=t2, item_type="Hotel",
        description="Marina Bay Sands",
        datetime_start=dt(11, 20, 0), datetime_end=dt(15, 8, 0),
        location="10 Bayfront Ave, Singapore", cost=0,
        confirmation_code="MBS-776655", notes="", is_confirmed=1,
        cost_currency="GBP", timezone="Asia/Singapore", cost_date=d(11),
    )
    db.add_itinerary_item(
        trip_id=t2, item_type="Meeting",
        description="APAC Partner Sync",
        datetime_start=dt(12, 9, 0), datetime_end=dt(12, 17, 0),
        location="Marina Bay Sands Expo, Level 4", cost=0,
        confirmation_code="", notes="", is_confirmed=1,
        cost_currency="GBP", timezone="Asia/Singapore", cost_date=d(12),
    )

    t3 = db.create_or_get_trip(
        priya, "NYC → SF", d(-30), d(-25),
        "Portfolio Review Week", base_currency="USD",
    )
    db.update_trip_budget(t3, 8000)
    db.update_trip_status(t3, "final")
    db.add_trip_stop(t3, 1, "New York", "United States", "",
                     d(-30), d(-28), "")
    db.add_trip_stop(t3, 2, "San Francisco", "United States", "",
                     d(-28), d(-25), "")
    db.add_itinerary_item(
        trip_id=t3, item_type="Flight",
        description="SFO → JFK (UA 2450)",
        datetime_start=dt(-30, 8, 0), datetime_end=dt(-30, 16, 30),
        location="SFO → JFK", cost=0,
        confirmation_code="UA-2450-XXX", notes="", is_confirmed=1,
        cost_currency="USD", timezone="America/Los_Angeles",
        cost_date=d(-30),
    )

    print("  → 3 trips")

    # =========================================================
    # 10. COMMUNICATIONS
    # =========================================================
    print("\n📞 Communications")

    db.add_communication(
        contact_id=nw_board_chair, comm_date=d(-2),
        channel="phone", direction="outbound",
        summary="Called to align on Q4 board agenda",
        next_action="Send updated deck", next_action_date=d(2),
        exec_id=sarah, linked_to_type="standalone",
    )
    db.add_communication(
        contact_id=nw_client_acme, comm_date=d(-15),
        channel="email", direction="outbound",
        summary="Sent proposal draft for review",
        next_action="Follow up on proposal", next_action_date=d(-3),
        exec_id=sarah, linked_to_type="standalone",
    )
    db.add_communication(
        contact_id=mh_podcast_host, comm_date=d(-5),
        channel="email", direction="inbound",
        summary="Rebecca sent recording link",
        exec_id=marcus, linked_to_type="standalone",
    )
    db.add_communication(
        contact_id=mh_tokyo, comm_date=d(-95),
        channel="email", direction="outbound",
        summary="Discussed Q1 Tokyo plans",
        exec_id=marcus, linked_to_type="standalone",
    )
    db.add_communication(
        contact_id=av_conference_org, comm_date=d(-3),
        channel="email", direction="inbound",
        summary="Sofia confirmed keynote slot and A/V requirements",
        next_action="Confirm slides deadline", next_action_date=d(10),
        exec_id=priya, linked_to_type="standalone",
    )
    db.add_communication(
        contact_id=av_founder_1, comm_date=d(-1),
        channel="phone", direction="outbound",
        summary="Coordination call for launch party",
        next_action="Send guest list", next_action_date=d(3),
        exec_id=priya, linked_to_type="standalone",
    )
    db.add_communication(
        contact_id=av_journalist, comm_date=d(-20),
        channel="email", direction="outbound",
        summary="Offered exclusive on the Series A",
        next_action="Confirm exclusive terms", next_action_date=d(-2),
        exec_id=priya, linked_to_type="standalone",
    )
    print("  → 7 communications")

    # =========================================================
    # 11. LIBRARY
    # =========================================================
    print("\n📚 Library")

    safe(db.add_destination_guide, country="United Kingdom",
         language="English", currency="GBP",
         emergency_police="999", emergency_ambulance="999",
         emergency_fire="999", notes="Tipping 10-15% at restaurants.")
    safe(db.add_destination_guide, country="France",
         language="French", currency="EUR",
         emergency_police="17", emergency_ambulance="15",
         emergency_fire="18")
    safe(db.add_destination_guide, country="Singapore",
         language="English", currency="SGD",
         emergency_police="999", emergency_ambulance="995",
         emergency_fire="995")
    safe(db.add_destination_guide, country="Japan",
         language="Japanese", currency="JPY",
         emergency_police="110", emergency_ambulance="119",
         emergency_fire="119")

    safe(db.add_visa_rule, passport_nationality="United States",
         destination_country="United Kingdom",
         visa_required=0, max_stay_days=180)
    safe(db.add_visa_rule, passport_nationality="United States",
         destination_country="France", visa_required=0, max_stay_days=90)
    safe(db.add_visa_rule, passport_nationality="United States",
         destination_country="India", visa_required=1,
         max_stay_days=90, processing_time="5-10 business days")

    safe(db.add_checklist_template,
         name="International Trip — Pre-Departure",
         description="Standard EA checklist",
         items_json=('["Confirm passport validity (>6 months)", '
                     '"Check visa requirements", '
                     '"Book airport transfer", '
                     '"Confirm hotel", '
                     '"Send itinerary to executive", '
                     '"Arrange currency", '
                     '"Download offline maps"]'))

    safe(db.add_packing_template,
         name="Business Travel — 3 Days", category="Clothing",
         items_json=('["2 suits", "3 shirts", "1 casual outfit", '
                     '"Gym clothes", "Laptop + charger", '
                     '"Business cards", "Presentation clicker"]'))

    safe(db.add_hospital, city="London", country="United Kingdom",
         name="The Harley Street Clinic", phone="+44-20-7935-7700")
    safe(db.add_hospital, city="Singapore", country="Singapore",
         name="Mount Elizabeth Hospital", phone="+65-6737-2666")
    safe(db.add_hospital, city="Tokyo", country="Japan",
         name="Tokyo Midtown Medical Center", phone="+81-3-5413-0080")

    safe(db.add_embassy, host_country="United Kingdom",
         representing_country="United States", city="London",
         name="US Embassy London", phone="+44-20-7499-9000")
    safe(db.add_embassy, host_country="France",
         representing_country="United States", city="Paris",
         name="US Embassy Paris", phone="+33-1-43-12-22-22")
    safe(db.add_embassy, host_country="Japan",
         representing_country="India", city="Tokyo",
         name="Embassy of India, Tokyo", phone="+81-3-3262-2391")

    print("  → 4 guides, 3 visa rules, 1 checklist, 1 packing, "
          "3 hospitals, 3 embassies")

    # =========================================================
    # SUMMARY
    # =========================================================
    print("\n" + "=" * 60)
    print("✅ Test data seeded successfully.")
    print("=" * 60)
    print("\nWhat's in the DB now:")
    print("  · 3 companies")
    print("  · 3 executives (Sarah, Marcus, Priya)")
    print("  · 14 contacts (staff, partners, local support, freelance)")
    print("  · 7 meetings (today + past + future + 1 recurring parent)")
    print("  · 4 events (client dinner, board retreat, launch, "
          "1 recurring)")
    print("  · 4 speaking engagements (confirmed, delivered, "
          "under review, invited)")
    print("  · 3 trips (UK roadshow, APAC tour, past portfolio week)")
    print("  · 5 venues")
    print("  · 5 vendors (some preferred)")
    print("  · 7 communications (including overdue follow-ups)")
    print("  · Library: 4 guides, 3 visa rules, 1 checklist, "
          "1 packing, 3 hospitals, 3 embassies")
    print("\nNew feature tests available:")
    print("  · Meeting 'Weekly Leadership Sync' — open it and click "
          "'Generate Occurrences'")
    print("  · Event 'Monthly Board Dinner' — open it and click "
          "'Generate Occurrences'")
    print("  · ICS export buttons on Meetings, Events, Speaking tabs")
    print("\nNext steps:")
    print("  1. Restart the app or refresh the browser")
    print("  2. Follow the test plan in the README")
    print("  3. Try a prep brief on Q4 Board Prep")
    print("  4. Try a mobile export")


def main():
    if "--reset" in sys.argv:
        print("Resetting database...")
        reset_db()
        print()

    if not os.path.exists(DB_PATH):
        print("Initializing database...")
        db.init_db()
        db.seed_default_budget_templates()

    seed()


if __name__ == "__main__":
    main()