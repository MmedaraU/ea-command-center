import requests
import json
import time
import os
from datetime import datetime

CACHE_FILE = "exchange_rates_cache.json"
CACHE_DURATION = 3600  # 1 hour in seconds


# =========================================================
# SUPPORTED CURRENCIES
# =========================================================
# Central source of truth for every currency dropdown in the app.
# Add or remove a code here and the change applies everywhere.

CURRENCIES_SUPPORTED = [
    # --- Major reserve / invoice currencies ---
    "USD",
    "EUR",
    "GBP",
    "JPY",
    "CHF",
    # --- Americas ---
    "CAD",
    "MXN",
    "BRL",
    "ARS",
    "CLP",
    "COP",
    # --- Europe (non-EUR) ---
    "NOK",
    "SEK",
    "DKK",
    "PLN",
    "CZK",
    "HUF",
    "RON",
    "TRY",
    "RUB",
    # --- Middle East ---
    "AED",
    "SAR",
    "QAR",
    "KWD",
    "ILS",
    "EGP",
    # --- Africa ---
    "NGN",
    "ZAR",
    "KES",
    "GHS",
    "MAD",
    # --- Asia-Pacific ---
    "CNY",
    "HKD",
    "SGD",
    "TWD",
    "KRW",
    "INR",
    "PKR",
    "THB",
    "MYR",
    "IDR",
    "PHP",
    "VND",
    # --- Oceania ---
    "AUD",
    "NZD",
]


# Symbol map. Anything not listed here falls back to the code
# itself — that's safer than guessing "$", which is wrong for
# AUD, CAD, HKD, etc.
_CURRENCY_SYMBOLS = {
    "USD": "$",
    "EUR": "€",
    "GBP": "£",
    "JPY": "¥",
    "CHF": "Fr",
    "CAD": "C$",
    "MXN": "MX$",
    "BRL": "R$",
    "ARS": "AR$",
    "CLP": "CLP$",
    "COP": "COL$",
    "NOK": "kr",
    "SEK": "kr",
    "DKK": "kr",
    "PLN": "zł",
    "CZK": "Kč",
    "HUF": "Ft",
    "RON": "lei",
    "TRY": "₺",
    "RUB": "₽",
    "AED": "AED",
    "SAR": "SAR",
    "QAR": "QAR",
    "KWD": "KWD",
    "ILS": "₪",
    "EGP": "EGP",
    "NGN": "₦",
    "ZAR": "R",
    "KES": "KSh",
    "GHS": "₵",
    "MAD": "MAD",
    "CNY": "¥",
    "HKD": "HK$",
    "SGD": "S$",
    "TWD": "NT$",
    "KRW": "₩",
    "INR": "₹",
    "PKR": "₨",
    "THB": "฿",
    "MYR": "RM",
    "IDR": "Rp",
    "PHP": "₱",
    "VND": "₫",
    "AUD": "A$",
    "NZD": "NZ$",
}


def get_currency_symbol(currency):
    """
    Return the symbol for a given currency code.

    Falls back to the code itself for currencies without a mapped
    symbol — that's clearer than returning "$", which would be
    wrong for many currencies.
    """
    if not currency:
        return ""
    return _CURRENCY_SYMBOLS.get(currency.upper(), currency.upper())


# =========================================================
# HTTP CACHE
# =========================================================


def get_exchange_rates(base_currency="USD"):
    """
    Fetch exchange rates with caching. Returns dict of {currency: rate}
    relative to base_currency.
    """
    # 1. Check cache
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            data = json.load(f)
            if time.time() - data["timestamp"] < CACHE_DURATION:
                return data["rates"]

    # 2. Fetch from exchangerate-api.com (free, no API key)
    url = f"https://api.exchangerate-api.com/v4/latest/{base_currency}"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            rates = response.json()["rates"]
            with open(CACHE_FILE, "w") as f:
                json.dump({"timestamp": time.time(), "rates": rates}, f)
            return rates
        else:
            if os.path.exists(CACHE_FILE):
                with open(CACHE_FILE, "r") as f:
                    data = json.load(f)
                    return data["rates"]
            raise Exception(
                f"API returned {response.status_code} and no cache available."
            )
    except Exception as e:
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, "r") as f:
                data = json.load(f)
                return data["rates"]
        raise e


def convert(amount, from_currency, to_currency, rates=None):
    """
    Convert an amount from one currency to another.
    If rates is None, fetches fresh rates.
    """
    if from_currency == to_currency:
        return amount
    if rates is None:
        rates = get_exchange_rates(from_currency)
    return amount * rates.get(to_currency, 1.0)


def get_snapshot_rate(base_currency, foreign_currency):
    """Fetch the current exchange rate to store as a snapshot."""
    if base_currency == foreign_currency:
        return 1.0
    rates = get_exchange_rates(base_currency)
    return rates.get(foreign_currency, 1.0)


# =========================================================
# STAGE 2 — Seed rates into the DB (exchange_rates table)
# =========================================================


def seed_rates_for_date(on_date=None, base_currency="USD"):
    """
    Fetch current rates for all supported currencies and store them
    in the exchange_rates table.

    For each target currency, stores the direct pair (base -> target)
    and the reverse pair (target -> base).

    Parameters
    ----------
    on_date : str, optional
        ISO date (YYYY-MM-DD). Defaults to today.
    base_currency : str
        The base currency for the fetch. Defaults to 'USD'.

    Returns
    -------
    dict
        {'stored': int, 'skipped': int, 'date': str}
    """
    import database as db  # local import to avoid circular dependency

    if on_date is None:
        on_date = datetime.now().date().isoformat()
    if isinstance(on_date, datetime):
        on_date = on_date.date().isoformat()
    if isinstance(on_date, str) and "T" in on_date:
        on_date = on_date.split("T")[0]

    try:
        rates = get_exchange_rates(base_currency)
    except Exception as e:
        return {"stored": 0, "skipped": 0, "date": on_date, "error": str(e)}

    stored = 0
    skipped = 0

    for target in CURRENCIES_SUPPORTED:
        if target == base_currency:
            continue
        rate = rates.get(target)
        if not rate or rate <= 0:
            skipped += 1
            continue

        # Direct pair
        db.save_exchange_rate(on_date, base_currency, target, rate, "api")
        stored += 1

        # Reverse pair (derived) – speeds up lookups in the opposite direction
        db.save_exchange_rate(on_date, target, base_currency, 1.0 / rate, "api_derived")
        stored += 1

    # Always store the self-pair for completeness
    for cur in CURRENCIES_SUPPORTED:
        db.save_exchange_rate(on_date, cur, cur, 1.0, "identity")
        stored += 1

    return {"stored": stored, "skipped": skipped, "date": on_date}


# =========================================================
# STAGE 3 — Conversion core
# =========================================================


def get_rate(from_currency, to_currency, on_date=None, base_currency="USD"):
    """
    Return the exchange rate from `from_currency` to `to_currency` on `on_date`.

    Resolution order:
      1. Direct in DB
      2. Reverse in DB
      3. Cross via base currency
      4. Fetch from API for that date, then retry
      5. Nearest date within ±7 days
      6. Raise ValueError

    Returns a float such that: amount_in_from * rate = amount_in_to.
    """
    import database as db  # local import to avoid circular dependency

    if not from_currency or not to_currency:
        raise ValueError("Both from_currency and to_currency are required.")

    if from_currency == to_currency:
        return 1.0

    # Normalize date
    if on_date is None:
        on_date = datetime.now().date().isoformat()
    if isinstance(on_date, datetime):
        on_date = on_date.date().isoformat()
    if hasattr(on_date, "isoformat") and not isinstance(on_date, str):
        on_date = on_date.isoformat()
    if isinstance(on_date, str) and "T" in on_date:
        on_date = on_date.split("T")[0]

    # ---- 1. Direct lookup ----
    row = db.get_exchange_rate(on_date, from_currency, to_currency)
    if row and row.get("rate"):
        return float(row["rate"])

    # ---- 2. Reverse lookup ----
    row = db.get_exchange_rate(on_date, to_currency, from_currency)
    if row and row.get("rate") and row["rate"] != 0:
        return 1.0 / float(row["rate"])

    # ---- 3. Cross via base currency ----
    if from_currency != base_currency and to_currency != base_currency:
        r1 = db.get_exchange_rate(on_date, from_currency, base_currency)
        r2 = db.get_exchange_rate(on_date, base_currency, to_currency)
        if r1 and r2 and r1.get("rate") and r2.get("rate"):
            return float(r1["rate"]) * float(r2["rate"])

    # ---- 4. Fetch from API for this date, then retry direct + reverse ----
    try:
        _fetch_rates_for_date(on_date, base_currency)
    except Exception:
        pass  # fall through to nearest

    row = db.get_exchange_rate(on_date, from_currency, to_currency)
    if row and row.get("rate"):
        return float(row["rate"])

    row = db.get_exchange_rate(on_date, to_currency, from_currency)
    if row and row.get("rate") and row["rate"] != 0:
        return 1.0 / float(row["rate"])

    # ---- 5. Nearest date within ±7 days ----
    row = db.get_nearest_exchange_rate(on_date, from_currency, to_currency, days=7)
    if row and row.get("rate"):
        return float(row["rate"])

    row = db.get_nearest_exchange_rate(on_date, to_currency, from_currency, days=7)
    if row and row.get("rate") and row["rate"] != 0:
        return 1.0 / float(row["rate"])

    # ---- 6. Give up ----
    raise ValueError(
        f"No exchange rate available for {from_currency} → {to_currency} "
        f"on or around {on_date}."
    )


def convert_amount(
    amount, from_currency, to_currency, on_date=None, base_currency="USD"
):
    """
    Convert `amount` from one currency to another.

    Returns 0.0 if amount is None. Returns the same amount if currencies
    match. Raises ValueError if no rate can be determined.
    """
    if amount is None:
        return 0.0
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        return 0.0

    if not from_currency or not to_currency:
        return amount
    if from_currency == to_currency:
        return amount

    rate = get_rate(from_currency, to_currency, on_date, base_currency)
    return amount * rate


def _fetch_rates_for_date(on_date, base_currency="USD"):
    """
    Fetch and store rates for a specific date.

    Note: The free exchangerate-api.com tier only provides TODAY's rates,
    not historical data. If `on_date` is not today, we fall back to
    fetching today's rates and store them under `on_date` so the app
    remains functional.

    For historically accurate rates, a paid API (or manual entry) is
    needed.
    """
    today = datetime.now().date().isoformat()

    # If the requested date is not today and we already have data, do nothing
    if on_date != today:
        existing = None
        try:
            import database as db

            existing = db.get_exchange_rate(on_date, base_currency, "EUR")
        except Exception:
            existing = None
        if existing:
            return  # already have data for this date

    # Seed rates for the requested date
    seed_rates_for_date(on_date, base_currency)
