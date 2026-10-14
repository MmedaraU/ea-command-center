"""
Weather module using Open-Meteo (https://open-meteo.com).
No API key required. Free for non-commercial use.
"""

import requests
import json
import os
import time
from datetime import datetime, timedelta

CACHE_FILE = "weather_cache.json"
CACHE_DURATION = 3600  # 1 hour


# WMO weather codes → human-readable descriptions
WMO_CODES = {
    0: ("Clear sky", "☀️"),
    1: ("Mainly clear", "🌤️"),
    2: ("Partly cloudy", "⛅"),
    3: ("Overcast", "☁️"),
    45: ("Fog", "🌫️"),
    48: ("Depositing rime fog", "🌫️"),
    51: ("Light drizzle", "🌦️"),
    53: ("Moderate drizzle", "🌦️"),
    55: ("Dense drizzle", "🌧️"),
    61: ("Slight rain", "🌦️"),
    63: ("Moderate rain", "🌧️"),
    65: ("Heavy rain", "🌧️"),
    71: ("Slight snow", "🌨️"),
    73: ("Moderate snow", "🌨️"),
    75: ("Heavy snow", "❄️"),
    77: ("Snow grains", "🌨️"),
    80: ("Slight showers", "🌦️"),
    81: ("Moderate showers", "🌧️"),
    82: ("Violent showers", "⛈️"),
    85: ("Slight snow showers", "🌨️"),
    86: ("Heavy snow showers", "❄️"),
    95: ("Thunderstorm", "⛈️"),
    96: ("Thunderstorm with hail", "⛈️"),
    99: ("Thunderstorm with heavy hail", "⛈️"),
}


def _load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _save_cache(cache):
    try:
        with open(CACHE_FILE, "w") as f:
            json.dump(cache, f)
    except Exception:
        pass


def geocode_city(city_name):
    """
    Convert a city name to (lat, lon, country, admin1, timezone).
    Returns None on failure.
    """
    if not city_name:
        return None
    cache = _load_cache()
    key = f"geo_{city_name.lower()}"
    cached = cache.get(key)
    if cached and (time.time() - cached.get("_ts", 0)) < CACHE_DURATION * 24:
        return cached.get("data")

    try:
        url = "https://geocoding-api.open-meteo.com/v1/search"
        params = {"name": city_name, "count": 1, "language": "en", "format": "json"}
        r = requests.get(url, params=params, timeout=8)
        if r.status_code != 200:
            return None
        data = r.json()
        results = data.get("results") or []
        if not results:
            return None
        first = results[0]
        result = {
            "name": first.get("name"),
            "latitude": first.get("latitude"),
            "longitude": first.get("longitude"),
            "country": first.get("country"),
            "admin1": first.get("admin1"),
            "timezone": first.get("timezone"),
        }
        cache[key] = {"_ts": time.time(), "data": result}
        _save_cache(cache)
        return result
    except Exception:
        return None


def get_weather(city_name):
    """
    Fetch current weather + 5-day forecast for a city.
    Returns a dict with `current` and `daily` keys, or None.
    """
    if not city_name:
        return None

    cache = _load_cache()
    key = f"weather_{city_name.lower()}"
    cached = cache.get(key)
    if cached and (time.time() - cached.get("_ts", 0)) < CACHE_DURATION:
        return cached.get("data")

    geo = geocode_city(city_name)
    if not geo:
        return None

    try:
        url = "https://api.open-meteo.com/v1/forecast"
        params = {
            "latitude": geo["latitude"],
            "longitude": geo["longitude"],
            "current": "temperature_2m,relative_humidity_2m,weather_code,wind_speed_10m",
            "daily": "temperature_2m_max,temperature_2m_min,precipitation_sum,uv_index_max,weather_code",
            "timezone": "auto",
            "forecast_days": 5,
        }
        r = requests.get(url, params=params, timeout=8)
        if r.status_code != 200:
            return None
        data = r.json()

        current = data.get("current") or {}
        daily = data.get("daily") or {}

        # Current
        cur_code = current.get("weather_code", 0)
        cur_desc, cur_icon = WMO_CODES.get(cur_code, ("Unknown", "❓"))

        # Daily forecast (up to 5 days)
        daily_forecast = []
        dates = daily.get("time") or []
        for i, date_str in enumerate(dates):
            d_code = (
                (daily.get("weather_code") or [0])[i]
                if i < len(daily.get("weather_code") or [])
                else 0
            )
            d_desc, d_icon = WMO_CODES.get(d_code, ("Unknown", "❓"))
            daily_forecast.append(
                {
                    "date": date_str,
                    "max": (
                        (daily.get("temperature_2m_max") or [None])[i]
                        if i < len(daily.get("temperature_2m_max") or [])
                        else None
                    ),
                    "min": (
                        (daily.get("temperature_2m_min") or [None])[i]
                        if i < len(daily.get("temperature_2m_min") or [])
                        else None
                    ),
                    "precip": (
                        (daily.get("precipitation_sum") or [None])[i]
                        if i < len(daily.get("precipitation_sum") or [])
                        else None
                    ),
                    "uv": (
                        (daily.get("uv_index_max") or [None])[i]
                        if i < len(daily.get("uv_index_max") or [])
                        else None
                    ),
                    "desc": d_desc,
                    "icon": d_icon,
                }
            )

        result = {
            "location": f"{geo['name']}, {geo.get('country') or ''}".strip(", "),
            "timezone": geo.get("timezone"),
            "current": {
                "temp": current.get("temperature_2m"),
                "humidity": current.get("relative_humidity_2m"),
                "wind": current.get("wind_speed_10m"),
                "desc": cur_desc,
                "icon": cur_icon,
            },
            "daily": daily_forecast,
            "fetched_at": datetime.now().isoformat(),
        }

        cache[key] = {"_ts": time.time(), "data": result}
        _save_cache(cache)
        return result
    except Exception:
        return None


def get_weather_for_range(city_name, start_date, end_date):
    """
    Fetch weather for a specific date range. Handles past, present, and
    future dates.

    Uses:
      - archive-api.open-meteo.com/v1/archive for ranges entirely in the
        deep past (older than ~92 days) — data goes back to 1940.
      - api.open-meteo.com/v1/forecast with past_days + forecast_days for
        recent past, present, and near-future ranges (past_days up to 92,
        forecast_days up to 16).

    Returns a dict with 'daily' list. Includes a 'current' block only when
    the range overlaps today. Returns None if the city can't be geocoded or
    no data is available.
    """
    geo = geocode_city(city_name)
    if not geo:
        return None

    # Normalize date arguments to ISO strings
    if hasattr(start_date, "isoformat") and not isinstance(start_date, str):
        start_date = start_date.isoformat()
    if hasattr(end_date, "isoformat") and not isinstance(end_date, str):
        end_date = end_date.isoformat()
    start_date = str(start_date)[:10]
    end_date = str(end_date)[:10]

    try:
        start = datetime.fromisoformat(start_date).date()
        end = datetime.fromisoformat(end_date).date()
    except Exception:
        return None

    if end < start:
        start, end = end, start
        start_date, end_date = end_date, start_date

    today = datetime.now().date()

    # Cache lookup
    cache = _load_cache()
    key = f"weather_range_{city_name.lower()}_{start_date}_{end_date}"
    cached = cache.get(key)
    if cached:
        age = time.time() - cached.get("_ts", 0)
        # Historical ranges cache for 30 days; anything touching today or
        # the future caches for the normal 1 hour.
        is_historical = end < today - timedelta(days=5)
        ttl = (60 * 60 * 24 * 30) if is_historical else CACHE_DURATION
        if age < ttl:
            return cached.get("data")

    days_back = (today - end).days  # positive if range is in the past

    try:
        if days_back > 92:
            # Entirely in the deep past → archive API
            url = "https://archive-api.open-meteo.com/v1/archive"
            params = {
                "latitude": geo["latitude"],
                "longitude": geo["longitude"],
                "start_date": start_date,
                "end_date": end_date,
                "daily": (
                    "temperature_2m_max,temperature_2m_min,"
                    "precipitation_sum,uv_index_max,weather_code"
                ),
                "timezone": "auto",
            }
        else:
            # Recent past, present, or future → forecast API
            url = "https://api.open-meteo.com/v1/forecast"
            past_days = max(0, min(92, (today - start).days))
            forecast_days = max(1, min(16, (end - today).days + 1))
            params = {
                "latitude": geo["latitude"],
                "longitude": geo["longitude"],
                "daily": (
                    "temperature_2m_max,temperature_2m_min,"
                    "precipitation_sum,uv_index_max,weather_code"
                ),
                "timezone": "auto",
                "past_days": past_days,
                "forecast_days": forecast_days,
            }
            # Only fetch current conditions if the range overlaps today
            if start <= today <= end:
                params["current"] = (
                    "temperature_2m,relative_humidity_2m," "weather_code,wind_speed_10m"
                )

        r = requests.get(url, params=params, timeout=12)
        if r.status_code != 200:
            return None
        data = r.json()
    except Exception:
        return None

    # Parse the daily block
    daily = data.get("daily") or {}
    dates = daily.get("time") or []

    def _at(lst, i):
        lst = lst or []
        return lst[i] if i < len(lst) else None

    daily_forecast = []
    for i, date_str in enumerate(dates):
        if date_str < start_date or date_str > end_date:
            continue
        d_code = _at(daily.get("weather_code"), i) or 0
        d_desc, d_icon = WMO_CODES.get(d_code, ("Unknown", "❓"))
        daily_forecast.append(
            {
                "date": date_str,
                "max": _at(daily.get("temperature_2m_max"), i),
                "min": _at(daily.get("temperature_2m_min"), i),
                "precip": _at(daily.get("precipitation_sum"), i),
                "uv": _at(daily.get("uv_index_max"), i),
                "desc": d_desc,
                "icon": d_icon,
            }
        )

    if not daily_forecast:
        # If the range sits entirely beyond the 16-day forecast window,
        # fall back to the archive API for the same month/day last year,
        # so the user sees typical conditions rather than an empty panel.
        if start > today + timedelta(days=16):
            try:
                ly_start = start.replace(year=start.year - 1)
                ly_end = end.replace(year=end.year - 1)
            except ValueError:
                # Feb 29 with no Feb 29 last year — skip the fallback
                return None
            last = get_weather_for_range(
                city_name,
                ly_start.isoformat(),
                ly_end.isoformat(),
            )
            if last:
                last["is_historical_fallback"] = True
                last["original_start_date"] = start_date
                last["original_end_date"] = end_date
            return last
        return None

    result = {
        "location": f"{geo['name']}, {geo.get('country') or ''}".strip(", "),
        "timezone": geo.get("timezone"),
        "start_date": start_date,
        "end_date": end_date,
        "daily": daily_forecast,
        "fetched_at": datetime.now().isoformat(),
    }

    # Attach current block if the API returned one
    cur = data.get("current") or {}
    if cur:
        cur_code = cur.get("weather_code", 0)
        cur_desc, cur_icon = WMO_CODES.get(cur_code, ("Unknown", "❓"))
        result["current"] = {
            "temp": cur.get("temperature_2m"),
            "humidity": cur.get("relative_humidity_2m"),
            "wind": cur.get("wind_speed_10m"),
            "desc": cur_desc,
            "icon": cur_icon,
        }

    cache[key] = {"_ts": time.time(), "data": result}
    _save_cache(cache)
    return result
