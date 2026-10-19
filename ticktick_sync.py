"""
TickTick integration.

Manual bidirectional sync between the local app and TickTick's Open API.

Auth flow:
  1. User registers an app at developer.ticktick.com
  2. Pastes Client ID + Secret into the sidebar
  3. Clicks "Connect" → browser opens TickTick's auth page
  4. TickTick redirects to http://localhost:8502/callback?code=...
  5. Local callback server captures the code
  6. We exchange it for access + refresh tokens
  7. Tokens stored in the SQLite DB

Sync:
  - pull_tasks()      — fetch all tasks from TickTick into local mirror
  - push_task(...)    — create a task in TickTick from a local action item
  - complete_task()   — mark a task complete in TickTick
  - full_sync()       — pull + reconcile
"""

import json
import os
import threading
import time
import urllib.parse as urlparse
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from typing import Optional

import requests

import database as db

TICKTICK_API_BASE = "https://api.ticktick.com/open/v1"
TICKTICK_AUTH_URL = "https://ticktick.com/oauth/authorize"
TICKTICK_TOKEN_URL = "https://ticktick.com/oauth/token"
CALLBACK_PORT = 8502
REDIRECT_URI = f"http://localhost:{CALLBACK_PORT}/callback"

# Rate limits: 100/min, 300/5min. We throttle conservatively.
_RATE_LIMIT_DELAY = 0.7  # seconds between requests
_last_request_time = 0.0
_rate_lock = threading.Lock()


def _throttle():
    global _last_request_time
    with _rate_lock:
        now = time.time()
        elapsed = now - _last_request_time
        if elapsed < _RATE_LIMIT_DELAY:
            time.sleep(_RATE_LIMIT_DELAY - elapsed)
        _last_request_time = time.time()


# =========================================================
# OAUTH CALLBACK SERVER
# =========================================================


# Shared slot: the callback handler stores the auth code here.
# We use a module-level dict because the handler runs in a thread.
_oauth_state = {
    "code": None,
    "error": None,
    "received_at": None,
    "active": False,
}


class _OAuthCallbackHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        # Silence the default access log to avoid cluttering the terminal
        return

    def do_GET(self):
        parsed = urlparse.urlparse(self.path)
        if parsed.path != "/callback":
            self.send_response(404)
            self.end_headers()
            return

        params = urlparse.parse_qs(parsed.query)
        code = params.get("code", [None])[0]
        error = params.get("error", [None])[0]

        _oauth_state["code"] = code
        _oauth_state["error"] = error
        _oauth_state["received_at"] = datetime.now().isoformat()

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        if error:
            msg = f"<h2>❌ Authorization failed</h2><p>{error}</p>"
        elif code:
            msg = (
                "<h2>✅ Connected to TickTick</h2>"
                "<p>You can close this tab and return to the app.</p>"
            )
        else:
            msg = "<h2>⚠️ No code received</h2>"
        self.wfile.write(
            f"<html><body style='font-family:sans-serif;"
            f"padding:40px;text-align:center;'>{msg}</body></html>".encode(
                "utf-8"
            )
        )


def start_oauth_callback_server():
    """
    Start a one-shot HTTP server in a background thread that waits for
    a single /callback request, then shuts down.

    Returns True if the server started, False if port already in use.
    """
    _oauth_state["code"] = None
    _oauth_state["error"] = None
    _oauth_state["received_at"] = None

    try:
        server = HTTPServer(("localhost", CALLBACK_PORT), _OAuthCallbackHandler)
    except OSError:
        return False

    def _serve():
        # Handle a single request, then shut down.
        server.handle_request()
        server.server_close()

    t = threading.Thread(target=_serve, daemon=True)
    t.start()
    return True


def get_oauth_state():
    return dict(_oauth_state)


# =========================================================
# OAUTH HELPERS
# =========================================================


def build_authorize_url(client_id: str) -> str:
    params = {
        "client_id": client_id,
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
        "scope": "tasks:read tasks:write",
        "state": "ea_command_center",
    }
    return f"{TICKTICK_AUTH_URL}?{urlparse.urlencode(params)}"


def exchange_code_for_tokens(client_id, client_secret, code):
    """
    POST to the token endpoint to exchange the auth code for
    access + refresh tokens.
    """
    auth = (client_id, client_secret)
    data = {
        "code": code,
        "grant_type": "authorization_code",
        "redirect_uri": REDIRECT_URI,
        "scope": "tasks:read tasks:write",
    }
    _throttle()
    try:
        r = requests.post(
            TICKTICK_TOKEN_URL,
            data=data,
            auth=auth,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=20,
        )
    except Exception as e:
        return {"error": str(e)}

    if r.status_code != 200:
        return {"error": f"HTTP {r.status_code}: {r.text[:300]}"}

    try:
        payload = r.json()
    except Exception as e:
        return {"error": f"Bad JSON: {e}"}

    access_token = payload.get("access_token")
    refresh_token = payload.get("refresh_token")
    expires_in = payload.get("expires_in") or 15552000  # ~180 days default

    if not access_token:
        return {"error": f"No access token in response: {payload}"}

    expires_at = (
        datetime.now() + timedelta(seconds=int(expires_in))
    ).isoformat()

    db.save_ticktick_config(
        client_id=client_id,
        client_secret=client_secret,
        access_token=access_token,
        refresh_token=refresh_token,
        expires_at=expires_at,
    )
    return {
        "access_token": access_token,
        "refresh_token": refresh_token,
        "expires_at": expires_at,
    }


def refresh_access_token():
    """
    Refresh the access token using the stored refresh_token.
    Returns the new access token or None on failure.
    """
    config = db.get_ticktick_config()
    if not config:
        return None
    rt = config.get("refresh_token")
    if not rt:
        return None

    auth = (config.get("client_id"), config.get("client_secret"))
    data = {
        "refresh_token": rt,
        "grant_type": "refresh_token",
        "scope": "tasks:read tasks:write",
    }
    _throttle()
    try:
        r = requests.post(
            TICKTICK_TOKEN_URL,
            data=data,
            auth=auth,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=20,
        )
    except Exception:
        return None

    if r.status_code != 200:
        return None

    try:
        payload = r.json()
    except Exception:
        return None

    access_token = payload.get("access_token")
    refresh_token = payload.get("refresh_token") or rt
    expires_in = payload.get("expires_in") or 15552000

    if not access_token:
        return None

    expires_at = (
        datetime.now() + timedelta(seconds=int(expires_in))
    ).isoformat()

    db.save_ticktick_config(
        access_token=access_token,
        refresh_token=refresh_token,
        expires_at=expires_at,
    )
    return access_token


def is_connected() -> bool:
    config = db.get_ticktick_config()
    if not config:
        return False
    return bool(config.get("access_token"))


# =========================================================
# API CLIENT
# =========================================================


def _headers():
    config = db.get_ticktick_config()
    if not config or not config.get("access_token"):
        return None
    return {
        "Authorization": f"Bearer {config['access_token']}",
        "Content-Type": "application/json",
    }


def _maybe_refresh():
    """Refresh the access token if it's close to expiry."""
    config = db.get_ticktick_config()
    if not config or not config.get("expires_at"):
        return
    try:
        exp = datetime.fromisoformat(config["expires_at"])
    except Exception:
        return
    if datetime.now() + timedelta(days=3) >= exp:
        refresh_access_token()


def _request(method, path, **kwargs):
    headers = _headers()
    if not headers:
        return None, "Not connected to TickTick."
    _maybe_refresh()
    headers = _headers()
    url = f"{TICKTICK_API_BASE}{path}"
    _throttle()
    try:
        r = requests.request(
            method, url, headers=headers, timeout=20, **kwargs
        )
    except Exception as e:
        return None, str(e)

    if r.status_code == 401:
        # Token expired — try refresh once
        new_token = refresh_access_token()
        if new_token:
            headers = _headers()
            _throttle()
            try:
                r = requests.request(
                    method, url, headers=headers, timeout=20, **kwargs
                )
            except Exception as e:
                return None, str(e)
        else:
            return None, "Auth failed; please reconnect."

    if r.status_code >= 400:
        return None, f"HTTP {r.status_code}: {r.text[:300]}"

    if r.status_code == 204 or not r.content:
        return {}, None

    try:
        return r.json(), None
    except Exception:
        return {}, None


# =========================================================
# PULL
# =========================================================


def pull_projects():
    """Fetch projects and mirror them locally."""
    data, err = _request("GET", "/project")
    if err or data is None:
        return err
    if not isinstance(data, list):
        return "Unexpected response from /project"
    for p in data:
        db.upsert_ticktick_project(
            project_id=p.get("id"),
            name=p.get("name"),
            color=p.get("color"),
            closed=1 if p.get("closed") else 0,
        )
    return None


def pull_tasks():
    """
    Pull all tasks across all projects, plus the Inbox.
    TickTick's /project endpoint does NOT return the Inbox, so we
    fetch it explicitly using the literal project id 'inbox'.
    """
    projects = db.get_ticktick_projects()
    if not projects:
        err = pull_projects()
        if err:
            return err
        projects = db.get_ticktick_projects()

    errors = []

    # Always include the Inbox first
    project_targets = [("Inbox", "inbox")]
    for p in projects:
        project_targets.append((p.get("name") or p["project_id"], p["project_id"]))

    for name, pid in project_targets:
        data, err = _request("GET", f"/project/{pid}/data")
        if err:
            errors.append(f"{name}: {err}")
            continue
        tasks = (data or {}).get("tasks") or []
        for t in tasks:
            _mirror_task(t)

    return "; ".join(errors) if errors else None


def _mirror_task(t):
    """Convert a TickTick task payload into a local row."""
    if not t or not t.get("id"):
        return
    due = t.get("dueDate")
    start = t.get("startDate")
    # TickTick returns ISO 8601 with timezone; store date portion only
    due_date = due[:10] if due else None
    start_date = start[:10] if start else None

    db.upsert_ticktick_task(
        ticktick_id=t["id"],
        project_id=t.get("projectId"),
        title=t.get("title") or "",
        content=t.get("content") or "",
        status=2 if t.get("status") == 2 else 0,
        priority=t.get("priority") or 0,
        due_date=due_date,
        start_date=start_date,
        is_all_day=1 if t.get("isAllDay") else 0,
        tags=json.dumps(t.get("tags") or []),
        modified_time=t.get("modifiedTime"),
        created_time=t.get("createdTime"),
        completed_time=t.get("completedTime"),
    )


# =========================================================
# PUSH
# =========================================================


DEFAULT_PROJECT_NAME = "EA Command Center"


def _find_or_create_default_project():
    """Find or create the 'EA Command Center' project."""
    projects = db.get_ticktick_projects()
    for p in projects:
        if (p.get("name") or "").strip() == DEFAULT_PROJECT_NAME:
            return p["project_id"]

    # Create it
    payload = {"name": DEFAULT_PROJECT_NAME, "kind": "TASK"}
    data, err = _request("POST", "/project", json=payload)
    if err:
        return None
    pid = (data or {}).get("id")
    if pid:
        db.upsert_ticktick_project(
            project_id=pid, name=DEFAULT_PROJECT_NAME
        )
    return pid


def push_action_item(action_item_id):
    """
    Push a meeting action item to TickTick.

    Creates the task under 'EA Command Center', links it back to
    the action item.
    """
    # Load the action item
    conn_items = db.get_action_items_due()  # all open items
    item = next(
        (a for a in conn_items if a["id"] == action_item_id), None
    )
    if not item:
        # Fallback: direct lookup
        import sqlite3
        conn = sqlite3.connect(db.DB_PATH, timeout=30)
        conn.row_factory = sqlite3.Row
        try:
            c = conn.cursor()
            c.execute(
                "SELECT * FROM meeting_action_items WHERE id = ?",
                (action_item_id,),
            )
            row = c.fetchone()
            item = dict(row) if row else None
        finally:
            conn.close()

    if not item:
        return None, "Action item not found."

    # Already pushed?
    existing = db.get_ticktick_tasks_for_link(
        "meeting_action_item", action_item_id
    )
    if existing:
        return existing[0].get("ticktick_id"), "Already pushed."

    project_id = _find_or_create_default_project()
    if not project_id:
        return None, "Could not find or create TickTick project."

    payload = {
        "title": item.get("description") or "Action item",
        "projectId": project_id,
        "priority": 3,  # Medium
    }
    if item.get("due_date"):
        payload["dueDate"] = f'{item["due_date"]}T09:00:00+0000'

    data, err = _request("POST", "/task", json=payload)
    if err:
        return None, err

    ticktick_id = (data or {}).get("id")
    if not ticktick_id:
        return None, "No task ID returned."

    # Mirror locally
    db.upsert_ticktick_task(
        ticktick_id=ticktick_id,
        project_id=project_id,
        title=payload["title"],
        status=0,
        priority=payload["priority"],
        due_date=item.get("due_date"),
        linked_to_type="meeting_action_item",
        linked_to_id=action_item_id,
    )
    # Mark the action item as pushed
    db.update_meeting_action_item(
        action_item_id, pushed_to_ticktick=1, ticktick_id=ticktick_id
    )
    return ticktick_id, None


def push_next_action(contact_id):
    """
    Push a contact's next action to TickTick.
    """
    contact = db.get_contact(contact_id)
    if not contact:
        return None, "Contact not found."
    if not contact.get("next_action"):
        return None, "This contact has no next action."

    existing = db.get_ticktick_tasks_for_link(
        "contact_next_action", contact_id
    )
    if existing:
        return existing[0].get("ticktick_id"), "Already pushed."

    project_id = _find_or_create_default_project()
    if not project_id:
        return None, "Could not find or create TickTick project."

    title = f'{contact.get("next_action")} — {contact.get("name")}'
    payload = {
        "title": title,
        "projectId": project_id,
        "priority": 3,
    }
    if contact.get("next_action_date"):
        payload["dueDate"] = (
            f'{contact["next_action_date"]}T09:00:00+0000'
        )

    data, err = _request("POST", "/task", json=payload)
    if err:
        return None, err
    ticktick_id = (data or {}).get("id")
    if not ticktick_id:
        return None, "No task ID returned."

    db.upsert_ticktick_task(
        ticktick_id=ticktick_id,
        project_id=project_id,
        title=title,
        status=0,
        priority=payload["priority"],
        due_date=contact.get("next_action_date"),
        linked_to_type="contact_next_action",
        linked_to_id=contact_id,
    )
    return ticktick_id, None


def complete_task(ticktick_id):
    """
    Mark a task complete in TickTick. Requires the project_id,
    which we have in the local mirror.
    """
    task = db.get_ticktick_task(ticktick_id)
    if not task:
        return "Task not found in local mirror."
    project_id = task.get("project_id")
    if not project_id:
        return "Missing project_id."

    data, err = _request(
        "POST", f"/project/{project_id}/task/{ticktick_id}/complete"
    )
    if err:
        return err

    # Update local mirror
    db.upsert_ticktick_task(
        ticktick_id=ticktick_id,
        status=2,
        completed_time=datetime.now().isoformat(),
    )

    # If linked to a local item, update that too
    if task.get("linked_to_type") == "meeting_action_item":
        db.update_meeting_action_item(
            task["linked_to_id"], status="done"
        )
    return None


def update_task(ticktick_id, title=None, due_date=None, priority=None):
    """
    Update an existing task in TickTick and sync the local mirror.
    Only the fields passed as non-None will be updated.
    """
    task = db.get_ticktick_task(ticktick_id)
    if not task:
        return "Task not found in local mirror."

    project_id = task.get("project_id")
    if not project_id:
        return "Missing project_id."

    payload = {
        "id": ticktick_id,
        "projectId": project_id,
    }

    if title is not None:
        payload["title"] = title
    if due_date is not None:
        # TickTick expects full ISO 8601 for dueDate
        payload["dueDate"] = f"{due_date}T09:00:00+0000" if due_date else None
    if priority is not None:
        payload["priority"] = int(priority)

    data, err = _request("POST", f"/task/{ticktick_id}", json=payload)
    if err:
        return err

    # Update local mirror
    db.upsert_ticktick_task(
        ticktick_id=ticktick_id,
        title=title if title is not None else task.get("title"),
        due_date=due_date if due_date is not None else task.get("due_date"),
        priority=priority if priority is not None else task.get("priority"),
    )
    return None


def delete_task(ticktick_id):
    task = db.get_ticktick_task(ticktick_id)
    if not task:
        return "Task not found."
    project_id = task.get("project_id")
    data, err = _request(
        "DELETE", f"/project/{project_id}/task/{ticktick_id}"
    )
    if err:
        return err
    db.delete_ticktick_task(ticktick_id)
    return None


# =========================================================
# FULL SYNC
# =========================================================


def full_sync():
    """
    Refresh projects, pull all tasks, and reconcile local action items.

    Returns a dict with counts.
    """
    result = {
        "projects": 0,
        "tasks_pulled": 0,
        "errors": [],
    }
    err = pull_projects()
    if err:
        result["errors"].append(f"Projects: {err}")
    result["projects"] = len(db.get_ticktick_projects())

    err = pull_tasks()
    if err:
        result["errors"].append(f"Tasks: {err}")
    result["tasks_pulled"] = len(db.get_ticktick_tasks(open_only=False))

    db.save_ticktick_config(last_sync_at=datetime.now().isoformat())
    return result
