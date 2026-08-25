"""Google Calendar helpers for khatna bookings."""

from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from typing import Any


def calendar_configured() -> bool:
    return bool(
        os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        and os.getenv("GOOGLE_CALENDAR_ID", "").strip()
    )


def _service_account_info() -> dict[str, Any]:
    """Parse GOOGLE_SERVICE_ACCOUNT_JSON, tolerating common Vercel paste issues."""
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON is empty")

    # Vercel / shell sometimes wraps the whole value in quotes.
    if (raw.startswith("'") and raw.endswith("'")) or (raw.startswith('"') and raw.endswith('"')):
        raw = raw[1:-1].strip()

    # If someone pasted with literal \n sequences instead of real newlines in the private key.
    if "\\n" in raw and "\n" not in raw[1:50]:
        # Only unescape when it looks like a single-line JSON blob.
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            raw = raw.replace("\\n", "\n")

    try:
        info = json.loads(raw)
    except json.JSONDecodeError:
        # Last resort: collapse accidental line breaks outside the private key.
        compact = " ".join(line.strip() for line in raw.splitlines() if line.strip())
        info = json.loads(compact)

    if not isinstance(info, dict) or info.get("type") != "service_account":
        raise ValueError("JSON is not a Google service account key")
    if not info.get("private_key") or not info.get("client_email"):
        raise ValueError("Service account JSON missing private_key or client_email")
    return info


def _service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    info = _service_account_info()
    scopes = ["https://www.googleapis.com/auth/calendar"]
    creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    return build("calendar", "v3", credentials=creds, cache_discovery=False)



def calendar_status_message() -> str:
    """Verify env vars AND that the app can actually reach the calendar."""
    if not calendar_configured():
        return "Google Calendar not connected yet"
    try:
        calendar_id = os.environ["GOOGLE_CALENDAR_ID"].strip()
        service = _service()
        service.calendars().get(calendarId=calendar_id).execute()
        return "Google Calendar connected"
    except json.JSONDecodeError:
        return "Calendar JSON key invalid — fix GOOGLE_SERVICE_ACCOUNT_JSON"
    except ValueError as exc:
        return f"Calendar key problem — {exc}"
    except Exception as exc:
        text = str(exc).lower()
        status = getattr(getattr(exc, "resp", None), "status", None)
        if status == 404 or "notfound" in text or "404" in text:
            return "Calendar ID not found — check GOOGLE_CALENDAR_ID"
        if status == 403 or "forbidden" in text or "403" in text:
            return "Share calendar with service account (Make changes to events)"
        return "Calendar setup error — check Vercel env vars & share settings"


def _event_body(booking: dict[str, str]) -> dict[str, Any]:
    tz = os.getenv("GOOGLE_CALENDAR_TIMEZONE", "Asia/Kolkata").strip() or "Asia/Kolkata"
    start = datetime.strptime(f"{booking['Date']} {booking['Time']}", "%Y-%m-%d %H:%M")
    end = start + timedelta(hours=1)
    description = (
        f"Baby: {booking.get('BabyName', '')}\n"
        f"Age: {booking.get('Age', '')}\n"
        f"Weight: {booking.get('Weight', '')}\n"
        f"Area: {booking.get('Area', '')}\n"
        f"Mobile: {booking.get('Mobile', '')}\n"
        f"Email: {booking.get('Email', '') or '—'}\n"
        f"Edit token: {booking.get('EditToken', '')}\n"
    )
    return {
        "summary": f"Khatna — {booking.get('BabyName', 'Baby')}",
        "description": description,
        "start": {"dateTime": start.isoformat(), "timeZone": tz},
        "end": {"dateTime": end.isoformat(), "timeZone": tz},
    }


def create_khatna_event(booking: dict[str, str]) -> str | None:
    """Create a calendar event. Returns Google event id (or None)."""
    if not calendar_configured():
        return None
    calendar_id = os.environ["GOOGLE_CALENDAR_ID"].strip()
    service = _service()
    event = (
        service.events()
        .insert(calendarId=calendar_id, body=_event_body(booking))
        .execute()
    )
    return event.get("id")


def update_khatna_event(event_id: str, booking: dict[str, str]) -> str | None:
    """Update an existing calendar event. Returns event id."""
    if not calendar_configured() or not event_id:
        return create_khatna_event(booking)
    calendar_id = os.environ["GOOGLE_CALENDAR_ID"].strip()
    service = _service()
    try:
        event = (
            service.events()
            .update(calendarId=calendar_id, eventId=event_id, body=_event_body(booking))
            .execute()
        )
        return event.get("id") or event_id
    except Exception:
        return create_khatna_event(booking)


def delete_khatna_event(event_id: str) -> None:
    if not calendar_configured() or not event_id:
        return
    calendar_id = os.environ["GOOGLE_CALENDAR_ID"].strip()
    service = _service()
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
    except Exception:
        pass
