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


def _service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    raw = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"].strip()
    info = json.loads(raw)
    scopes = ["https://www.googleapis.com/auth/calendar"]
    creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    return build("calendar", "v3", credentials=creds, cache_discovery=False)


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
