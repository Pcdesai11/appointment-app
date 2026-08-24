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


def create_khatna_event(booking: dict[str, str]) -> str | None:
    """
    Create a calendar event for the booking.
    Returns event htmlLink or None if calendar is not configured.
    """
    if not calendar_configured():
        return None

    calendar_id = os.environ["GOOGLE_CALENDAR_ID"].strip()
    tz = os.getenv("GOOGLE_CALENDAR_TIMEZONE", "Asia/Kolkata").strip() or "Asia/Kolkata"
    date_str = booking["Date"]
    time_str = booking["Time"]  # HH:MM
    start = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
    end = start + timedelta(hours=1)

    summary = f"Khatna — {booking.get('BabyName', 'Baby')}"
    description = (
        f"Baby: {booking.get('BabyName', '')}\n"
        f"Age: {booking.get('Age', '')}\n"
        f"Weight: {booking.get('Weight', '')}\n"
        f"Area: {booking.get('Area', '')}\n"
        f"Mobile: {booking.get('Mobile', '')}\n"
        f"Email: {booking.get('Email', '') or '—'}\n"
    )

    body: dict[str, Any] = {
        "summary": summary,
        "description": description,
        "start": {"dateTime": start.isoformat(), "timeZone": tz},
        "end": {"dateTime": end.isoformat(), "timeZone": tz},
    }
    mobile = (booking.get("Mobile") or "").strip()
    if mobile:
        body["description"] += f"\nCall: {mobile}"

    service = _service()
    event = (
        service.events()
        .insert(calendarId=calendar_id, body=body)
        .execute()
    )
    return event.get("htmlLink") or event.get("id")
