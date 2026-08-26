"""Google Calendar helpers for khatna bookings."""

from __future__ import annotations

import base64
import json
import os
import re
from datetime import datetime, timedelta, timezone
from typing import Any


def calendar_configured() -> bool:
    return bool(
        os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON_BASE64", "").strip()
    )


def resolved_calendar_id() -> str:
    """Prefer GOOGLE_CALENDAR_ID; otherwise use the service account's own calendar."""
    explicit = os.getenv("GOOGLE_CALENDAR_ID", "").strip().strip('"').strip("'")
    if explicit:
        return explicit
    return str(_service_account_info().get("client_email", "")).strip()


def _raw_service_account_text() -> str:
    b64 = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON_BASE64", "").strip()
    if b64:
        # Allow accidental whitespace/newlines in the base64 value.
        b64 = re.sub(r"\s+", "", b64)
        return base64.b64decode(b64).decode("utf-8")

    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON is empty")

    # Vercel / shell sometimes wraps the whole value in quotes.
    if (raw.startswith("'") and raw.endswith("'")) or (raw.startswith('"') and raw.endswith('"')):
        raw = raw[1:-1].strip()

    # If the value is base64 instead of raw JSON, decode it.
    if not raw.lstrip().startswith("{"):
        try:
            compact = re.sub(r"\s+", "", raw)
            decoded = base64.b64decode(compact).decode("utf-8")
            if decoded.lstrip().startswith("{"):
                return decoded
        except Exception:
            pass

    return raw


def _service_account_info() -> dict[str, Any]:
    """Parse service-account JSON, tolerating common Vercel paste issues."""
    raw = _raw_service_account_text()

    info = None
    try:
        info = json.loads(raw)
    except json.JSONDecodeError:
        # Common case: private_key newlines were turned into real line breaks
        # inside the JSON string. Repair by escaping those breaks.
        repaired = raw
        if "-----BEGIN PRIVATE KEY-----" in repaired and "\\n" not in repaired:
            repaired = repaired.replace(
                "-----BEGIN PRIVATE KEY-----\n",
                "-----BEGIN PRIVATE KEY-----\\n",
            )
            repaired = repaired.replace(
                "\n-----END PRIVATE KEY-----",
                "\\n-----END PRIVATE KEY-----",
            )
            # Escape remaining newlines that sit inside the private key block.
            def _escape_pem(match: re.Match[str]) -> str:
                return match.group(0).replace("\n", "\\n")

            repaired = re.sub(
                r"-----BEGIN PRIVATE KEY-----.*?-----END PRIVATE KEY-----\\n?",
                _escape_pem,
                repaired,
                flags=re.DOTALL,
            )
        try:
            info = json.loads(repaired)
        except json.JSONDecodeError as exc:
            raise ValueError(
                "JSON key is corrupted in Vercel. Re-add it as one line or use BASE64 "
                "(see doctor desk tip)."
            ) from exc

    if not isinstance(info, dict) or info.get("type") != "service_account":
        raise ValueError("JSON is not a Google service account key")
    if not info.get("private_key") or not info.get("client_email"):
        raise ValueError("Service account JSON missing private_key or client_email")

    # Ensure PEM newlines are real newlines for google-auth.
    pk = info["private_key"]
    if "\\n" in pk and "\n" not in pk:
        info["private_key"] = pk.replace("\\n", "\n")
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
        calendar_id = resolved_calendar_id()
        if not calendar_id:
            return "Calendar key problem — missing client_email"
        service = _service()
        meta = service.calendars().get(calendarId=calendar_id).execute()
        summary = meta.get("summary") or calendar_id
        return f"Connected — writing to {calendar_id} ({summary})"
    except json.JSONDecodeError:
        return "Calendar JSON key invalid — re-paste GOOGLE_SERVICE_ACCOUNT_JSON as one line"
    except ValueError as exc:
        return f"Calendar key problem — {exc}"
    except Exception as exc:
        text = str(exc).lower()
        status = getattr(getattr(exc, "resp", None), "status", None)
        calendar_id = ""
        try:
            calendar_id = resolved_calendar_id()
        except Exception:
            pass
        if status == 404 or "notfound" in text or "404" in text:
            return f"Calendar ID not found ({calendar_id or 'empty'}) — set GOOGLE_CALENDAR_ID to pcdesai02@gmail.com"
        if status == 403 or "forbidden" in text or "403" in text:
            sa = ""
            try:
                sa = str(_service_account_info().get("client_email", ""))
            except Exception:
                sa = "service-account"
            return (
                f"No write access to {calendar_id or 'calendar'}. "
                f"Share that calendar with {sa} as Make changes to events"
            )
        return f"Calendar setup error — {type(exc).__name__}: {exc}"


# India Standard Time is always UTC+05:30 (no DST). Avoid zoneinfo/tzdata on Vercel.
_IST = timezone(timedelta(hours=5, minutes=30))
_DEFAULT_TZ_NAME = "Asia/Kolkata"


def _calendar_timezone_name() -> str:
    raw = os.getenv("GOOGLE_CALENDAR_TIMEZONE", _DEFAULT_TZ_NAME) or _DEFAULT_TZ_NAME
    raw = raw.strip().strip('"').strip("'")
    aliases = {
        "ist": _DEFAULT_TZ_NAME,
        "india": _DEFAULT_TZ_NAME,
        "kolkata": _DEFAULT_TZ_NAME,
        "asia/calcutta": _DEFAULT_TZ_NAME,
        "asia/kolkata": _DEFAULT_TZ_NAME,
        "calcutta": _DEFAULT_TZ_NAME,
        "in": _DEFAULT_TZ_NAME,
    }
    candidate = aliases.get(raw.lower(), raw) or _DEFAULT_TZ_NAME
    # Only Asia/Kolkata is guaranteed valid for this clinic app.
    if candidate != _DEFAULT_TZ_NAME:
        # Still accept other IANA names as labels if user set them, but attach IST offset.
        return candidate if "/" in candidate else _DEFAULT_TZ_NAME
    return _DEFAULT_TZ_NAME


def _as_google_time(dt: datetime) -> dict[str, str]:
    """Build Google Calendar start/end using a fixed +05:30 offset."""
    tz_name = _calendar_timezone_name()
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=_IST)
    else:
        dt = dt.astimezone(_IST)
    # Example: 2026-08-26T10:00:00+05:30
    return {"dateTime": dt.isoformat(), "timeZone": tz_name}


def test_calendar_write() -> str:
    """Insert then delete a short test event. Returns a human status string."""
    if not calendar_configured():
        return "Google Calendar is not configured."
    calendar_id = resolved_calendar_id()
    service = _service()
    tz_name = _calendar_timezone_name()
    start = datetime.now(_IST).replace(microsecond=0) + timedelta(minutes=5)
    end = start + timedelta(minutes=15)
    body = {
        "summary": "Khatna app test (safe to delete)",
        "description": "Temporary test from doctor desk",
        "start": _as_google_time(start),
        "end": _as_google_time(end),
    }
    try:
        event = (
            service.events()
            .insert(calendarId=calendar_id, body=body)
            .execute()
        )
        event_id = event.get("id")
        if event_id:
            try:
                service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
            except Exception:
                pass
        return (
            f"Success — can write to {calendar_id} (timezone {tz_name}). "
            "New patient bookings should appear there."
        )
    except Exception as exc:
        status = getattr(getattr(exc, "resp", None), "status", None)
        sa = ""
        try:
            sa = str(_service_account_info().get("client_email", ""))
        except Exception:
            sa = "service-account"
        if status == 403:
            return (
                f"Write blocked for {calendar_id}. "
                f"In Google Calendar (pcdesai02), share the calendar with {sa} "
                "using permission Make changes to events, then try again."
            )
        if status == 404:
            return (
                f"Calendar not found: {calendar_id}. "
                "Set GOOGLE_CALENDAR_ID=pcdesai02@gmail.com in Vercel and redeploy."
            )
        return f"Write test failed for {calendar_id}: {exc}"


def _event_body(booking: dict[str, str]) -> dict[str, Any]:
    start = datetime.strptime(
        f"{booking['Date']} {booking['Time']}", "%Y-%m-%d %H:%M"
    ).replace(tzinfo=_IST)
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
        "start": _as_google_time(start),
        "end": _as_google_time(end),
    }


def create_khatna_event(booking: dict[str, str]) -> str | None:
    """Create a calendar event. Returns Google event id (or None)."""
    if not calendar_configured():
        return None
    calendar_id = resolved_calendar_id()
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
    calendar_id = resolved_calendar_id()
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
    calendar_id = resolved_calendar_id()
    service = _service()
    try:
        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
    except Exception:
        pass
