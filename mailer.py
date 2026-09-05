"""Send appointment confirmation emails via Resend or SMTP."""

from __future__ import annotations

import json
import os
import smtplib
import ssl
import urllib.error
import urllib.request
from email.message import EmailMessage
from typing import Any


def mail_configured() -> bool:
    if os.getenv("RESEND_API_KEY", "").strip():
        return True
    return bool(
        os.getenv("SMTP_HOST", "").strip()
        and os.getenv("SMTP_USER", "").strip()
        and os.getenv("SMTP_PASSWORD", "").strip()
    )


def _from_address() -> str:
    return (
        os.getenv("RESEND_FROM", "").strip()
        or os.getenv("SMTP_FROM", "").strip()
        or os.getenv("SMTP_USER", "").strip()
        or "onboarding@resend.dev"
    )


def _time_label(time_str: str, time_labels: dict[str, str] | None = None) -> str:
    if time_labels and time_str in time_labels:
        label = time_labels[time_str]
        if "·" in label:
            return label.split("·", 1)[0].strip()
        return label
    plain = {"10:00": "10:00 AM", "13:30": "1:30 PM", "17:00": "5:00 PM"}
    return plain.get(time_str, time_str or "—")


def build_confirmation_bodies(
    booking: dict[str, str],
    time_labels: dict[str, str] | None = None,
) -> tuple[str, str]:
    name = booking.get("BabyName") or "your child"
    date = booking.get("Date") or "—"
    time_txt = _time_label(booking.get("Time", ""), time_labels)
    area = booking.get("Area") or "—"
    mobile = booking.get("Mobile") or "—"

    plain = f"""Appointment confirmed / અપોઇન્ટમેન્ટ કન્ફર્મ

Dear parent/guardian,

Your khatna appointment is confirmed.

Baby: {name}
Date: {date}
Time: {time_txt}
Area: {area}
Mobile: {mobile}

Thank you.
"""

    html = f"""<!DOCTYPE html>
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.5; color: #1c2b33;">
  <h2>Appointment confirmed</h2>
  <p>Dear parent/guardian,</p>
  <p>Your khatna appointment is confirmed.</p>
  <table style="border-collapse: collapse; margin: 1rem 0;">
    <tr><td style="padding: 4px 12px 4px 0;"><strong>Baby</strong></td><td>{name}</td></tr>
    <tr><td style="padding: 4px 12px 4px 0;"><strong>Date</strong></td><td>{date}</td></tr>
    <tr><td style="padding: 4px 12px 4px 0;"><strong>Time</strong></td><td>{time_txt}</td></tr>
    <tr><td style="padding: 4px 12px 4px 0;"><strong>Area</strong></td><td>{area}</td></tr>
    <tr><td style="padding: 4px 12px 4px 0;"><strong>Mobile</strong></td><td>{mobile}</td></tr>
  </table>
  <p>Thank you.</p>
</body>
</html>
"""
    return plain, html


def _send_via_resend(
    *,
    api_key: str,
    from_addr: str,
    to_addr: str,
    subject: str,
    plain: str,
    html: str,
) -> None:
    payload = {
        "from": from_addr,
        "to": [to_addr],
        "subject": subject,
        "text": plain,
        "html": html,
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=data,
        method="POST",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Resend HTTP {exc.code}: {detail}") from exc


def _send_via_smtp(
    *,
    from_addr: str,
    to_addr: str,
    subject: str,
    plain: str,
    html: str,
) -> None:
    host = os.getenv("SMTP_HOST", "").strip()
    port = int(os.getenv("SMTP_PORT", "587").strip() or "587")
    user = os.getenv("SMTP_USER", "").strip()
    password = os.getenv("SMTP_PASSWORD", "").strip()
    use_ssl = os.getenv("SMTP_SSL", "").strip().lower() in {"1", "true", "yes"}

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")

    if use_ssl or port == 465:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as server:
            server.login(user, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
            server.login(user, password)
            server.send_message(msg)


def send_confirmation_email(
    booking: dict[str, Any],
    time_labels: dict[str, str] | None = None,
) -> bool:
    """
    Send confirmation if patient email is present and Resend/SMTP is configured.
    Returns True if sent, False if skipped/failed (never raises to caller).
    """
    to_addr = (booking.get("Email") or "").strip()
    if not to_addr or not mail_configured():
        return False

    plain, html = build_confirmation_bodies(booking, time_labels)
    subject = f"Appointment confirmed — {booking.get('BabyName') or 'Khatna'}"
    from_addr = _from_address()
    resend_key = os.getenv("RESEND_API_KEY", "").strip()

    try:
        if resend_key:
            _send_via_resend(
                api_key=resend_key,
                from_addr=from_addr,
                to_addr=to_addr,
                subject=subject,
                plain=plain,
                html=html,
            )
        else:
            _send_via_smtp(
                from_addr=from_addr,
                to_addr=to_addr,
                subject=subject,
                plain=plain,
                html=html,
            )
        return True
    except Exception as exc:
        print(f"Confirmation email failed: {exc}")
        return False
