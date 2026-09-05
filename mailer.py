"""Send appointment confirmation emails via SMTP."""

from __future__ import annotations

import os
import smtplib
import ssl
from email.message import EmailMessage
from typing import Any


def mail_configured() -> bool:
    return bool(
        os.getenv("SMTP_HOST", "").strip()
        and os.getenv("SMTP_USER", "").strip()
        and os.getenv("SMTP_PASSWORD", "").strip()
    )


def _from_address() -> str:
    return (
        os.getenv("SMTP_FROM", "").strip()
        or os.getenv("SMTP_USER", "").strip()
    )


def _time_label(time_str: str, time_labels: dict[str, str] | None = None) -> str:
    if time_labels and time_str in time_labels:
        # Prefer English-only for email when label has Gujarati separator
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

Important precautions / મહત્વની સાવચેતીઓ:
- The child should not eat anything for 2 hours before the appointment.
  અપોઇન્ટમેન્ટ પહેલાં બે કલાક સુધી બાળકને કંઈ ખવડાવશો નહીં.
- Keep the baby clean and dry.
- Inform the doctor about any fever, cough, cold, allergy, or medicine.
- Arrive 10–15 minutes early.
- Follow the doctor's advice after the procedure.

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
  <h3>Important precautions · મહત્વની સાવચેતીઓ</h3>
  <ul>
    <li><strong>The child should not eat anything for 2 hours before the appointment.</strong><br/>
      અપોઇન્ટમેન્ટ પહેલાં બે કલાક સુધી બાળકને કંઈ ખવડાવશો નહીં.</li>
    <li>Keep the baby clean and dry.</li>
    <li>Inform the doctor about any fever, cough, cold, allergy, or medicine.</li>
    <li>Arrive 10–15 minutes early.</li>
    <li>Follow the doctor's advice after the procedure.</li>
  </ul>
  <p>Thank you.</p>
</body>
</html>
"""
    return plain, html


def send_confirmation_email(
    booking: dict[str, Any],
    time_labels: dict[str, str] | None = None,
) -> bool:
    """
    Send confirmation if patient email is present and SMTP is configured.
    Returns True if sent, False if skipped/failed (never raises to caller).
    """
    to_addr = (booking.get("Email") or "").strip()
    if not to_addr or not mail_configured():
        return False

    host = os.getenv("SMTP_HOST", "").strip()
    port = int(os.getenv("SMTP_PORT", "587").strip() or "587")
    user = os.getenv("SMTP_USER", "").strip()
    password = os.getenv("SMTP_PASSWORD", "").strip()
    from_addr = _from_address()
    use_ssl = os.getenv("SMTP_SSL", "").strip().lower() in {"1", "true", "yes"}

    plain, html = build_confirmation_bodies(booking, time_labels)
    msg = EmailMessage()
    msg["Subject"] = f"Appointment confirmed — {booking.get('BabyName') or 'Khatna'}"
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")

    try:
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
        return True
    except Exception as exc:
        print(f"Confirmation email failed: {exc}")
        return False
