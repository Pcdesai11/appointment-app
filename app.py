"""
Appointment form + Ask AI app.
- Local: stores in data/appointments.xlsx
- Vercel: stores the same Excel file in Vercel Blob (persistent)
"""

from __future__ import annotations

import io
import json
import os
import re
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook, load_workbook

load_dotenv()

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
EXCEL_PATH = DATA_DIR / "appointments.xlsx"
BLOB_PATHNAME = "appointments.xlsx"
COLUMNS = ["Timestamp", "Name", "Date", "Time", "Reason"]

app = Flask(
    __name__,
    template_folder=str(APP_DIR / "templates"),
    static_folder=str(APP_DIR / "static"),
)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "appointment-local-dev-key")


def using_blob() -> bool:
    return bool(os.getenv("BLOB_READ_WRITE_TOKEN", "").strip())


def storage_label() -> str:
    if using_blob():
        return "Vercel Blob (Excel in the cloud)"
    return str(EXCEL_PATH)


# --- Excel helpers -----------------------------------------------------------------

def _blank_workbook_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Appointments"
    ws.append(COLUMNS)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _rows_from_workbook_bytes(data: bytes) -> list[dict[str, str]]:
    wb = load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.active
    rows: list[dict[str, str]] = []
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    # Normalize headers
    header_map = {}
    for idx, h in enumerate(headers):
        name = str(h or "").strip()
        header_map[idx] = name if name in COLUMNS else (COLUMNS[idx] if idx < len(COLUMNS) else name)

    for excel_row in ws.iter_rows(min_row=2, values_only=True):
        if excel_row is None or all(v is None or str(v).strip() == "" for v in excel_row):
            continue
        item: dict[str, str] = {c: "" for c in COLUMNS}
        for idx, value in enumerate(excel_row):
            key = header_map.get(idx)
            if key not in COLUMNS:
                continue
            if value is None:
                item[key] = ""
            elif isinstance(value, datetime):
                item[key] = (
                    value.strftime("%Y-%m-%d")
                    if key == "Date"
                    else value.strftime("%H:%M")
                    if key == "Time"
                    else value.strftime("%Y-%m-%d %H:%M:%S")
                )
            elif isinstance(value, time):
                item[key] = value.strftime("%H:%M")
            else:
                text = str(value).strip()
                if key == "Time" and re.fullmatch(r"\d{1,2}:\d{2}:\d{2}", text):
                    text = text[:5]
                item[key] = text
        if any(item[c] for c in ("Name", "Date", "Time", "Reason")):
            rows.append(item)
    return rows


def _workbook_bytes_from_rows(rows: list[dict[str, str]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Appointments"
    ws.append(COLUMNS)
    for row in rows:
        ws.append([row.get(c, "") for c in COLUMNS])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --- Blob storage ------------------------------------------------------------------

BLOB_API = "https://vercel.com/api/blob"


def _blob_list() -> list[dict[str, Any]]:
    import urllib.parse
    import urllib.request

    token = os.environ["BLOB_READ_WRITE_TOKEN"]
    qs = urllib.parse.urlencode({"prefix": BLOB_PATHNAME})
    req = urllib.request.Request(
        f"{BLOB_API}?{qs}",
        headers={
            "Authorization": f"Bearer {token}",
            "x-api-version": "7",
        },
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    return payload.get("blobs") or []


def _blob_download(url: str) -> bytes:
    import urllib.request

    with urllib.request.urlopen(url, timeout=30) as resp:
        return resp.read()


def _blob_upload(data: bytes) -> None:
    import urllib.parse
    import urllib.request

    token = os.environ["BLOB_READ_WRITE_TOKEN"]
    qs = urllib.parse.urlencode({"pathname": BLOB_PATHNAME})
    req = urllib.request.Request(
        f"{BLOB_API}?{qs}",
        data=data,
        headers={
            "Authorization": f"Bearer {token}",
            "x-api-version": "7",
            "x-vercel-blob-access": "public",
            "x-allow-overwrite": "1",
            "x-add-random-suffix": "0",
            "x-content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        method="PUT",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        resp.read()


def _read_excel_bytes() -> bytes:
    if using_blob():
        blobs = _blob_list()
        match = next((b for b in blobs if b.get("pathname") == BLOB_PATHNAME), None)
        if not match:
            data = _blank_workbook_bytes()
            _blob_upload(data)
            return data
        return _blob_download(match["url"])

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not EXCEL_PATH.exists():
        EXCEL_PATH.write_bytes(_blank_workbook_bytes())
    return EXCEL_PATH.read_bytes()


def _write_excel_bytes(data: bytes) -> None:
    if using_blob():
        _blob_upload(data)
        return
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    EXCEL_PATH.write_bytes(data)


def read_appointments() -> list[dict[str, str]]:
    return _rows_from_workbook_bytes(_read_excel_bytes())


def append_appointment(name: str, date_str: str, time_str: str, reason: str) -> None:
    rows = read_appointments()
    rows.append(
        {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Name": name.strip(),
            "Date": date_str,
            "Time": time_str,
            "Reason": reason.strip(),
        }
    )
    _write_excel_bytes(_workbook_bytes_from_rows(rows))


# --- Ask / filter ------------------------------------------------------------------

def _normalize_time_token(raw: str) -> time | None:
    raw = raw.strip().lower().replace(".", ":")
    raw = re.sub(r"\s+", "", raw)

    m = re.fullmatch(r"(\d{1,2}):(\d{2})\s*(am|pm)?", raw)
    if m:
        hour = int(m.group(1))
        minute = int(m.group(2))
        meridiem = m.group(3)
        if meridiem == "pm" and hour != 12:
            hour += 12
        if meridiem == "am" and hour == 12:
            hour = 0
        if 0 <= hour <= 23 and 0 <= minute <= 59:
            return time(hour, minute)
        return None

    m = re.fullmatch(r"(\d{1,2})\s*(am|pm)", raw)
    if m:
        hour = int(m.group(1))
        meridiem = m.group(2)
        if meridiem == "pm" and hour != 12:
            hour += 12
        if meridiem == "am" and hour == 12:
            hour = 0
        if 0 <= hour <= 23:
            return time(hour, 0)
        return None

    return None


def _row_time(value: str) -> time | None:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p", "%I %p", "%I%p"):
        try:
            return datetime.strptime(text, fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            continue
    return _normalize_time_token(text)


def _row_date(value: str) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text[:10], fmt)
        except ValueError:
            continue
    return None


def filter_with_rules(question: str, rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], str]:
    q = question.strip().lower()
    if not rows:
        return [], "No appointments saved yet."

    filtered = list(rows)
    notes: list[str] = []

    time_matches = re.findall(
        r"(?:at\s+)?(\d{1,2}:\d{2}\s*(?:am|pm)?|\d{1,2}\s*(?:am|pm))",
        q,
        flags=re.IGNORECASE,
    )
    target_times: list[time] = []
    for token in time_matches:
        t = _normalize_time_token(token)
        if t:
            target_times.append(t)

    if target_times:
        ambiguous = []
        for token in time_matches:
            cleaned = re.sub(r"\s+", "", token.lower())
            ambiguous.append(not cleaned.endswith(("am", "pm")))

        def matches_time(val: str) -> bool:
            row_t = _row_time(val)
            if row_t is None:
                return False
            for t, is_ambiguous in zip(target_times, ambiguous):
                if row_t.minute != t.minute:
                    continue
                if is_ambiguous and 1 <= t.hour <= 12:
                    if row_t.hour % 12 == t.hour % 12:
                        return True
                elif row_t.hour == t.hour:
                    return True
            return False

        filtered = [r for r in filtered if matches_time(r.get("Time", ""))]
        pretty = ", ".join(
            (t.strftime("%I:%M").lstrip("0") + (" (am/pm)" if amb else " " + t.strftime("%p")))
            for t, amb in zip(target_times, ambiguous)
        )
        notes.append(f"time = {pretty}")

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    if "today" in q:
        filtered = [r for r in filtered if _row_date(r.get("Date", "")) == today]
        notes.append("date = today")
    elif "tomorrow" in q:
        tomorrow = today + timedelta(days=1)
        filtered = [r for r in filtered if _row_date(r.get("Date", "")) == tomorrow]
        notes.append("date = tomorrow")
    else:
        date_match = re.search(r"(\d{4}-\d{2}-\d{2})", q)
        if date_match:
            target = datetime.strptime(date_match.group(1), "%Y-%m-%d")
            filtered = [r for r in filtered if _row_date(r.get("Date", "")) == target]
            notes.append(f"date = {date_match.group(1)}")

    name_match = re.search(r"(?:for|named|name(?:\s+is)?)\s+([a-zA-Z][a-zA-Z\s'-]{0,40})", q)
    if name_match:
        name = name_match.group(1).strip()
        name = re.split(r"\b(?:at|on|about|with|tomorrow|today)\b", name)[0].strip()
        if name:
            filtered = [r for r in filtered if name.lower() in r.get("Name", "").lower()]
            notes.append(f"name contains '{name}'")

    about_match = re.search(r"\babout\s+([a-zA-Z0-9][\w\s'-]{0,40})", q)
    if about_match:
        keyword = about_match.group(1).strip()
        keyword = re.split(r"\b(?:at|on|named|for)\b", keyword)[0].strip()
        if keyword and keyword not in {"appointment", "appointments"}:
            filtered = [r for r in filtered if keyword.lower() in r.get("Reason", "").lower()]
            notes.append(f"reason contains '{keyword}'")

    if not notes:
        stop = {
            "tell", "show", "list", "what", "when", "have", "appointment",
            "appointments", "please", "find", "give", "with", "from", "that",
            "this", "my", "the", "and", "are", "who", "booked", "schedule",
        }
        words = [w for w in re.findall(r"[a-zA-Z]{3,}", q) if w not in stop]
        if words:
            filtered = [
                r
                for r in filtered
                if any(
                    w.lower() in r.get("Name", "").lower() or w.lower() in r.get("Reason", "").lower()
                    for w in words
                )
            ]
            notes.append(f"text match: {', '.join(words)}")

    if notes:
        summary = "Filtered by " + "; ".join(notes) + f". Found {len(filtered)} row(s)."
    else:
        summary = f"Showing all {len(filtered)} appointment(s). Try asking e.g. “appointments at 6:00”."

    return filtered, summary


def ask_openai(question: str, rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], str] | None:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key or not rows:
        return None

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        prompt = (
            "You answer questions about appointment rows. "
            "Return ONLY a JSON object with keys: "
            '"indices" (array of 0-based row indexes that match) and '
            '"summary" (short plain-English answer). '
            "If none match, indices should be [].\n\n"
            f"Question: {question}\n\nRows:\n{rows}"
        )
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            response_format={"type": "json_object"},
        )
        payload = json.loads(response.choices[0].message.content or "{}")
        indices = payload.get("indices") or []
        summary = payload.get("summary") or "Here are the matching appointments."
        valid = [i for i in indices if isinstance(i, int) and 0 <= i < len(rows)]
        return [rows[i] for i in valid], str(summary)
    except Exception:
        return None


# --- Routes ------------------------------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    ask_result = None
    ask_summary = None
    question = ""

    if request.method == "POST":
        action = request.form.get("action")

        if action == "add":
            name = request.form.get("name", "").strip()
            date_str = request.form.get("date", "").strip()
            time_str = request.form.get("time", "").strip()
            reason = request.form.get("reason", "").strip()

            if not all([name, date_str, time_str, reason]):
                flash("Please fill in all fields.", "error")
            else:
                try:
                    append_appointment(name, date_str, time_str, reason)
                    flash("Appointment saved to Excel.", "success")
                except Exception as exc:
                    flash(f"Could not save appointment: {exc}", "error")
            return redirect(url_for("index"))

        if action == "ask":
            question = request.form.get("question", "").strip()
            rows = read_appointments()
            if not question:
                flash("Type a question first.", "error")
            else:
                ai = ask_openai(question, rows)
                if ai is not None:
                    ask_result, ask_summary = ai
                else:
                    ask_result, ask_summary = filter_with_rules(question, rows)

    appointments = read_appointments()
    return render_template(
        "index.html",
        appointments=appointments,
        ask_result=ask_result,
        ask_summary=ask_summary,
        question=question,
        excel_path=storage_label(),
        using_cloud=using_blob(),
    )


@app.route("/download.xlsx")
def download_excel():
    data = _read_excel_bytes()
    return send_file(
        io.BytesIO(data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="appointments.xlsx",
    )


# WSGI entry for Vercel
app.debug = False

if __name__ == "__main__":
    print(f"Storage: {storage_label()}")
    print("Open http://127.0.0.1:5000")
    app.run(debug=True, host="127.0.0.1", port=5000)
