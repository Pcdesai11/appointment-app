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
JSON_LEGACY_PATHNAME = "appointments.json"
COLUMNS = ["Timestamp", "Name", "Date", "Time", "Reason"]
XLSX_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

app = Flask(
    __name__,
    template_folder=str(APP_DIR / "templates"),
    static_folder=str(APP_DIR / "static"),
)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "appointment-local-dev-key")


class _VercelPathMiddleware:
    """Restore the browser path when Vercel rewrites to /api/index."""

    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        original = (
            environ.get("HTTP_X_VERCEL_ORIGINAL_PATH")
            or environ.get("HTTP_X_INVOKE_PATH")
            or environ.get("HTTP_X_FORWARDED_URI")
            or ""
        )
        path = original.split("?", 1)[0]
        if path.startswith("http"):
            from urllib.parse import urlparse

            path = urlparse(path).path
        # If rewrite collapsed everything to /api/index, prefer the request URI path.
        req_uri = environ.get("REQUEST_URI") or environ.get("RAW_URI") or ""
        if (not path or path.startswith("/api/index")) and req_uri:
            path = req_uri.split("?", 1)[0]
        if path and path != environ.get("PATH_INFO"):
            environ["PATH_INFO"] = path
            environ["SCRIPT_NAME"] = ""
        return self.wsgi_app(environ, start_response)


app.wsgi_app = _VercelPathMiddleware(app.wsgi_app)

try:
    from werkzeug.middleware.proxy_fix import ProxyFix

    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
except Exception:
    pass


def using_blob() -> bool:
    return bool(os.getenv("BLOB_READ_WRITE_TOKEN", "").strip())


def storage_label() -> str:
    if using_blob():
        return "One shared file: appointments.xlsx (cloud)"
    return f"One shared file: {EXCEL_PATH}"


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


def _blob_list(prefix: str | None = None) -> list[dict[str, Any]]:
    import urllib.parse
    import urllib.request

    token = os.environ["BLOB_READ_WRITE_TOKEN"]
    qs = urllib.parse.urlencode({"prefix": prefix or BLOB_PATHNAME})
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


def _blob_upload(
    data: bytes,
    pathname: str = BLOB_PATHNAME,
    content_type: str = XLSX_CONTENT_TYPE,
) -> None:
    import urllib.error
    import urllib.parse
    import urllib.request

    token = os.environ["BLOB_READ_WRITE_TOKEN"]
    # Path-style PUT matches the Blob API contract used by x-api-version 7.
    url = f"{BLOB_API}/{urllib.parse.quote(pathname, safe='/')}"
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": content_type,
            "x-api-version": "7",
            "x-vercel-blob-access": "public",
            "x-allow-overwrite": "1",
            "x-add-random-suffix": "0",
            "x-content-type": content_type,
        },
        method="PUT",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            resp.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Blob upload failed ({exc.code}): {detail}") from exc


def _find_blob(pathname: str) -> dict[str, Any] | None:
    blobs = _blob_list(pathname)
    return next((b for b in blobs if b.get("pathname") == pathname), None)


def _migrate_legacy_json_if_needed() -> bytes | None:
    """If an older JSON blob exists and Excel does not, convert once into Excel."""
    excel_blob = _find_blob(BLOB_PATHNAME)
    if excel_blob and excel_blob.get("url"):
        return None
    json_blob = _find_blob(JSON_LEGACY_PATHNAME)
    if not json_blob or not json_blob.get("url"):
        return None
    raw = _blob_download(json_blob["url"]).decode("utf-8")
    data = json.loads(raw or "[]")
    rows: list[dict[str, str]] = []
    if isinstance(data, list):
        for item in data:
            if isinstance(item, dict):
                rows.append({c: str(item.get(c, "") or "") for c in COLUMNS})
    excel_bytes = _workbook_bytes_from_rows(rows)
    _blob_upload(excel_bytes)
    return excel_bytes


def read_excel_bytes() -> bytes:
    """Return the single shared appointments.xlsx bytes (local or cloud)."""
    if using_blob():
        match = _find_blob(BLOB_PATHNAME)
        if match and match.get("url"):
            return _blob_download(match["url"])
        migrated = _migrate_legacy_json_if_needed()
        if migrated is not None:
            return migrated
        blank = _blank_workbook_bytes()
        _blob_upload(blank)
        return blank

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not EXCEL_PATH.exists():
        EXCEL_PATH.write_bytes(_blank_workbook_bytes())
    return EXCEL_PATH.read_bytes()


def write_excel_bytes(data: bytes) -> None:
    """Overwrite the same appointments.xlsx file (never create a new name)."""
    if using_blob():
        _blob_upload(data)
        return
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    EXCEL_PATH.write_bytes(data)


def read_appointments() -> list[dict[str, str]]:
    return _rows_from_workbook_bytes(read_excel_bytes())


def append_appointment(name: str, date_str: str, time_str: str, reason: str) -> None:
    """Append one row into the same Excel workbook and save it back."""
    data = read_excel_bytes()
    wb = load_workbook(io.BytesIO(data))
    ws = wb.active
    if ws.max_row == 0 or ws.cell(1, 1).value is None:
        ws.append(COLUMNS)
    ws.append(
        [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            name.strip(),
            date_str,
            time_str,
            reason.strip(),
        ]
    )
    buf = io.BytesIO()
    wb.save(buf)
    write_excel_bytes(buf.getvalue())


# --- Ask / filter ------------------------------------------------------------------

STOP_WORDS = {
    "tell", "show", "list", "what", "when", "have", "appointment", "appointments",
    "please", "find", "give", "with", "from", "that", "this", "my", "the", "and",
    "are", "who", "booked", "schedule", "schedules", "me", "any", "get", "see",
    "display", "looking", "look", "for", "can", "you", "your", "our", "there",
    "here", "about", "all", "everything", "everyone", "rows", "data", "excel",
    "do", "i", "is", "a", "an", "of", "on", "to", "in", "at", "or", "be", "was",
    "were", "am", "pm", "o", "clock", "time", "date", "name", "reason", "which",
    "ones", "those", "these", "some", "them", "their", "has", "had", "will",
}


def _normalize_time_token(raw: str) -> time | None:
    raw = raw.strip().lower().replace(".", ":")
    raw = re.sub(r"\s+", "", raw)
    raw = raw.replace("o'clock", "").replace("oclock", "")

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

    # Bare hour: "6", "18"
    m = re.fullmatch(r"(\d{1,2})", raw)
    if m:
        hour = int(m.group(1))
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
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%m-%d-%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text[:10], fmt)
        except ValueError:
            continue
    try:
        # Lenient parse for values like "Aug 16, 2026"
        parsed = datetime.strptime(text, "%b %d, %Y")
        return parsed.replace(hour=0, minute=0, second=0, microsecond=0)
    except ValueError:
        return None


def _extract_time_queries(q: str) -> list[tuple[time, bool, bool]]:
    """
    Returns list of (time, ambiguous_ampm, hour_only).
    hour_only=True means match any minutes in that hour.
    """
    found: list[tuple[time, bool, bool]] = []

    patterns = [
        # at 6:00 / 6:00pm / 18:00
        r"(?:at|around|about|by|for)?\s*(\d{1,2}:\d{2})\s*(am|pm)?",
        # 6pm / 6 pm
        r"(?:at|around|about|by|for)?\s*(\d{1,2})\s*(am|pm)\b",
        # at 6 / at 6 o'clock
        r"(?:at|around|about|by)\s+(\d{1,2})(?:\s*o'?clock)?\b",
        # bare evening-style: "6 o'clock"
        r"\b(\d{1,2})\s*o'?clock\b",
    ]

    seen: set[tuple[int, int, bool, bool]] = set()
    for pattern in patterns:
        for m in re.finditer(pattern, q, flags=re.IGNORECASE):
            groups = [g for g in m.groups() if g is not None]
            if not groups:
                continue
            token = "".join(groups)
            # Rebuild nicer token
            if len(groups) == 2 and groups[1] in {"am", "pm"}:
                token = f"{groups[0]}{groups[1]}"
            elif len(groups) == 1:
                token = groups[0]
            t = _normalize_time_token(token)
            if not t:
                continue
            cleaned = re.sub(r"\s+", "", token.lower())
            has_meridiem = cleaned.endswith(("am", "pm"))
            hour_only = ":" not in cleaned and not has_meridiem
            # "6pm" is hour_only for minutes (any minute in that hour) but not ambiguous am/pm
            if has_meridiem and ":" not in cleaned:
                hour_only = True
            ambiguous = not has_meridiem
            key = (t.hour, t.minute, ambiguous, hour_only)
            if key in seen:
                continue
            seen.add(key)
            found.append((t, ambiguous, hour_only))
    return found


def _time_matches_row(row_t: time, target: time, ambiguous: bool, hour_only: bool) -> bool:
    if hour_only:
        if ambiguous and 1 <= target.hour <= 12:
            return row_t.hour % 12 == target.hour % 12
        return row_t.hour == target.hour

    if ambiguous and 1 <= target.hour <= 12:
        return row_t.hour % 12 == target.hour % 12 and row_t.minute == target.minute
    return row_t.hour == target.hour and row_t.minute == target.minute


def _fuzzy_contains(haystack: str, needle: str) -> bool:
    h = haystack.lower().strip()
    n = needle.lower().strip()
    if not n:
        return False
    if n in h:
        return True
    # Soft: all needle tokens appear somewhere
    parts = [p for p in re.split(r"[\s_-]+", n) if p]
    return bool(parts) and all(p in h for p in parts)


def filter_with_rules(question: str, rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], str]:
    q = question.strip().lower()
    if not rows:
        return [], "No appointments saved yet."

    # Broad "show everything" style questions
    if re.search(
        r"\b(all|everything|everyone|entire|full\s+list|show\s+me\s+all|list\s+all)\b",
        q,
    ) and not re.search(r"\b(\d{1,2}|today|tomorrow|named|about)\b", q):
        return rows, f"Showing all {len(rows)} appointment(s)."

    filtered = list(rows)
    notes: list[str] = []

    time_queries = _extract_time_queries(q)
    if time_queries:
        def matches_time(val: str) -> bool:
            row_t = _row_time(val)
            if row_t is None:
                return False
            return any(
                _time_matches_row(row_t, t, ambiguous, hour_only)
                for t, ambiguous, hour_only in time_queries
            )

        filtered = [r for r in filtered if matches_time(r.get("Time", ""))]
        pretty_bits = []
        for t, ambiguous, hour_only in time_queries:
            label = t.strftime("%I").lstrip("0")
            if not hour_only:
                label += t.strftime(":%M")
            if ambiguous:
                label += " (am/pm)"
            else:
                label += " " + t.strftime("%p")
            if hour_only:
                label += " any minute"
            pretty_bits.append(label)
        notes.append("time ~ " + ", ".join(pretty_bits))

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    if re.search(r"\btoday\b", q):
        filtered = [r for r in filtered if _row_date(r.get("Date", "")) == today]
        notes.append("date = today")
    elif re.search(r"\btomorrow\b", q):
        tomorrow = today + timedelta(days=1)
        filtered = [r for r in filtered if _row_date(r.get("Date", "")) == tomorrow]
        notes.append("date = tomorrow")
    else:
        date_match = re.search(r"(\d{4}-\d{2}-\d{2})", q)
        if date_match:
            target = datetime.strptime(date_match.group(1), "%Y-%m-%d")
            filtered = [r for r in filtered if _row_date(r.get("Date", "")) == target]
            notes.append(f"date = {date_match.group(1)}")

    # Names: "named X", "for X", "with X", or any capitalized-looking token handled via words below
    name_match = re.search(
        r"(?:for|named|name(?:\s+is)?|with|by)\s+([a-zA-Z][a-zA-Z\s'-]{0,40})",
        q,
    )
    if name_match:
        name = name_match.group(1).strip()
        name = re.split(
            r"\b(?:at|on|about|with|tomorrow|today|around|appointment|appointments)\b",
            name,
        )[0].strip()
        if name and name not in STOP_WORDS:
            filtered = [r for r in filtered if _fuzzy_contains(r.get("Name", ""), name)]
            notes.append(f"name ~ '{name}'")

    about_match = re.search(r"\b(?:about|regarding|reason|for)\s+([a-zA-Z0-9][\w\s'-]{0,40})", q)
    if about_match:
        keyword = about_match.group(1).strip()
        keyword = re.split(r"\b(?:at|on|named|with|tomorrow|today)\b", keyword)[0].strip()
        if keyword and keyword not in STOP_WORDS and keyword not in {"appointment", "appointments"}:
            # Only apply reason filter if it actually hits something OR user said about/reason explicitly
            reason_hits = [r for r in filtered if _fuzzy_contains(r.get("Reason", ""), keyword)]
            name_hits = [r for r in filtered if _fuzzy_contains(r.get("Name", ""), keyword)]
            if reason_hits:
                filtered = reason_hits
                notes.append(f"reason ~ '{keyword}'")
            elif name_hits and "about" not in q and "reason" not in q:
                filtered = name_hits
                notes.append(f"name ~ '{keyword}'")
            elif "about" in q or "reason" in q or "regarding" in q:
                filtered = reason_hits
                notes.append(f"reason ~ '{keyword}'")

    # If nothing specific matched yet, soft-search leftover words across all columns
    if not notes:
        words = [w for w in re.findall(r"[a-zA-Z0-9]{2,}", q) if w not in STOP_WORDS]
        # Drop pure numbers already handled as times unless no time filter ran
        words = [w for w in words if not w.isdigit()]
        if words:
            def row_hits(r: dict[str, str]) -> bool:
                blob = " ".join(r.get(c, "") for c in ("Name", "Date", "Time", "Reason")).lower()
                return any(w in blob for w in words)

            hits = [r for r in filtered if row_hits(r)]
            if hits:
                filtered = hits
                notes.append(f"text ~ {', '.join(words)}")
            else:
                # Lenient fallback: show all instead of empty when the chatbot is unsure
                return (
                    rows,
                    f"Couldn't tightly match \"{question.strip()}\", so showing all {len(rows)} appointment(s). "
                    f"Try \"at 6\", \"today\", or a name.",
                )

    if notes:
        summary = "Matched " + "; ".join(notes) + f". Found {len(filtered)} row(s)."
    else:
        summary = f"Showing all {len(filtered)} appointment(s)."

    # Extra leniency: if filters wiped everything, fall back to all with a hint
    if notes and not filtered:
        return (
            rows,
            f"No exact matches for \"{question.strip()}\". Showing all {len(rows)} appointment(s) instead.",
        )

    return filtered, summary


def ask_free_ai(question: str, rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], str] | None:
    """
    Free AI first (Groq / Llama), optional OpenAI, else None -> built-in filter.
    Get a free Groq key at https://console.groq.com
    """
    if not rows:
        return None

    prompt = (
        "You help a doctor review appointment bookings. "
        "Return ONLY a JSON object with keys: "
        '"indices" (array of 0-based row indexes that match the question) and '
        '"summary" (short plain-English answer for the doctor). '
        "Be helpful and inclusive — if the question is vague, return likely matches. "
        "If none match, indices should be [].\n\n"
        f"Question: {question}\n\nRows:\n{rows}"
    )

    providers: list[tuple[str, str, str]] = []
    groq_key = os.getenv("GROQ_API_KEY", "").strip()
    if groq_key:
        providers.append(
            (groq_key, "https://api.groq.com/openai/v1", "llama-3.3-70b-versatile")
        )
    openai_key = os.getenv("OPENAI_API_KEY", "").strip()
    if openai_key:
        providers.append((openai_key, "https://api.openai.com/v1", "gpt-4o-mini"))

    for api_key, base_url, model in providers:
        try:
            from openai import OpenAI

            client = OpenAI(api_key=api_key, base_url=base_url)
            response = client.chat.completions.create(
                model=model,
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
            continue
    return None


def doctor_password() -> str:
    return os.getenv("DOCTOR_PASSWORD", "doctor123").strip() or "doctor123"


def doctor_logged_in() -> bool:
    from flask import session

    return bool(session.get("doctor_ok"))


def require_doctor(view):
    from functools import wraps

    from flask import session

    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("doctor_ok"):
            return redirect(url_for("doctor_login"))
        return view(*args, **kwargs)

    return wrapped


@app.route("/", methods=["GET", "POST"])
@app.route("/book", methods=["GET", "POST"])
@app.route("/api/index", methods=["GET", "POST"])
def patient_form():
    """Public patient link — form only."""
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        date_str = request.form.get("date", "").strip()
        time_str = request.form.get("time", "").strip()
        reason = request.form.get("reason", "").strip()

        if not all([name, date_str, time_str, reason]):
            flash("Please fill in all fields.", "error")
        else:
            try:
                append_appointment(name, date_str, time_str, reason)
                flash("Thanks — your appointment request was submitted.", "success")
            except Exception as exc:
                flash(f"Could not submit right now. Please try again. ({exc})", "error")
        return redirect(url_for("patient_form"))

    return render_template("patient.html")


@app.route("/doctor/login", methods=["GET", "POST"])
def doctor_login():
    from flask import session

    if session.get("doctor_ok"):
        return redirect(url_for("doctor_dashboard"))

    if request.method == "POST":
        password = request.form.get("password", "")
        if password == doctor_password():
            session["doctor_ok"] = True
            flash("Welcome back, doctor.", "success")
            return redirect(url_for("doctor_dashboard"))
        flash("Incorrect password.", "error")

    return render_template("doctor_login.html")


@app.route("/doctor/logout", methods=["POST", "GET"])
def doctor_logout():
    from flask import session

    session.clear()
    flash("Signed out.", "success")
    return redirect(url_for("doctor_login"))


@app.route("/doctor", methods=["GET", "POST"])
@require_doctor
def doctor_dashboard():
    ask_result = None
    ask_summary = None
    question = ""
    ai_status = (
        "AI assistant connected"
        if os.getenv("GROQ_API_KEY", "").strip() or os.getenv("OPENAI_API_KEY", "").strip()
        else "Smart filter active"
    )

    if request.method == "POST" and request.form.get("action") == "ask":
        question = request.form.get("question", "").strip()
        rows = read_appointments()
        if not question:
            flash("Type a question first.", "error")
        else:
            ai = ask_free_ai(question, rows)
            if ai is not None:
                ask_result, ask_summary = ai
            else:
                ask_result, ask_summary = filter_with_rules(question, rows)

    appointments = read_appointments()
    return render_template(
        "doctor.html",
        appointments=appointments,
        ask_result=ask_result,
        ask_summary=ask_summary,
        question=question,
        excel_path=storage_label(),
        ai_status=ai_status,
        patient_link=url_for("patient_form", _external=True),
    )


@app.route("/download.xlsx")
@require_doctor
def download_excel():
    data = read_excel_bytes()
    return send_file(
        io.BytesIO(data),
        mimetype=XLSX_CONTENT_TYPE,
        as_attachment=True,
        download_name="appointments.xlsx",
    )


# WSGI entry for Vercel
app.debug = False

if __name__ == "__main__":
    print(f"Storage: {storage_label()}")
    print("Patient form: http://127.0.0.1:5000/")
    print("Doctor desk:  http://127.0.0.1:5000/doctor")
    app.run(debug=True, host="127.0.0.1", port=5000)
