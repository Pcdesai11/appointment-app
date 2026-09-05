# Khatna appointments — patient form + doctor desk

## Links

| Who | URL | What they see |
|---|---|---|
| **Patients** | `/` | Gujarati/English khatna booking form |
| **Doctor** | `/doctor` | Ask AI, Excel download, all bookings |

Patient times are fixed: **10 AM, 1:30 PM, 5 PM** only.

## Form fields

- Baby full name / બાળકનું પૂરું નામ
- Age / ઉંમર
- Weight / વજન
- Date / તારીખ
- Time / સમય (10 AM / 1:30 PM / 5 PM)
- Area / વિસ્તાર
- Mobile / મોબાઇલ
- Email optional / ઈમેઈલ

## Google Calendar setup

1. In [Google Cloud Console](https://console.cloud.google.com/) create a project.
2. Enable **Google Calendar API**.
3. Create a **Service Account** → download JSON key.
4. Open Google Calendar → Settings → share calendar with the service account email (**Make changes to events**).
5. Copy Calendar ID (often your email, or from Calendar settings).
6. In Vercel env vars set:
   - `GOOGLE_CALENDAR_ID` = your calendar id
   - `GOOGLE_CALENDAR_TIMEZONE` = `Asia/Kolkata`
   - Prefer **base64** (avoids paste breakage):
     - PowerShell: `[Convert]::ToBase64String([IO.File]::ReadAllBytes("C:\path\to\key.json"))`
     - Set `GOOGLE_SERVICE_ACCOUNT_JSON_BASE64` to that output
   - Or set `GOOGLE_SERVICE_ACCOUNT_JSON` to the JSON as **one single line** (no outer quotes)

When a patient submits, the booking is saved to Excel **and** added to the calendar.

## Confirmation email (optional)

If the patient enters an email, a confirmation is sent when SMTP is configured in Vercel:

| Name | Example |
|---|---|
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | your sending Gmail |
| `SMTP_PASSWORD` | Gmail App Password |
| `SMTP_FROM` | optional display From (defaults to SMTP_USER) |

For Gmail: enable 2-Step Verification → create an [App Password](https://myaccount.google.com/apppasswords), use that as `SMTP_PASSWORD`.

## Local run

```powershell
cd appointment-app
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```
