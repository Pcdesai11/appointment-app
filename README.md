# Khatna appointments — patient form + doctor desk

## Links

| Who | URL | What they see |
|---|---|---|
| **Patients** | `/` | Gujarati/English khatna booking form |
| **Doctor** | `/doctor` | Ask AI, Excel download, all bookings |

Patient times are fixed: **12 PM, 3 PM, 6 PM** only.

## Form fields

- Baby full name / બાળકનું પૂરું નામ
- Age / ઉંમર
- Weight / વજન
- Date / તારીખ
- Time / સમય (12 / 3 / 6 PM)
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
   - `GOOGLE_SERVICE_ACCOUNT_JSON` = full JSON key (one line)
   - `GOOGLE_CALENDAR_ID` = your calendar id
   - `GOOGLE_CALENDAR_TIMEZONE` = `Asia/Kolkata`

When a patient submits, the booking is saved to Excel **and** added to the calendar.

## Local run

```powershell
cd appointment-app
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```
