# Appointments (Form + Ask AI + Excel)

Web app with:
- Form: Name, Date, Time, Reason
- Cloud storage on Vercel Blob (local Excel file when run on your PC)
- **Download Excel** button exports `appointments.xlsx`
- Ask AI box for questions like “appointments at 6:00”

## Live

- App: https://appointment-app-ivory.vercel.app
- Repo: https://github.com/Pcdesai11/appointment-app

## Local run

```powershell
cd appointment-app
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000  
Local Excel path: `data/appointments.xlsx`

## Vercel notes

A public Blob store (`appointments-excel`) is linked to the project.  
`BLOB_READ_WRITE_TOKEN` is set automatically.

Optional env vars:
- `OPENAI_API_KEY` — freer natural-language answers
- `FLASK_SECRET_KEY` — flash/session secret

## Ask AI examples

- `appointments at 6:00`
- `tomorrow`
- `named Priya`
- `about dental`
