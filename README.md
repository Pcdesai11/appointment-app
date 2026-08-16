# Appointments (Form + Ask AI + Excel)

Web app with:
- Form: Name, Date, Time, Reason
- Excel storage (local file, or Vercel Blob when hosted)
- Ask AI box for questions like “appointments at 6:00”
- Public shareable URL on Vercel

## Local run

```powershell
cd appointment-app
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -c "from app import app; app.run(debug=True)"
```

Open http://127.0.0.1:5000  
Local Excel path: `data/appointments.xlsx`

## Vercel

1. Create a **Blob** store in the Vercel project (Storage → Blob).
2. Ensure `BLOB_READ_WRITE_TOKEN` is set in Project → Settings → Environment Variables.
3. Redeploy.

Optional: `OPENAI_API_KEY`, `FLASK_SECRET_KEY`

## Ask AI examples

- `appointments at 6:00`
- `tomorrow`
- `named Priya`
- `about dental`
