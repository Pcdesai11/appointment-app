# Appointments — patient form + doctor desk

## Two links

| Who | URL | What they see |
|---|---|---|
| **Patients** | `/` | Booking form only |
| **You (doctor)** | `/doctor` | Ask AI, Excel download, all appointments |

Live example:
- Patient: https://appointment-app-ivory.vercel.app/
- Doctor: https://appointment-app-ivory.vercel.app/doctor

Default doctor password: `doctor123`  
Change it with env var `DOCTOR_PASSWORD`.

## Free AI (Groq / Llama)

1. Create a free key at https://console.groq.com/keys  
2. Set `GROQ_API_KEY` in Vercel → Project → Settings → Environment Variables  
3. Redeploy  

Without a key, Ask AI still works with the built-in lenient filter.

## Local run

```powershell
cd appointment-app
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

## Env vars

- `DOCTOR_PASSWORD` — doctor desk login  
- `GROQ_API_KEY` — free Llama AI  
- `BLOB_READ_WRITE_TOKEN` — set automatically by Vercel Blob  
- `FLASK_SECRET_KEY` — session cookie secret  
