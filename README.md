# Appointments (Form + Ask AI + Excel)

One shared **`appointments.xlsx`** file:
- Each form submit **appends a row** to that same workbook
- Download always returns that same file (updated contents)

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

Local file: `data/appointments.xlsx`  
Hosted: the same filename in Vercel Blob (`appointments.xlsx`), overwritten in place on every save.
