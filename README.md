# wilson-capital-financial

Financial tooling for Wilson Capital's property portfolio.

## Setup

```powershell
python -m venv venv
venv\Scripts\python.exe -m pip install -r requirements.txt
copy .env.example .env   # then fill in real values
```

## Notes

- Isolated git repo (see `C:\Users\daisy\Projects\SETUP.md` for conventions).
- Real secrets live in `.env` (git-ignored); `.env.example` documents the shape.
