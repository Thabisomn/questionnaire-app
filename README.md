# Stakeholder Engagement Questionnaire – Flask Mini-System

A simple Flask web app to collect questionnaire responses and export them as Excel.

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
python app.py
```

The app runs at: http://localhost:5000

## Pages

| URL        | Purpose                                      |
|------------|----------------------------------------------|
| `/`        | The questionnaire form (share this link)     |
| `/admin`   | Admin dashboard – see total response count   |
| `/download`| Download all responses as a .xlsx file       |

## How to Share

- Run `python app.py` on your computer
- Share the link with respondents on the same network (use your IP: `http://192.168.x.x:5000`)
- Or deploy to Render/Railway/PythonAnywhere for internet access

## Deployment on Render (recommended)

1. Push this folder to a GitHub repo
2. Create a new **Web Service** on Render
3. Set **Start Command**: `gunicorn app:app`
4. Install gunicorn: add `gunicorn` to requirements.txt

> **Note:** On Render's free tier, the SQLite `responses.db` file is ephemeral (resets on redeploy).
> For persistent storage, switch to PostgreSQL (SQLAlchemy + psycopg2) — same pattern as your DLEMS system.

## Excel Output

Each downloaded file contains:
- **Meta columns**: ID, Submitted At, Consent
- **Section A**: Demographic info (Gender, Age, District, Category, Education)
- **Sections B–F**: All Likert scale responses (1–5), organized by section with colored headers
