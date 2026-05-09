# MOM Creator Web App

A web-based application for creating Minutes of Meeting (MOM).

## Features

- Create MOM with Date, Time, Attendees (with roles), Points Discussed, Action Items (with dates and responsible persons), and Signatory section
- Grid-based tables for Attendees and Action Items with add/remove functionality
- Manual entry option
- Export to Word (.docx) and PDF
- Date formatting in exported documents uses DD/MM/YYYY

## Requirements

- Python 3.x
- Dependencies: streamlit, python-docx, fpdf

## Installation

1. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

Run the Flask website (React UI + API):
```
python app.py
```

Open [http://127.0.0.1:5000](http://127.0.0.1:5000).

Optional: production-style local run:
```
gunicorn app:app --bind 127.0.0.1:5000
```

- Fill in Meeting Date and Time.
- Add Attendees using the grid and remove rows as needed.
- Enter Points Discussed manually.
- Add Action Items using the action grid.
- Export to Word or PDF using the download buttons.

## Deploy as a website

This app is **Flask** (`app.py`) serving the React + Tailwind UI and `/api/export/*` endpoints. Do **not** deploy with Streamlit Cloud for this UI.

### Render (recommended; blueprint included)

1. Push the repo to GitHub.
2. In [Render](https://render.com): **New +** → **Blueprint**.
3. Connect the repo; Render reads `render.yaml`.
4. Deploy. It runs: `gunicorn app:app --bind 0.0.0.0:$PORT`.

Your service URL will look like `https://mom-creator.onrender.com` (name may vary).

### Railway / Fly.io / PythonAnywhere

Same idea: **Python web service**, install `requirements.txt`, start with:

```bash
gunicorn app:app --bind 0.0.0.0:$PORT
```

Use the platform’s **PORT** env var / binding docs.

### Checklist before go-live

- [ ] Repo includes `requirements.txt` (includes `flask`, `gunicorn`, `python-docx`, `fpdf`).
- [ ] **`uploads/`** is not required for current exports (in-memory response); no persistent disk needed unless you add file storage later.
- [ ] PDF uses Verdana on Windows paths in code; on Linux hosts it **falls back to Arial** (already handled in `app.py`).
- [ ] Custom domain: add in host DNS (CNAME) + enable HTTPS in the host dashboard.

### Local production-style check

```bash
pip install -r requirements.txt
gunicorn app:app --bind 127.0.0.1:5000
```

Open [http://127.0.0.1:5000](http://127.0.0.1:5000) and test Word/PDF export.
