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

Run the Streamlit app:
```
streamlit run streamlit_app.py
```

Open the local Streamlit URL shown in the terminal.

- Fill in Meeting Date and Time.
- Add Attendees using the grid and remove rows as needed.
- Enter Points Discussed manually.
- Add Action Items using the action grid.
- Export to Word or PDF using the download buttons.

## Deploy as a Website

### Option 1: Streamlit Community Cloud (fastest)

1. Push this project to a GitHub repository.
2. Go to [https://share.streamlit.io/](https://share.streamlit.io/) and sign in.
3. Click **New app** and select:
   - Repository: your repo
   - Branch: `main` (or your branch)
   - Main file path: `streamlit_app.py`
4. Click **Deploy**.

Your app will get a public URL like:
`https://<your-app-name>.streamlit.app`

### Option 2: Render

This repo includes `render.yaml` for blueprint deployment.

1. Push this project to GitHub.
2. In Render, choose **New +** -> **Blueprint**.
3. Select your repository and deploy.
4. Render will automatically:
   - install `requirements.txt`
   - run `streamlit run streamlit_app.py --server.port $PORT --server.address 0.0.0.0`

### Local verify before deploying

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

If you want to keep the previous Flask app as a legacy version, `app.py` remains in the repo but is no longer the main website entry point.
