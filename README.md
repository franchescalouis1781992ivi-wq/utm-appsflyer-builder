# UTM / AppsFlyer Builder

Shared Streamlit tool for:
- campaign parameter setup
- web/app link generation
- AppsFlyer bulk sheet preview
- local archive/history with reopen, duplicate, update, delete, and search
- CSV/XLSX export

## Files
- `app.py`: Streamlit app
- `start.sh`: deployment start command
- `render.yaml`: Render deployment config
- `.streamlit/config.toml`: Streamlit runtime config
- `requirements.txt`: Python dependencies
- `data/link_builder_archive.db`: local SQLite archive, created automatically

## Local run
```bash
cd /Users/chengxinyuan/Documents/Playground
pip3 install -r requirements.txt
streamlit run app.py
```

## Deploy options

### Option 1: Streamlit Community Cloud
1. Push this folder to a GitHub repo.
2. Open [https://share.streamlit.io/](https://share.streamlit.io/).
3. Create a new app from that repo.
4. Main file path: `app.py`
5. Deploy.

Recommended when:
- you want the fastest shareable URL
- the team just needs browser access

### Option 2: Render
This repo already includes `render.yaml` and `start.sh`.

Steps:
1. Push this folder to GitHub.
2. In Render, create a new `Web Service` from the repo.
3. Render can detect `render.yaml`, or you can set manually:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `bash start.sh`
4. Deploy.

Recommended when:
- you want a simple Python hosting option
- you may later add custom domains or environment variables

## Important note about archive/history

The current history feature uses local SQLite:
- local file: `data/link_builder_archive.db`
- on your laptop: history persists normally
- on cloud free hosting: the file may be ephemeral and can reset after redeploy/restart

If you want shared team history that stays stable online, the next step is to move archive storage to:
- Supabase / Postgres
- MySQL
- a managed SQLite disk on your hosting platform

## Current capabilities
- campaign-level parameter input
- batch web/app link generation
- AppsFlyer bulk sheet preview
- archive/history with reopen, duplicate, update, delete, and search
- CSV/XLSX export
- copy-ready column/table output

## Recommended next upgrade
If you want this to behave like a true team tool instead of a personal tool deployed online, upgrade archive storage from local SQLite to a shared database.
