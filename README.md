# BRIMI Web App

BRI-managed investment fund performance and AUM reporting pipeline. Produces ranked fund tables by category and performance summaries for BRI-managed funds.

## Architecture

```
User uploads 3 Excel files → Flask app → Compiled workbook → Output Excel
                              │
                              └─ Fetches peer/index data from Investdata API
```

### Pipeline (3 steps)

1. **API Fetch** (`brimi_engine.py:fetch_api_data`) — calls Infovesta API for latest peer fund NAV/performance and index benchmark data. Writes to temporary D-1, D-2, INDEKS Excel files in `/tmp`.

2. **Build Compiled** (`brimi_engine.py:build_compiled`) — merges 6 sources into one workbook:
   | Sheet | Source | Description |
   |---|---|---|
   | D-1 | API | Peer fund data (latest date) |
   | D-2 | API | Peer fund data (previous date) |
   | INDEKS | API | Benchmark/index data |
   | BRIMI D-1 | User upload | BRI HistoricalNAV (T-1) |
   | BRIMI D-2 | User upload | BRI HistoricalNAV (T-2) |
   | BloombergIndex | User upload | Bloomberg indexes |

3. **Process** (`process_brimi.py:process`) — reads compiled workbook, applies `fund_universe.json` config, ranks funds by quartile, writes output Excel.

### Modules

| File | Purpose |
|---|---|
| `api/index.py` | Flask web app + Vercel entrypoint. Upload page, `/process` endpoint, admin page |
| `brimi_engine.py` | Pipeline orchestrator: API fetch → compile → process |
| `process_brimi.py` | Core processor: page table + performance table generation with styling |
| `fetch_investdata.py` | Investdata API client (OAuth, NAV, AUM, scoring, index data) |
| `raw_to_compiled.py` | Standalone compiler for local dev (without web UI) |
| `manage_funds.py` | CLI for managing fund universe (add/remove funds) |
| `parity_check.py` | Validation script comparing output against reference |

### Config Files

| File | Purpose |
|---|---|
| `fund_universe.json` | Defines page_table sections (26 sections, 19 unique names) and performance_table funds (28 BRI-managed funds) |
| `fund_map.json` | Name mapping overrides for cross-source fund matching (5 entries) |
| `quartile_group_mapping.json` | Reference groups for quartile ranking |
| `vercel.json` | Vercel deployment config (`@vercel/python` builder, fluid compute) |

### Peer Override System

Admin page (`/admin`) allows editing peer lists without touching code:
- Edits saved to `peer_overrides.json` via GitHub API (commits to git)
- Format: `{"0": [funds], "5": [funds]}` — dict keyed by section index
- `process_brimi.py` reads and applies overrides at line 223
- Only changed sections are stored; auto-deletes if all match base

## Local Development

```bash
pip install -r requirements.txt
python api/index.py          # Flask dev server on :5000
python process_brimi.py compiled.xlsx  # Direct processing
python fetch_investdata.py   # Fetch latest API data to very_raw_data/
```

## Deployment

- **Platform**: Vercel Fluid Compute (Python 3.12)
- **Production URL**: https://brimimiperi.vercel.app
- **Admin**: https://brimimiperi.vercel.app/admin

See `.claude/WORKFLOW.md` for deployment rules and guidelines.

## Env Vars Required

| Variable | Purpose |
|---|---|
| `INVESTDATA_USERNAME` | Infovesta API auth |
| `INVESTDATA_PASSWORD` | Infovesta API auth |
| `INVESTDATA_CLIENT_ID` | Infovesta API OAuth client |
| `INVESTDATA_CLIENT_SECRET` | Infovesta API OAuth secret |
| `GITHUB_TOKEN` | GitHub API (PAT with `repo` scope) for admin save |
| `GITHUB_REPO_OWNER` | GitHub repo owner (`cevinkidambi`) |
| `GITHUB_REPO_NAME` | GitHub repo name (`brimi-webapp`) |
| `GITHUB_BRANCH` | Git branch (`main`) |
