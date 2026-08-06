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
| `fund_universe.json` | Defines page_table sections (26 sections, incl. 3 duration-based fixed income) and performance_table funds (28 BRI-managed funds) |
| `fund_map.json` | Name mapping overrides for cross-source fund matching |
| `quartile_group_mapping.json` | Reference groups for quartile ranking (26 groups) |
| `insurance.xlsx` | Historical AUM series for insurance-fund rolling returns (tracked in git) |
| `vercel.json` | Vercel deployment config (`@vercel/python` builder, fluid compute) |

### Peer Override System

Admin page (`/admin`) allows editing peer lists without touching code:
- Edits saved to `peer_overrides.json` via GitHub API (commits to git)
- Format: `{"0": [funds], "5": [funds]}` — dict keyed by section **index**, not name (handles duplicate section names)
- `process_brimi._apply_peer_overrides()` reads and applies overrides at runtime (~line 269)
- Only changed sections are stored; auto-deletes the file if all sections match base
- New funds added via admin only work if they exist in the API D-1/D-2 data; otherwise they show empty performance

> **Note on deploy flow**: `peer_overrides.json` is gitignored locally, so a local `vercel` CLI deploy carries no overrides file — the committed GitHub copy is what a git-integration deploy uses. When regrouping sections, all funds are baked into `fund_universe.json` and the stale overrides file deleted to avoid index-mismatch clobbering.

## Local Development

```bash
pip install -r requirements.txt
cp investdata_api/.env.example investdata_api/.env   # then fill real credentials
python app.py           # Flask dev server on :5001 (uses templates/)
# or, processing a pre-compiled workbook directly:
python process_brimi.py compiled.xlsx  # CLI: <input> [output] [universe] [fund_map]
```

Local runs write to `./tmp/`. The Vercel entrypoint `api/index.py` differs from local `app.py` (API
entrypoint inlines all HTML; app.py renders `templates/index.html`).

## Deployment

- **Platform**: Vercel Fluid Compute (Python 3.12)
- **Production URL**: https://brimimiperi.vercel.app
- **Admin**: https://brimimiperi.vercel.app/admin
- **GitHub integration**: pushing to `main` auto-deploys production. Use local preview deploys to test before any prod push.
- **Account/domain migration**: the project is owned by a personal Vercel account
  (`cevinkidambi`). On handover the project moves to the team account and the
  production domain WILL change. See `HANDOVER.md` §9 for the full migration runbook
  (re-provision project, re-add all 8 env vars, re-attach custom domain + DNS, re-link
  GitHub, re-verify, then update the URL everywhere).

### Deploy workflow (preview first, always)

```bash
vercel --yes           # preview deploy (SSO-protected URL)
# verify: main page 200, /admin/config returns sections
vercel --prod --yes    # promote to production only after verifying
```

- `vercel.json` MUST have `"fluid": true` — without it the Python builder is skipped (build completes in ~100ms).
- Use `/tmp` for working files, not `/var/task` (read-only on Vercel).
- Preview deployments are protected by Vercel Authentication (SSO) — open them in a logged-in browser.
- `vercel env add`: use `printf '%s' "value" | vercel env add <name> ...` — a `<<<` heredoc adds a trailing newline that breaks the value.

See `.claude/WORKFLOW.md` and `HANDOVER.md` for operational details.

## Env Vars Required

Set on **both** Production and Preview Vercel environments. Placeholders + the same set are in `.env.example`.

| Variable | Purpose | Default/local |
|---|---|---|
| `INVESTDATA_USERNAME` | Infovesta API auth | from `investdata_api/.env` |
| `INVESTDATA_PASSWORD` | Infovesta API auth | from `investdata_api/.env` |
| `INVESTDATA_CLIENT_ID` | Infovesta API OAuth client | `api2` |
| `INVESTDATA_CLIENT_SECRET` | Infovesta API OAuth secret | `api2` |
| `GITHUB_TOKEN` | GitHub API (PAT with `repo` scope) for admin save | — |
| `GITHUB_REPO_OWNER` | GitHub repo owner | `cevinkidambi` |
| `GITHUB_REPO_NAME` | GitHub repo name | `brimi-webapp` |
| `GITHUB_BRANCH` | Git branch to commit overrides | `main` |

## Common Pitfalls

| Symptom | Likely cause |
|---|---|
| Build completes in ~100ms | Missing `"fluid": true` in vercel.json |
| 401 on all pages | Vercel SSO / Deployment Protection |
| 404 after deploy | Python builder skipped, no output generated |
| Admin save "Read-only file system" | Missing `GITHUB_TOKEN` env var (falls back to local file) |
| Env var not working in deployed app | Trailing newline from `<<<` heredoc |
| Duplicate sections get same overrides | Must match by index, not name (fixed) |
