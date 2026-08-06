# BRIMI Web App — Handover & Operations Runbook

This doc hands the project to the team. Read it end-to-end once, then use the
relevant sections per task. Accounts/credentials cannot be transferred from a
repo — see the Access Transfer Checklist at the end.

- **Production**: https://brimimiperi.vercel.app
- **Admin (peer group editor)**: https://brimimiperi.vercel.app/admin
- **Repo**: https://github.com/cevinkidambi/brimi-webapp
- **Platform**: Vercel Fluid Compute (Python 3.12), Flask

---

## 1. What the app does

Weekly reporting pipeline for BRI-managed investment funds. A user uploads
3 Excel files; the app pulls peer/index data from the Infovesta API, compiles
everything, and returns a styled Excel report with two sheets:

1. **NEW PAGE TABLE (output)** — ranked fund tables grouped by category
   (Equity, Money Market, Balanced, Fixed Income, discretionary/insurance).
2. **NEW PERFORMANCE TABLE (output)** — BRI-managed funds with Score & Quartile.

Top-group banners (e.g. "Fixed Income Fund", "Index & ETF Fund") are derived
from section names via `top_group()` in `process_brimi.py`.

---

## 2. Weekly report run (operator)

User-facing flow. No code changes normally needed.

1. Open production URL, upload the 3 files:
   - `HistoricalNAV` T-1 (BRI, newest date)
   - `HistoricalNAV` T-2 (BRI, previous date)
   - `INDEKS Bloomberg` (Bloomberg indexes, sheet `Sheet1`)
2. The app auto-fetches Investdata API data (D-1/D-2 peers, AUM, scoring,
   INDEKS) and returns `BRIMI_Output_<date>.xlsx` (via `/download`).
3. Download and check the two output sheets.

**Do NOT confuse the two D-1/D-2 pairs:**
- `D-1` / `D-2` sheets = peer funds from the API.
- `BRIMI D-1` / `BRIMI D-2` sheets = BRI funds from the uploaded HistoricalNAV.

---

## 3. Setup a developer's machine

```bash
git clone <your-repo-url>        # repo moves to the team org — clone from there
cd brimi-webapp
pip install -r requirements.txt
cp investdata_api/.env.example investdata_api/.env   # fill real creds
python app.py          # local server on :5001, uses ./tmp
```

Alternative via Docker (host-agnostic, no local Python needed):
```bash
docker build -t brimi .
docker run -p 5001:5001 -e INVESTDATA_USERNAME=... -e INVESTDATA_PASSWORD=... brimi
```

Local processing of an existing compiled workbook:
```bash
python process_brimi.py tmp/compiled_input.xlsx out.xlsx
```

---

## 4. Deployment (CRITICAL RULES)

These rules are for the CURRENT Vercel hosting. See §9b for moving to any other host.

- **NEVER deploy straight to production.**
- **Pushing to `main` auto-deploys production** via Vercel GitHub integration.
  So test on preview BEFORE pushing to main.

Local deploy flow (deploys local working-tree files, including uncommitted ones):
```bash
vercel --yes          # preview deploy → SSO-protected URL
# verify preview: main page 200, /admin/config returns sections
vercel --prod --yes   # only after verifying preview
```

- `vercel.json` MUST keep `"fluid": true`. Without it the Python builder is
  skipped and deploys complete in ~100ms with no output.
- Working files must use `/tmp`. `/var/task` is read-only on Vercel.
- Preview URLs are SSO-protected (Vercel Authentication) — open in a browser
  logged into Vercel.
- Set/update env vars on **both** Production and Preview:
  `printf '%s' "value" | vercel env add <name> production` (and `preview`).
  Never use `<<<` heredoc — it appends a trailing newline that breaks values.

---

## 5. Config files that matter

| File | What it controls |
|---|---|
| `fund_universe.json` | `page_table` = the 26 ranked sections + their funds; `performance_table` = 28 BRI funds w/ Score & Quartile config |
| `fund_map.json` | Cross-source name re-mapping (BRI data name vs API name, base vs Total Return variant, Darlink alias, etc.) |
| `quartile_group_mapping.json` | 26 reference peer groups used to compute quartiles |
| `insurance.xlsx` | Historical AUM series for the 3 insurance funds' rolling returns (keep in git) |
| `.claude/CLAUDE.md` | Project context reference (data flow, modules, override system) |
| `.claude/WORKFLOW.md` | Deployment rules + pitfalls table |

### Key business rules baked into `process_brimi.py`
- **Total Return (`*`) funds use base NAV** for NAV but Total Return variant
  for performance columns (TR-variant lookup logic in `get_perf`, `lookup_*`).
- **Since Inception** uses base NAV 1000 (IDR) / 1 (USD).
- **Bloomberg values** are decimal fractions and are ×100 to percent.
- **Insurance funds** compute rolling returns from `insurance.xlsx` historical
  AUM using trading-day windows (1Hr=1, 1Mgg=5, 1Bln=19, 3Bln=56, 6Bln=119, 1Thn=235).
- **Units**: IDR AUM in Rp Miliar (÷1e9), USD AUM in Juta USD (÷1e6).

---

## 6. Changing funds

Three ways, in order of preference:

1. **Admin page** `/admin` — edit peer fund lists per section. Saves ONLY diffs
   to `peer_overrides.json`, committed to GitHub via API (needs `GITHUB_TOKEN`).
   Overrides are keyed by **section index**.
2. **CLI** `manage_funds.py` — add/remove/rename funds in `fund_universe.json`.
   See `python manage_funds.py --help` (`--add-page`, `--add-perf`, `--rename`,
   `--alias`, `--brimi-alias`, `--is-index`, `--usd`, `--price-return`, ...).
3. **Direct edit** of `fund_universe.json` (careful with duplicate section names).

If a fund name differs between the BRI upload and the API/peer data, add a
mapping to `fund_map.json`.

> After rearranging sections, check for a stale `peer_overrides.json` on GitHub
> — its indices may no longer match and will clobber the new sections. Regrouped
> sections should have all funds baked into `fund_universe.json` and the stale
> override file deleted.

---

## 7. Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| Build completes in ~100ms | Missing `"fluid": true` in vercel.json |
| 401 on pages | Vercel SSO / Deployment Protection — open in logged-in browser |
| 404 after deploy | Python builder skipped (see fluid above) |
| Admin save "Read-only file system" | Missing `GITHUB_TOKEN` env var |
| Env var looks wrong at runtime | Trailing newline from `<<<` heredoc |
| Fund shows empty performance | Fund not present in API D-1/D-2 data, or missing `fund_map.json` alias |
| Same override applied to 2 sections | Overrides matched by name; must be keyed by index |

Parity of output vs a reference Excel: `python parity_check.py <output> <reference>`

---

## 8. Access Transfer Checklist (TODO — accounts can't be moved by repo)

Hand these to the new maintainer / manager:

- [ ] **GitHub**: transfer `cevinkidambi/brimi-webapp` to the team org (Settings → Danger Zone → Transfer ownership).
- [ ] **Host account**: transfer/re-provision the hosting project (currently Vercel — `cevinkidambi`, login `safourkidambi@gmail.com`) OR set up a new host (see §9b).
- [ ] **Infovesta API creds**: REISSUE fresh `INVESTDATA_USERNAME`/`PASSWORD` on the team's Infovesta account — the current ones belong to the departing employee and will stop working.
- [ ] **Host env vars**: set all 8 vars (`INVESTDATA_*` + `GITHUB_*`) on the new host from the new creds.
- [ ] **GITHUB_TOKEN**: reissue as a machine/team PAT with `repo` scope (a personal PAT of the departing account dies with it, and admin saves break).
- [ ] **Deploy access**: developer/admin role on the new host for the new maintainer.

## 9. Account / Domain Migration (IMPORTANT — hand the codebase + infrastructure)

**The GitHub and Vercel accounts are personal (`cevinkidambi` / login
`safourkidambi@gmail.com`) and will be retired.** The team must take over BOTH the
codebase (the git repo) AND the infrastructure (deployment config, env vars, secrets,
and the hosting account). The current production deployment stays live in the meantime,
so there is no deadline pressure — the new host goes up first, then it is switched over.

**The app is NOT Vercel-locked.** It is a plain Python/Flask app. `vercel.json` + the
Vercel CLI are simply one way to host it (used here only because that was the departing
engineer's choice). The exact same code runs on any WSGI-capable host via
`gunicorn app:app` (see `Procfile`) or as a container (see new `Dockerfile`). So the team
is free to pick any host (AWS, GCP, Render, Railway, a VM, K8s, etc.). What MUST be
transferred are the **credentials and the deploy configuration**, listed below.

### What "codebase + infra" means here (transfer checklist)

**Codebase** (portable by definition — it is in git):
- Everything in `https://github.com/cevinkidambi/brimi-webapp.git` — code,
  `fund_universe.json`, `fund_map.json`, `quartile_group_mapping.json`,
  `insurance.xlsx`, `requirements.txt`, `Procfile`, `vercel.json`, docs.

**Infrastructure** (NOT in git — must be recreated / reissued on the team side):
- Host / project (currently Vercel) + its env config.
- The 8 env vars (credentials) — `INVESTDATA_*` (4) + `GITHUB_*` (4).
- `investdata_api/.env` (real Infovesta creds) — gitignored.
- A `GITHUB_TOKEN` PAT (owner of the retiring account, must be reissued).
- A custom domain, if any.

### 9a. GitHub — transfer the repo (source of truth for code)

1. **Create the destination**: a new GitHub org (recommended) or a team member's account.
2. **Transfer the repo**: GitHub → repo `Settings` → `Danger Zone` → `Transfer ownership`
   → new owner. History and branches come along; old URL auto-redirects.
3. **Re-issue `GITHUB_TOKEN`**: the PAT that `/admin` save uses must be owned by the new
   org/a service account (a personal PAT under the retiring account dies with it).
4. **Update `GITHUB_REPO_OWNER`** wherever the new owner differs — the code default is
   hard-coded in `api/index.py` lines 630/661 (`cevinkidambi`). Prefer the env var; if
   the org name changes, also update the code default + `.env.example`.
5. Team clones/pushes as normal afterward.

### 9b. Re-provision the project on a NEW host (Vercel is one option)

The team can host anywhere. Two paths:

**If the team stays on Vercel:**
1. **Create the destination**: new Vercel team + project on the team account.
2. A fresh project gets a NEW default domain — update every reference to the old URL
   in README / this doc / end users, and re-attach any custom domain (Settings → Domains).
3. Recreate deployment settings: `vercel.json` (keep `"fluid": true`) is already in the
   repo; re-link the GitHub integration for auto-deploy.
4. **Re-add all 8 env vars** (Production AND Preview): `INVESTDATA_*` (4) + `GITHUB_*` (4).
   See `.env.example`.
5. Deploy + verify: main page 200, `/admin` loads, `/admin/config` returns the 26 sections.

**If the team uses ANY other host (recommended since Vercel was just the chooser):**
1. It is a standard Flask app — no Vercel-specific code beyond `vercel.json`, which is
   ignored on other hosts.
2. Run it with the bundled `Procfile` (`gunicorn app:app`) or a container
   (`Docker build -t brimi . && docker run -p 5001:5001 brimi`).
3. Provide the same 8 env vars as environment variables on that host. Where a host
   stores secrets (e.g. a `.env` file or a secrets manager), mirror
   `.env.example` / `investdata_api/.env`.
4. Share the new URL — the app has no hard-coded domain dependency, only what you point
   at it.

In BOTH cases the domain WILL change from `https://brimimiperi.vercel.app`.

### 9c. Secrets — will CHANGE, so re-issue (not just copy)

Transfer these out-of-band (secure channel, never in a public drive link).
**Note: the current `INVESTDATA_*` credentials belong to the departing user's
employee account and WILL be replaced.** Do not plan to keep them working.

- `investdata_api/.env` (gitignored): real `INVESTDATA_USERNAME`/`PASSWORD`.
  **Request fresh Infovesta API creds** on the team's Infovesta account.
- `investdata_api/token.json` (gitignored): cached OAuth token — safe to delete; the
  app re-fetches via `.env`.
- The 8 env var values on the host — set the NEW creds, not the old ones.
- `GITHUB_TOKEN`: a PAT with `repo` scope used by `/admin` save. Must be a machine/team
  token, not a personal PAT of the departing account.

Local CLI deploys on the new account also land on the new domain — run
`vercel link` (or `vercel --token <team-token>`) under the team account once.

---

*Generated at handover. Current hosting: Vercel, account `cevinkidambi` (login
`safourkidambi@gmail.com`), branch `main` is deployable; live config = fixed-income
duration sections + fund renames (BRI-MI Anagata) + Index & ETF banner. The current
deployment stays live until the team's new host is up. Infovesta API creds WILL be
re-issued (departing employee account); GitHub and Vercel accounts both retire.*