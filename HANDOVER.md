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
git clone https://github.com/cevinkidambi/brimi-webapp.git
cd brimi-webapp
pip install -r requirements.txt
cp investdata_api/.env.example investdata_api/.env   # fill real creds
python app.py          # local server on :5001, uses ./tmp
```

Local processing of an existing compiled workbook:
```bash
python process_brimi.py tmp/compiled_input.xlsx out.xlsx
```

---

## 4. Deployment (CRITICAL RULES)

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

- [ ] **Vercel**: add team members to the `brimi-webapp` project (Settings → Members / Team). Owner / account `cevinkidambi` (login `safourkidambi@gmail.com`).
- [ ] **GitHub**: transfer/maintain access to the `cevinkidambi/brimi-webapp` repo (owner `cevinkidambi`).
- [ ] **Infovesta API creds**: the values in `investdata_api/.env` (`INVESTDATA_USERNAME`/`PASSWORD`) and the same 4 `INVESTDATA_*` Vercel env vars. Reissue if the departing user is the credential owner.
- [ ] **Vercel env vars**: confirm all 8 vars are present on Production AND Preview, and update them if credentials rotate.
- [ ] **GITHUB_TOKEN**: a PAT with `repo` scope used by the admin save endpoint. If owned by the departing user's GitHub account, it must be reissued as a machine/team token, or admin saves will break.
- [ ] **Deploy access**: developer or admin role on Vercel CLI.

---

*Generated at handover. CLI account: `cevinkidambi` (Vercel login `safourkidambi@gmail.com`). Branch `main` is deployable; current live config = fixed-income duration sections + fund renames (BRI-MI Anagata) + Index & ETF banner.*