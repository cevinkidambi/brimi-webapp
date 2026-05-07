# BRIMI Web App — Project Context

## What This Project Does

Produces ranked investment fund performance reports for BRI (Bank Rakyat Indonesia). Two outputs:
1. **Page Table** — full ranked fund table by category (equity, fixed income, money market, etc.)
2. **Performance Table** — BRI-managed funds with Score & Quartile

## Data Flow

```
User uploads:                          API (auto):
  1. HistoricalNAV T-1 (BRI)            1. D-1 peer fund data
  2. HistoricalNAV T-2 (BRI)            2. D-2 peer fund data (prev day)
  3. indeks_bloomberg.xlsx              3. INDEKS benchmark data
                                               ↓
                                      Compiled Excel (6 sheets)
                                               ↓
                                  process_brimi.py → Output Excel
```

**Key distinction**: There are TWO sets of D-1/D-2:
- **D-1 / D-2** sheets → peer fund data from API (auto-generated in `/tmp`)
- **BRIMI D-1 / BRIMI D-2** sheets → BRI fund data from user-uploaded HistoricalNAV Excel

Never confuse them. BRIMI = BRI funds from upload. D-1/D-2 (without "BRIMI") = peers from API.

## Module Map

| Module | What it does | Called by |
|---|---|---|
| `api/index.py` | Flask web server, upload UI, admin page, `/process` endpoint | Vercel / browser |
| `brimi_engine.py` | Orchestrates: fetch API → build compiled workbook → run process_brimi | `api/index.py` `/process` |
| `process_brimi.py` | Reads compiled Excel, ranks funds, writes styled output Excel | `brimi_engine.py` or CLI |
| `fetch_investdata.py` | OAuth + API calls to Infovesta (NAV, AUM, scoring, index) | `brimi_engine.py` or CLI |
| `raw_to_compiled.py` | Standalone compiler for local dev (no web UI) | CLI only |
| `manage_funds.py` | CLI for managing fund_universe.json | CLI only |

## Config Files

| File | Purpose |
|---|---|
| `fund_universe.json` | 26 page_table sections (19 unique names, some duplicated like "Money Market" x4). 28 performance_table BRI funds. |
| `fund_map.json` | Name mapping overrides for cross-source matching. 5 entries. Handles "Danareksa" → "BRI" rebranding, casing diffs, etc. |
| `quartile_group_mapping.json` | Reference groups for quartile ranking (26 groups) |
| `vercel.json` | Vercel config. Must have `"fluid": true` and `builds` array with `@vercel/python`. |
| `.gitignore` | Excludes `peer_overrides.json` (committed via admin page API) |

## Peer Override System

Admin page at `/admin` lets users add/remove peer funds per section:
- Loads config from `fund_universe.json` + any existing `peer_overrides.json`
- Saves diffs only (not full copy) — compares by section **index**, not name (handles duplicate section names)
- Commits `peer_overrides.json` to GitHub via API
- Format: `{"0": [funds], "5": [funds]}` — dict keyed by section index
- `process_brimi._apply_peer_overrides()` reads and applies at runtime
- If all sections match base, overrides file is deleted from GitHub

## Deployment

- **Platform**: Vercel Fluid Compute
- **Python**: 3.12
- **Framework**: Flask (not Next.js)
- **Build config**: `"fluid": true` in `vercel.json` — CRITICAL, without it Python builder skips
- **Production URL**: https://brimimiperi.vercel.app

## Env Vars (Vercel project settings)

All 8 vars set on both Production and Preview (fix-token-cache branch). See README.md for list.
