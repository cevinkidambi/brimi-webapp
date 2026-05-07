# BRIMI Web App — Workflow & Operational Rules

## Deployment Rule: Preview First, Always

**NEVER deploy directly to production.** Always test on preview first.

```bash
# Preview deploy (from feature branch or main)
vercel --yes

# After verifying on preview, promote to production
vercel --prod --yes
```

The production URL https://brimimiperi.vercel.app is user-facing. Broken deployments go directly to users.

## Data Sources — Do Not Confuse

| Source | What it provides | How it arrives |
|---|---|---|
| **Investdata API** | Peer fund NAV, AUM, scoring, index data | Auto-fetched at runtime via `fetch_investdata.py` |
| **HistoricalNAV (user upload)** | BRI fund AUM data | User uploads Excel files on the web UI |
| **indeks_bloomberg.xlsx (user upload)** | Bloomberg global indexes | User uploads Excel file on the web UI |

**Two sets of D-1/D-2 exist:**
- `D-1` / `D-2` sheets = peer data from API (auto-generated in `/tmp`)
- `BRIMI D-1` / `BRIMI D-2` sheets = BRI fund data from uploaded HistoricalNAV

Never use "D-1" without the "BRIMI" prefix when referring to BRI upload data. The naming overlap causes bugs.

## Vercel Configuration

- `vercel.json` must have `"fluid": true` — without it, the Python builder is skipped entirely (builds complete in ~100ms instead of ~2s)
- Python version detected automatically (3.12), no `.python-version` file needed
- Use `/tmp` for all working files — `/var/task` is read-only on Vercel

## Env Vars

Set via `vercel env add`. Required on both Production and Preview environments:
- `INVESTDATA_*` (4 vars) — API credentials
- `GITHUB_*` (4 vars) — Admin page commit to GitHub

**Do NOT set env vars with `<<<` heredoc** — it adds trailing newlines. Use `printf '%s' "value" | vercel env add ...` instead.

## Admin Page

- URL: `/admin` — peer group management
- Save endpoint commits to GitHub via API, requires `GITHUB_TOKEN` env var
- Saves **only diffs** vs `fund_universe.json` — compares by section index
- New funds added via admin only work if they exist in the API's D-1/D-2 data (peer funds)
- If a fund is not in the API data, it will show empty performance

## Testing

- Main page must return 200 with upload form
- `/admin/config` must return 200 with 26 sections
- After any change, test: upload page loads → admin page loads → admin/config returns data
- For SSO-protected deployments, use `vercel url get_access` or MCP `get_access_to_vercel_url` for authenticated testing

## When Adding/Removing Funds

1. Use `manage_funds.py --help` for CLI management of `fund_universe.json`
2. Or use the admin page at `/admin` for peer groups
3. If a fund name differs between BRI data and API data, add mapping to `fund_map.json`
4. Deploy and verify output contains the new/removed fund

## Common Pitfalls

| Symptom | Likely cause |
|---|---|
| Build completes in ~100ms | Missing `"fluid": true` in vercel.json |
| 401 on all pages | Vercel SSO / Deployment Protection |
| 404 after deploy | Python builder skipped, no output generated |
| Admin save returns "Read-only file system" | Missing `GITHUB_TOKEN` env var (falls back to local file) |
| Env var not working in deployed app | Trailing newline from `<<<` heredoc |
| Duplicate sections get same overrides | Must match by index, not name (fixed) |
