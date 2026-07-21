# Infra Deal Pipeline

Monitors Google News RSS for infrastructure allocation headlines mentioning a watchlist
of ~50 global pension funds and sovereign wealth funds (LPs). Classifies each headline
with the Anthropic API, then appends new mentions to the "Market Intel" sheet of
`Infra_Deal_Pipeline.xlsx`. Safe to re-run — deduplicates by article URL.

Google News is the *retrieval* layer; Claude is the *filtering and enrichment* layer.
Claude catches indirect references the old regex missed ("Norway's oil fund"), rejects
noise the query pulls in (returns reports, personnel moves), and extracts deal size,
asset class, and counterparty into their own columns.

See parent `CLAUDE.md` for shared environment and package info.

## Key Files

| File | Purpose |
|---|---|
| `update_market_intel.py` | Main script — fetches RSS, filters by LP, appends to Excel |
| `update.bat` | Launcher for `update_market_intel.py` |
| `Infra_Deal_Pipeline.xlsx` | Main workbook (deal pipeline + Market Intel sheet) |
| `Infra_Deal_Pipeline (First Version).xlsx` | Archived first version |

## How to Run

```
# Via launcher:
update.bat

# Direct:
python update_market_intel.py

# Regex matching only — no API calls, no key needed:
python update_market_intel.py --no-ai

# Override the model:
python update_market_intel.py --model claude-sonnet-5
```

The script reads and writes `Infra_Deal_Pipeline.xlsx` **from its own directory**
(resolved via `__file__`), so moving the project folder is safe. Keep the workbook
next to the script.

## How It Works

1. Sends several themed Google News RSS queries (general infra / renewables+energy /
   digital / utilities), each an AND of a fund term, a sector term, and an action word,
   then merges and dedupes them by URL (~250-300 headlines per run). See
   `search_terms.md` for the sector vocabulary and `SECTOR_THEMES` in the script.
2. Sends every headline to Claude in **one batched call** with the watchlist, using
   structured outputs (`messages.parse` + a pydantic schema) so the response is
   validated before it reaches Excel. Claude returns, per headline: investor,
   event type, deal size, asset class, counterparty, and confidence.
3. Appends matched articles to the "Market Intel" sheet.
4. Skips URLs already present in the sheet (deduplication).

Headlines where Claude identifies no watchlist investor are dropped. Everything else
is written and *labelled* rather than filtered — noise like `returns_report` lands in
the sheet with that tag so you can filter in Excel instead of losing the row silently.

## Fallback Behaviour

The tool degrades to the original regex alias matching (leaving enrichment columns
blank) when: `--no-ai` is passed, `ANTHROPIC_API_KEY` is unset, or the API call fails.
It never hard-fails on a Claude error — you always get the Google News rows.

Note the interaction with dedupe: a row added by a regex-fallback run is deduped by
URL and will **not** be re-enriched on a later AI run. Prefer `--no-ai` for testing
against a throwaway copy of the workbook, not the real one.

## Market Intel Columns

| # | Column | Source |
|---|---|---|
| 1-7 | Date, Investor, Investor Type, AUM ($M), Headline, Source, URL | RSS + watchlist |
| 8 | Event Type | Claude — `commitment`, `direct_investment`, `acquisition`, `allocation_target`, `mandate`, `returns_report`, `personnel`, `other` |
| 9 | Deal Size ($M) | Claude — null unless stated in the headline |
| 10 | Asset Class | Claude — `digital_infrastructure`, `renewables`, `energy`, `transport`, `utilities`, `social_infrastructure`, `diversified`, `unknown` |
| 11 | Counterparty | Claude — manager/fund/asset on the other side |
| 12 | Confidence | Claude — `high`, `medium`, `low` |

Columns 8-12 were added after the sheet already had data. The script migrates the
header automatically on first run; pre-existing rows keep blank enrichment cells.

## Watchlist

Edit `WATCHLIST` in `update_market_intel.py` to add/remove LPs.
Each entry is a tuple: `(Display Name, Type, AUM_in_$M, [alias list])`.

Current watchlist covers ~50 institutions including:

**Sovereign Wealth Funds**: Norway GPFG, CIC, ADIA, KIA, PIF, GIC, QIA, Temasek,
Mubadala, KIC, Future Fund, NZ Super Fund, Khazanah, Alaska Permanent Fund

**Canadian Pensions**: CPP Investments, CDPQ, Ontario Teachers, PSP Investments,
BCI, AIMCo, OMERS, HOOPP, IMCO

**US Pensions**: CalPERS, CalSTRS, NY State Common, NYC Retirement, Florida SBA,
Texas TRS (and others)

## Google News Queries

Each query is an AND of three groups:

```
(fund term)  AND  (sector term — one theme)  AND  (action word)
```

- `FUND_GROUP` and `ACTION_GROUP` are shared across all themes.
- `SECTOR_THEMES` holds the per-theme sector vocabulary (general / renewables_energy /
  digital / utilities). Each theme is fetched as its own query and results are merged
  and deduped by URL, because Google News caps a single query at ~100 results.
- Multi-word sector terms **must stay quoted** — Google reads an unquoted space as AND.

To broaden/narrow: edit `SECTOR_THEMES` (add/remove terms or a whole theme). The full
vocabulary, including terms held in reserve, lives in `search_terms.md`.

## Dependencies

```
pip install openpyxl anthropic
```

`anthropic` 0.95.0 and `pydantic` 2.12.5 are installed and support the
`messages.parse` + `output_format` structured-output path this script relies on.

## API Key

Resolved by `load_api_key()` in this order:

1. `ANTHROPIC_API_KEY` (or `ANTHROPIC_AUTH_TOKEN`) environment variable
2. `%USERPROFILE%\.anthropic\api_key.txt` — first non-comment, non-blank line

Without either, the script prints a notice and falls back to regex matching rather
than failing. The key is never printed to the console or written into the workbook.

**The key file is deliberately outside this folder.** This project lives under
OneDrive, so anything stored here syncs to Microsoft's cloud and to every device on
the account. `~\.anthropic\` is local disk only, and the file is ACL'd to the current
user (`icacls /inheritance:r`). Do not move the key into the project directory, and
do not hardcode it into the script.

To rotate: overwrite the line in `api_key.txt`, or revoke the key at
https://console.anthropic.com/settings/keys

## Cost

One batched call per run. At ~85 headlines that's roughly 4-6k input tokens and 3-5k
output tokens on `claude-opus-4-8` — a few cents per run. The watchlist prompt is
stable but sits below the 4096-token minimum cacheable prefix for Opus, so prompt
caching is not wired up; it would not pay off at this size.
