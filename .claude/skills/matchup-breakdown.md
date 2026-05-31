# Matchup Breakdown

Fetches real game-by-game tournament results from melee.gg (the official SWU tournament platform) and builds a color-coded head-to-head win rate matrix spreadsheet for the current meta.

## Usage

```
/matchup-breakdown [--timeframe YYYY-MM-DD:YYYY-MM-DD] [--num-decks N] [--min-players N] [--output filename.xlsx]
```

## Parameters

| Param | Default | Description |
|---|---|---|
| `--timeframe` | `2026-03-13:<today>` | Date range. Post-rotation data only begins 2026-03-13. |
| `--num-decks` | `12` | How many top archetypes to include in the matrix. |
| `--min-players` | `32` | Minimum event size. Matches the SWU Competitive Hub filter standard. |
| `--output` | `matchup_matrix.xlsx` | Output file path. |
| `--dry-run` | off | List qualifying tournaments without fetching match data. Useful for verifying API connectivity. |
| `--verbose` | off | Print each API request for debugging. |

## What the Script Does

1. Fetches qualifying SWU Premier tournaments from melee.gg within the timeframe
2. For each tournament, iterates over Swiss rounds (top-cut excluded to avoid win-skew)
3. For each pairing, fetches both players' decklists and identifies their archetype via leader + base card
4. Accumulates win/loss records per ordered archetype pair
5. Ranks archetypes by total game appearances, selects the top N
6. Computes head-to-head win rates and writes a color-coded Excel matrix

## Instructions for Claude

When this skill is invoked:

### 1. Resolve parameters

Parse any arguments from the invocation. If `--timeframe` is not given, use `2026-03-13:<today's date>`. Warn the user if the start date is before 2026-03-13 (pre-rotation data uses different card legality and should not be mixed).

### 2. Run the script

No installation needed — stdlib only (Python 3.8+).

```bash
cd /home/user/MetaAnalysis && python scripts/matchup_breakdown.py <args>
```

Always pass `--verbose` on the first run of a session so API issues surface clearly.

### 4. Handle API failures

If the script exits with a non-zero code, read the error message carefully:

- **"game slug may be incorrect"** → The `SWU_GAME_SLUG` constant at the top of `scripts/matchup_breakdown.py` needs updating. Tell the user to visit melee.gg in a browser, open DevTools → Network, navigate to the SWU tournament list, and find the XHR call to identify the correct game filter value. Then edit the constant.
- **"No match data collected"** → melee.gg may not expose pairing/decklist data publicly for these events. Suggest the fallback workflow below.
- **Rate limit / 429** → Add `--verbose` and retry with a longer `REQUEST_DELAY` (edit the constant in the script).

### 5. Fallback if melee.gg API is unavailable

If direct API access fails, tell the user about these alternatives ranked by quality:

1. **melee.gg web scraping**: melee.gg tournament pages are public. Tournament IDs appear in URLs (`/Tournament/View/{id}`). The script can be extended to scrape standings pages — ask the user if they'd like that approach.
2. **SWU Competitive Hub** (swu-competitivehub.com): Has structured tournament results. Check if they expose a public API or scrapable pages.
3. **Manual export**: melee.gg allows tournament organizers to export results as CSV. If the user has organizer access or knows organizers, this is the cleanest path to raw data.
4. **swumetastats.com API** (`swumetastats.com/api-docs`): Public, no auth. Returns aggregated matchup data — **not raw games**, but useful for cross-checking. Note clearly that this is aggregated, not game-by-game.

### 6. Interpret and report results

After a successful run, summarize:

- How many events and total games were processed
- Top 3 archetypes by average win rate
- Any standout favorable matchups (≥60% WR with ≥20 games)
- Any cells marked `*` (low sample, <20 games) that should be treated cautiously
- Whether any archetypes had fewer than 50 total matchup games — flag as low-confidence overall

Then send the `.xlsx` file to the user with `SendUserFile`.

## Data Source Notes

| Source | Role | Auth |
|---|---|---|
| melee.gg API | Primary — raw pairings + decklists | None (public data) |
| swu-competitivehub.com | Cross-check / fallback aggregation | None |
| swumetastats.com/api-docs | Archetype list reference, cross-check | None |
| The SWU Report (swu.report) | Narrative context for anomalous results | N/A |

## Important Caveats

- **Post-rotation only**: Filter start date to no earlier than 2026-03-13. Pre-rotation data (SOR/SHD/TWI era) uses different card pools.
- **Swiss rounds only**: Top-cut rounds are excluded. Top-8 brackets have strong selection bias (only the best records play in cut), which inflates win rates for popular decks.
- **Mirror matches**: Excluded from the matrix (always 50/50 by definition at scale).
- **Low sample**: Any cell with `*` has <20 games. The Vader/Aurra matchup is a known extreme (78% in small samples); treat outliers skeptically.
- **Event weighting**: The script treats all events equally. For deeper analysis, consider weighting Sector Qualifiers (200-500p) and Regionals (500-700p) more heavily than PQs.
- **Archetype discovery**: Archetypes are derived purely from tournament data — the script reads each player's leader and base card name from their melee.gg decklist and forms `"Leader / Base"` labels automatically. `--num-decks` controls how many top archetypes (ranked by total game count) appear in the matrix. No deck list is hardcoded.
