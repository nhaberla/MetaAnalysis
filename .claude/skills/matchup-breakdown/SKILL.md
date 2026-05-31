---
description: Builds a color-coded head-to-head win rate spreadsheet from real melee.gg tournament results for Star Wars Unlimited. Use when the user wants matchup data, head-to-head win rates, or a meta breakdown. Accepts optional --timeframe YYYY-MM-DD:YYYY-MM-DD and --num-decks N arguments.
---

Today's date: !`date +%Y-%m-%d`

Build a SWU matchup matrix using: $ARGUMENTS

## Steps

1. **Resolve arguments from `$ARGUMENTS`.**
   - `--timeframe`: default `2026-03-13:` + today's date above. Warn if start date is before `2026-03-13` (pre-rotation data uses different card legality).
   - `--num-decks`: default `12`.
   - `--min-players`: default `32`.
   - `--output`: default `matchup_matrix.xlsx`.

2. **Run the script** (stdlib only, no install needed):
   ```
   python scripts/matchup_breakdown.py --timeframe <timeframe> --num-decks <n> [other args]
   ```
   Pass `--verbose` on the first run of a session so API issues surface clearly.

3. **Handle failures.**
   - *403 / network error*: The environment may block outbound requests. Tell the user to run the script locally.
   - *404 / 0 results*: `SWU_GAME_SLUG` at the top of `scripts/matchup_breakdown.py` may be wrong. Ask the user to inspect melee.gg XHR traffic in DevTools to find the correct game filter value.
   - *No match data collected*: melee.gg may require auth for pairing/decklist endpoints. Suggest these fallbacks in order:
     1. Tournament organiser CSV export from melee.gg (raw data, cleanest option)
     2. swu-competitivehub.com structured pages
     3. swumetastats.com/api-docs public API (pre-aggregated, not raw — note this clearly)

4. **Summarise results** after a successful run:
   - Events and total matchup games processed
   - Top 3 archetypes by average win rate
   - Any standout matchups ≥60% WR with ≥20 games
   - Any archetypes with <50 total games (flag as low-confidence)
   - Cells marked `*` have <20 games — call these out

5. **Send the file** to the user with `SendUserFile`.

## Notes

- Archetypes are discovered dynamically from melee.gg decklist data (leader + base card names). No deck list is hardcoded; `--num-decks` picks the top N by game count.
- Swiss rounds only — top-cut is excluded to avoid selection bias.
- Mirror matches are excluded from the matrix.
