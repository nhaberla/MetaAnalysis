---
description: Builds a stacked area trend chart showing how each deck's meta share percentage has evolved over time, using real melee.gg tournament results for Star Wars Unlimited. Use when the user wants meta trends, meta share over time, deck popularity trends, or how the meta has evolved. Accepts optional --timeframe YYYY-MM-DD:YYYY-MM-DD and --num-decks N arguments.
---

Today's date: !`date +%Y-%m-%d`

Build a SWU meta share trend chart using: $ARGUMENTS

## Steps

1. **Resolve arguments from `$ARGUMENTS`.**
   - `--timeframe`: default is the last 30 days (computed by the script at runtime). Warn if start date is before `2026-03-13` (pre-rotation data uses different card legality).
   - `--num-decks`: default `10` (top N archetypes shown individually; all remaining decks are grouped into "Other").
   - `--min-players`: default `32`.
   - `--output`: default `meta_trend.xlsx`.

2. **Run the script** (stdlib only, no install needed):
   ```
   python ${CLAUDE_SKILL_DIR}/meta_breakdown.py --timeframe <timeframe> --num-decks <n> [other args]
   ```
   Pass `--verbose` on the first run of a session so API issues surface clearly.

3. **Handle failures.**
   - *403 / network error*: The environment may block outbound requests. Tell the user to run the script locally.
   - *404 / 0 results*: `SWU_GAME_SLUG` at the top of the script may be wrong. Ask the user to inspect melee.gg XHR traffic in DevTools to find the correct game filter value.
   - *No match data collected*: melee.gg may require auth for pairing/decklist endpoints. Suggest these fallbacks in order:
     1. Tournament organiser CSV export from melee.gg (raw data, cleanest option)
     2. swu-competitivehub.com structured pages
     3. swumetastats.com/api-docs public API (pre-aggregated, not raw — note this clearly)

4. **Summarise results** after a successful run:
   - Events and total deck appearances processed
   - Top 3 archetypes by overall meta share across the full period
   - Notable trends: archetypes that grew or shrank by ≥5 percentage points from first to last week
   - Size of the "Other" bucket — if it exceeds 20%, suggest increasing `--num-decks`
   - Number of weekly data points in the chart

5. **Send the file** to the user with `SendUserFile`.

## Notes

- Meta share per week = (archetype appearances that week) / (total deck appearances that week) × 100.
- Archetypes are discovered dynamically from melee.gg decklist data. No deck list is hardcoded.
- Decks outside the top N are grouped into "Other" so the stacked areas always sum to 100%.
- Data is bucketed by calendar week (Monday-start). Each tournament is assigned to the week of its start date.
- Swiss rounds only — top-cut is excluded to avoid selection bias.
- Appearances are counted per match-competitor per round (not per unique player), so high-round-count events naturally have more weight.
