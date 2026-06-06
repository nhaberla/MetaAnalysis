---
description: Breaks down what cards people run inside a single Star Wars Unlimited archetype, using real melee.gg tournament decklists. Reports per-card inclusion rate, average copies, and copy distribution, with full-field vs. top-cut columns side by side. Use when the user wants a deck's card breakdown, "what's in" a deck, inclusion rates, how many copies of a card people run, tech-card choices, or maindeck/sideboard composition for one leader. Requires --deck "Leader / X" (X = a colour or an exact base name). Also accepts optional --timeframe YYYY-MM-DD:YYYY-MM-DD, --min-players N, and --max-lists N.
---

Today's date: !`date +%Y-%m-%d`

Break down the cards in a single SWU archetype using: $ARGUMENTS

## Steps

1. **Resolve arguments from `$ARGUMENTS`.**
   - `--deck` (**required**): `"Leader / X"`. The leader is matched against the
     leaders actually present in the field; `X` is **either** a colour keyword
     (`yellow`/`blue`/`green`/`red`/`white`/`black`) — which folds together every
     list with that leader whose **base** is that aspect (e.g. all of Vader's
     generic yellow bases) — **or** an exact base name (e.g. `Lake Country`),
     which selects that one leader+base only. Aspectless high-HP bases (e.g. Lake
     Country) belong to no colour group and must be named exactly.
   - `--timeframe`: default last 30 days. Warn if the start date is before
     `2026-03-13` (pre-rotation card legality differs).
   - `--min-players`: default `32`.
   - `--max-lists`: default `0` (no cap). For very popular archetypes, set this
     (e.g. `60`) to keep runtime down — each decklist is one HTTP request.
   - `--output`: default `decklist_breakdown.md`.

2. **Run the script** (stdlib only, no install needed):
   ```
   python ${CLAUDE_SKILL_DIR}/decklist_breakdown.py --deck "<leader> / <X>" [other args]
   ```
   Pass `--verbose` on the first run of a session so API issues surface clearly.

3. **Handle the disambiguation exit (code 3).** Leaders are never merged across
   subtitles. If the leader in `--deck` matches several leaders (e.g. `"Darth
   Vader"` matches *Victor Squadron Leader*, *Unstoppable*, …), the script prints
   the candidate leaders with list counts and exits. **Present those options to
   the user, ask which they mean, then re-run** `--deck` with the exact leader
   name (subtitle included). Likewise, exit 4 means the leader wasn't found
   (the script lists available leaders) and exit 5 means an exact base name
   wasn't found (it lists the leader's bases) — relay these and retry.

4. **Handle other failures.**
   - *403 / network error from melee.gg*: the environment may block outbound
     requests — tell the user to run the script locally. (Note: melee's data
     endpoints usually work even when the Cloudflare-protected root 403s.)
   - *No decklist data collected*: all qualifying events may have decklists
     disabled, or the page structure changed — re-run with `--verbose`.
   - *SWUDB colour misses*: if a base's colour can't be resolved live, the script
     warns and excludes that base from a colour group rather than guessing.
     Re-running usually resolves transient SWUDB hiccups.

5. **Summarise results** after a successful run:
   - The resolved archetype (leader + base/colour group) and, for colour groups,
     the distinct bases that were folded together with their counts.
   - Sample sizes: number of full-field lists and top-cut lists, across how many
     events.
   - The standout cards: the highest-inclusion non-obvious cards, and any cards
     where the **top-cut** inclusion notably exceeds (or trails) the full field —
     those are the lists' tech/flex differentiators.
   - Flag small samples (e.g. top-cut N < 10) as low-confidence.

6. **Send the file** to the user with `SendUserFile`.

## Notes

- **Field basis:** every player on the archetype is counted once via **round 1**
  of Swiss — the full active field before any drops — so there's no survivorship
  bias. Top-cut lists are the subset whose decklist appeared in an elimination
  (non-"Round N") round.
- **Per-card stats:** *inclusion rate* = % of lists running ≥1 copy; *average
  copies* = mean copies among lists that run it; *copy distribution* = how many
  lists run exactly 1 / 2 / 3. Field and top-cut are shown side by side.
- **Maindeck vs. sideboard** are tallied separately (melee exposes a Sideboard
  category), so a card's maindeck and sideboard usage don't blur together.
- **Base colours are resolved live** from the SWUDB API (`api.swu-db.com`) per
  run and memoised only in memory — there is no stored colour table to go stale.
  Resolution filters to `Type == Base` with an exact name match before reading
  `Aspects` (Cunning=yellow, Command=green, Aggression=red, Vigilance=blue,
  Heroism=white, Villainy=black; an empty `Aspects` = aspectless base).
- Swiss round 1 only for the field; top-cut rounds only to flag top finishers.
  Premier events only, post-rotation (2026-03-13+).
- Output is Markdown for now; an XLSX renderer can follow once the numbers are
  verified, mirroring the meta-breakdown / matchup-breakdown skills.
