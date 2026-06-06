#!/usr/bin/env python3
"""
SWU Decklist Breakdown — stdlib only, no external dependencies.

Scrapes public melee.gg pages (no API key required) and reports what cards
people actually run inside a *single* archetype: inclusion rate, average
copies, and copy distribution — with full-field vs. top-cut columns side by
side. Base colors are resolved live from the SWUDB API (no stored table).

The archetype is selected with --deck "Leader / X" where X is either:
  • a colour keyword (yellow/blue/green/red/white/black) — folds together every
    list with that leader whose BASE is that aspect (e.g. all of Vader's generic
    yellow bases), or
  • an exact base name (e.g. "Lake Country") — that one leader+base only.

Leaders are never merged across subtitles. If the leader you name is ambiguous
(e.g. "Darth Vader" when several Vader leaders are in the field) the script
lists the candidates and stops so you can re-run against exactly one.

Usage:
    python decklist_breakdown.py --deck "Darth Vader, Victor Squadron Leader / yellow"
    python decklist_breakdown.py --deck "Boba Fett, Daimyo / Lake Country" --timeframe 2026-05-01:2026-06-06
    python decklist_breakdown.py --deck "Sabine Wren / green" --output sabine.md --verbose
"""

import argparse
import html
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Set, Tuple

# ── Constants ─────────────────────────────────────────────────────────────────

MELEE_BASE = "https://melee.gg"
SWUDB_SEARCH = "https://api.swu-db.com/cards/search"
POST_ROTATION_DATE = "2026-03-13"
REQUEST_DELAY = 1.0        # seconds between melee requests (respects crawl delay)
MIN_PLAYERS_DEFAULT = 32   # matches the other skills' filter standard
SWUDB_DELAY = 0.34         # gentle pacing for the SWUDB API

# DataTables column names used by /Match/GetRoundMatches (must match pairings-section.min.js)
_MATCH_COLUMNS = ["TableNumber", "PodNumber", "Teams", "Decklists", "ResultString"]

# Colour word ↔ SWU aspect. Bases carry one of these aspects, or none at all
# (aspectless high-HP bases like Lake Country return an empty Aspects array).
COLOR_TO_ASPECT = {
    "blue": "Vigilance",
    "green": "Command",
    "red": "Aggression",
    "yellow": "Cunning",
    "white": "Heroism",
    "black": "Villainy",
}
ASPECT_TO_COLOR = {v: k for k, v in COLOR_TO_ASPECT.items()}

# Card-category display order for the report. "Leader" is dropped (fixed by the
# archetype); "Base" is shown as its own distribution section, not a card table;
# "Sideboard" is always rendered last.
_CATEGORY_ORDER = ["Ground Unit", "Space Unit", "Unit", "Event", "Upgrade"]
_SKIP_CATEGORIES = {"Leader", "Base"}


# ── HTTP helpers ──────────────────────────────────────────────────────────────

_HTTP_CACHE: Dict[str, object] = {}
_BASE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


def _get_html(path: str, verbose: bool = False) -> str:
    url = MELEE_BASE + path
    if url in _HTTP_CACHE:
        return _HTTP_CACHE[url]  # type: ignore[return-value]
    time.sleep(REQUEST_DELAY)
    if verbose:
        print(f"  GET {url}", file=sys.stderr)
    req = urllib.request.Request(url, headers={
        **_BASE_HEADERS,
        "Accept": "text/html,application/xhtml+xml",
        "Referer": MELEE_BASE + "/",
    })
    with urllib.request.urlopen(req, timeout=20) as r:
        body = r.read().decode("utf-8", errors="replace")
    _HTTP_CACHE[url] = body
    return body


def _post_json(path: str, params: list, verbose: bool = False) -> object:
    url = MELEE_BASE + path
    data = urllib.parse.urlencode(params).encode()
    if verbose:
        shown = [(k, v) for k, v in params if not k.startswith("columns")]
        print(f"  POST {url} {dict(shown)}", file=sys.stderr)
    time.sleep(REQUEST_DELAY)
    req = urllib.request.Request(url, data=data, headers={
        **_BASE_HEADERS,
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": MELEE_BASE + "/",
    })
    with urllib.request.urlopen(req, timeout=20) as r:
        body = r.read()
    if body.strip().startswith(b"<"):
        raise RuntimeError(
            f"melee.gg returned HTML instead of JSON for POST {path}. "
            "The endpoint may have changed or require authentication."
        )
    return json.loads(body)


def _dt_params(start: int = 0, length: int = 250) -> list:
    """DataTables POST params for /Match/GetRoundMatches (flat keys)."""
    params = [
        ("draw", "1"), ("start", str(start)), ("length", str(length)),
        ("order[0][column]", "0"), ("order[0][dir]", "asc"),
        ("search[value]", ""), ("search[regex]", "false"),
    ]
    for i, col in enumerate(_MATCH_COLUMNS):
        params += [
            (f"columns[{i}][data]", col),
            (f"columns[{i}][name]", ""),
            (f"columns[{i}][searchable]", "true"),
            (f"columns[{i}][orderable]", "true"),
            (f"columns[{i}][search][value]", ""),
            (f"columns[{i}][search][regex]", "false"),
        ]
    return params


def _search_dt_params(start: int = 0, length: int = 250) -> list:
    """DataTables POST params for /Tournament/TournamentSearch (variables[...] keys)."""
    return [
        ("variables[draw]", "1"),
        ("variables[start]", str(start)),
        ("variables[length]", str(length)),
        ("variables[order][0][column]", "0"),
        ("variables[order][0][dir]", "asc"),
        ("variables[search][value]", ""),
        ("variables[search][regex]", "false"),
    ]


# ── Tournament fetching ───────────────────────────────────────────────────────

def _t_id(t: dict) -> int:
    return int(t.get("id") or t.get("ID") or 0)

def _t_name(t: dict) -> str:
    return str(t.get("name") or t.get("Name") or "Unnamed")

def _t_players(t: dict) -> int:
    return int(t.get("enrolledPlayerCount") or t.get("playerCount") or 0)

def _t_date(t: dict) -> str:
    return str(t.get("startDate") or t.get("StartDate") or "")[:10]


def fetch_tournaments(
    start_date: str, end_date: str, min_players: int, verbose: bool,
) -> List[dict]:
    """
    Return all SWU Premier ended tournaments in the date range with >= min_players.

    melee.gg returns events only approximately sorted by start date, so we page
    backward from the newest events with a jitter buffer to make sure no in-range
    event is missed. (Lifted from the meta-breakdown skill.)
    """
    base = [
        ("ordering", "StartDate"), ("mode", "Table"),
        ("filters[]", "Ended"), ("filters[]", "Premier"),
    ]

    probe = _post_json("/Tournament/TournamentSearch",
                       base + _search_dt_params(0, 25), verbose=verbose)
    total = int(probe.get("recordsTotal") or 0)  # type: ignore[union-attr]
    if verbose:
        print(f"  Total Ended+Premier SWU events indexed: {total}", file=sys.stderr)

    JITTER_BUFFER_DAYS = 35
    cutoff = (datetime.strptime(start_date, "%Y-%m-%d") -
              timedelta(days=JITTER_BUFFER_DAYS)).strftime("%Y-%m-%d")

    results: List[dict] = []
    seen_ids: Set[int] = set()
    batch_size = 250
    offset = max(0, total - batch_size)
    while offset >= 0:
        resp = _post_json("/Tournament/TournamentSearch",
                          base + _search_dt_params(offset, batch_size), verbose=verbose)
        items = resp.get("data", [])  # type: ignore[union-attr]

        batch_max_date = ""
        for t in items:
            d = _t_date(t)
            if d > batch_max_date:
                batch_max_date = d
            if d < start_date or d > end_date:
                continue
            if _t_players(t) < min_players:
                continue
            tid = _t_id(t)
            if tid in seen_ids:
                continue
            seen_ids.add(tid)
            results.append(t)

        if batch_max_date and batch_max_date < cutoff:
            break
        if offset == 0:
            break
        offset = max(0, offset - batch_size)

    if verbose:
        print(f"  Paged back to ~{cutoff} cutoff; {len(results)} events in range", file=sys.stderr)
    return results


# ── Round + match fetching ────────────────────────────────────────────────────

def _is_swiss_round_name(name: str) -> bool:
    """Swiss rounds are named 'Round N'; top-cut rounds have descriptive names."""
    return bool(re.match(r"^Round\s+\d+$", name.strip(), re.IGNORECASE))


def fetch_rounds(tournament_id: int, verbose: bool) -> Tuple[List[str], List[str], bool]:
    """
    Parse /Tournament/View/{id} for (swiss_round_ids, topcut_round_ids, decklist_enabled).

    Swiss = 'Round N' buttons (used for the full field); every other round button
    is treated as top-cut / elimination (used to flag top finishers).
    """
    html_doc = _get_html(f"/Tournament/View/{tournament_id}", verbose=verbose)

    round_buttons = re.findall(
        r'<button[^>]+class="[^"]*round-selector[^"]*"[^>]*'
        r'data-id="(\d+)"[^>]*data-name="([^"]+)"',
        html_doc,
    )
    if not round_buttons:
        round_buttons = [
            (rid, name)
            for name, rid in re.findall(
                r'<button[^>]+data-name="([^"]+)"[^>]+data-id="(\d+)"', html_doc
            )
        ]

    swiss_ids = [rid for rid, name in round_buttons if _is_swiss_round_name(name)]
    topcut_ids = [rid for rid, name in round_buttons if not _is_swiss_round_name(name)]

    dl_match = re.search(r'name="DecklistEnabled"[^>]+value="([^"]+)"', html_doc)
    decklist_enabled = bool(dl_match and dl_match.group(1).lower() == "true")

    if verbose:
        names = [name for _, name in round_buttons]
        print(f"  [debug] rounds: {names} | decklists: {decklist_enabled}", file=sys.stderr)

    return swiss_ids, topcut_ids, decklist_enabled


def fetch_round_matches(round_id: str, verbose: bool) -> List[dict]:
    """Fetch all matches for a round via POST /Match/GetRoundMatches/{roundId}."""
    all_matches: List[dict] = []
    start = 0
    while True:
        resp = _post_json(f"/Match/GetRoundMatches/{round_id}", _dt_params(start, 250),
                          verbose=verbose)
        batch = resp.get("data", [])  # type: ignore[union-attr]
        all_matches.extend(batch)
        total = int(resp.get("recordsTotal") or 0)  # type: ignore[union-attr]
        if len(all_matches) >= total or not batch:
            break
        start += 250
    return all_matches


# ── Archetype parsing ─────────────────────────────────────────────────────────

def parse_decklist_name(name: str) -> Optional[Tuple[str, str]]:
    """Parse 'Leader, Subtitle - Base Name' into (leader, base). Splits on the last ' - '."""
    if not name:
        return None
    sep = " - "
    idx = name.rfind(sep)
    if idx < 0:
        return None
    leader = name[:idx].strip()
    base = name[idx + len(sep):].strip()
    if not leader or not base:
        return None
    return leader, base


# ── Field collection ──────────────────────────────────────────────────────────

class Entry:
    """One player's list in one tournament's round 1."""
    __slots__ = ("tid", "tname", "player_id", "decklist_id", "leader", "base", "is_top")

    def __init__(self, tid, tname, player_id, decklist_id, leader, base):
        self.tid = tid
        self.tname = tname
        self.player_id = player_id
        self.decklist_id = decklist_id
        self.leader = leader
        self.base = base
        self.is_top = False


def collect_field(tournaments: List[dict], verbose: bool) -> List[Entry]:
    """
    Build the full field of entries (round 1, one per player) across all events,
    flagging which decklists also appeared in a top-cut round.
    """
    entries: List[Entry] = []
    top_decklist_ids: Set[str] = set()

    for t in sorted(tournaments, key=_t_date):
        tid = _t_id(t)
        tname = _t_name(t)
        if not tid:
            continue
        try:
            swiss_ids, topcut_ids, enabled = fetch_rounds(tid, verbose)
        except Exception as exc:  # noqa: BLE001
            print(f"  ✗  {tname[:50]}: rounds error: {exc}", file=sys.stderr)
            continue
        if not enabled:
            if verbose:
                print(f"  [skip] {tname[:50]}: decklists not public", file=sys.stderr)
            continue
        if not swiss_ids:
            if verbose:
                print(f"  [skip] {tname[:50]}: no Swiss rounds", file=sys.stderr)
            continue

        # Round 1 = the full active field, one entry per player.
        try:
            matches = fetch_round_matches(swiss_ids[0], verbose)
        except Exception as exc:  # noqa: BLE001
            print(f"  ✗  {tname[:50]}: round 1 error: {exc}", file=sys.stderr)
            continue

        added = 0
        for m in matches:
            for c in m.get("Competitors", []):
                dl = (c.get("Decklists") or [{}])[0]
                did = dl.get("DecklistId")
                pair = parse_decklist_name(dl.get("DecklistName", ""))
                if not did or not pair:
                    continue
                entries.append(Entry(tid, tname, c.get("PlayerId"), did, pair[0], pair[1]))
                added += 1

        # Top-cut decklist IDs → flag top finishers.
        for rid in topcut_ids:
            try:
                tc_matches = fetch_round_matches(rid, verbose)
            except Exception:  # noqa: BLE001
                continue
            for m in tc_matches:
                for c in m.get("Competitors", []):
                    dl = (c.get("Decklists") or [{}])[0]
                    if dl.get("DecklistId"):
                        top_decklist_ids.add(dl["DecklistId"])

        print(f"  ✓  {tname[:50]:<50}  {added} entries"
              f"{'  (+top cut)' if topcut_ids else ''}")

    for e in entries:
        if e.decklist_id in top_decklist_ids:
            e.is_top = True
    return entries


# ── SWUDB base-colour resolution (live, memoised per run only) ─────────────────

def _swudb_get(query: str, verbose: bool) -> Optional[list]:
    url = f"{SWUDB_SEARCH}?q={urllib.parse.quote(query)}"
    time.sleep(SWUDB_DELAY)
    if verbose:
        print(f"  SWUDB {url}", file=sys.stderr)
    req = urllib.request.Request(url, headers={
        "User-Agent": "swu-decklist-breakdown/1.0 (melee.gg analysis script)",
        "Accept": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            payload = json.loads(r.read().decode("utf-8", errors="replace"))
    except Exception as exc:  # noqa: BLE001
        if verbose:
            print(f"  SWUDB error for {query!r}: {exc}", file=sys.stderr)
        return None
    return payload.get("data") or []


def resolve_base_aspects(
    base_name: str, cache: Dict[str, Optional[List[str]]], verbose: bool,
) -> Optional[List[str]]:
    """
    Live-resolve a base's aspect list from SWUDB. Returns:
      • a list of aspects (possibly empty [] for aspectless bases like Lake Country),
      • or None if the base could not be resolved (network error / not found).
    Results are memoised for the duration of the run only — never persisted.
    """
    key = base_name.lower()
    if key in cache:
        return cache[key]

    data = _swudb_get(base_name, verbose)
    aspects: Optional[List[str]] = None
    if data is not None:
        for c in data:
            if str(c.get("Type", "")).lower() == "base" and \
               str(c.get("Name", "")).lower() == key:
                aspects = list(c.get("Aspects") or [])
                break
    cache[key] = aspects
    return aspects


# ── Decklist detail fetching + parsing ────────────────────────────────────────

_CAT_TITLE_RE = re.compile(r'decklist-category-title">([^<]+)</div>')
_RECORD_RE = re.compile(
    r'decklist-record-quantity">\s*(\d+)\s*</span>\s*'
    r'<(?:a|span)[^>]*decklist-record-name[^>]*>([^<]+)</(?:a|span)>'
)


def parse_decklist_detail(html_doc: str) -> Dict[str, Dict[str, int]]:
    """
    Parse a /Decklist/View page into {category: {card_name: quantity}}.

    The page is server-rendered with one '.decklist-category' block per zone
    (Leader, Base, Ground Unit, Space Unit, Unit, Event, Upgrade, Sideboard).
    """
    titles = list(_CAT_TITLE_RE.finditer(html_doc))
    out: Dict[str, Dict[str, int]] = {}
    for i, m in enumerate(titles):
        # Strip the trailing "(count)" from the category title.
        cat = re.sub(r"\s*\(\d+\)\s*$", "", m.group(1)).strip()
        seg_start = m.end()
        seg_end = titles[i + 1].start() if i + 1 < len(titles) else len(html_doc)
        segment = html_doc[seg_start:seg_end]
        cards = out.setdefault(cat, {})
        for q, name in _RECORD_RE.findall(segment):
            card = html.unescape(name).strip()
            cards[card] = cards.get(card, 0) + int(q)
    return out


def fetch_decklist(decklist_id: str, cache: Dict[str, Dict[str, Dict[str, int]]],
                   verbose: bool) -> Dict[str, Dict[str, int]]:
    """Fetch + parse a decklist by ID, memoised per run."""
    if decklist_id in cache:
        return cache[decklist_id]
    html_doc = _get_html(f"/Decklist/View/{decklist_id}", verbose=verbose)
    parsed = parse_decklist_detail(html_doc)
    cache[decklist_id] = parsed
    return parsed


# ── Aggregation ───────────────────────────────────────────────────────────────

class CardAgg:
    """Per-card tallies within one scope (field or top)."""
    __slots__ = ("running", "copies", "dist")

    def __init__(self):
        self.running = 0                 # lists running ≥1 copy
        self.copies = 0                  # total copies across those lists
        self.dist = defaultdict(int)     # {n_copies: n_lists}

    def add(self, qty: int):
        self.running += 1
        self.copies += qty
        self.dist[qty] += 1

    def incl(self, total_lists: int) -> float:
        return 100.0 * self.running / total_lists if total_lists else 0.0

    def avg(self) -> float:
        return self.copies / self.running if self.running else 0.0

    def dist_str(self) -> str:
        return "/".join(str(self.dist.get(n, 0)) for n in (1, 2, 3))


def aggregate(
    lists_parsed: List[Tuple[Dict[str, Dict[str, int]], bool]],
) -> Tuple[Dict[str, Dict[str, CardAgg]], Dict[str, Dict[str, CardAgg]]]:
    """
    Build {category: {card: CardAgg}} for the field scope and the top scope.

    `lists_parsed` is a list of (parsed_decklist, is_top) tuples.
    """
    field: Dict[str, Dict[str, CardAgg]] = defaultdict(lambda: defaultdict(CardAgg))
    top: Dict[str, Dict[str, CardAgg]] = defaultdict(lambda: defaultdict(CardAgg))

    for parsed, is_top in lists_parsed:
        for cat, cards in parsed.items():
            if cat in _SKIP_CATEGORIES:
                continue
            for card, qty in cards.items():
                field[cat][card].add(qty)
                if is_top:
                    top[cat][card].add(qty)
    return field, top


def ordered_categories(cats: Set[str]) -> List[str]:
    """Preferred category order, with any unknown categories appended, Sideboard last."""
    ordered = [c for c in _CATEGORY_ORDER if c in cats]
    extras = sorted(c for c in cats if c not in _CATEGORY_ORDER and c != "Sideboard")
    tail = ["Sideboard"] if "Sideboard" in cats else []
    return ordered + extras + tail


# ── Markdown writer ───────────────────────────────────────────────────────────

def _md(s: str) -> str:
    """Escape a value for a GFM table cell. Card names contain ' | ' (name | subtitle)
    which would otherwise be read as a column separator."""
    return s.replace("|", "\\|")


def _fmt_pct(agg: Optional[CardAgg], total: int) -> str:
    if not agg or total == 0:
        return "—"
    return f"{agg.incl(total):.0f}%"


def _fmt_avg(agg: Optional[CardAgg]) -> str:
    if not agg or agg.running == 0:
        return "—"
    return f"{agg.avg():.1f}"


def build_markdown(
    leader: str,
    group_label: str,
    folded_bases: List[Tuple[str, int]],
    start_date: str,
    end_date: str,
    n_events: int,
    n_events_public: int,
    field_lists: List[Entry],
    top_lists: List[Entry],
    field_agg: Dict[str, Dict[str, CardAgg]],
    top_agg: Dict[str, Dict[str, CardAgg]],
    color_mode: bool,
) -> str:
    n_field = len(field_lists)
    n_top = len(top_lists)
    lines: List[str] = []

    lines.append(f"# Decklist Breakdown — {leader} / {group_label}")
    lines.append("")
    lines.append(f"- **Timeframe:** {start_date} → {end_date}")
    lines.append(f"- **Events:** {n_events} in range ({n_events_public} with public decklists)")
    lines.append(f"- **Lists analysed:** {n_field} full field · {n_top} top-cut")
    if color_mode:
        folded = ", ".join(f"{b} ({c})" for b, c in folded_bases)
        lines.append(f"- **Bases folded into '{group_label}':** {folded or '—'}")
    lines.append("")
    lines.append(
        "> Full field = every player on this archetype counted once via round 1 "
        "(no survivorship bias). Top-cut = the subset whose list reached an "
        "elimination round. Copy distribution columns show how many lists ran "
        "1 / 2 / 3 copies. Post-rotation legality only.")
    lines.append("")

    # Base distribution (only meaningful when a colour folded several bases).
    if color_mode and folded_bases:
        lines.append("## Bases")
        lines.append("")
        lines.append("| Base | Field % (n) | Top-cut % (n) |")
        lines.append("|---|---|---|")
        top_base_counts: Dict[str, int] = defaultdict(int)
        for e in top_lists:
            top_base_counts[e.base] += 1
        for base, count in sorted(folded_bases, key=lambda x: -x[1]):
            fpct = 100.0 * count / n_field if n_field else 0.0
            tc = top_base_counts.get(base, 0)
            tpct = 100.0 * tc / n_top if n_top else 0.0
            lines.append(f"| {_md(base)} | {fpct:.0f}% ({count}) | {tpct:.0f}% ({tc}) |")
        lines.append("")

    # Card tables per category.
    all_cats = set(field_agg.keys())
    for cat in ordered_categories(all_cats):
        cards = field_agg.get(cat, {})
        if not cards:
            continue
        lines.append(f"## {cat}")
        lines.append("")
        lines.append("| Card | Field incl% | Top incl% | Avg (field) | Avg (top) | Copies 1/2/3 (field) |")
        lines.append("|---|---|---|---|---|---|")
        ranked = sorted(
            cards.items(),
            key=lambda kv: (-kv[1].incl(n_field), -kv[1].copies, kv[0]),
        )
        for card, agg in ranked:
            tagg = top_agg.get(cat, {}).get(card)
            lines.append(
                f"| {_md(card)} | {_fmt_pct(agg, n_field)} | {_fmt_pct(tagg, n_top)} "
                f"| {_fmt_avg(agg)} | {_fmt_avg(tagg)} | {agg.dist_str()} |")
        lines.append("")

    lines.append("---")
    lines.append(f"*Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} from melee.gg "
                 f"round-1 data; base colours resolved live via SWUDB.*")
    lines.append("")
    return "\n".join(lines)


# ── Deck-argument resolution ──────────────────────────────────────────────────

def split_deck_arg(deck: str) -> Tuple[str, str]:
    """Split --deck 'Leader / X' into (leader_query, base_or_color)."""
    if "/" not in deck:
        sys.exit(
            "Error: --deck must be 'Leader / X' where X is a colour "
            "(yellow/blue/green/red/white/black) or an exact base name.\n"
            f"Got: {deck!r}")
    leader_q, x = deck.rsplit("/", 1)
    return leader_q.strip(), x.strip()


def resolve_leader(entries: List[Entry], leader_query: str) -> str:
    """
    Resolve leader_query to exactly one leader string present in the field.

    Leaders are never merged: if the query matches several distinct leaders
    (different subtitles), print them and exit 3 for disambiguation.
    """
    lq = leader_query.lower()
    matches: Dict[str, int] = defaultdict(int)
    for e in entries:
        if lq in e.leader.lower():
            matches[e.leader] += 1

    if not matches:
        leaders = sorted({e.leader for e in entries})
        print(f"\nNo leader matching {leader_query!r} found in the field.\n"
              f"Leaders present ({len(leaders)}):", file=sys.stderr)
        for lead in leaders:
            print(f"  • {lead}", file=sys.stderr)
        sys.exit(4)

    if len(matches) > 1:
        print(f"\nDISAMBIGUATION NEEDED — {leader_query!r} matches "
              f"{len(matches)} leaders. Re-run --deck with one of these exact "
              f"leader names:\n", file=sys.stderr)
        for lead, ct in sorted(matches.items(), key=lambda kv: -kv[1]):
            print(f"  • {lead}   ({ct} lists)", file=sys.stderr)
        sys.exit(3)

    return next(iter(matches))


def resolve_bases(
    entries: List[Entry], leader: str, x: str, verbose: bool,
) -> Tuple[Set[str], List[Tuple[str, int]], str, bool]:
    """
    Decide which base names belong to the selected archetype.

    Returns (included_bases, folded_bases_with_counts, group_label, color_mode).
    """
    leader_entries = [e for e in entries if e.leader == leader]
    base_counts: Dict[str, int] = defaultdict(int)
    for e in leader_entries:
        base_counts[e.base] += 1

    x_low = x.lower()
    target_aspect: Optional[str] = None
    if x_low in COLOR_TO_ASPECT:
        target_aspect = COLOR_TO_ASPECT[x_low]
    elif x in ASPECT_TO_COLOR:                 # accept an aspect name directly
        target_aspect = x
    elif x_low.capitalize() in ASPECT_TO_COLOR:
        target_aspect = x_low.capitalize()

    if target_aspect is None:
        # Exact base name mode.
        match = next((b for b in base_counts if b.lower() == x_low), None)
        if match is None:
            print(f"\nBase {x!r} not found for {leader!r}.\n"
                  f"Bases present for this leader:", file=sys.stderr)
            for b, c in sorted(base_counts.items(), key=lambda kv: -kv[1]):
                print(f"  • {b}   ({c} lists)", file=sys.stderr)
            sys.exit(5)
        return {match}, [(match, base_counts[match])], match, False

    # Colour mode: live-resolve each base's aspect via SWUDB.
    color_word = ASPECT_TO_COLOR[target_aspect]
    print(f"\nResolving base colours via SWUDB for {len(base_counts)} distinct "
          f"base(s)...", file=sys.stderr)
    aspect_cache: Dict[str, Optional[List[str]]] = {}
    included: Set[str] = set()
    folded: List[Tuple[str, int]] = []
    unresolved: List[str] = []
    for base, count in base_counts.items():
        aspects = resolve_base_aspects(base, aspect_cache, verbose)
        if aspects is None:
            unresolved.append(base)
            continue
        if target_aspect in aspects:
            included.add(base)
            folded.append((base, count))

    if unresolved:
        print(f"  ⚠  Could not resolve colour for: {', '.join(sorted(unresolved))} "
              f"(SWUDB miss) — excluded.", file=sys.stderr)
    folded.sort(key=lambda x: -x[1])
    return included, folded, f"{color_word} bases", True


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Break down the cards run inside one SWU archetype from melee.gg data.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--deck", required=True, metavar='"Leader / X"',
                   help="Archetype. X is a colour (yellow/blue/green/red/white/black) "
                        "to fold all of that leader's bases of that aspect, or an exact "
                        "base name.")
    p.add_argument("--timeframe", default=None, metavar="YYYY-MM-DD:YYYY-MM-DD",
                   help="Date range (default: last 30 days).")
    p.add_argument("--min-players", type=int, default=MIN_PLAYERS_DEFAULT,
                   help=f"Minimum event size (default: {MIN_PLAYERS_DEFAULT}).")
    p.add_argument("--max-lists", type=int, default=0,
                   help="Cap the number of decklists fetched (0 = no cap). Useful for "
                        "very popular archetypes to keep runtime down.")
    p.add_argument("--output", default="decklist_breakdown.md")
    p.add_argument("--verbose", "-v", action="store_true")
    return p.parse_args()


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()
    leader_query, x = split_deck_arg(args.deck)

    today = datetime.now().date()
    if args.timeframe is None:
        start_date, end_date = today - timedelta(days=30), today
    else:
        try:
            s, e = args.timeframe.split(":")
            start_date = datetime.strptime(s.strip(), "%Y-%m-%d").date()
            end_date = datetime.strptime(e.strip(), "%Y-%m-%d").date()
        except ValueError:
            sys.exit("Error: --timeframe must be YYYY-MM-DD:YYYY-MM-DD")

    if str(start_date) < POST_ROTATION_DATE:
        print(f"Warning: {start_date} predates post-rotation ({POST_ROTATION_DATE}). "
              "Pre-rotation card legality differs.", file=sys.stderr)

    # ── Fetch the field ────────────────────────────────────────────────────────
    print(f"Fetching SWU tournaments: {start_date} → {end_date} "
          f"(min {args.min_players} players)...")
    try:
        tournaments = fetch_tournaments(str(start_date), str(end_date),
                                        args.min_players, args.verbose)
    except urllib.error.URLError as exc:
        sys.exit(f"\nNetwork error: {exc}\nIf this environment blocks outbound "
                 "requests, run the script locally.\n")
    except RuntimeError as exc:
        sys.exit(f"\n{exc}\n")

    if not tournaments:
        sys.exit(f"No SWU Premier tournaments found ({start_date} to {end_date}, "
                 f"min {args.min_players} players). Try widening --timeframe.")

    print(f"Found {len(tournaments)} event(s). Collecting round-1 field...\n")
    entries = collect_field(tournaments, args.verbose)
    if not entries:
        sys.exit("\nNo decklist data collected (decklists may be disabled on all events).")

    n_events_public = len({e.tid for e in entries})

    # ── Resolve archetype ──────────────────────────────────────────────────────
    leader = resolve_leader(entries, leader_query)
    included_bases, folded_bases, group_label, color_mode = \
        resolve_bases(entries, leader, x, args.verbose)

    selected = [e for e in entries
                if e.leader == leader and e.base in included_bases]
    if not selected:
        sys.exit(f"\nNo lists found for {leader} / {group_label} in this timeframe.")

    # Dedup by decklist id (one fetch per unique list).
    by_did: Dict[str, Entry] = {}
    for e in selected:
        by_did.setdefault(e.decklist_id, e)
    unique = list(by_did.values())
    if args.max_lists and len(unique) > args.max_lists:
        # Keep every top-cut list (the scarce, high-value signal), then fill the
        # remaining slots with full-field lists, so capping never zeroes out the
        # top-cut column.
        tops = [e for e in unique if e.is_top]
        rest = [e for e in unique if not e.is_top]
        unique = (tops + rest[:max(0, args.max_lists - len(tops))])[:args.max_lists]
        print(f"\nCapping at --max-lists {args.max_lists} (keeping all "
              f"{len(tops)} top-cut list(s)).", file=sys.stderr)

    field_lists = unique
    top_lists = [e for e in unique if e.is_top]
    print(f"\nSelected {len(field_lists)} list(s) for {leader} / {group_label} "
          f"({len(top_lists)} in top cut). Fetching decklists...\n")

    # ── Fetch + parse decklists ────────────────────────────────────────────────
    dl_cache: Dict[str, Dict[str, Dict[str, int]]] = {}
    lists_parsed: List[Tuple[Dict[str, Dict[str, int]], bool]] = []
    for i, e in enumerate(field_lists, 1):
        try:
            parsed = fetch_decklist(e.decklist_id, dl_cache, args.verbose)
        except Exception as exc:  # noqa: BLE001
            print(f"  ✗  list {i}/{len(field_lists)}: {exc}", file=sys.stderr)
            continue
        lists_parsed.append((parsed, e.is_top))
        if i % 10 == 0 or i == len(field_lists):
            print(f"  ...{i}/{len(field_lists)} decklists parsed")

    if not lists_parsed:
        sys.exit("\nNo decklists could be parsed.")

    field_agg, top_agg = aggregate(lists_parsed)

    # ── Write report ───────────────────────────────────────────────────────────
    md = build_markdown(
        leader, group_label, folded_bases, str(start_date), str(end_date),
        len(tournaments), n_events_public, field_lists, top_lists,
        field_agg, top_agg, color_mode,
    )
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(md)

    print(f"\nSaved: {args.output}")
    print(f"\n{leader} / {group_label}: {len(field_lists)} lists "
          f"({len(top_lists)} top-cut) across {n_events_public} events.")


if __name__ == "__main__":
    main()
