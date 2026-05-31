#!/usr/bin/env python3
"""
SWU Matchup Breakdown — stdlib only, no external dependencies.

Fetches real game-by-game results from melee.gg and builds a color-coded
head-to-head win rate matrix spreadsheet for the Star Wars Unlimited meta.
Archetypes are discovered dynamically from tournament decklist data.

Usage:
    python matchup_breakdown.py --timeframe 2026-03-13:2026-05-31 --num-decks 12
    python matchup_breakdown.py --timeframe 2026-05-01:2026-05-31 --num-decks 5 --output may.xlsx
    python matchup_breakdown.py --timeframe 2026-03-13:2026-05-31 --dry-run --verbose
"""

import argparse
import io
import json
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.sax.saxutils
import zipfile
from collections import defaultdict
from datetime import datetime
from typing import Dict, FrozenSet, List, Optional, Tuple

# ── Constants ─────────────────────────────────────────────────────────────────

MELEE_API = "https://melee.gg/api/v1"
POST_ROTATION_DATE = "2026-03-13"
REQUEST_DELAY = 0.5        # seconds between API calls
MIN_PLAYERS_DEFAULT = 32   # matches swu-competitivehub.com filter standard
LOW_SAMPLE_THRESHOLD = 20  # matchups below this count get an asterisk

# ── Demo data (SWU LAW Meta Guide, post-rotation LAW Premier) ─────────────────
# Used by --demo to validate spreadsheet output without live API calls.

_DEMO_ARCHETYPES = [
    "Lando Calrissian / Lake Country",
    "Boba Fett / Lake Country",
    "Aurra Sing / Data Vault",
    "Dedra Meero / Colossus",
    "Obi-Wan Kenobi / Blue Force",
    "Mother Talzin / Yellow Force",
    "Chewbacca / Cunning",
    "Col. Yularen / Aggression",
    "Darth Vader / Cunning",
    "Luke Skywalker / Data Vault",
    "Admiral Piett / Blue",
    "Tobias Beckett / Red",
]

# Meta share % per archetype (from the guide) — used to estimate sample sizes.
# The product of two shares approximates relative matchup frequency.
_DEMO_META_SHARE = [9.7, 12.9, 3.3, 2.4, 10.2, 5.2, 3.6, 3.6, 3.9, 6.6, 3.7, 2.0]

# Win rate of row archetype vs column archetype (None = mirror).
# Source: SWU LAW Meta Guide matchup matrix.
_DEMO_WIN_PCT: List[List[Optional[int]]] = [
    #        Lando  Boba  Aurra  Dedra  Obi   Talzin Chewb  Yular  Vader  Luke  Piett  Tobias
    [None,   53,    44,   50,    53,    55,   55,    56,    70,    62,   55,    57  ],  # Lando
    [47,     None,  45,   50,    65,    53,   52,    55,    55,    58,   54,    56  ],  # Boba
    [56,     55,    None, 52,    45,    53,   50,    52,    22,    54,   52,    54  ],  # Aurra
    [50,     50,    48,   None,  54,    52,   50,    54,    55,    57,   53,    55  ],  # Dedra
    [47,     35,    55,   46,    None,  42,   45,    37,    68,    55,   50,    52  ],  # Obi-Wan
    [45,     47,    47,   48,    58,    None, 48,    50,    52,    54,   50,    52  ],  # Talzin
    [45,     48,    50,   50,    57,    52,   None,  52,    55,    56,   52,    53  ],  # Chewbacca
    [44,     45,    48,   46,    63,    50,   48,    None,  55,    63,   52,    54  ],  # Yularen
    [30,     45,    78,   45,    32,    48,   45,    45,    None,  55,   50,    52  ],  # Vader
    [38,     42,    46,   43,    45,    46,   44,    37,    45,    None, 48,    50  ],  # Luke
    [45,     46,    48,   47,    50,    50,   48,    48,    50,    52,   None,  51  ],  # Piett
    [43,     44,    46,   45,    48,    48,   47,    46,    48,    50,   49,    None],  # Tobias
]

# melee.gg game identifier for Star Wars Unlimited.
# If you see 0 results or a 404, open melee.gg in a browser, start DevTools →
# Network, navigate to the SWU tournament list, and inspect the XHR request to
# find the current value. It may be a string slug or a numeric gameId.
SWU_GAME_SLUG = "sw-unlimited"

# ── HTTP helpers ──────────────────────────────────────────────────────────────

_HTTP_CACHE: Dict[str, object] = {}


def _http_get(url: str, params: Optional[dict] = None, verbose: bool = False) -> object:
    if params:
        url = f"{url}?{urllib.parse.urlencode(params)}"
    if url in _HTTP_CACHE:
        return _HTTP_CACHE[url]
    time.sleep(REQUEST_DELAY)
    if verbose:
        print(f"  GET {url}", file=sys.stderr)
    req = urllib.request.Request(url, headers={
        "User-Agent": "SWU-MetaAnalysis/1.0",
        "Accept": "application/json",
    })
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = json.loads(resp.read().decode())
    _HTTP_CACHE[url] = data
    return data


def _unwrap(data: object) -> list:
    """Extract the list payload from common melee.gg envelope shapes."""
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("data", "tournaments", "rounds", "pairings"):
            if key in data and isinstance(data[key], list):
                return data[key]
    return []


# ── Tournament fetching ───────────────────────────────────────────────────────

def fetch_tournaments(
    start_date: str, end_date: str, min_players: int, verbose: bool
) -> List[dict]:
    """
    Return all SWU Premier tournaments in the date range with >= min_players.

    melee.gg API (reverse-engineered by community tools):
        GET /api/v1/Tournaments
        params: game, startDate, endDate, pageSize, pageNumber, format
    """
    results: List[dict] = []
    page, page_size = 1, 50

    while True:
        try:
            raw = _http_get(f"{MELEE_API}/Tournaments", params={
                "game": SWU_GAME_SLUG,
                "startDate": start_date,
                "endDate": end_date,
                "pageSize": page_size,
                "pageNumber": page,
                "format": "Premier",
            }, verbose=verbose)
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                raise RuntimeError(
                    f"melee.gg returned 404. SWU_GAME_SLUG='{SWU_GAME_SLUG}' may be wrong.\n"
                    f"Inspect melee.gg XHR traffic to find the correct game filter value."
                ) from exc
            raise

        items = _unwrap(raw)
        if not items:
            break
        for t in items:
            count = int(t.get("playerCount") or t.get("registrations") or 0)
            if count >= min_players:
                results.append(t)
        if len(items) < page_size:
            break
        page += 1

    return results


# ── Archetype discovery ───────────────────────────────────────────────────────

def _clean(name: str) -> str:
    return " ".join(name.split()) if name else "Unknown"


def fetch_archetype(decklist_id: str, verbose: bool) -> Optional[Tuple[str, str]]:
    """
    Fetch a decklist from melee.gg and return (leader_name, base_name).

    Each SWU deck has exactly one Leader and one Base card. melee.gg may
    expose these as top-level dict fields or inside a 'cards' / 'mainDeck'
    array. Both shapes are tried.
    """
    try:
        dl = _http_get(f"{MELEE_API}/Decklists/{decklist_id}", verbose=verbose)
    except urllib.error.HTTPError:
        return None
    if not isinstance(dl, dict):
        return None

    leader, base = "", ""

    if isinstance(dl.get("leader"), dict):
        leader = _clean(dl["leader"].get("name", ""))
    if isinstance(dl.get("base"), dict):
        base = _clean(dl["base"].get("name", ""))

    if not leader or not base:
        for card in (dl.get("cards") or dl.get("mainDeck") or []):
            ctype = (card.get("type") or card.get("cardType") or "").lower()
            if ctype == "leader" and not leader:
                leader = _clean(card.get("name", ""))
            elif ctype == "base" and not base:
                base = _clean(card.get("name", ""))

    return (leader or "Unknown", base or "Unknown") if (leader or base) else None


def archetype_label(leader: str, base: str) -> str:
    return f"{leader} / {base}"


# ── Match processing ──────────────────────────────────────────────────────────

def _is_top_cut(rnd: dict) -> bool:
    rtype = (rnd.get("type") or rnd.get("roundType") or "").lower()
    return rnd.get("isTopCut", False) or rtype in (
        "top8", "top4", "top2", "final", "semifinal", "quarterfinal"
    )


def process_tournament(
    tournament: dict,
    wins: "defaultdict[Tuple[str, str], int]",
    pair_games: "defaultdict[FrozenSet[str], int]",
    verbose: bool,
) -> int:
    """
    Walk all Swiss rounds of a tournament. For each completed pairing, resolve
    both players' archetypes from their decklists and record the result.
    Returns the number of matchup records added.
    """
    tid = str(tournament.get("id") or tournament.get("tournamentId", ""))

    try:
        raw_rounds = _http_get(f"{MELEE_API}/Tournaments/{tid}/Rounds", verbose=verbose)
    except urllib.error.HTTPError:
        return 0

    swiss_rounds = [r for r in _unwrap(raw_rounds) if not _is_top_cut(r)]
    player_arch: Dict[str, Optional[str]] = {}  # player_id → "Leader / Base"
    added = 0

    for rnd in swiss_rounds:
        rid = str(rnd.get("id") or rnd.get("roundId", ""))
        try:
            raw_pairings = _http_get(
                f"{MELEE_API}/Pairings", params={"roundId": rid}, verbose=verbose
            )
        except urllib.error.HTTPError:
            continue

        for p in _unwrap(raw_pairings):
            p1_id = str(p.get("player1Id") or p.get("teamOneId") or "")
            p2_id = str(p.get("player2Id") or p.get("teamTwoId") or "")
            p1_score = int(p.get("player1Score") or p.get("teamOneScore") or 0)
            p2_score = int(p.get("player2Score") or p.get("teamTwoScore") or 0)

            if not p1_id or not p2_id or p1_score == p2_score:
                continue  # skip byes and draws

            winner_id = p1_id if p1_score > p2_score else p2_id
            loser_id = p2_id if winner_id == p1_id else p1_id

            for pid, side in [(p1_id, "player1"), (p2_id, "player2")]:
                if pid in player_arch:
                    continue
                dl_id = str(
                    p.get(f"{side}DecklistId")
                    or p.get(f"team{'One' if side == 'player1' else 'Two'}DecklistId")
                    or ""
                )
                if dl_id:
                    pair = fetch_archetype(dl_id, verbose)
                    player_arch[pid] = archetype_label(*pair) if pair else None
                else:
                    player_arch[pid] = None

            w_arch = player_arch.get(winner_id)
            l_arch = player_arch.get(loser_id)

            if w_arch and l_arch and w_arch != l_arch:
                wins[(w_arch, l_arch)] += 1
                pair_games[frozenset({w_arch, l_arch})] += 1
                added += 1

    return added


# ── Ranking and matrix ────────────────────────────────────────────────────────

def rank_archetypes(
    pair_games: "defaultdict[FrozenSet[str], int]", top_n: int
) -> List[str]:
    totals: "defaultdict[str, int]" = defaultdict(int)
    for pair, g in pair_games.items():
        for arch in pair:
            totals[arch] += g
    return [a for a, _ in sorted(totals.items(), key=lambda x: -x[1])][:top_n]


def get_wr(
    a: str, b: str,
    wins: "defaultdict[Tuple[str, str], int]",
    pair_games: "defaultdict[FrozenSet[str], int]",
) -> Optional[float]:
    g = pair_games[frozenset({a, b})]
    return wins[(a, b)] / g * 100 if g else None


def get_games(
    a: str, b: str, pair_games: "defaultdict[FrozenSet[str], int]"
) -> int:
    return pair_games[frozenset({a, b})]


# ── Demo data builder ─────────────────────────────────────────────────────────

def _build_demo_data(
    top_n: int,
) -> Tuple[
    "defaultdict[Tuple[str, str], int]",
    "defaultdict[FrozenSet[str], int]",
    List[str],
]:
    """
    Populate wins/pair_games from the guide's matchup matrix.

    Sample sizes are estimated as max(8, round(share_a * share_b * 0.8)).
    This naturally mirrors reality: high-meta-share matchups get ~100 games
    (full colour coding) while fringe-vs-fringe matchups fall below
    LOW_SAMPLE_THRESHOLD and render with an asterisk.
    """
    wins: "defaultdict[Tuple[str, str], int]" = defaultdict(int)
    pair_games: "defaultdict[FrozenSet[str], int]" = defaultdict(int)
    n = min(top_n, len(_DEMO_ARCHETYPES))
    archetypes = _DEMO_ARCHETYPES[:n]

    for i, arch_a in enumerate(archetypes):
        for j, arch_b in enumerate(archetypes):
            if i >= j:
                continue
            wr_a = _DEMO_WIN_PCT[i][j]
            if wr_a is None:
                continue
            n_games = max(8, round(_DEMO_META_SHARE[i] * _DEMO_META_SHARE[j] * 0.8))
            wins_a = round(wr_a / 100 * n_games)
            wins[(arch_a, arch_b)] = wins_a
            wins[(arch_b, arch_a)] = n_games - wins_a
            pair_games[frozenset({arch_a, arch_b})] = n_games

    return wins, pair_games, archetypes


# ── Minimal stdlib XLSX writer ────────────────────────────────────────────────
# XLSX is a ZIP of XML files. We write six files:
#   [Content_Types].xml, _rels/.rels, xl/workbook.xml,
#   xl/_rels/workbook.xml.rels, xl/styles.xml, xl/worksheets/sheet1.xml

# Cell style indices — must match <cellXfs> order in _STYLES_XML below.
_S_DEFAULT    = 0
_S_COL_HEADER = 1   # rotated column header
_S_ROW_LABEL  = 2   # row label / corner
_S_MIRROR     = 3   # mirror match diagonal
_S_WR_60P     = 4   # ≥60% dark green, white text
_S_WR_55      = 5   # 55–59% green
_S_WR_50      = 6   # 50–54% light green
_S_WR_45      = 7   # 45–49% light orange
_S_WR_40      = 8   # 40–44% orange
_S_WR_LOW     = 9   # <40% red, white text
_S_LOW_SAMPLE = 10  # <20 games light grey
_S_NO_DATA    = 11  # no matchup data
_S_TITLE      = 12  # title row
_S_AVG_HDR    = 13  # "Avg WR%" header


_STYLES_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="5">
    <font><sz val="9"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><sz val="9"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><i/><sz val="8"/><color rgb="FF888888"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="12">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1C2833"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD0D0D0"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1A6B3A"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF52BE80"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFABEBC6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFAD7A0"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE59866"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFB03A2E"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE8E8E8"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7F7F7"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="14">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="bottom" textRotation="45"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="9" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="10" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="11" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
</styleSheet>"""

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>'
)

_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)

_WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Matchup Matrix" sheetId="1" r:id="rId1"/></sheets>'
    '</workbook>'
)

_WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>'
)


def _col_letter(n: int) -> str:
    """Convert 1-based column index to Excel column letter (1→A, 27→AA)."""
    name = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        name = chr(65 + r) + name
    return name


def _xe(s: str) -> str:
    return xml.sax.saxutils.escape(str(s))


def _cell(ref: str, value: object, style: int) -> str:
    if isinstance(value, (int, float)):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    return f'<c r="{ref}" s="{style}" t="str"><v>{_xe(str(value))}</v></c>'


def _wr_style(wr: float, n_games: int) -> int:
    if n_games < LOW_SAMPLE_THRESHOLD:
        return _S_LOW_SAMPLE
    if wr >= 60:
        return _S_WR_60P
    if wr >= 55:
        return _S_WR_55
    if wr >= 50:
        return _S_WR_50
    if wr >= 45:
        return _S_WR_45
    if wr >= 40:
        return _S_WR_40
    return _S_WR_LOW


def _build_sheet_xml(
    archetypes: List[str],
    wins: "defaultdict[Tuple[str, str], int]",
    pair_games: "defaultdict[FrozenSet[str], int]",
    start_date: str,
    end_date: str,
    tournament_count: int,
    total_match_count: int,
) -> str:
    n = len(archetypes)
    avg_col = n + 2
    rows: List[str] = []

    # Row 1: title
    title = (
        f"SWU Matchup Matrix | {start_date} – {end_date} | "
        f"{tournament_count} events | {total_match_count} matchups | top {n} archetypes"
    )
    rows.append(f'<row r="1" ht="20" customHeight="1">{_cell("A1", title, _S_TITLE)}</row>')

    # Row 2: column headers
    header_cells = [_cell("A2", "Deck ↓  vs  →", _S_ROW_LABEL)]
    for ci, arch in enumerate(archetypes, start=2):
        header_cells.append(_cell(f"{_col_letter(ci)}2", arch.split(" / ")[0], _S_COL_HEADER))
    header_cells.append(_cell(f"{_col_letter(avg_col)}2", "Avg WR%", _S_AVG_HDR))
    rows.append(f'<row r="2" ht="55" customHeight="1">{"".join(header_cells)}</row>')

    # Data rows
    for ri, row_arch in enumerate(archetypes, start=3):
        cells = [_cell(f"A{ri}", row_arch, _S_ROW_LABEL)]
        row_wrs: List[float] = []

        for ci, col_arch in enumerate(archetypes, start=2):
            ref = f"{_col_letter(ci)}{ri}"
            if row_arch == col_arch:
                cells.append(_cell(ref, "—", _S_MIRROR))
                continue
            wr = get_wr(row_arch, col_arch, wins, pair_games)
            g = get_games(row_arch, col_arch, pair_games)
            if wr is None:
                cells.append(_cell(ref, "N/A", _S_NO_DATA))
            else:
                label = f"{wr:.1f}%*" if g < LOW_SAMPLE_THRESHOLD else f"{wr:.1f}%"
                cells.append(_cell(ref, label, _wr_style(wr, g)))
                if g >= LOW_SAMPLE_THRESHOLD:
                    row_wrs.append(wr)

        avg_ref = f"{_col_letter(avg_col)}{ri}"
        if row_wrs:
            avg = sum(row_wrs) / len(row_wrs)
            cells.append(_cell(avg_ref, f"{avg:.1f}%", _wr_style(avg, 999)))
        else:
            cells.append(_cell(avg_ref, "—", _S_MIRROR))

        rows.append(f'<row r="{ri}">{"".join(cells)}</row>')

    # Legend
    legend_start = n + 4
    legend = [
        ("≥60%  Heavily favoured",      _S_WR_60P),
        ("55–59%  Favoured",            _S_WR_55),
        ("50–54%  Slight edge",          _S_WR_50),
        ("45–49%  Slight disadvantage",  _S_WR_45),
        ("40–44%  Unfavoured",           _S_WR_40),
        ("<40%    Heavily unfavoured",        _S_WR_LOW),
        ("*  < 20 games (low sample)",        _S_LOW_SAMPLE),
        ("N/A  No data",                      _S_NO_DATA),
    ]
    rows.append(
        f'<row r="{legend_start}">'
        f'{_cell(f"A{legend_start}", "Legend", _S_ROW_LABEL)}'
        f'</row>'
    )
    for i, (label, style) in enumerate(legend, start=legend_start + 1):
        rows.append(f'<row r="{i}">{_cell(f"A{i}", label, style)}</row>')

    note_row = legend_start + len(legend) + 2
    rows.append(
        f'<row r="{note_row}">'
        f'{_cell(f"A{note_row}", "Swiss rounds only. Top-cut excluded. Mirrors excluded. Post-rotation data (from 2026-03-13) only.", _S_NO_DATA)}'
        f'</row>'
    )

    cols_xml = (
        f'<cols>'
        f'<col min="1" max="1" width="33" customWidth="1"/>'
        f'<col min="2" max="{n + 1}" width="7" customWidth="1"/>'
        f'<col min="{n + 2}" max="{n + 2}" width="9" customWidth="1"/>'
        f'</cols>'
    )
    freeze_xml = (
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>'
        '</sheetView></sheetViews>'
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + freeze_xml
        + cols_xml
        + f'<sheetData>{"".join(rows)}</sheetData>'
        + '</worksheet>'
    )


def write_xlsx(
    archetypes: List[str],
    wins: "defaultdict[Tuple[str, str], int]",
    pair_games: "defaultdict[FrozenSet[str], int]",
    output_path: str,
    start_date: str,
    end_date: str,
    tournament_count: int,
    total_match_count: int,
) -> None:
    sheet_xml = _build_sheet_xml(
        archetypes, wins, pair_games,
        start_date, end_date, tournament_count, total_match_count,
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("xl/workbook.xml", _WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", _WORKBOOK_RELS)
        zf.writestr("xl/styles.xml", _STYLES_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    with open(output_path, "wb") as f:
        f.write(buf.getvalue())
    print(f"Saved: {output_path}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Build a SWU matchup win-rate matrix from raw melee.gg tournament data.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--timeframe", default=None, metavar="YYYY-MM-DD:YYYY-MM-DD",
                   help="Date range for live data. Required unless --demo is used.")
    p.add_argument("--num-decks", type=int, default=12,
                   help="Top N archetypes by game count (default: 12).")
    p.add_argument("--min-players", type=int, default=MIN_PLAYERS_DEFAULT,
                   help=f"Minimum event size (default: {MIN_PLAYERS_DEFAULT}).")
    p.add_argument("--output", default="matchup_matrix.xlsx")
    p.add_argument("--demo", action="store_true",
                   help="Use guide data instead of live API calls. Validates spreadsheet output.")
    p.add_argument("--dry-run", action="store_true",
                   help="List qualifying tournaments; skip match data fetch.")
    p.add_argument("--verbose", "-v", action="store_true")
    return p.parse_args()


def main() -> None:
    args = parse_args()

    # ── Demo mode — skip all API calls ────────────────────────────────────────
    if args.demo:
        print(f"Demo mode: using SWU LAW Meta Guide matchup matrix (top {args.num_decks} archetypes)\n")
        wins, pair_games, archetypes = _build_demo_data(args.num_decks)
        total_match_count = sum(pair_games.values())
        print(f"Archetypes ({len(archetypes)}):")
        for i, arch in enumerate(archetypes, 1):
            g = sum(get_games(arch, other, pair_games) for other in archetypes if other != arch)
            print(f"  {i:2}.  {arch:<45}  {g} matchup games")
        write_xlsx(
            archetypes, wins, pair_games, args.output,
            "DEMO", "SWU LAW Meta Guide", 0, total_match_count,
        )
        return

    if not args.timeframe:
        sys.exit("Error: --timeframe is required unless --demo is used.")

    try:
        start_str, end_str = args.timeframe.split(":")
        start_date = datetime.strptime(start_str.strip(), "%Y-%m-%d").date()
        end_date = datetime.strptime(end_str.strip(), "%Y-%m-%d").date()
    except ValueError:
        sys.exit("Error: --timeframe must be YYYY-MM-DD:YYYY-MM-DD")

    if str(start_date) < POST_ROTATION_DATE:
        print(
            f"Warning: {start_date} predates post-rotation ({POST_ROTATION_DATE}). "
            f"Pre-rotation card legality differs — consider adjusting --timeframe.",
            file=sys.stderr,
        )

    # ── Fetch tournaments ──────────────────────────────────────────────────────
    print(f"Fetching SWU tournaments: {start_date} → {end_date}  (min {args.min_players} players)...")

    try:
        tournaments = fetch_tournaments(
            str(start_date), str(end_date), args.min_players, args.verbose
        )
    except RuntimeError as exc:
        sys.exit(f"\n{exc}\n")
    except urllib.error.URLError as exc:
        sys.exit(
            f"\nNetwork error: {exc}\n"
            f"Possible causes:\n"
            f"  • No internet access in this environment\n"
            f"  • melee.gg is temporarily unavailable\n"
            f"  • melee.gg requires session authentication for API access\n"
            f"If running in a restricted environment, run this script locally instead.\n"
        )

    if not tournaments:
        sys.exit(
            f"No tournaments found (game='{SWU_GAME_SLUG}', {start_date}–{end_date}).\n"
            f"Check SWU_GAME_SLUG or widen the date range.\n"
        )

    print(f"Found {len(tournaments)} qualifying tournament(s):\n")
    for t in sorted(tournaments, key=lambda x: x.get("date") or x.get("startDate") or ""):
        count = t.get("playerCount") or t.get("registrations") or "?"
        d = t.get("date") or t.get("startDate") or "unknown date"
        print(f"  • {t.get('name', 'Unnamed'):<55}  {count!s:>4} players  {d}")

    if args.dry_run:
        print("\nDry run complete. Re-run without --dry-run to process match data.")
        return

    # ── Process match data ─────────────────────────────────────────────────────
    wins: "defaultdict[Tuple[str, str], int]" = defaultdict(int)
    pair_games: "defaultdict[FrozenSet[str], int]" = defaultdict(int)
    total_match_count = 0

    print(f"\nProcessing match results ({len(tournaments)} events)...\n")
    for t in tournaments:
        name = t.get("name", "Unnamed")
        try:
            added = process_tournament(t, wins, pair_games, args.verbose)
            total_match_count += added
            print(f"  ✓  {name:<55}  {added} matchups")
        except Exception as exc:  # noqa: BLE001
            print(f"  ✗  {name:<55}  ERROR: {exc}", file=sys.stderr)

    if total_match_count == 0:
        sys.exit(
            "\nNo match data collected.\n"
            "Possible causes:\n"
            "  • melee.gg pairings/decklists may not be publicly accessible\n"
            "  • Pairing API endpoint may have changed (try --verbose)\n"
            "  • Authentication may be required\n"
            "\nFallback options:\n"
            "  • Ask a tournament organiser for a melee.gg export CSV\n"
            "  • Check swu-competitivehub.com for structured tournament data\n"
            "  • Use swumetastats.com/api-docs for aggregated (not raw) matchup data\n"
        )

    print(f"\nTotal matchup records: {total_match_count}")

    # ── Rank archetypes (purely from data) ────────────────────────────────────
    archetypes = rank_archetypes(pair_games, args.num_decks)
    print(f"\nTop {len(archetypes)} archetypes by game count:\n")
    for i, arch in enumerate(archetypes, 1):
        g = sum(get_games(arch, other, pair_games) for other in archetypes if other != arch)
        print(f"  {i:2}.  {arch:<45}  {g} matchup games")

    low_confidence = [
        (arch, sum(get_games(arch, o, pair_games) for o in archetypes if o != arch))
        for arch in archetypes
        if sum(get_games(arch, o, pair_games) for o in archetypes if o != arch) < 50
    ]
    if low_confidence:
        print("\nLow-confidence archetypes (< 50 total games):")
        for arch, g in low_confidence:
            print(f"  ⚠  {arch}  ({g} games)")

    # ── Write spreadsheet ──────────────────────────────────────────────────────
    write_xlsx(
        archetypes, wins, pair_games,
        args.output, str(start_date), str(end_date),
        len(tournaments), total_match_count,
    )

    # ── Summary ────────────────────────────────────────────────────────────────
    print("\n── Average win rates (excluding low-sample matchups) ──\n")
    summary = []
    for arch in archetypes:
        high_conf = [
            get_wr(arch, opp, wins, pair_games)
            for opp in archetypes
            if opp != arch and get_games(arch, opp, pair_games) >= LOW_SAMPLE_THRESHOLD
        ]
        high_conf_wrs = [w for w in high_conf if w is not None]
        avg = sum(high_conf_wrs) / len(high_conf_wrs) if high_conf_wrs else None
        summary.append((arch, avg, len(high_conf_wrs)))
    summary.sort(key=lambda x: -(x[1] or 0))
    for arch, avg, n_matchups in summary:
        avg_str = f"{avg:.1f}%" if avg is not None else "N/A"
        print(f"  {arch:<45}  avg {avg_str}  ({n_matchups} high-conf matchups)")

    print(f"\n* = fewer than {LOW_SAMPLE_THRESHOLD} games — treat as low-confidence\n")


if __name__ == "__main__":
    main()
