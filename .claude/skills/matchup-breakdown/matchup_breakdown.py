#!/usr/bin/env python3
"""
SWU Matchup Breakdown — stdlib only, no external dependencies.

Scrapes public melee.gg pages (no API key required) and builds a color-coded
head-to-head win rate matrix spreadsheet for the Star Wars Unlimited meta.
Archetypes are discovered dynamically from DecklistName fields in match data.

Usage:
    python matchup_breakdown.py --timeframe 2026-03-13:2026-05-31 --num-decks 12
    python matchup_breakdown.py --timeframe 2026-05-01:2026-05-31 --num-decks 5 --output may.xlsx
    python matchup_breakdown.py --timeframe 2026-03-13:2026-05-31 --dry-run --verbose
"""

import argparse
import io
import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.sax.saxutils
import zipfile
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Dict, FrozenSet, List, Optional, Tuple

# ── Constants ─────────────────────────────────────────────────────────────────

MELEE_BASE = "https://melee.gg"
POST_ROTATION_DATE = "2026-03-13"
REQUEST_DELAY = 1.0        # seconds between requests (respects robots.txt crawl delay)
MIN_PLAYERS_DEFAULT = 32   # matches swu-competitivehub.com filter standard
LOW_SAMPLE_THRESHOLD = 20  # matchups below this count get an asterisk

# DataTables column names used by /Match/GetRoundMatches (must match pairings-section.min.js)
_MATCH_COLUMNS = ["TableNumber", "PodNumber", "Teams", "Decklists", "ResultString"]

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
        html = r.read().decode("utf-8", errors="replace")
    _HTTP_CACHE[url] = html
    return html


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
    start_date: str, end_date: str, min_players: int, verbose: bool
) -> List[dict]:
    """
    Return all SWU Premier ended tournaments in the date range with >= min_players.

    Uses POST /Tournament/TournamentSearch — no API key required.
    Filters: Ended + Premier (server-side); date range and player count are client-side.
    Paginates from near the end of the sorted result set to find recent events.
    """
    search_params_base = [
        ("ordering", "StartDate"), ("mode", "Table"),
        ("filters[]", "Ended"), ("filters[]", "Premier"),
    ]

    # Get total record count first (one cheap request)
    probe = _post_json("/Tournament/TournamentSearch",
                       search_params_base + _search_dt_params(0, 25), verbose=verbose)
    total = int(probe.get("recordsTotal") or 0)  # type: ignore[union-attr]
    if verbose:
        print(f"  Total Ended+Premier SWU events indexed: {total}", file=sys.stderr)

    # Estimate how far back we need to go: ~36 events/day is conservative
    days = max(1, (datetime.strptime(end_date, "%Y-%m-%d") -
                   datetime.strptime(start_date, "%Y-%m-%d")).days)
    fetch_window = max(500, days * 50)
    start_offset = max(0, total - fetch_window)

    results: List[dict] = []
    batch_size = 250

    while start_offset <= total:
        resp = _post_json("/Tournament/TournamentSearch",
                          search_params_base + _search_dt_params(start_offset, batch_size),
                          verbose=verbose)
        items = resp.get("data", [])  # type: ignore[union-attr]
        if not items:
            break

        for t in items:
            d = _t_date(t)
            if d < start_date or d > end_date:
                continue
            if _t_players(t) >= min_players:
                results.append(t)

        start_offset += batch_size

    return results


# ── Archetype discovery ───────────────────────────────────────────────────────

def parse_decklist_name(name: str) -> Optional[Tuple[str, str]]:
    """
    Parse 'Leader Name, Subtitle - Base Name' into (leader, base).

    melee.gg stores the archetype directly in DecklistName as 'Leader - Base'.
    We split on the last ' - ' to handle leaders whose subtitles contain dashes.
    """
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


def archetype_label(leader: str, base: str) -> str:
    return f"{leader} / {base}"


# ── Round and match fetching ──────────────────────────────────────────────────

def _is_swiss_round_name(name: str) -> bool:
    """Swiss rounds are named 'Round N'; top-cut rounds have descriptive names."""
    return bool(re.match(r"^Round\s+\d+$", name.strip(), re.IGNORECASE))


def fetch_swiss_round_ids(tournament_id: int, verbose: bool) -> Tuple[List[str], bool]:
    """
    Parse /Tournament/View/{id} HTML to extract Swiss round IDs and decklist visibility.
    Returns (swiss_round_ids, decklist_enabled).
    """
    html = _get_html(f"/Tournament/View/{tournament_id}", verbose=verbose)

    # Round selector buttons: <button ... data-id="NNN" data-name="Round N">
    round_buttons = re.findall(
        r'<button[^>]+class="[^"]*round-selector[^"]*"[^>]*'
        r'data-id="(\d+)"[^>]*data-name="([^"]+)"',
        html,
    )
    # Also try with attributes in reversed order
    if not round_buttons:
        round_buttons = [
            (rid, name)
            for name, rid in re.findall(
                r'<button[^>]+data-name="([^"]+)"[^>]+data-id="(\d+)"', html
            )
        ]

    swiss_ids = [rid for rid, name in round_buttons if _is_swiss_round_name(name)]

    dl_match = re.search(r'name="DecklistEnabled"[^>]+value="([^"]+)"', html)
    decklist_enabled = bool(dl_match and dl_match.group(1).lower() == "true")

    if verbose:
        all_names = [name for _, name in round_buttons]
        print(f"  [debug] rounds: {all_names} | decklists: {decklist_enabled}", file=sys.stderr)

    return swiss_ids, decklist_enabled


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


# ── Match processing ──────────────────────────────────────────────────────────

def process_tournament(
    tournament: dict,
    wins: "defaultdict[Tuple[str, str], int]",
    pair_games: "defaultdict[FrozenSet[str], int]",
    verbose: bool,
) -> int:
    """
    Fetch Swiss round matches for a tournament. For each completed head-to-head,
    extract archetypes from DecklistName and record the result.
    Returns the number of matchup records added.
    """
    tid = _t_id(tournament)
    if not tid:
        return 0

    swiss_round_ids, decklist_enabled = fetch_swiss_round_ids(tid, verbose)

    if not decklist_enabled:
        if verbose:
            print(f"  [skip] {_t_name(tournament)}: decklists not public", file=sys.stderr)
        return 0

    if not swiss_round_ids:
        if verbose:
            print(f"  [skip] {_t_name(tournament)}: no Swiss rounds found", file=sys.stderr)
        return 0

    added = 0
    for round_id in swiss_round_ids:
        try:
            matches = fetch_round_matches(round_id, verbose)
        except Exception as exc:
            if verbose:
                print(f"  [warn] round {round_id}: {exc}", file=sys.stderr)
            continue

        for m in matches:
            comps = m.get("Competitors", [])
            if len(comps) != 2 or not m.get("HasResult"):
                continue  # bye or unfinished

            c1, c2 = comps
            w1 = int(c1.get("GameWins") or 0)
            w2 = int(c2.get("GameWins") or 0)
            if w1 == w2:
                continue  # draw

            winner, loser = (c1, c2) if w1 > w2 else (c2, c1)

            dl_w = (winner.get("Decklists") or [{}])[0]
            dl_l = (loser.get("Decklists") or [{}])[0]

            w_pair = parse_decklist_name(dl_w.get("DecklistName", ""))
            l_pair = parse_decklist_name(dl_l.get("DecklistName", ""))
            if not w_pair or not l_pair:
                continue

            w_arch = archetype_label(*w_pair)
            l_arch = archetype_label(*l_pair)
            if w_arch == l_arch:
                continue  # mirror

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
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="bottom" textRotation="90"/></xf>
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
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
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
    '<sheets>'
    '<sheet name="Matchup Matrix" sheetId="1" r:id="rId1"/>'
    '<sheet name="Matchup Counts" sheetId="2" r:id="rId2"/>'
    '</sheets>'
    '</workbook>'
)

_WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
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


def _build_counts_sheet_xml(
    archetypes: List[str],
    pair_games: "defaultdict[FrozenSet[str], int]",
    start_date: str,
    end_date: str,
    tournament_count: int,
    total_match_count: int,
) -> str:
    n = len(archetypes)
    total_col = n + 2
    rows: List[str] = []

    title = (
        f"SWU Matchup Counts | {start_date} – {end_date} | "
        f"{tournament_count} events | {total_match_count} matchups | top {n} archetypes"
    )
    rows.append(f'<row r="1" ht="20" customHeight="1">{_cell("A1", title, _S_TITLE)}</row>')

    header_cells = [_cell("A2", "Deck ↓  vs  →", _S_ROW_LABEL)]
    for ci, arch in enumerate(archetypes, start=2):
        header_cells.append(_cell(f"{_col_letter(ci)}2", arch.split(" / ")[0], _S_COL_HEADER))
    header_cells.append(_cell(f"{_col_letter(total_col)}2", "Total", _S_AVG_HDR))
    rows.append(f'<row r="2" ht="55" customHeight="1">{"".join(header_cells)}</row>')

    for ri, row_arch in enumerate(archetypes, start=3):
        cells = [_cell(f"A{ri}", row_arch, _S_ROW_LABEL)]
        row_total = 0
        for ci, col_arch in enumerate(archetypes, start=2):
            ref = f"{_col_letter(ci)}{ri}"
            if row_arch == col_arch:
                cells.append(_cell(ref, "—", _S_MIRROR))
                continue
            g = get_games(row_arch, col_arch, pair_games)
            if g == 0:
                cells.append(_cell(ref, "—", _S_NO_DATA))
            else:
                cells.append(_cell(ref, g, _S_DEFAULT))
                row_total += g
        total_ref = f"{_col_letter(total_col)}{ri}"
        cells.append(_cell(total_ref, row_total if row_total else "—",
                           _S_AVG_HDR if row_total else _S_MIRROR))
        rows.append(f'<row r="{ri}">{"".join(cells)}</row>')

    note_row = n + 5
    rows.append(
        f'<row r="{note_row}">'
        f'{_cell(f"A{note_row}", "Game counts from Swiss rounds only. Mirror matches excluded.", _S_NO_DATA)}'
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
    sheet1_xml = _build_sheet_xml(
        archetypes, wins, pair_games,
        start_date, end_date, tournament_count, total_match_count,
    )
    sheet2_xml = _build_counts_sheet_xml(
        archetypes, pair_games,
        start_date, end_date, tournament_count, total_match_count,
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("xl/workbook.xml", _WORKBOOK)
        zf.writestr("xl/_rels/workbook.xml.rels", _WORKBOOK_RELS)
        zf.writestr("xl/styles.xml", _STYLES_XML)
        zf.writestr("xl/worksheets/sheet1.xml", sheet1_xml)
        zf.writestr("xl/worksheets/sheet2.xml", sheet2_xml)
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
                   help="Date range (default: last 30 days).")
    p.add_argument("--num-decks", type=int, default=10,
                   help="Top N archetypes by game count (default: 10).")
    p.add_argument("--min-players", type=int, default=MIN_PLAYERS_DEFAULT,
                   help=f"Minimum event size (default: {MIN_PLAYERS_DEFAULT}).")
    p.add_argument("--output", default="matchup_matrix.xlsx")
    p.add_argument("--dry-run", action="store_true",
                   help="List qualifying tournaments; skip match data fetch.")
    p.add_argument("--verbose", "-v", action="store_true")
    return p.parse_args()


def main() -> None:
    # Ensure UTF-8 output on Windows where the default cp1252 codec can't
    # encode the special characters used in progress/summary lines.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    args = parse_args()

    today = datetime.now().date()
    if args.timeframe is None:
        start_date = today - timedelta(days=30)
        end_date = today
    else:
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
            "Possible causes:\n"
            "  • No internet access in this environment\n"
            "  • melee.gg is temporarily unavailable\n"
            "If running in a restricted environment, run this script locally instead.\n"
        )

    if not tournaments:
        sys.exit(
            f"No SWU Premier tournaments found ({start_date} to {end_date}, min {args.min_players} players).\n"
            "Possible causes:\n"
            "  • No events in the date range — try widening --timeframe\n"
            "  • Game/format client-side filtering is too strict (try --verbose to see raw data)\n"
        )

    print(f"Found {len(tournaments)} qualifying tournament(s):\n")
    for t in sorted(tournaments, key=_t_date):
        count = _t_players(t)
        d = _t_date(t)
        print(f"  • {_t_name(t):<55}  {count:>4} players  {d}")

    if args.dry_run:
        print("\nDry run complete. Re-run without --dry-run to process match data.")
        return

    # ── Process match data ─────────────────────────────────────────────────────
    wins: "defaultdict[Tuple[str, str], int]" = defaultdict(int)
    pair_games: "defaultdict[FrozenSet[str], int]" = defaultdict(int)
    total_match_count = 0

    print(f"\nProcessing match results ({len(tournaments)} events)...\n")
    for t in tournaments:
        name = _t_name(t)
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
            "  • All qualifying tournaments had decklists disabled (not made public by organizer)\n"
            "  • melee.gg page structure may have changed (try --verbose)\n"
            "\nFallback options:\n"
            "  • Widen the date range to include more events\n"
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
