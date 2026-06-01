#!/usr/bin/env python3
"""
SWU Meta Breakdown — stdlib only, no external dependencies.

Scrapes public melee.gg pages and produces a stacked area chart XLSX
showing how each archetype's meta share percentage has evolved week by week.

Usage:
    python meta_breakdown.py --timeframe 2026-03-13:2026-05-31 --num-decks 10
    python meta_breakdown.py --timeframe 2026-05-01:2026-05-31 --output may_meta.xlsx
    python meta_breakdown.py --timeframe 2026-03-13:2026-05-31 --dry-run --verbose
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
from typing import Dict, List, Optional, Tuple

# ── Constants ─────────────────────────────────────────────────────────────────

MELEE_BASE = "https://melee.gg"
POST_ROTATION_DATE = "2026-03-13"
REQUEST_DELAY = 1.0        # seconds between requests (respects robots.txt crawl delay)
MIN_PLAYERS_DEFAULT = 32   # matches swu-competitivehub.com filter standard

# DataTables column names used by /Match/GetRoundMatches (must match pairings-section.min.js)
_MATCH_COLUMNS = ["TableNumber", "PodNumber", "Teams", "Decklists", "ResultString"]

# OOXML-compatible RGB hex colors for archetype series (no leading alpha byte)
_ARCH_COLORS = [
    "4472C4",  # blue
    "ED7D31",  # orange
    "70AD47",  # green
    "FFC000",  # gold
    "5B9BD5",  # light blue
    "FF0000",  # red
    "7030A0",  # purple
    "00B0F0",  # cyan
    "FF6600",  # dark orange
    "A9D18E",  # light green
    "C9AB2E",  # yellow-brown
    "FF0066",  # pink
]
_OTHER_COLOR = "BFBFBF"    # gray for the "Other" bucket

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
    """
    search_params_base = [
        ("ordering", "StartDate"), ("mode", "Table"),
        ("filters[]", "Ended"), ("filters[]", "Premier"),
    ]

    probe = _post_json("/Tournament/TournamentSearch",
                       search_params_base + _search_dt_params(0, 25), verbose=verbose)
    total = int(probe.get("recordsTotal") or 0)  # type: ignore[union-attr]
    if verbose:
        print(f"  Total Ended+Premier SWU events indexed: {total}", file=sys.stderr)

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

    Splits on the last ' - ' to handle leaders whose subtitles contain dashes.
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

    round_buttons = re.findall(
        r'<button[^>]+class="[^"]*round-selector[^"]*"[^>]*'
        r'data-id="(\d+)"[^>]*data-name="([^"]+)"',
        html,
    )
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


# ── Meta share tracking ───────────────────────────────────────────────────────

def week_start(date_str: str) -> str:
    """Return the ISO date string (YYYY-MM-DD) of the Monday starting the week."""
    d = datetime.strptime(date_str[:10], "%Y-%m-%d").date()
    monday = d - timedelta(days=d.weekday())
    return str(monday)


def process_tournament_meta(
    tournament: dict,
    week_arch_counts: "defaultdict[str, defaultdict[str, int]]",
    verbose: bool,
) -> int:
    """
    Count archetype appearances per week bucket.

    Each competitor in each Swiss match counts as 1 appearance for their archetype.
    The tournament's start date determines which week bucket it falls in.
    Returns total appearances added.
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

    week = week_start(_t_date(tournament))
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
            for c in comps:
                dl = (c.get("Decklists") or [{}])[0]
                pair = parse_decklist_name(dl.get("DecklistName", ""))
                if pair:
                    arch = archetype_label(*pair)
                    week_arch_counts[week][arch] += 1
                    added += 1

    return added


def compute_top_archetypes(
    week_arch_counts: "defaultdict[str, defaultdict[str, int]]", top_n: int
) -> List[str]:
    """Return top N archetypes by total appearances across all weeks."""
    totals: "defaultdict[str, int]" = defaultdict(int)
    for week_data in week_arch_counts.values():
        for arch, count in week_data.items():
            totals[arch] += count
    return [a for a, _ in sorted(totals.items(), key=lambda x: -x[1])][:top_n]


def compute_meta_shares(
    weeks: List[str],
    week_arch_counts: "defaultdict[str, defaultdict[str, int]]",
    top_archs: List[str],
) -> Dict[str, Dict[str, float]]:
    """
    Compute meta share % for each archetype (plus "Other") per week.
    Returns {arch: {week: pct}} where values sum to ~100% per week.
    """
    shares: Dict[str, Dict[str, float]] = {arch: {} for arch in top_archs}
    shares["Other"] = {}

    for week in weeks:
        week_data = week_arch_counts.get(week, {})
        total = sum(week_data.values())
        if total == 0:
            for arch in top_archs:
                shares[arch][week] = 0.0
            shares["Other"][week] = 0.0
            continue

        for arch in top_archs:
            shares[arch][week] = week_data.get(arch, 0) / total * 100.0
        other_count = total - sum(week_data.get(arch, 0) for arch in top_archs)
        shares["Other"][week] = max(0.0, other_count / total * 100.0)

    return shares


# ── XLSX writer ───────────────────────────────────────────────────────────────
# XLSX is a ZIP of XML files. We produce:
#   [Content_Types].xml, _rels/.rels, xl/workbook.xml,
#   xl/_rels/workbook.xml.rels, xl/styles.xml,
#   xl/worksheets/sheet1.xml  (data table),
#   xl/worksheets/sheet2.xml  (chart host),
#   xl/worksheets/_rels/sheet2.xml.rels,
#   xl/drawings/drawing1.xml, xl/drawings/_rels/drawing1.xml.rels,
#   xl/charts/chart1.xml

# Cell style indices — must match <cellXfs> order in _STYLES_XML below.
_S_DEFAULT   = 0
_S_TITLE     = 1   # bold title row
_S_COL_HDR   = 2   # dark header, rotated
_S_ROW_LABEL = 3   # gray row label
_S_NUMBER    = 4   # percentage number (custom fmt "0.0%")
_S_OTHER_LBL = 5   # "Other" row label, bold
_S_OTHER_NUM = 6   # "Other" row number
_S_NOTE      = 7   # italic footer note


_STYLES_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.0&quot;%&quot;"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="9"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="9"/><name val="Calibri"/></font>
    <font><sz val="9"/><name val="Calibri"/></font>
    <font><i/><sz val="8"/><color rgb="FF888888"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1C2833"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD5D8DC"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="bottom" textRotation="45"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="4" fillId="0" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="2" fillId="2" borderId="0" xfId="0"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0"><alignment horizontal="left" vertical="center"/></xf>
  </cellXfs>
</styleSheet>"""


def _col_letter(n: int) -> str:
    """Convert 1-based column index to Excel letter (1→A, 27→AA)."""
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


def _format_week_label(week_str: str) -> str:
    """Format a week-start date string as 'Mon DD' (e.g. 'May 04')."""
    d = datetime.strptime(week_str, "%Y-%m-%d").date()
    return d.strftime("%b %d")


def _build_data_sheet_xml(
    archetypes_with_other: List[str],
    weeks: List[str],
    shares: Dict[str, Dict[str, float]],
    start_date: str,
    end_date: str,
    tournament_count: int,
    total_appearances: int,
) -> str:
    n_weeks = len(weeks)
    n_archs = len(archetypes_with_other)
    rows: List[str] = []

    # Row 1: title
    title = (
        f"SWU Meta Share Trend | {start_date} – {end_date} | "
        f"{tournament_count} events | {total_appearances} deck-round appearances"
    )
    rows.append(f'<row r="1" ht="20" customHeight="1">{_cell("A1", title, _S_TITLE)}</row>')

    # Row 2: column headers (week labels)
    header_cells = [_cell("A2", "Deck / Week", _S_COL_HDR)]
    for ci, week in enumerate(weeks, start=2):
        header_cells.append(_cell(f"{_col_letter(ci)}2", _format_week_label(week), _S_COL_HDR))
    rows.append(f'<row r="2" ht="48" customHeight="1">{"".join(header_cells)}</row>')

    # Data rows: one per archetype, "Other" last
    for ri, arch in enumerate(archetypes_with_other, start=3):
        is_other = arch == "Other"
        lbl_style = _S_OTHER_LBL if is_other else _S_ROW_LABEL
        num_style = _S_OTHER_NUM if is_other else _S_NUMBER
        cells = [_cell(f"A{ri}", arch, lbl_style)]
        for ci, week in enumerate(weeks, start=2):
            val = round(shares.get(arch, {}).get(week, 0.0), 2)
            cells.append(_cell(f"{_col_letter(ci)}{ri}", val, num_style))
        rows.append(f'<row r="{ri}">{"".join(cells)}</row>')

    # Footer note
    note_row = n_archs + 4
    rows.append(
        f'<row r="{note_row}">'
        f'{_cell(f"A{note_row}", "Swiss rounds only. Top-cut excluded. Post-rotation (2026-03-13+) only. Each match-competitor counted as 1 appearance per round.", _S_NOTE)}'
        f'</row>'
    )

    cols_xml = (
        f'<cols>'
        f'<col min="1" max="1" width="42" customWidth="1"/>'
        f'<col min="2" max="{n_weeks + 1}" width="9" customWidth="1"/>'
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


def _build_chart_xml(
    archetypes_with_other: List[str],
    weeks: List[str],
    shares: Dict[str, Dict[str, float]],
    chart_title: str,
) -> str:
    """
    Build OOXML stacked area chart XML referencing 'Meta Share Data' sheet.

    Produces a 100%-visual stacked area where series values are pre-computed
    percentages that sum to ~100 per week bucket.
    """
    n_weeks = len(weeks)
    first_col = _col_letter(2)
    last_col = _col_letter(n_weeks + 1)

    series_parts: List[str] = []
    for i, arch in enumerate(archetypes_with_other):
        color = _ARCH_COLORS[i % len(_ARCH_COLORS)] if arch != "Other" else _OTHER_COLOR
        row = i + 3  # row 3 = first archetype in data sheet

        name_f = f"'Meta Share Data'!$A${row}"
        cat_f = f"'Meta Share Data'!${first_col}$2:${last_col}$2"
        val_f = f"'Meta Share Data'!${first_col}${row}:${last_col}${row}"

        cat_pts = "".join(
            f'<c:pt idx="{j}"><c:v>{_xe(_format_week_label(w))}</c:v></c:pt>'
            for j, w in enumerate(weeks)
        )
        val_pts = "".join(
            f'<c:pt idx="{j}"><c:v>{shares.get(arch, {}).get(w, 0.0):.4f}</c:v></c:pt>'
            for j, w in enumerate(weeks)
        )

        series_parts.append(
            f'<c:ser>'
            f'<c:idx val="{i}"/><c:order val="{i}"/>'
            f'<c:tx>'
            f'<c:strRef>'
            f'<c:f>{name_f}</c:f>'
            f'<c:strCache>'
            f'<c:ptCount val="1"/><c:pt idx="0"><c:v>{_xe(arch)}</c:v></c:pt>'
            f'</c:strCache>'
            f'</c:strRef>'
            f'</c:tx>'
            f'<c:spPr>'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:ln><a:noFill/></a:ln>'
            f'</c:spPr>'
            f'<c:cat>'
            f'<c:strRef>'
            f'<c:f>{cat_f}</c:f>'
            f'<c:strCache><c:ptCount val="{n_weeks}"/>{cat_pts}</c:strCache>'
            f'</c:strRef>'
            f'</c:cat>'
            f'<c:val>'
            f'<c:numRef>'
            f'<c:f>{val_f}</c:f>'
            f'<c:numCache>'
            f'<c:formatCode>0.0</c:formatCode>'
            f'<c:ptCount val="{n_weeks}"/>{val_pts}'
            f'</c:numCache>'
            f'</c:numRef>'
            f'</c:val>'
            f'</c:ser>'
        )

    # Axis IDs must be unique integers referenced by both chart and axis elements
    ax_cat = "1"
    ax_val = "2"

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace'
        ' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:lang val="en-US"/>'
        '<c:chart>'
        '<c:title>'
        '<c:tx><c:rich>'
        '<a:bodyPr/><a:lstStyle/>'
        f'<a:p><a:r><a:t>{_xe(chart_title)}</a:t></a:r></a:p>'
        '</c:rich></c:tx>'
        '<c:overlay val="0"/>'
        '</c:title>'
        '<c:autoTitleDeleted val="0"/>'
        '<c:plotArea>'
        '<c:layout/>'
        '<c:areaChart>'
        '<c:grouping val="stacked"/>'
        '<c:varyColors val="0"/>'
        + "".join(series_parts) +
        f'<c:axId val="{ax_cat}"/>'
        f'<c:axId val="{ax_val}"/>'
        '</c:areaChart>'
        f'<c:catAx>'
        f'<c:axId val="{ax_cat}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="b"/>'
        '<c:tickMark val="out"/>'
        '<c:tickLblPos val="nextTo"/>'
        f'<c:crossAx val="{ax_val}"/>'
        '</c:catAx>'
        f'<c:valAx>'
        f'<c:axId val="{ax_val}"/>'
        '<c:scaling>'
        '<c:orientation val="minMax"/>'
        '<c:max val="100"/>'
        '</c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="l"/>'
        '<c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0"/>'
        '<c:tickMark val="out"/>'
        '<c:tickLblPos val="nextTo"/>'
        f'<c:crossAx val="{ax_cat}"/>'
        '</c:valAx>'
        '</c:plotArea>'
        '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/>'
        '</c:chart>'
        '</c:chartSpace>'
    )


def write_xlsx(
    archetypes_with_other: List[str],
    weeks: List[str],
    shares: Dict[str, Dict[str, float]],
    output_path: str,
    start_date: str,
    end_date: str,
    tournament_count: int,
    total_appearances: int,
) -> None:
    chart_title = (
        f"SWU Meta Share Trend — {start_date} to {end_date} "
        f"({tournament_count} events)"
    )

    data_sheet = _build_data_sheet_xml(
        archetypes_with_other, weeks, shares,
        start_date, end_date, tournament_count, total_appearances,
    )
    chart_xml = _build_chart_xml(archetypes_with_other, weeks, shares, chart_title)

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/worksheets/sheet2.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/drawings/drawing1.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        '<Override PartName="/xl/charts/chart1.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        '</Types>'
    )

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
        ' Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets>'
        '<sheet name="Meta Share Data" sheetId="1" r:id="rId1"/>'
        '<sheet name="Meta Trend Chart" sheetId="2" r:id="rId2"/>'
        '</sheets>'
        '</workbook>'
    )

    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
        ' Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
        ' Target="worksheets/sheet2.xml"/>'
        '<Relationship Id="rId3"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
        ' Target="styles.xml"/>'
        '</Relationships>'
    )

    # sheet2 hosts the chart via a drawing reference
    sheet2 = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetData/>'
        '<drawing r:id="rId1"/>'
        '</worksheet>'
    )

    sheet2_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"'
        ' Target="../drawings/drawing1.xml"/>'
        '</Relationships>'
    )

    drawing = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr'
        ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<xdr:twoCellAnchor>'
        '<xdr:from>'
        '<xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>'
        '</xdr:from>'
        '<xdr:to>'
        '<xdr:col>18</xdr:col><xdr:colOff>0</xdr:colOff>'
        '<xdr:row>30</xdr:row><xdr:rowOff>0</xdr:rowOff>'
        '</xdr:to>'
        '<xdr:graphicFrame macro="">'
        '<xdr:nvGraphicFramePr>'
        '<xdr:cNvPr id="2" name="Chart 1"/>'
        '<xdr:cNvGraphicFramePr/>'
        '</xdr:nvGraphicFramePr>'
        '<xdr:xfrm>'
        '<a:off x="0" y="0"/>'
        '<a:ext cx="0" cy="0"/>'
        '</xdr:xfrm>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart'
        ' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' r:id="rId1"/>'
        '</a:graphicData>'
        '</a:graphic>'
        '</xdr:graphicFrame>'
        '<xdr:clientData/>'
        '</xdr:twoCellAnchor>'
        '</xdr:wsDr>'
    )

    drawing_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1"'
        ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"'
        ' Target="../charts/chart1.xml"/>'
        '</Relationships>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", _STYLES_XML)
        zf.writestr("xl/worksheets/sheet1.xml", data_sheet)
        zf.writestr("xl/worksheets/sheet2.xml", sheet2)
        zf.writestr("xl/worksheets/_rels/sheet2.xml.rels", sheet2_rels)
        zf.writestr("xl/drawings/drawing1.xml", drawing)
        zf.writestr("xl/drawings/_rels/drawing1.xml.rels", drawing_rels)
        zf.writestr("xl/charts/chart1.xml", chart_xml)

    with open(output_path, "wb") as f:
        f.write(buf.getvalue())
    print(f"Saved: {output_path}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Build a SWU meta share trend chart from raw melee.gg tournament data.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("--timeframe", default=None, metavar="YYYY-MM-DD:YYYY-MM-DD",
                   help="Date range (default: last 30 days).")
    p.add_argument("--num-decks", type=int, default=10,
                   help="Top N archetypes shown individually; rest grouped into 'Other' (default: 10).")
    p.add_argument("--min-players", type=int, default=MIN_PLAYERS_DEFAULT,
                   help=f"Minimum event size (default: {MIN_PLAYERS_DEFAULT}).")
    p.add_argument("--output", default="meta_trend.xlsx")
    p.add_argument("--dry-run", action="store_true",
                   help="List qualifying tournaments; skip match data fetch.")
    p.add_argument("--verbose", "-v", action="store_true")
    return p.parse_args()


def main() -> None:
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
            "Pre-rotation card legality differs — consider adjusting --timeframe.",
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
            f"No SWU Premier tournaments found ({start_date} to {end_date}, "
            f"min {args.min_players} players).\n"
            "Possible causes:\n"
            "  • No events in the date range — try widening --timeframe\n"
            "  • Game/format filtering may be too strict (try --verbose)\n"
        )

    print(f"Found {len(tournaments)} qualifying tournament(s):\n")
    for t in sorted(tournaments, key=_t_date):
        print(f"  • {_t_name(t):<55}  {_t_players(t):>4} players  {_t_date(t)}")

    if args.dry_run:
        print("\nDry run complete. Re-run without --dry-run to process match data.")
        return

    # ── Process match data ─────────────────────────────────────────────────────
    week_arch_counts: "defaultdict[str, defaultdict[str, int]]" = \
        defaultdict(lambda: defaultdict(int))
    total_appearances = 0

    print(f"\nProcessing match results ({len(tournaments)} events)...\n")
    for t in tournaments:
        name = _t_name(t)
        try:
            added = process_tournament_meta(t, week_arch_counts, args.verbose)
            total_appearances += added
            print(f"  ✓  {name:<55}  {added} appearances")
        except Exception as exc:  # noqa: BLE001
            print(f"  ✗  {name:<55}  ERROR: {exc}", file=sys.stderr)

    if total_appearances == 0:
        sys.exit(
            "\nNo deck appearance data collected.\n"
            "Possible causes:\n"
            "  • All qualifying tournaments had decklists disabled\n"
            "  • melee.gg page structure may have changed (try --verbose)\n"
            "\nFallback options:\n"
            "  • Widen the date range to include more events\n"
            "  • Use swumetastats.com/api-docs for aggregated meta data\n"
        )

    print(f"\nTotal deck appearances recorded: {total_appearances}")

    # ── Compute meta shares ────────────────────────────────────────────────────
    top_archs = compute_top_archetypes(week_arch_counts, args.num_decks)
    weeks = sorted(week_arch_counts.keys())
    shares = compute_meta_shares(weeks, week_arch_counts, top_archs)
    archetypes_with_other = top_archs + ["Other"]

    # Aggregate totals for summary
    grand_totals: "defaultdict[str, int]" = defaultdict(int)
    for week_data in week_arch_counts.values():
        for arch, count in week_data.items():
            grand_totals[arch] += count
    grand_total = sum(grand_totals.values())

    print(f"\nTop {len(top_archs)} archetypes by total appearances:\n")
    for i, arch in enumerate(top_archs, 1):
        pct = grand_totals[arch] / grand_total * 100 if grand_total else 0.0
        print(f"  {i:2}.  {arch:<45}  {pct:.1f}%  ({grand_totals[arch]} appearances)")

    other_total = grand_total - sum(grand_totals[a] for a in top_archs)
    other_pct = other_total / grand_total * 100 if grand_total else 0.0
    print(f"\n  Other (all remaining decks):  {other_pct:.1f}%  ({other_total} appearances)")
    if other_pct > 20:
        print(f"  ⚠  'Other' is large ({other_pct:.1f}%). Consider --num-decks {args.num_decks + 5}.")

    # ── Trend summary ──────────────────────────────────────────────────────────
    if len(weeks) >= 2:
        print(f"\n── Meta share trend ({weeks[0]} → {weeks[-1]}) ──\n")
        first_week, last_week = weeks[0], weeks[-1]
        for arch in top_archs:
            first_pct = shares[arch].get(first_week, 0.0)
            last_pct = shares[arch].get(last_week, 0.0)
            delta = last_pct - first_pct
            arrow = f"↑ +{delta:.1f}%" if delta >= 0 else f"↓ {delta:.1f}%"
            flag = "  ⚡" if abs(delta) >= 5 else ""
            print(f"  {arch:<45}  {first_pct:.1f}% → {last_pct:.1f}%  ({arrow}){flag}")
    else:
        print(f"\n  Only {len(weeks)} week bucket(s) found — trend comparison requires at least 2 weeks.")

    # ── Write spreadsheet ──────────────────────────────────────────────────────
    write_xlsx(
        archetypes_with_other, weeks, shares,
        args.output, str(start_date), str(end_date),
        len(tournaments), total_appearances,
    )

    print(
        f"\n  Chart: stacked area by week (Monday-start). "
        f"Data sheet: 'Meta Share Data'. Chart sheet: 'Meta Trend Chart'."
    )


if __name__ == "__main__":
    main()
