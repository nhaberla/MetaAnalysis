#!/usr/bin/env python3
"""
SWU Matchup Breakdown
Fetches real game-by-game results from melee.gg and builds a head-to-head
win rate matrix spreadsheet for the Star Wars Unlimited meta.

Usage:
    python matchup_breakdown.py --timeframe 2026-03-13:2026-05-31 --num-decks 12
    python matchup_breakdown.py --timeframe 2026-04-01:2026-05-31 --num-decks 8 --output april.xlsx
    python matchup_breakdown.py --timeframe 2026-03-13:2026-05-31 --dry-run --verbose
"""

import argparse
import sys
import time
from collections import defaultdict
from datetime import datetime, date
from pathlib import Path
from typing import Dict, FrozenSet, List, Optional, Tuple

import requests
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ── Constants ─────────────────────────────────────────────────────────────────

MELEE_API = "https://melee.gg/api/v1"
POST_ROTATION_DATE = "2026-03-13"
REQUEST_DELAY = 0.5  # seconds between API calls — respect the server
MIN_PLAYERS_DEFAULT = 32  # matches swu-competitivehub.com filter standard

# melee.gg game identifier for Star Wars Unlimited.
# If you get 0 results or a 404, open melee.gg in a browser, filter by SWU,
# and inspect the XHR requests in DevTools → Network to find the correct value.
# It may be a string slug or a numeric gameId depending on API version.
SWU_GAME_SLUG = "sw-unlimited"

LOW_SAMPLE_THRESHOLD = 20  # games below this get an asterisk

# ── Archetype map ─────────────────────────────────────────────────────────────
# Maps (leader_fragment, base_fragment) → canonical archetype label.
# Matching is case-insensitive substring. Order matters: first match wins.
# Add new archetypes here as the meta evolves.

ARCHETYPE_MAP: List[Tuple[Tuple[str, str], str]] = [
    # S Tier (LAW)
    (("Lando Calrissian", "Lake Country"),       "Lando / Lake Country"),
    (("Boba Fett", "Lake Country"),              "Boba Fett / Lake Country"),
    (("Aurra Sing", "Data Vault"),               "Aurra Sing / Data Vault"),
    (("Dedra Meero", "Colossus"),                "Dedra Meero / Colossus"),
    # A Tier
    (("Obi-Wan Kenobi", "Blue"),                 "Obi-Wan / Blue Force"),
    (("Mother Talzin", "Yellow"),                "Mother Talzin / Yellow Force"),
    (("Chewbacca", "Cunning"),                   "Chewbacca / Cunning"),
    (("Colonel Yularen", "Aggression"),          "Col. Yularen / Aggression"),
    (("Darth Vader", "Cunning"),                 "Darth Vader / Cunning"),
    (("Luke Skywalker", "Data Vault"),           "Luke Skywalker / Data Vault"),
    (("Admiral Piett", "Blue"),                  "Admiral Piett / Blue"),
    (("Tobias Beckett", "Red"),                  "Tobias Beckett / Red"),
    # B Tier
    (("Sabé", "Data Vault"),                     "Sabé / Data Vault"),
    (("Kazuda Xiono", "Data Vault"),             "Kazuda Xiono / Data Vault"),
    (("Lando Calrissian", "Data Vault"),         "Lando / Data Vault"),
    # C Tier
    (("The Client", "Red"),                      "The Client / Red"),
    (("Qui-Gon Jinn", "Green"),                  "Qui-Gon Jinn / Green Force"),
    (("Darth Maul", "Blue"),                     "Darth Maul / Blue Force"),
    (("Grand Admiral Thrawn", "Yellow"),         "Thrawn / Yellow Force"),
    (("Admiral Ackbar", "Data Vault"),           "Admiral Ackbar / Data Vault"),
]


def normalize_archetype(leader: str, base: str) -> str:
    for (leader_frag, base_frag), label in ARCHETYPE_MAP:
        if leader_frag.lower() in leader.lower() and base_frag.lower() in base.lower():
            return label
    return f"{leader} / {base}"


# ── melee.gg API client ───────────────────────────────────────────────────────

class MeleeClient:
    def __init__(self, verbose: bool = False):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "SWU-MetaAnalysis/1.0",
            "Accept": "application/json",
        })
        self.verbose = verbose
        self._cache: Dict[str, object] = {}

    def _get(self, url: str, params: Optional[dict] = None) -> object:
        key = f"{url}|{params}"
        if key in self._cache:
            return self._cache[key]
        time.sleep(REQUEST_DELAY)
        if self.verbose:
            print(f"  GET {url} params={params}", file=sys.stderr)
        resp = self.session.get(url, params=params, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        self._cache[key] = data
        return data

    def _unwrap(self, data: object) -> list:
        """Extract a list from common melee.gg response envelope shapes."""
        if isinstance(data, list):
            return data
        for key in ("data", "tournaments", "rounds", "pairings", "standings"):
            if isinstance(data, dict) and key in data:
                v = data[key]
                if isinstance(v, list):
                    return v
        return []

    # ── Endpoints ─────────────────────────────────────────────────────────────

    def get_swu_tournaments(self, start_date: str, end_date: str,
                             min_players: int) -> List[dict]:
        """
        Fetch SWU Premier tournaments within the date range.

        Known melee.gg endpoint (reverse-engineered by community tools):
            GET /api/v1/Tournaments
            params: game, startDate, endDate, pageSize, pageNumber, format

        If this returns 0 results, check SWU_GAME_SLUG. Possible values seen
        in the wild: "sw-unlimited", "swu", or a numeric gameId. Inspect the
        XHR calls on melee.gg to confirm the current value.
        """
        tournaments: List[dict] = []
        page = 1
        page_size = 50

        while True:
            try:
                raw = self._get(f"{MELEE_API}/Tournaments", params={
                    "game": SWU_GAME_SLUG,
                    "startDate": start_date,
                    "endDate": end_date,
                    "pageSize": page_size,
                    "pageNumber": page,
                    "format": "Premier",
                })
            except requests.HTTPError as exc:
                code = exc.response.status_code
                if code == 404:
                    raise RuntimeError(
                        f"melee.gg returned 404 for tournament list.\n"
                        f"SWU_GAME_SLUG='{SWU_GAME_SLUG}' may be wrong.\n"
                        f"Inspect melee.gg XHR traffic to find the correct game filter."
                    ) from exc
                raise

            items = self._unwrap(raw)
            if not items:
                break

            for t in items:
                count = t.get("playerCount") or t.get("registrations") or 0
                if int(count) >= min_players:
                    tournaments.append(t)

            if len(items) < page_size:
                break
            page += 1

        return tournaments

    def get_rounds(self, tournament_id: str) -> List[dict]:
        raw = self._get(f"{MELEE_API}/Tournaments/{tournament_id}/Rounds")
        return self._unwrap(raw)

    def get_pairings(self, round_id: str) -> List[dict]:
        raw = self._get(f"{MELEE_API}/Pairings", params={"roundId": round_id})
        return self._unwrap(raw)

    def get_decklist(self, decklist_id: str) -> Optional[dict]:
        try:
            return self._get(f"{MELEE_API}/Decklists/{decklist_id}")  # type: ignore[return-value]
        except requests.HTTPError:
            return None

    def extract_leader_base(self, decklist: dict) -> Tuple[str, str]:
        """
        Pull leader and base names out of a melee.gg decklist object.

        melee.gg may use any of these shapes (observed across different games):
          • Top-level "leader"/"base" dict fields
          • "cards" array with "type" == "Leader" / "Base"
          • "mainDeck" array with similar type annotations
        """
        leader, base = "", ""

        if isinstance(decklist.get("leader"), dict):
            leader = decklist["leader"].get("name", "")
        if isinstance(decklist.get("base"), dict):
            base = decklist["base"].get("name", "")

        if not leader or not base:
            for card in (decklist.get("cards") or decklist.get("mainDeck") or []):
                ctype = (card.get("type") or card.get("cardType") or "").lower()
                if ctype == "leader" and not leader:
                    leader = card.get("name", "")
                elif ctype == "base" and not base:
                    base = card.get("name", "")

        return leader or "Unknown", base or "Unknown"


# ── Tournament processing ─────────────────────────────────────────────────────

def is_top_cut_round(rnd: dict) -> bool:
    rtype = (rnd.get("type") or rnd.get("roundType") or "").lower()
    return rnd.get("isTopCut", False) or rtype in ("top8", "top4", "top2", "final", "semifinal", "quarterfinal")


def process_tournament(
    client: MeleeClient,
    tournament: dict,
    wins: "defaultdict[Tuple[str, str], int]",
    pair_games: "defaultdict[FrozenSet[str], int]",
    verbose: bool,
) -> int:
    """
    Walk all Swiss rounds of a tournament, recording wins and game totals
    per archetype pair. Returns the number of matchup records added.
    """
    tid = str(tournament.get("id") or tournament.get("tournamentId", ""))
    name = tournament.get("name", f"Tournament {tid}")

    rounds = client.get_rounds(tid)
    swiss_rounds = [r for r in rounds if not is_top_cut_round(r)]

    if verbose:
        print(f"    {len(swiss_rounds)} Swiss rounds", file=sys.stderr)

    # Per-tournament cache: player → (archetype, leader, base)
    player_arch: Dict[str, Optional[str]] = {}
    added = 0

    for rnd in swiss_rounds:
        rid = str(rnd.get("id") or rnd.get("roundId", ""))
        pairings = client.get_pairings(rid)

        for p in pairings:
            # Score fields vary by melee.gg API version
            p1_id = str(p.get("player1Id") or p.get("teamOneId") or "")
            p2_id = str(p.get("player2Id") or p.get("teamTwoId") or "")
            p1_score = int(p.get("player1Score") or p.get("teamOneScore") or 0)
            p2_score = int(p.get("player2Score") or p.get("teamTwoScore") or 0)

            if not p1_id or not p2_id or p1_score == p2_score:
                continue  # bye or draw

            winner_id = p1_id if p1_score > p2_score else p2_id
            loser_id = p2_id if winner_id == p1_id else p1_id

            # Resolve archetypes (cached per player within tournament)
            for pid, side in [(p1_id, "player1"), (p2_id, "player2")]:
                if pid in player_arch:
                    continue
                dl_id = str(
                    p.get(f"{side}DecklistId")
                    or p.get(f"team{'One' if side == 'player1' else 'Two'}DecklistId")
                    or ""
                )
                if not dl_id:
                    player_arch[pid] = None
                    continue
                dl = client.get_decklist(dl_id)
                if dl:
                    leader, base = client.extract_leader_base(dl)
                    player_arch[pid] = normalize_archetype(leader, base)
                else:
                    player_arch[pid] = None

            w_arch = player_arch.get(winner_id)
            l_arch = player_arch.get(loser_id)

            if w_arch and l_arch and w_arch != l_arch:
                wins[(w_arch, l_arch)] += 1
                pair_games[frozenset({w_arch, l_arch})] += 1
                added += 1

    return added


# ── Matrix logic ──────────────────────────────────────────────────────────────

def rank_archetypes(
    pair_games: "defaultdict[FrozenSet[str], int]",
    wins: "defaultdict[Tuple[str, str], int]",
    top_n: int,
) -> List[str]:
    """Rank archetypes by total games played (proxy for meta presence)."""
    totals: "defaultdict[str, int]" = defaultdict(int)
    for pair, g in pair_games.items():
        for arch in pair:
            totals[arch] += g
    return [arch for arch, _ in sorted(totals.items(), key=lambda x: -x[1])][:top_n]


def win_rate(arch_a: str, arch_b: str,
             wins: "defaultdict[Tuple[str, str], int]",
             pair_games: "defaultdict[FrozenSet[str], int]") -> Optional[float]:
    total = pair_games[frozenset({arch_a, arch_b})]
    if total == 0:
        return None
    return wins[(arch_a, arch_b)] / total * 100


def total_games(arch_a: str, arch_b: str,
                pair_games: "defaultdict[FrozenSet[str], int]") -> int:
    return pair_games[frozenset({arch_a, arch_b})]


# ── Color scale ───────────────────────────────────────────────────────────────

_FILLS = {
    "mirror":     PatternFill("solid", fgColor="D0D0D0"),
    "no_data":    PatternFill(fill_type=None),
    "low_sample": PatternFill("solid", fgColor="E0E0E0"),
    "wr_60p":     PatternFill("solid", fgColor="1A6B3A"),  # dark green
    "wr_55":      PatternFill("solid", fgColor="52BE80"),  # green
    "wr_50":      PatternFill("solid", fgColor="ABEBC6"),  # light green
    "wr_45":      PatternFill("solid", fgColor="FAD7A0"),  # light orange
    "wr_40":      PatternFill("solid", fgColor="E59866"),  # orange
    "wr_low":     PatternFill("solid", fgColor="B03A2E"),  # red
}

_FONT_DARK = Font(color="FFFFFF", size=9)
_FONT_NORMAL = Font(color="000000", size=9)


def wr_fill(wr: Optional[float], n_games: int) -> Tuple[PatternFill, Font]:
    if wr is None:
        return _FILLS["no_data"], _FONT_NORMAL
    if n_games < LOW_SAMPLE_THRESHOLD:
        return _FILLS["low_sample"], _FONT_NORMAL
    if wr >= 60:
        return _FILLS["wr_60p"], _FONT_DARK
    if wr >= 55:
        return _FILLS["wr_55"], _FONT_NORMAL
    if wr >= 50:
        return _FILLS["wr_50"], _FONT_NORMAL
    if wr >= 45:
        return _FILLS["wr_45"], _FONT_NORMAL
    if wr >= 40:
        return _FILLS["wr_40"], _FONT_NORMAL
    return _FILLS["wr_low"], _FONT_DARK


# ── Excel writer ──────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill("solid", fgColor="1C2833")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=9)
AVG_FILL = PatternFill("solid", fgColor="2E4057")
LEGEND = [
    ("≥60%  Heavily favoured",        "1A6B3A", "FFFFFF"),
    ("55–59% Favoured",               "52BE80", "000000"),
    ("50–54% Slight edge",            "ABEBC6", "000000"),
    ("45–49% Slight disadvantage",    "FAD7A0", "000000"),
    ("40–44% Unfavoured",             "E59866", "000000"),
    ("<40%   Heavily unfavoured",     "B03A2E", "FFFFFF"),
    ("*  < 20 games (low sample)",    "E0E0E0", "000000"),
    ("N/A  No data",                  "FFFFFF", "808080"),
]


def write_excel(
    archetypes: List[str],
    wins: "defaultdict[Tuple[str, str], int]",
    pair_games: "defaultdict[FrozenSet[str], int]",
    output_path: str,
    start_date: str,
    end_date: str,
    tournament_count: int,
    total_match_count: int,
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matchup Matrix"

    n = len(archetypes)
    avg_col = n + 2  # column index for the Avg WR column

    # ── Title ──────────────────────────────────────────────────────────────────
    title_range = f"A1:{get_column_letter(avg_col + 1)}1"
    ws.merge_cells(title_range)
    tc = ws["A1"]
    tc.value = (
        f"SWU Matchup Matrix  |  {start_date} – {end_date}  |  "
        f"{tournament_count} events  |  {total_match_count} matchups  |  "
        f"Top {n} archetypes by meta presence"
    )
    tc.font = Font(bold=True, size=11)
    tc.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 20

    # ── Column headers (row 2) ─────────────────────────────────────────────────
    corner = ws.cell(row=2, column=1, value="Deck ↓  vs  →")
    corner.fill = HEADER_FILL
    corner.font = HEADER_FONT
    corner.alignment = Alignment(horizontal="center", vertical="center")

    for ci, arch in enumerate(archetypes, start=2):
        c = ws.cell(row=2, column=ci)
        c.value = arch.split(" / ")[0]  # leader name only to keep headers short
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(textRotation=45, horizontal="center", vertical="bottom")

    avg_header = ws.cell(row=2, column=avg_col)
    avg_header.value = "Avg WR%"
    avg_header.fill = HEADER_FILL
    avg_header.font = HEADER_FONT
    avg_header.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 55

    # ── Data rows ──────────────────────────────────────────────────────────────
    for ri, row_arch in enumerate(archetypes, start=3):
        # Row label
        lbl = ws.cell(row=ri, column=1, value=row_arch)
        lbl.fill = HEADER_FILL
        lbl.font = HEADER_FONT
        lbl.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

        row_wrs: List[float] = []

        for ci, col_arch in enumerate(archetypes, start=2):
            cell = ws.cell(row=ri, column=ci)
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if row_arch == col_arch:
                cell.value = "—"
                cell.fill = _FILLS["mirror"]
                cell.font = Font(color="808080", size=9)
                continue

            wr = win_rate(row_arch, col_arch, wins, pair_games)
            n_games = total_games(row_arch, col_arch, pair_games)
            fill, font = wr_fill(wr, n_games)
            cell.fill = fill
            cell.font = font

            if wr is None:
                cell.value = "N/A"
                cell.font = Font(color="AAAAAA", size=9)
            elif n_games < LOW_SAMPLE_THRESHOLD:
                cell.value = f"{wr:.0f}%*"
            else:
                cell.value = f"{wr:.1f}%"
                row_wrs.append(wr)

        # Average WR for this archetype
        avg_cell = ws.cell(row=ri, column=avg_col)
        avg_cell.alignment = Alignment(horizontal="center", vertical="center")
        if row_wrs:
            avg = sum(row_wrs) / len(row_wrs)
            avg_fill, avg_font = wr_fill(avg, 999)
            avg_cell.fill = avg_fill
            avg_cell.font = Font(bold=True, color=avg_font.color, size=9)
            avg_cell.value = f"{avg:.1f}%"
        else:
            avg_cell.value = "—"
            avg_cell.fill = _FILLS["mirror"]

    # ── Legend ─────────────────────────────────────────────────────────────────
    legend_start = n + 4
    ws.cell(row=legend_start, column=1, value="Legend").font = Font(bold=True, size=9)
    for i, (label, bg, fg) in enumerate(LEGEND, start=legend_start + 1):
        c = ws.cell(row=i, column=1, value=label)
        c.fill = PatternFill("solid", fgColor=bg)
        c.font = Font(color=fg, size=9)
        c.alignment = Alignment(horizontal="left")

    # Note about Swiss-only
    note_row = legend_start + len(LEGEND) + 2
    note = ws.cell(row=note_row, column=1,
                   value="Note: Swiss rounds only. Top-cut excluded. Mirrors excluded. Post-rotation data (≥2026-03-13) only.")
    note.font = Font(italic=True, size=8, color="606060")
    ws.merge_cells(f"A{note_row}:{get_column_letter(avg_col)}{note_row}")

    # ── Column widths ──────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 32
    for ci in range(2, n + 2):
        ws.column_dimensions[get_column_letter(ci)].width = 7
    ws.column_dimensions[get_column_letter(avg_col)].width = 9

    ws.freeze_panes = "B3"

    wb.save(output_path)


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a SWU matchup win-rate matrix from raw melee.gg tournament data.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--timeframe", required=True,
                        metavar="YYYY-MM-DD:YYYY-MM-DD",
                        help="Date range for tournament data.")
    parser.add_argument("--num-decks", type=int, default=12,
                        help="Top N archetypes to include (default: 12).")
    parser.add_argument("--min-players", type=int, default=MIN_PLAYERS_DEFAULT,
                        help=f"Minimum event size (default: {MIN_PLAYERS_DEFAULT}).")
    parser.add_argument("--output", default="matchup_matrix.xlsx",
                        help="Output Excel file (default: matchup_matrix.xlsx).")
    parser.add_argument("--dry-run", action="store_true",
                        help="List qualifying tournaments; skip match data fetch.")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Print each API request.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    # Parse timeframe
    try:
        start_str, end_str = args.timeframe.split(":")
        start_date = datetime.strptime(start_str.strip(), "%Y-%m-%d").date()
        end_date = datetime.strptime(end_str.strip(), "%Y-%m-%d").date()
    except ValueError:
        sys.exit("Error: --timeframe must be YYYY-MM-DD:YYYY-MM-DD")

    if str(start_date) < POST_ROTATION_DATE:
        print(
            f"Warning: {start_date} is before the post-rotation date ({POST_ROTATION_DATE}).\n"
            f"Pre-rotation data uses different card legality. Consider adjusting --timeframe.",
            file=sys.stderr,
        )

    client = MeleeClient(verbose=args.verbose)

    # ── Fetch tournaments ──────────────────────────────────────────────────────
    print(f"Fetching SWU tournaments: {start_date} → {end_date}  (min {args.min_players} players)...")

    try:
        tournaments = client.get_swu_tournaments(str(start_date), str(end_date), args.min_players)
    except RuntimeError as exc:
        sys.exit(f"\nFailed to fetch tournaments:\n{exc}\n")
    except requests.RequestException as exc:
        sys.exit(f"\nNetwork error fetching tournaments: {exc}\n")

    if not tournaments:
        print(
            f"No tournaments found for SWU_GAME_SLUG='{SWU_GAME_SLUG}' in this date range.\n"
            f"Possible causes:\n"
            f"  • Wrong game slug — inspect melee.gg XHR traffic to verify\n"
            f"  • Date range too narrow — try widening it\n"
            f"  • melee.gg may require a different 'format' param for Premier events\n"
        )
        sys.exit(1)

    print(f"Found {len(tournaments)} qualifying tournament(s):\n")
    for t in sorted(tournaments, key=lambda x: x.get("date") or x.get("startDate") or ""):
        count = t.get("playerCount") or t.get("registrations") or "?"
        d = t.get("date") or t.get("startDate") or "unknown date"
        print(f"  • {t.get('name', 'Unnamed'):<55}  {count:>4} players  {d}")

    if args.dry_run:
        print("\nDry run complete. Re-run without --dry-run to process match data.")
        return

    # ── Process match results ──────────────────────────────────────────────────
    wins: "defaultdict[Tuple[str, str], int]" = defaultdict(int)
    pair_games: "defaultdict[FrozenSet[str], int]" = defaultdict(int)
    total_match_count = 0

    print(f"\nProcessing match results ({len(tournaments)} events)...\n")

    for t in tournaments:
        name = t.get("name", "Unnamed")
        try:
            added = process_tournament(client, t, wins, pair_games, args.verbose)
            total_match_count += added
            print(f"  ✓  {name:<55}  {added} matchups")
        except Exception as exc:  # noqa: BLE001
            print(f"  ✗  {name:<55}  ERROR: {exc}", file=sys.stderr)

    if total_match_count == 0:
        print(
            "\nNo match data collected. Possible causes:\n"
            "  • melee.gg pairings/decklists may not be publicly accessible\n"
            "  • Pairing API endpoint has changed — run with --verbose to inspect\n"
            "  • Decklists may require authentication\n"
            "\nAlternatives:\n"
            "  • Check swu-competitivehub.com for structured data\n"
            "  • Ask a tournament organizer for their melee.gg export CSV\n"
            "  • Use swumetastats.com/api-docs for pre-aggregated matchup data\n"
            "    (note: that source is aggregated, not raw game-by-game)\n"
        )
        sys.exit(1)

    print(f"\nTotal matchup records: {total_match_count}")

    # ── Rank and select top archetypes ─────────────────────────────────────────
    archetypes = rank_archetypes(pair_games, wins, args.num_decks)

    print(f"\nTop {len(archetypes)} archetypes by game count:\n")
    for i, arch in enumerate(archetypes, 1):
        g = sum(pair_games[frozenset({arch, other})] for other in archetypes if other != arch)
        print(f"  {i:2}.  {arch:<40}  {g} matchup games")

    # Flag any archetype with very low total sample
    low_sample_archs = []
    for arch in archetypes:
        g = sum(pair_games[frozenset({arch, other})] for other in archetypes if other != arch)
        if g < 50:
            low_sample_archs.append((arch, g))
    if low_sample_archs:
        print("\nLow-sample archetypes (< 50 total games, treat with caution):")
        for arch, g in low_sample_archs:
            print(f"  ⚠  {arch}  ({g} games)")

    # ── Write Excel ────────────────────────────────────────────────────────────
    output_path = args.output
    write_excel(
        archetypes, wins, pair_games,
        output_path, str(start_date), str(end_date),
        len(tournaments), total_match_count,
    )
    print(f"\nSpreadsheet saved: {output_path}")

    # ── Summary table ──────────────────────────────────────────────────────────
    print("\n── Average win rates ────────────────────────────────────────────\n")
    summary = []
    for arch in archetypes:
        wrs = [
            win_rate(arch, opp, wins, pair_games)
            for opp in archetypes
            if opp != arch and win_rate(arch, opp, wins, pair_games) is not None
            and total_games(arch, opp, pair_games) >= LOW_SAMPLE_THRESHOLD
        ]
        avg = sum(wrs) / len(wrs) if wrs else None
        summary.append((arch, avg, len(wrs)))
    summary.sort(key=lambda x: -(x[1] or 0))
    for arch, avg, n_matchups in summary:
        avg_str = f"{avg:.1f}%" if avg is not None else "N/A"
        print(f"  {arch:<40}  avg {avg_str}  ({n_matchups} matchups with ≥{LOW_SAMPLE_THRESHOLD} games)")

    print(f"\n* = fewer than {LOW_SAMPLE_THRESHOLD} games (low confidence)\n")


if __name__ == "__main__":
    main()
