#!/usr/bin/env python3
"""
FantaNBA lottery generator for 12 teams.

Input team order: team 1 is Seed 1 (last place / worst record),
team 2 is Seed 2, ..., team 12 is Seed 12 (best record).

The script reads the odds matrix from Odds.xlsx when available. The matrix
is treated as a doubly-stochastic matrix of marginal probabilities:
rows = seeds, columns = picks. A Birkhoff-von Neumann decomposition is used
so every draw is a valid permutation and the long-run marginal probabilities
match the matrix.

Example:
    python fantanba_lottery.py --teams A B C D E F G H I J K L --output lottery.html
"""

from __future__ import annotations

import argparse
import json
import math
import random
import re
import secrets
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET


DEFAULT_TEAMS = list("ABCDEFGHIJKL")

# Extracted from the uploaded Odds.xlsx file, sheet "Odds Matrix", range B5:M16.
# Rows are Seed 1..12. Columns are Pick 1..12.
DEFAULT_ODDS: List[List[float]] = [
    [0.14, 0.3551853664681282, 0.24380836766850866, 0.1607036925727639, 0.10030257329059931, 0, 0, 0, 0, 0, 0, 0],
    [0.14, 0.5607684347662873, 0.1280081352422556, 0.08391664006491963, 0.053205529008597985, 0.0341012609179395, 0, 0, 0, 0, 0, 0],
    [0.14, 0.05603976578419088, 0.5366790459475915, 0.1178049556311759, 0.07423784690883185, 0.048111124738829024, 0.027127260989380814, 0, 0, 0, 0, 0],
    [0.12, 0.01803016679653696, 0.058021279495446124, 0.5356304981166196, 0.11246842813168241, 0.07223633838935592, 0.0412330056059481, 0.02217073804160677, 0.011117110769874462, 0.00505082535399096, 0.0030317342312659603, 0.0010098750676727738],
    [0.1, 0.005992159478347605, 0.02094515007988616, 0.06276910419466933, 0.5526581185797038, 0.12003540084943723, 0.06818281424589019, 0.03617136051760376, 0.018137466638080437, 0.009064413953770348, 0.004030275529830518, 0.002013735932780601],
    [0.08, 0.0019970583536947388, 0.0069805697073939054, 0.02291194001725045, 0.06506680464141398, 0.5940776205880944, 0.11228265651685958, 0.060275696889583064, 0.030224144897447392, 0.015104874689279837, 0.007051823483507223, 0.004026810215475505],
    [0.07, 0.0009955825474607433, 0.0029828443960331747, 0.008939084703474669, 0.0249518440268299, 0.07678280756556512, 0.6057365728551547, 0.10817608523747259, 0.054242917848719784, 0.027108541190027445, 0.013057596993604061, 0.007026122635657808],
    [0.06, 0.00019872259883806258, 0.0009923144859946744, 0.003965066924973964, 0.009960992794541036, 0.030851404813342212, 0.08279979316175125, 0.6177830209031209, 0.10426117385753118, 0.051103723747481455, 0.0250610870196944, 0.013022699692730764],
    [0.05, 0.00019790507597188458, 0.0009882322135833653, 0.0019743775383594666, 0.003968005747278816, 0.013875574112458329, 0.035765420504013684, 0.08860274456183947, 0.6249903034190579, 0.10278488886480071, 0.050914296479229314, 0.025938251483406893],
    [0.04, 0.00019826986433148255, 0.00019801075435074754, 0.0009890084040221154, 0.001987659885215076, 0.00595763582245875, 0.015925042259702447, 0.03989485909837686, 0.09302114283001378, 0.6428398556812305, 0.1060169271858198, 0.052971588214478456],
    [0.035, 0.0001981632424327636, 0.00019790427179139865, 0.0001976953106767771, 0.000993295499131513, 0.0029772160176551555, 0.006963459293509158, 0.017943032330681726, 0.04298664671580028, 0.09892211812559887, 0.6797428532818742, 0.11387761591084812],
    [0.025, 0.0001984050237793372, 0.00019814573716478992, 0.00019793652109420846, 0.0001989014861744232, 0.000993616184864218, 0.003983974567790018, 0.008982462419714945, 0.02101909302347478, 0.048020758393819814, 0.1110934057951745, 0.7801133008469491],
]


class LotteryError(Exception):
    """Raised for invalid odds or lottery input."""


def xml_text(node: ET.Element) -> str:
    """Return all text contained in an XML node."""
    return "".join(node.itertext()) if node is not None else ""


def local_name(tag: str) -> str:
    """Strip an XML namespace from a tag."""
    return tag.rsplit("}", 1)[-1]


def column_letters_to_index(letters: str) -> int:
    """Convert Excel letters to a 1-based column index."""
    value = 0
    for ch in letters.upper():
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value


def parse_cell_ref(ref: str) -> Tuple[int, int]:
    """Return (row, column) as 1-based indexes from an Excel cell ref."""
    match = re.fullmatch(r"([A-Za-z]+)([0-9]+)", ref)
    if not match:
        raise ValueError(f"Invalid Excel cell reference: {ref}")
    col = column_letters_to_index(match.group(1))
    row = int(match.group(2))
    return row, col


def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    """Read xl/sharedStrings.xml from an xlsx archive."""
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    values: List[str] = []
    for item in root:
        if local_name(item.tag) == "si":
            values.append(xml_text(item))
    return values


def workbook_sheet_targets(zf: zipfile.ZipFile) -> Dict[str, str]:
    """Map workbook sheet names to zip paths such as xl/worksheets/sheet1.xml."""
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_map: Dict[str, str] = {}
    for rel in rels:
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if not rel_id or not target:
            continue
        if target.startswith("/"):
            path = target.lstrip("/")
        elif target.startswith("xl/"):
            path = target
        else:
            path = "xl/" + target
        rel_map[rel_id] = path

    result: Dict[str, str] = {}
    relationship_key = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    for node in workbook.iter():
        if local_name(node.tag) != "sheet":
            continue
        name = node.attrib.get("name")
        rel_id = node.attrib.get(relationship_key)
        if name and rel_id and rel_id in rel_map:
            result[name] = rel_map[rel_id]
    return result


def read_worksheet_cells(zf: zipfile.ZipFile, sheet_path: str, shared_strings: Sequence[str]) -> Dict[Tuple[int, int], object]:
    """Read non-empty cells from a worksheet XML file."""
    root = ET.fromstring(zf.read(sheet_path))
    cells: Dict[Tuple[int, int], object] = {}

    for cell in root.iter():
        if local_name(cell.tag) != "c":
            continue
        ref = cell.attrib.get("r")
        if not ref:
            continue
        row, col = parse_cell_ref(ref)
        cell_type = cell.attrib.get("t")

        value_node = None
        inline_node = None
        for child in cell:
            name = local_name(child.tag)
            if name == "v":
                value_node = child
            elif name == "is":
                inline_node = child

        value: object = None
        if cell_type == "s" and value_node is not None and value_node.text is not None:
            idx = int(value_node.text)
            value = shared_strings[idx]
        elif cell_type == "inlineStr":
            value = xml_text(inline_node)
        elif value_node is not None and value_node.text is not None:
            raw = value_node.text
            if cell_type == "str":
                value = raw
            elif cell_type == "b":
                value = raw == "1"
            else:
                try:
                    number = float(raw)
                    value = int(number) if number.is_integer() else number
                except ValueError:
                    value = raw

        if value is not None:
            cells[(row, col)] = value
    return cells


def load_odds_from_xlsx(path: Path, sheet_name: str = "Odds Matrix") -> List[List[float]]:
    """Load a Seed x Pick odds matrix from an xlsx file."""
    with zipfile.ZipFile(path) as zf:
        shared_strings = load_shared_strings(zf)
        targets = workbook_sheet_targets(zf)
        if sheet_name not in targets:
            available = ", ".join(sorted(targets)) or "none"
            raise LotteryError(f"Sheet '{sheet_name}' not found in {path}. Available sheets: {available}")
        cells = read_worksheet_cells(zf, targets[sheet_name], shared_strings)

    seed_cells = [pos for pos, value in cells.items() if str(value).strip().lower() == "seed"]
    if not seed_cells:
        raise LotteryError("Could not find the 'Seed' header in the odds sheet.")
    header_row, seed_col = min(seed_cells)

    pick_cols: List[int] = []
    for (row, col), value in sorted(cells.items()):
        if row != header_row or col <= seed_col:
            continue
        text = str(value).strip().lower()
        if re.fullmatch(r"pick\s+\d+", text):
            pick_cols.append(col)
    if not pick_cols:
        raise LotteryError("Could not find Pick columns in the odds sheet.")

    n = len(pick_cols)
    rows_by_seed: Dict[int, int] = {}
    for (row, col), value in cells.items():
        if col != seed_col or row <= header_row:
            continue
        if isinstance(value, (int, float)) and float(value).is_integer():
            seed = int(value)
            if 1 <= seed <= n:
                rows_by_seed[seed] = row

    missing = [seed for seed in range(1, n + 1) if seed not in rows_by_seed]
    if missing:
        raise LotteryError(f"Missing seed rows in the odds sheet: {missing}")

    odds: List[List[float]] = []
    for seed in range(1, n + 1):
        row = rows_by_seed[seed]
        odds_row: List[float] = []
        for col in pick_cols:
            value = cells.get((row, col), 0)
            try:
                odds_row.append(float(value))
            except (TypeError, ValueError):
                raise LotteryError(f"Non-numeric odds value at seed {seed}, column {col}: {value!r}")
        odds.append(odds_row)
    return odds


def validate_odds_matrix(matrix: Sequence[Sequence[float]], tol: float = 1e-7) -> None:
    """Validate that matrix is square and approximately doubly-stochastic."""
    n = len(matrix)
    if n == 0:
        raise LotteryError("Odds matrix is empty.")
    if any(len(row) != n for row in matrix):
        raise LotteryError("Odds matrix must be square: same number of seeds and picks.")
    for i, row in enumerate(matrix, start=1):
        if any(value < -tol for value in row):
            raise LotteryError(f"Odds matrix contains a negative value on row {i}.")
        total = sum(row)
        if abs(total - 1.0) > tol:
            raise LotteryError(f"Seed row {i} sums to {total:.12f}, not 1.0.")
    for col in range(n):
        total = sum(matrix[row][col] for row in range(n))
        if abs(total - 1.0) > tol:
            raise LotteryError(f"Pick column {col + 1} sums to {total:.12f}, not 1.0.")


def find_perfect_matching(support_matrix: Sequence[Sequence[float]], eps: float) -> List[int]:
    """Find one perfect matching in the positive support graph.

    Returns a list perm where perm[row] = matched column.
    """
    n = len(support_matrix)
    col_to_row = [-1] * n

    def dfs(row: int, seen_cols: List[bool]) -> bool:
        for col, value in enumerate(support_matrix[row]):
            if value <= eps or seen_cols[col]:
                continue
            seen_cols[col] = True
            if col_to_row[col] == -1 or dfs(col_to_row[col], seen_cols):
                col_to_row[col] = row
                return True
        return False

    for row in range(n):
        seen = [False] * n
        if not dfs(row, seen):
            raise LotteryError("Could not find a perfect matching in the residual odds matrix.")

    perm = [-1] * n
    for col, row in enumerate(col_to_row):
        if row == -1:
            raise LotteryError("Internal matching error: unmatched column.")
        perm[row] = col
    return perm


def birkhoff_decomposition(matrix: Sequence[Sequence[float]], eps: float = 1e-12) -> List[Tuple[float, Tuple[int, ...]]]:
    """Decompose a doubly-stochastic matrix into weighted permutation matrices."""
    validate_odds_matrix(matrix)
    n = len(matrix)
    residual = [[max(0.0, float(value)) for value in row] for row in matrix]
    decomposition: List[Tuple[float, Tuple[int, ...]]] = []

    # A safe cap for this tiny 12x12 case. A normal decomposition is much shorter.
    max_steps = n * n * 4
    for _ in range(max_steps):
        remaining = sum(sum(row) for row in residual)
        if remaining <= eps * n:
            break

        perm = find_perfect_matching(residual, eps)
        weight = min(residual[row][perm[row]] for row in range(n))
        if weight <= eps:
            break

        decomposition.append((weight, tuple(perm)))
        for row in range(n):
            col = perm[row]
            residual[row][col] -= weight
            if abs(residual[row][col]) <= eps:
                residual[row][col] = 0.0
            elif residual[row][col] < 0:
                residual[row][col] = 0.0
    else:
        raise LotteryError("Birkhoff decomposition did not converge.")

    total_weight = sum(weight for weight, _ in decomposition)
    if not decomposition or abs(total_weight - 1.0) > 1e-7:
        raise LotteryError(f"Invalid decomposition total weight: {total_weight:.12f}")

    # Normalize tiny floating-point drift.
    return [(weight / total_weight, perm) for weight, perm in decomposition]


def choose_weighted_permutation(
    decomposition: Sequence[Tuple[float, Tuple[int, ...]]],
    rng: random.Random,
) -> Tuple[int, ...]:
    """Randomly choose one permutation according to decomposition weights."""
    threshold = rng.random()
    cumulative = 0.0
    for weight, perm in decomposition:
        cumulative += weight
        if threshold <= cumulative:
            return perm
    return decomposition[-1][1]


def draw_lottery(teams: Sequence[str], odds: Sequence[Sequence[float]], rng: random.Random) -> List[str]:
    """Return teams ordered by pick 1..N."""
    n = len(odds)
    if len(teams) != n:
        raise LotteryError(f"Expected {n} teams, received {len(teams)}.")

    decomposition = birkhoff_decomposition(odds)
    seed_to_pick = choose_weighted_permutation(decomposition, rng)

    pick_order: List[Optional[str]] = [None] * n
    for seed_index, pick_index in enumerate(seed_to_pick):
        pick_order[pick_index] = teams[seed_index]

    if any(team is None for team in pick_order):
        raise LotteryError("Internal error: incomplete pick order.")
    return [str(team) for team in pick_order]


def build_html(pick_order: Sequence[str], teams: Sequence[str], odds_source: str) -> str:
    """Build a small self-contained HTML page."""
    n = len(pick_order)
    final_order = [{"pick": index + 1, "team": team} for index, team in enumerate(pick_order)]
    reveal_order = list(reversed(final_order))
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload_reveal = json.dumps(reveal_order, ensure_ascii=False)
    payload_final = json.dumps(final_order, ensure_ascii=False)
    payload_teams = json.dumps(list(teams), ensure_ascii=False)

    return f"""<!doctype html>
<html lang="it">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>FantaNBA Lottery</title>
  <style>
    body {{
      margin: 0;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #f6f7fb;
      color: #151821;
    }}
    main {{
      max-width: 760px;
      margin: 36px auto;
      padding: 0 18px;
    }}
    .panel {{
      background: white;
      border: 1px solid #e4e7ef;
      border-radius: 18px;
      padding: 24px;
      box-shadow: 0 10px 28px rgba(20, 25, 40, 0.08);
    }}
    h1 {{ margin: 0 0 8px; font-size: 32px; }}
    p {{ line-height: 1.5; color: #51576a; }}
    button {{
      width: 100%;
      margin: 18px 0 10px;
      padding: 16px 18px;
      border: 0;
      border-radius: 14px;
      cursor: pointer;
      font-size: 18px;
      font-weight: 800;
      background: #151821;
      color: white;
    }}
    button:disabled {{ opacity: 0.55; cursor: not-allowed; }}
    .revealed {{ display: grid; gap: 10px; margin-top: 14px; }}
    .card {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 14px;
      padding: 14px 16px;
      border: 1px solid #e4e7ef;
      border-radius: 14px;
      background: #fbfcff;
      animation: pop 170ms ease-out;
    }}
    .pick {{ color: #626b80; font-weight: 700; }}
    .team {{ font-size: 20px; font-weight: 900; text-align: right; }}
    #finale {{ margin-top: 24px; }}
    ol {{ padding-left: 24px; }}
    li {{ padding: 5px 0; font-size: 18px; }}
    .meta {{ font-size: 13px; color: #7a8194; margin-top: 20px; }}
    @keyframes pop {{ from {{ transform: scale(0.98); opacity: 0; }} to {{ transform: scale(1); opacity: 1; }} }}
  </style>
</head>
<body>
<main>
  <section class="panel">
    <h1>FantaNBA Lottery</h1>
    <p>Premi il pulsante per rivelare le squadre una per una, partendo dalla pick #{n}. Alla fine comparira la lista ordinata dalla pick #1 alla pick #{n}.</p>

    <button id="revealButton">Rivela pick #{n}</button>
    <div id="revealed" class="revealed" aria-live="polite"></div>

    <section id="finale" hidden>
      <h2>Ordine finale</h2>
      <ol id="finalList"></ol>
    </section>

    <p class="meta">Generato: {generated_at}. Odds: {odds_source}. Seed input: <span id="seedInput"></span></p>
  </section>
</main>
<script>
const revealOrder = {payload_reveal};
const finalOrder = {payload_final};
const seedInput = {payload_teams};
let revealIndex = 0;

const button = document.getElementById('revealButton');
const revealed = document.getElementById('revealed');
const finale = document.getElementById('finale');
const finalList = document.getElementById('finalList');
document.getElementById('seedInput').textContent = seedInput.join(', ');

function revealNext() {{
  if (revealIndex >= revealOrder.length) return;

  const item = revealOrder[revealIndex];
  const card = document.createElement('div');
  card.className = 'card';

  const pick = document.createElement('span');
  pick.className = 'pick';
  pick.textContent = 'Pick #' + item.pick;

  const team = document.createElement('span');
  team.className = 'team';
  team.textContent = item.team;

  card.appendChild(pick);
  card.appendChild(team);
  revealed.appendChild(card);

  revealIndex += 1;

  if (revealIndex < revealOrder.length) {{
    button.textContent = 'Rivela pick #' + revealOrder[revealIndex].pick;
  }} else {{
    button.textContent = 'Lottery completa';
    button.disabled = true;
    showFinalOrder();
  }}
}}

function showFinalOrder() {{
  finalList.replaceChildren();
  for (const item of finalOrder) {{
    const li = document.createElement('li');
    li.textContent = item.team;
    finalList.appendChild(li);
  }}
  finale.hidden = false;
}}

button.addEventListener('click', revealNext);
</script>
</body>
</html>
"""


def load_teams(args: argparse.Namespace) -> List[str]:
    """Load teams from --teams, --teams-file, or use A-L default."""
    if args.teams_file:
        path = Path(args.teams_file)
        if not path.exists():
            raise LotteryError(f"Teams file not found: {path}")
        teams = [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    elif args.teams:
        teams = args.teams
    else:
        teams = DEFAULT_TEAMS
    return [str(team).strip() for team in teams]


def load_odds(args: argparse.Namespace) -> Tuple[List[List[float]], str]:
    """Load odds from xlsx or fallback to embedded matrix."""
    if args.use_default_odds:
        return [row[:] for row in DEFAULT_ODDS], "matrice incorporata"

    odds_path = Path(args.odds_file)
    if odds_path.exists():
        return load_odds_from_xlsx(odds_path, args.sheet), str(odds_path)

    # Make the script runnable immediately in demo mode even without the xlsx.
    if args.odds_file == "Odds.xlsx":
        return [row[:] for row in DEFAULT_ODDS], "matrice incorporata (Odds.xlsx non trovato)"

    raise LotteryError(f"Odds file not found: {odds_path}")


def make_rng(seed: Optional[int]) -> random.Random:
    """Use deterministic RNG when seed is given; otherwise SystemRandom."""
    if seed is not None:
        return random.Random(seed)
    return secrets.SystemRandom()


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create a FantaNBA lottery HTML page.")
    parser.add_argument("--odds-file", default="Odds.xlsx", help="Path to the xlsx odds file. Default: Odds.xlsx")
    parser.add_argument("--sheet", default="Odds Matrix", help="Sheet containing the odds matrix. Default: Odds Matrix")
    parser.add_argument("--teams", nargs="*", help="Team names in seed order: first is last place, second is penultimate, etc.")
    parser.add_argument("--teams-file", help="Optional text file with one team per line, in seed order.")
    parser.add_argument("--output", "-o", default="lottery.html", help="Output HTML path. Default: C:/Users/994944/Desktop/Progetti/lottery.html")
    parser.add_argument("--seed", type=int, help="Optional random seed for reproducible tests.")
    parser.add_argument("--use-default-odds", action="store_true", help="Use the embedded odds matrix instead of reading xlsx.")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    try:
        args = parse_args(argv)
        teams = load_teams(args)
        odds, odds_source = load_odds(args)
        validate_odds_matrix(odds)

        if len(teams) != len(odds):
            raise LotteryError(f"The odds matrix has {len(odds)} seeds, but {len(teams)} teams were provided.")

        rng = make_rng(args.seed)
        pick_order = draw_lottery(teams, odds, rng)
        html = build_html(pick_order, teams, odds_source)

        output_path = Path(args.output)
        output_path.write_text(html, encoding="utf-8")

        print(f"HTML created: {output_path}")
        print("Final order:")
        for idx, team in enumerate(pick_order, start=1):
            print(f"  Pick {idx}: {team}")
        return 0
    except LotteryError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
