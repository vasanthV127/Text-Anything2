"""
Process `leaderboard.xlsx` into normalized CSV/JSON outputs.

This script uses a lightweight XLSX parser (zip + XML) so it has no
hard dependency on pandas/openpyxl. It will normalize "-" and empty
cells to None and convert numeric strings to numbers where appropriate.

Usage:
    python -m Test2.process_leaderboard --input "leaderboard.xlsx"

Outputs:
    Test2/output/leaderboard_normalized.csv
    Test2/output/leaderboard_normalized.json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, Any
from functools import cmp_to_key
from collections import Counter


def parse_xlsx_to_rows(path: str) -> List[Dict[str, Optional[str]]]:
    """Return rows as list of dicts keyed by column letter (A, B, C...).

    The first row corresponds to index 0 and usually holds headers.
    """
    with zipfile.ZipFile(path) as z:
        namelist = z.namelist()
        if 'xl/sharedStrings.xml' in namelist:
            ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            shared = [
                ''.join([t.text or '' for t in si.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')])
                for si in ss_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')
            ]
        else:
            shared = []

        sheet_name = 'xl/worksheets/sheet1.xml'
        if sheet_name not in namelist:
            raise FileNotFoundError(f"{sheet_name} not found in {path}")

        sheet_root = ET.fromstring(z.read(sheet_name))
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        rows: List[Dict[str, Optional[str]]] = []
        for r in sheet_root.findall('.//ns:row', ns):
            row_cells: Dict[str, Optional[str]] = {}
            for c in r.findall('ns:c', ns):
                ref = c.get('r')  # like A1, B2
                m = re.match(r'([A-Z]+)', ref)
                col = m.group(1) if m else ref
                t = c.get('t')
                v = c.find('ns:v', ns)
                val: Optional[str] = None
                if v is not None and v.text is not None:
                    text = v.text
                    if t == 's':
                        idx = int(text)
                        val = shared[idx] if 0 <= idx < len(shared) else text
                    else:
                        val = text
                else:
                    is_elem = c.find('ns:is', ns)
                    if is_elem is not None:
                        t_e = is_elem.find('.//ns:t', ns)
                        if t_e is not None:
                            val = t_e.text
                row_cells[col] = val
            rows.append(row_cells)

        return rows


def col_letter_range(start: str, end: str) -> List[str]:
    """Return list of column letters from start to end (inclusive)."""
    def to_index(col: str) -> int:
        res = 0
        for ch in col:
            res = res * 26 + (ord(ch) - 64)
        return res

    def to_letter(idx: int) -> str:
        out = ''
        while idx > 0:
            idx, rem = divmod(idx - 1, 26)
            out = chr(65 + rem) + out
        return out

    return [to_letter(i) for i in range(to_index(start), to_index(end) + 1)]


def normalize_rows(rows: List[Dict[str, Optional[str]]]) -> List[Dict[str, Any]]:
    """Convert parsed rows into list of dicts keyed by header names.

    Heuristics used because spreadsheet has irregular headers:
    - The header row contains 'Pos' and 'Player' in first two columns.
    - Round columns are labeled like R01, R02, ...
    - Trailing columns include Total / Points / Spent ($m) / $m/Pt
    """
    if not rows:
        return []

    header_row = rows[0]
    # Determine column letters in order
    cols = sorted(header_row.keys(), key=lambda c: sum((ord(ch) - 64) * (26 ** i) for i, ch in enumerate(reversed(c))))
    headers = [header_row.get(c) or f'COL_{c}' for c in cols]

    # Build list of dicts mapping header->value
    data = []
    for r in rows[1:]:
        rec: Dict[str, Any] = {}
        for i, c in enumerate(cols):
            h = headers[i]
            v = r.get(c)
            if v is None:
                rec[h] = None
                continue
            v_str = str(v).strip()
            if v_str == '-' or v_str == '':
                rec[h] = None
                continue
            # try int
            if re.fullmatch(r"-?\d+", v_str):
                rec[h] = int(v_str)
                continue
            # try float
            if re.fullmatch(r"-?\d+\.\d+", v_str):
                try:
                    rec[h] = float(v_str)
                    continue
                except Exception:
                    pass
            rec[h] = v_str
        data.append(rec)

    return data


def detect_round_columns(headers: List[str]) -> List[str]:
    rounds = [h for h in headers if isinstance(h, str) and re.match(r'^R\d{2}$', h)]
    if rounds:
        return rounds
    # fallback: detect repeated 'Pts' columns by position
    return [h for h in headers if h.startswith('R')]


def run(input_path: str, out_dir: str) -> Dict[str, Any]:
    os.makedirs(out_dir, exist_ok=True)
    rows = parse_xlsx_to_rows(input_path)
    norm = normalize_rows(rows)
    if not norm:
        raise RuntimeError('No data found in spreadsheet')

    headers = list(norm[0].keys())
    # detect round columns
    round_cols = [h for h in headers if re.match(r'^R\d{2}$', h)]
    # fallback: common columns "Pts" occurrences
    if not round_cols:
        # try to find contiguous numeric columns after 'Player'
        if 'Player' in headers:
            idx = headers.index('Player')
            possible = headers[idx + 1: idx + 1 + 30]
            round_cols = [h for h in possible if any(isinstance(r[h], int) for r in norm if h in r and r[h] is not None)]

    # compute totals from round columns
    for r in norm:
        if round_cols:
            pts = 0
            counted = 0
            for rc in round_cols:
                v = r.get(rc)
                # Treat missing or dash/'D$Q' as zero for round scoring per spec
                if v is None:
                    val = 0
                elif isinstance(v, (int, float)):
                    val = v
                else:
                    try:
                        val = float(str(v))
                    except Exception:
                        # non-numeric like 'D$Q' or '-' treat as zero
                        val = 0
                pts += val
                # count only actual numeric entries (non-empty strings that parse)
                counted += 1
            r['computed_rounds_total'] = pts
            r['computed_rounds_count'] = counted
        else:
            r['computed_rounds_total'] = None
            r['computed_rounds_count'] = 0

    # Determine canonical total column (prefer explicit total/pts if present)
    total_col = None
    for cand in ('Total', 'Pts', 'Points', 'TOTAL'):
        if cand in headers:
            total_col = cand
            break

    # Detect spending column heuristically
    spend_col = None
    for h in headers:
        if isinstance(h, str) and re.search(r'spend|spent|\$m|\$|spent', h, re.I):
            spend_col = h
            break

    # Build helper values for ranking
    for r in norm:
        # primary total: prefer explicit total if numeric, else computed_rounds_total
        tot = None
        if total_col and isinstance(r.get(total_col), (int, float)):
            tot = r.get(total_col)
        else:
            tot = r.get('computed_rounds_total')
        r['_rank_total'] = float(tot) if tot is not None else 0.0

        # spending: parse to float if present, else treat as 0
        spend_val = None
        if spend_col and r.get(spend_col) is not None:
            try:
                spend_val = float(str(r.get(spend_col)).replace(',', '').replace('$', ''))
            except Exception:
                spend_val = 0.0
        else:
            spend_val = 0.0
        r['_rank_spend'] = float(spend_val)

        # countback list: sorted round scores descending (including repeats)
        scores = []
        for rc in round_cols:
            v = r.get(rc)
            if v is None:
                sval = 0.0
            elif isinstance(v, (int, float)):
                sval = float(v)
            else:
                try:
                    sval = float(str(v))
                except Exception:
                    sval = 0.0
            scores.append(sval)
        # sort descending
        scores_sorted = sorted(scores, reverse=True)
        r['_countback'] = scores_sorted

    # write outputs
    csv_path = os.path.join(out_dir, 'leaderboard_normalized.csv')
    json_path = os.path.join(out_dir, 'leaderboard_normalized.json')

    # Apply ranking and tiebreakers
    try:
        norm = rank_rows(norm)
    except Exception:
        # if ranking fails, still attempt to write raw normalized data
        pass

    # write CSV header from keys
    keys = list(norm[0].keys()) + ['computed_rounds_total', 'computed_rounds_count']
    with open(csv_path, 'w', encoding='utf-8') as fh:
        fh.write(','.join(keys) + '\n')
        for r in norm:
            rowvals = [str(r.get(k)) if r.get(k) is not None else '' for k in keys]
            fh.write(','.join(rowvals) + '\n')

    with open(json_path, 'w', encoding='utf-8') as fh:
        json.dump({'rows': norm, 'round_columns': round_cols}, fh, indent=2)

    return {'csv': csv_path, 'json': json_path, 'rows': len(norm), 'round_columns': round_cols}


def _compare_rows(a: Dict[str, Any], b: Dict[str, Any]) -> int:
    # total descending
    if a.get('_rank_total', 0.0) != b.get('_rank_total', 0.0):
        return -1 if a['_rank_total'] > b['_rank_total'] else 1
    # spending ascending (lower spend wins)
    if a.get('_rank_spend', 0.0) != b.get('_rank_spend', 0.0):
        return -1 if a['_rank_spend'] < b['_rank_spend'] else 1
    # countback: compare sorted score lists lexicographically
    al = a.get('_countback', [])
    bl = b.get('_countback', [])
    L = max(len(al), len(bl))
    for i in range(L):
        av = al[i] if i < len(al) else 0.0
        bv = bl[i] if i < len(bl) else 0.0
        if av != bv:
            return -1 if av > bv else 1
    # numeric tiebreakers equal
    return 0


def rank_rows(norm: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Sort and annotate rows with `rank`, `tie_group`, and `tie_highlight`.

    Expects each row to already have `_rank_total`, `_rank_spend`, and `_countback` fields.
    Returns a new list (sorted) with rank metadata added in-place.
    """
    # Sort using comparator but keep deterministic alphabetical order for stable output
    norm_sorted = sorted(norm, key=cmp_to_key(lambda a, b: _compare_rows(a, b) or (-1 if (a.get('Player') or '') < (b.get('Player') or '') else 1)))

    def key_for(r: Dict[str, Any]) -> tuple:
        return (round(float(r.get('_rank_total', 0.0)), 6), round(float(r.get('_rank_spend', 0.0)), 6), tuple([round(x, 6) for x in r.get('_countback', [])]))

    counts: Dict[tuple, int] = Counter(key_for(r) for r in norm_sorted)
    tie_group_id = 0
    group_map: Dict[tuple, int] = {}
    for idx, r in enumerate(norm_sorted):
        k = key_for(r)
        if k not in group_map and counts[k] > 1:
            tie_group_id += 1
            group_map[k] = tie_group_id
        if idx == 0:
            r['rank'] = 1
        else:
            prev = norm_sorted[idx - 1]
            if key_for(prev) == k:
                r['rank'] = prev['rank']
            else:
                r['rank'] = idx + 1
        # tie metadata
        if counts[k] > 1:
            r['tie_group'] = group_map[k]
            r['tie_highlight'] = True
        else:
            r['tie_group'] = None
            r['tie_highlight'] = False

    return norm_sorted

    


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--input', '-i', default='leaderboard.xlsx', help='Path to leaderboard.xlsx')
    p.add_argument('--out', '-o', default='Test2/output', help='Output directory')
    args = p.parse_args()
    inp = args.input
    # if relative, resolve from workspace root
    if not os.path.isabs(inp):
        inp = os.path.join(os.getcwd(), inp)
    result = run(inp, args.out)
    print('Wrote:', result['csv'], result['json'])


if __name__ == '__main__':
    main()
