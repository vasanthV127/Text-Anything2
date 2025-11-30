Test2 — Leaderboard processing
===============================

What this folder contains
- `process_leaderboard.py`: script that reads `leaderboard.xlsx` and
  produces `Test2/output/leaderboard_normalized.csv` and
  `Test2/output/leaderboard_normalized.json`.
- `requirements.txt`: optional dependencies (`pandas`, `openpyxl`) if you
  want to switch to a pandas implementation.

Quick start (PowerShell)
-------------------------
From workspace root:

```powershell
# (optional) Create/activate venv
# Install optional deps
pip install -r Test2/requirements.txt

# Run the parser (it will read leaderboard.xlsx from cwd if relative)
python -m Test2.process_leaderboard --input leaderboard.xlsx --out Test2/output
```

Outputs
- `Test2/output/leaderboard_normalized.csv`
- `Test2/output/leaderboard_normalized.json`

Notes
- The included parser intentionally avoids heavy dependencies and works
  by reading the XLSX ZIP contents and extracting `sheet1.xml` and
  `sharedStrings.xml`. It uses heuristics based on the file layout found
  in the supplied `leaderboard.xlsx`.

Ranking rules implemented
-------------------------
- Primary sort: total points (prefers explicit `Total`/`Pts` column if numeric, otherwise computed from round columns).
- Tiebreaker 1: lower spending wins (spending parsed from any header matching `spend|spent|$m|$` — missing spend treated as 0).
- Tiebreaker 2: countback — compare players' round scores in descending order lexicographically (this effectively compares highest score first, then considers repeated occurrences of that score, then next highest, etc.).
- Final fallback: if all numeric tiebreakers are equal, players are shown in alphabetical order for deterministic output but are marked as tied.

Assumptions & edge-cases
------------------------
- Cells containing `-`, empty, or `D$Q` are treated as zero for totals and countback.
- Spending values are parsed by removing commas and `$` and converting to float; unparsable values default to 0.0.
- The script adds `rank`, `tie_group`, and `tie_highlight` fields to the JSON output.

Testing
-------
Run the unit tests which include focused tiebreaker checks:

```powershell
python -m pytest -q
```

If you'd like me to push these changes to your repository or add an API to serve the ranked JSON, tell me and I'll proceed.
