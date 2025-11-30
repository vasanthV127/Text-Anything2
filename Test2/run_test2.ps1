# PowerShell helper to run Test2 processing
Set-Location -LiteralPath (Resolve-Path ..).Path
if (-Not (Test-Path -Path .venv)) {
    python -m venv .venv
}
. .\.venv\Scripts\Activate.ps1
pip install -r Test2/requirements.txt --disable-pip-version-check
python -m Test2.process_leaderboard --input leaderboard.xlsx --out Test2/output
