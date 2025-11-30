import os
import json
from Test2 import process_leaderboard


def test_parse_and_write(tmp_path):
    workspace = os.getcwd()
    src = os.path.join(workspace, 'leaderboard.xlsx')
    out = tmp_path / 'out'
    out = str(out)
    res = process_leaderboard.run(src, out)
    assert 'csv' in res and 'json' in res
    assert os.path.exists(res['csv'])
    assert os.path.exists(res['json'])
    with open(res['json'], 'r', encoding='utf-8') as fh:
        j = json.load(fh)
        assert 'rows' in j and isinstance(j['rows'], list)
