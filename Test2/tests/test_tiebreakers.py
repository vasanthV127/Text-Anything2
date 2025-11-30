import pytest
from Test2.process_leaderboard import rank_rows


def make_row(name, total, spend, countback):
    return {
        'Player': name,
        '_rank_total': float(total),
        '_rank_spend': float(spend),
        '_countback': list(countback),
    }


def test_spend_tiebreaker():
    # Two players same total, lower spend should win
    a = make_row('Alpha', 100.0, 10.0, [50, 50])
    b = make_row('Beta', 100.0, 20.0, [50, 50])
    out = rank_rows([b, a])
    assert out[0]['Player'] == 'Alpha'
    assert out[0]['rank'] == 1
    assert out[1]['rank'] == 2


def test_countback_tiebreaker():
    # Same total and spend, but Alpha has higher top score
    a = make_row('Alpha', 100.0, 10.0, [60, 40])
    b = make_row('Beta', 100.0, 10.0, [50, 50])
    out = rank_rows([b, a])
    assert out[0]['Player'] == 'Alpha'


def test_countback_occurrence():
    # Same total and spend and same highest, but Alpha has two occurrences of top score
    a = make_row('Alpha', 100.0, 10.0, [60, 60, 40])
    b = make_row('Beta', 100.0, 10.0, [60, 50, 40])
    out = rank_rows([b, a])
    assert out[0]['Player'] == 'Alpha'


def test_full_tie_alphabetical():
    # Fully tied numeric values -> alphabetical decides order, but both marked tied
    a = make_row('Alpha', 100.0, 10.0, [50, 50])
    b = make_row('Beta', 100.0, 10.0, [50, 50])
    out = rank_rows([b, a])
    # alphabetical: Alpha before Beta
    assert out[0]['Player'] == 'Alpha'
    # ranks equal since numeric tiebreakers same
    assert out[0]['rank'] == out[1]['rank']
    assert out[0]['tie_group'] == out[1]['tie_group']
    assert out[0]['tie_highlight'] is True
