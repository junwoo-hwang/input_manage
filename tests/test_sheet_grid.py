"""격자가 올려준 것을 표로 되돌리는 부분.

격자 자체(끌어서 선택, 복사/붙여넣기, 행·열 넣고 빼기)는 브라우저에서만
도는 자바스크립트라 tests/test_browser.py 가 진짜 브라우저로 확인한다.
여기는 그 결과를 파이썬이 어떻게 받아들이는지만 본다.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from src.input_manage.sheet_grid import to_frame


def test_a_plain_grid_becomes_the_same_table():
    df = to_frame({"cols": ["code", "desc"], "rows": [["C1", "재작업"], ["C2", "폐기"]]})
    assert list(df.columns) == ["code", "desc"]
    assert df["desc"].tolist() == ["재작업", "폐기"]


def test_duplicate_column_names_are_made_unique():
    """겹친 채로 두면 pandas 에서 df["a"] 가 Series 가 아니라 DataFrame 이 되고,
    엑셀로 내보낼 때 값 대신 칸 이름이 실린 파일이 오류 없이 만들어진다."""
    df = to_frame({"cols": ["a", "a", "a"], "rows": [["1", "2", "3"]]})
    assert list(df.columns) == ["a", "a_1", "a_2"]
    assert df.iloc[0].tolist() == ["1", "2", "3"]


def test_a_short_row_is_padded_not_dropped():
    """붙여넣기로 줄마다 칸 수가 달라질 수 있다. 짧다고 버리면 값이 사라진다."""
    df = to_frame({"cols": ["a", "b", "c"], "rows": [["1"], ["1", "2", "3"]]})
    assert df.iloc[0].tolist() == ["1", "", ""]
    assert len(df) == 2


def test_a_long_row_is_cut_to_the_columns_that_exist():
    df = to_frame({"cols": ["a"], "rows": [["1", "2", "3"]]})
    assert df.iloc[0].tolist() == ["1"]


def test_an_empty_grid_does_not_crash():
    df = to_frame({"cols": [], "rows": []})
    assert df.empty and list(df.columns) == []


def test_values_stay_text_so_ids_do_not_lose_their_leading_zeros():
    """wafer '03' 이나 코드 '0012' 가 3, 12 로 바뀌면 안 된다."""
    df = to_frame({"cols": ["wafer"], "rows": [["03"], ["0012"]]})
    assert df["wafer"].tolist() == ["03", "0012"]
