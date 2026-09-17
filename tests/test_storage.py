"""기준 정보가 조용히 사라지거나 덮어써지지 않는가 (S3 는 가짜로 대신한다)."""
import io
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from src.input_manage import storage
from tests import fake_s3


@pytest.fixture(autouse=True)
def fake(monkeypatch):
    fake_s3.reset()
    monkeypatch.setattr(storage, "s3io", fake_s3)
    monkeypatch.setattr(storage, "PREFIX", "2GAPU/input")
    return fake_s3


def sheets(**kw):
    return {k: pd.DataFrame(v) for k, v in kw.items()}


def put_book(book, data):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for n, df in data.items():
            df.to_excel(w, sheet_name=n, index=False)
    fake_s3.put_object(f"2GAPU/input/{book}.xlsx", buf.getvalue())


# ---------------------------------------------------------- 목록 / 왕복

def test_only_real_workbooks_are_listed():
    put_book("OCAP기준", sheets(S=[{"a": 1}]))
    put_book("설비기준", sheets(S=[{"a": 1}]))
    fake_s3.put_object("2GAPU/input/_audit.csv", b"x")
    fake_s3.put_object("2GAPU/input/_history/OCAP기준/20260101_000000_a.xlsx", b"x")
    fake_s3.put_object("2GAPU/input/읽어보기.txt", b"x")
    assert storage.list_workbooks() == ["OCAP기준", "설비기준"]


def test_a_round_trip_keeps_every_sheet_and_value():
    storage.save_workbook("OCAP기준", sheets(
        코드표=[{"code": "C1", "desc": "재작업"}, {"code": "C2", "desc": "폐기"}],
        담당자=[{"item": "item1", "owner": "홍길동"}]), "hong")
    back, _ = storage.load_workbook("OCAP기준")
    assert list(back) == ["코드표", "담당자"]
    assert back["코드표"]["desc"].tolist() == ["재작업", "폐기"]


def test_workbooks_do_not_bleed_into_each_other():
    storage.save_workbook("A", sheets(S=[{"v": "a"}]), "hong")
    storage.save_workbook("B", sheets(S=[{"v": "b"}]), "hong")
    assert storage.load_workbook("A")[0]["S"]["v"].tolist() == ["a"]
    assert storage.load_workbook("B")[0]["S"]["v"].tolist() == ["b"]


def test_a_missing_workbook_opens_empty_instead_of_raising():
    got, stamp = storage.load_workbook("없는파일")
    assert got == {} and stamp == ""


# ------------------------------------------------------ 동시 편집 / 실패

def test_a_second_person_saving_first_is_refused_not_overwritten():
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    _mine, stamp = storage.load_workbook("A")

    storage.save_workbook("A", sheets(S=[{"a": 2}]), "kim")       # 옆 사람이 먼저

    with pytest.raises(storage.ConcurrentEdit):
        storage.save_workbook("A", sheets(S=[{"a": 3}]), "hong", base_stamp=stamp)
    assert storage.load_workbook("A")[0]["S"]["a"].tolist() == [2], \
        "먼저 저장한 값이 덮어써졌습니다"


def test_the_guard_is_per_workbook_not_global():
    """다른 파일을 누가 저장했다고 내 저장이 막히면 안 된다."""
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    _s, stamp_a = storage.load_workbook("A")
    storage.save_workbook("B", sheets(S=[{"b": 1}]), "kim")
    storage.save_workbook("A", sheets(S=[{"a": 2}]), "hong", base_stamp=stamp_a)
    assert storage.load_workbook("A")[0]["S"]["a"].tolist() == [2]


def test_saving_with_the_returned_stamp_goes_through():
    stamp = storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    stamp = storage.save_workbook("A", sheets(S=[{"a": 2}]), "hong", base_stamp=stamp)
    storage.save_workbook("A", sheets(S=[{"a": 3}]), "hong", base_stamp=stamp)
    assert storage.load_workbook("A")[0]["S"]["a"].tolist() == [3]


def test_a_failed_save_leaves_the_stored_file_intact():
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    before = fake_s3.STORE["2GAPU/input/A.xlsx"]

    def explode(*_a, **_k):
        raise RuntimeError("S3 가 응답하지 않는다")

    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(pd.DataFrame, "to_excel", explode)
        with pytest.raises(Exception):
            storage.save_workbook("A", sheets(S=[{"a": 2}]), "hong")
    assert fake_s3.STORE["2GAPU/input/A.xlsx"] == before
    assert storage.load_workbook("A")[0]["S"]["a"].tolist() == [1]


# ------------------------------------------------------------- 이력

def test_every_save_keeps_a_copy_named_by_who_saved_it():
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    storage.save_workbook("A", sheets(S=[{"a": 2}]), "kim.lee")
    keys = storage.history_keys("A")
    assert len(keys) == 2
    assert any("hong" in k for k in keys) and any("kim.lee" in k for k in keys)
    assert all(k.startswith("2GAPU/input/_history/A/") for k in keys), keys


def test_history_is_trimmed_so_it_cannot_grow_forever(monkeypatch):
    monkeypatch.setattr(storage, "HISTORY_KEEP", 3)
    for i in range(6):
        storage.save_workbook("A", sheets(S=[{"a": i}]), f"u{i}")
    assert len(storage.history_keys("A")) == 3


def test_a_weird_user_id_cannot_escape_the_history_folder():
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "../../etc/passwd")
    keys = storage.history_keys("A")
    assert len(keys) == 1 and keys[0].startswith("2GAPU/input/_history/A/"), keys


def test_the_history_copy_exists_even_if_the_main_put_fails():
    """되돌릴 판이 없는 순간이 생기면 안 된다."""
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    calls = {"n": 0}
    real = fake_s3.put_object

    def flaky(key, data):
        calls["n"] += 1
        if key.endswith("/A.xlsx") and calls["n"] > 1:
            raise RuntimeError("본 파일 올리기 실패")
        return real(key, data)

    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(fake_s3, "put_object", flaky)
        with pytest.raises(RuntimeError):
            storage.save_workbook("A", sheets(S=[{"a": 2}]), "hong")
    assert len(storage.history_keys("A")) == 2, "이력이 안 남았습니다"


# -------------------------------------------------------------- 감사

def test_the_audit_says_who_changed_what_in_which_file():
    storage.save_workbook("OCAP기준", sheets(코드표=[{"a": 1, "b": 2}]), "hong")
    storage.save_workbook("OCAP기준", sheets(코드표=[{"a": 1, "b": 9}]), "kim")
    log = storage.read_audit("OCAP기준")
    assert log.iloc[0]["user_id"] == "kim"
    assert log.iloc[0]["workbook"] == "OCAP기준"
    assert log.iloc[0]["sheet"] == "코드표"
    assert log.iloc[0]["changed_cells"] == "1", log.to_dict("records")


def test_the_audit_of_one_file_does_not_show_another():
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    storage.save_workbook("B", sheets(S=[{"b": 1}]), "kim")
    assert set(storage.read_audit("A")["workbook"]) == {"A"}


def test_the_audit_survives_many_saves_without_losing_earlier_lines():
    for i in range(5):
        storage.save_workbook("A", sheets(S=[{"a": i}]), f"u{i}")
    log = storage.read_audit("A")
    assert len(log) == 5, log.to_dict("records")
    assert list(log["user_id"]) == ["u4", "u3", "u2", "u1", "u0"]


def test_a_save_that_changed_nothing_is_not_logged():
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    n = len(storage.read_audit("A"))
    storage.save_workbook("A", sheets(S=[{"a": 1}]), "kim")
    assert len(storage.read_audit("A")) == n


# ------------------------------------------------------------ 값 다듬기

def test_blank_rows_added_by_the_grid_are_not_saved():
    storage.save_workbook("A", sheets(
        S=[{"a": "x", "b": "y"}, {"a": None, "b": None}, {"a": "", "b": "  "}]), "hong")
    assert len(storage.load_workbook("A")[0]["S"]) == 1


def test_missing_values_are_saved_blank_not_as_the_word_nan():
    storage.save_workbook("A", sheets(S=[{"a": "x", "b": None}]), "hong")
    assert storage.load_workbook("A")[0]["S"]["b"].tolist() == [None]


@pytest.mark.parametrize("before,after,want", [
    ([{"a": 1}], [{"a": 1}], 0),
    ([{"a": 1}], [{"a": 2}], 1),
    ([{"a": 1}], [{"a": 1}, {"a": 2}], 1),              # 줄이 늘었다
    ([{"a": 1}, {"a": 2}], [{"a": 1}], 1),              # 줄이 줄었다
    ([{"a": 1}], [{"a": 1, "b": 2}], 1),                # 칸이 늘었다
])
def test_changed_cells_counts_what_a_person_would_count(before, after, want):
    assert storage.changed_cells(pd.DataFrame(before), pd.DataFrame(after)) == want


@pytest.mark.parametrize("name", ["아주긴시트이름" * 10, "재고/현황", "a[b]c", "매출:합"])
def test_sheet_names_excel_would_refuse_do_not_break_the_save(name):
    storage.save_workbook("A", {name: pd.DataFrame([{"a": 1}])}, "hong")
    got = list(storage.load_workbook("A")[0])[0]
    assert len(got) <= 31 and not set(got) & set(':\\/?*[]'), got
