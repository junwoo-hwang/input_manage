"""기준 정보가 조용히 사라지거나 덮어써지지 않는가 (S3 는 가짜로 대신한다)."""
import io
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from src.input_manage import input_manage as im
from tests import fake_s3


@pytest.fixture(autouse=True)
def fake(monkeypatch):
    fake_s3.reset()
    monkeypatch.setattr(im, "s3", fake_s3)
    monkeypatch.setattr(im, "FOLDER_PATH", "2GAPU/input")
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
    assert im.list_workbooks() == ["OCAP기준", "설비기준"]


def test_a_round_trip_keeps_every_sheet_and_value():
    im.save_workbook("OCAP기준", sheets(
        코드표=[{"code": "C1", "desc": "재작업"}, {"code": "C2", "desc": "폐기"}],
        담당자=[{"item": "item1", "owner": "홍길동"}]), "hong")
    back, _ = im.load_workbook("OCAP기준")
    assert list(back) == ["코드표", "담당자"]
    assert back["코드표"]["desc"].tolist() == ["재작업", "폐기"]


def test_workbooks_do_not_bleed_into_each_other():
    im.save_workbook("A", sheets(S=[{"v": "a"}]), "hong")
    im.save_workbook("B", sheets(S=[{"v": "b"}]), "hong")
    assert im.load_workbook("A")[0]["S"]["v"].tolist() == ["a"]
    assert im.load_workbook("B")[0]["S"]["v"].tolist() == ["b"]


def test_a_missing_workbook_opens_empty_instead_of_raising():
    got, stamp = im.load_workbook("없는파일")
    assert got == {} and stamp == ""


# ------------------------------------------------------ 동시 편집 / 실패

def test_a_second_person_saving_first_is_refused_not_overwritten():
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    _mine, stamp = im.load_workbook("A")

    im.save_workbook("A", sheets(S=[{"a": 2}]), "kim")       # 옆 사람이 먼저

    with pytest.raises(im.ConcurrentEdit):
        im.save_workbook("A", sheets(S=[{"a": 3}]), "hong", base_stamp=stamp)
    assert im.load_workbook("A")[0]["S"]["a"].tolist() == [2], \
        "먼저 저장한 값이 덮어써졌습니다"


def test_the_guard_is_per_workbook_not_global():
    """다른 파일을 누가 저장했다고 내 저장이 막히면 안 된다."""
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    _s, stamp_a = im.load_workbook("A")
    im.save_workbook("B", sheets(S=[{"b": 1}]), "kim")
    im.save_workbook("A", sheets(S=[{"a": 2}]), "hong", base_stamp=stamp_a)
    assert im.load_workbook("A")[0]["S"]["a"].tolist() == [2]


def test_saving_with_the_returned_stamp_goes_through():
    stamp = im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    stamp = im.save_workbook("A", sheets(S=[{"a": 2}]), "hong", base_stamp=stamp)
    im.save_workbook("A", sheets(S=[{"a": 3}]), "hong", base_stamp=stamp)
    assert im.load_workbook("A")[0]["S"]["a"].tolist() == [3]


def test_a_failed_save_leaves_the_stored_file_intact():
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    before = fake_s3.STORE["2GAPU/input/A.xlsx"]

    def explode(*_a, **_k):
        raise RuntimeError("S3 가 응답하지 않는다")

    with pytest.MonkeyPatch.context() as mp:
        mp.setattr(im, "xlsx_write", explode)
        with pytest.raises(Exception):
            im.save_workbook("A", sheets(S=[{"a": 2}]), "hong")
    assert fake_s3.STORE["2GAPU/input/A.xlsx"] == before
    assert im.load_workbook("A")[0]["S"]["a"].tolist() == [1]


# ------------------------------------------------------------- 이력

def test_every_save_keeps_a_copy_named_by_who_saved_it():
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    im.save_workbook("A", sheets(S=[{"a": 2}]), "kim.lee")
    keys = im.history_keys("A")
    assert len(keys) == 2
    assert any("hong" in k for k in keys) and any("kim.lee" in k for k in keys)
    assert all(k.startswith("2GAPU/input/_history/A/") for k in keys), keys


def test_history_is_trimmed_so_it_cannot_grow_forever(monkeypatch):
    monkeypatch.setattr(im, "HISTORY_KEEP", 3)
    for i in range(6):
        im.save_workbook("A", sheets(S=[{"a": i}]), f"u{i}")
    assert len(im.history_keys("A")) == 3


def test_a_weird_user_id_cannot_escape_the_history_folder():
    im.save_workbook("A", sheets(S=[{"a": 1}]), "../../etc/passwd")
    keys = im.history_keys("A")
    assert len(keys) == 1 and keys[0].startswith("2GAPU/input/_history/A/"), keys


def test_the_history_copy_exists_even_if_the_main_put_fails():
    """되돌릴 판이 없는 순간이 생기면 안 된다."""
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
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
            im.save_workbook("A", sheets(S=[{"a": 2}]), "hong")
    assert len(im.history_keys("A")) == 2, "이력이 안 남았습니다"


# -------------------------------------------------------------- 감사

def test_the_audit_says_who_changed_what_in_which_file():
    im.save_workbook("OCAP기준", sheets(코드표=[{"a": 1, "b": 2}]), "hong")
    im.save_workbook("OCAP기준", sheets(코드표=[{"a": 1, "b": 9}]), "kim")
    log = im.read_audit("OCAP기준")
    assert log.iloc[0]["user_id"] == "kim"
    assert log.iloc[0]["workbook"] == "OCAP기준"
    assert log.iloc[0]["sheet"] == "코드표"
    assert log.iloc[0]["changed_cells"] == "1", log.to_dict("records")


def test_the_audit_of_one_file_does_not_show_another():
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    im.save_workbook("B", sheets(S=[{"b": 1}]), "kim")
    assert set(im.read_audit("A")["workbook"]) == {"A"}


def test_the_audit_survives_many_saves_without_losing_earlier_lines():
    for i in range(5):
        im.save_workbook("A", sheets(S=[{"a": i}]), f"u{i}")
    log = im.read_audit("A")
    assert len(log) == 5, log.to_dict("records")
    assert list(log["user_id"]) == ["u4", "u3", "u2", "u1", "u0"]


def test_a_save_that_changed_nothing_is_not_logged():
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    n = len(im.read_audit("A"))
    im.save_workbook("A", sheets(S=[{"a": 1}]), "kim")
    assert len(im.read_audit("A")) == n


# ------------------------------------------------------------ 값 다듬기

def test_blank_rows_added_by_the_grid_are_not_saved():
    im.save_workbook("A", sheets(
        S=[{"a": "x", "b": "y"}, {"a": None, "b": None}, {"a": "", "b": "  "}]), "hong")
    assert len(im.load_workbook("A")[0]["S"]) == 1


def test_missing_values_are_saved_blank_not_as_the_word_nan():
    im.save_workbook("A", sheets(S=[{"a": "x", "b": None}]), "hong")
    assert im.load_workbook("A")[0]["S"]["b"].tolist() == [None]


@pytest.mark.parametrize("before,after,want", [
    ([{"a": 1}], [{"a": 1}], 0),
    ([{"a": 1}], [{"a": 2}], 1),
    ([{"a": 1}], [{"a": 1}, {"a": 2}], 1),              # 줄이 늘었다
    ([{"a": 1}, {"a": 2}], [{"a": 1}], 1),              # 줄이 줄었다
    ([{"a": 1}], [{"a": 1, "b": 2}], 1),                # 칸이 늘었다
])
def test_changed_cells_counts_what_a_person_would_count(before, after, want):
    assert im.changed_cells(pd.DataFrame(before), pd.DataFrame(after)) == want


@pytest.mark.parametrize("name", ["아주긴시트이름" * 10, "재고/현황", "a[b]c", "매출:합"])
def test_sheet_names_excel_would_refuse_do_not_break_the_save(name):
    im.save_workbook("A", {name: pd.DataFrame([{"a": 1}])}, "hong")
    got = list(im.load_workbook("A")[0])[0]
    assert len(got) <= 31 and not set(got) & set(':\\/?*[]'), got


def test_adding_an_empty_column_counts_as_a_change():
    """값만 견주면 0 이 나와서 열을 넣고 저장을 누를 수가 없다."""
    before = pd.DataFrame([{"a": "1"}, {"a": "2"}])
    after = pd.DataFrame([{"a": "1", "새칸": ""}, {"a": "2", "새칸": ""}])
    assert im.changed_cells(before, after) == 1


def test_removing_an_empty_column_counts_as_a_change():
    before = pd.DataFrame([{"a": "1", "빈칸": ""}])
    after = pd.DataFrame([{"a": "1"}])
    assert im.changed_cells(before, after) == 1


def test_a_column_added_with_content_is_not_counted_twice():
    """내용이 있으면 그 값이 이미 세어졌으므로 더하지 않는다."""
    before = pd.DataFrame([{"a": "1"}])
    after = pd.DataFrame([{"a": "1", "b": "2"}])
    assert im.changed_cells(before, after) == 1


# -------------------------------------- 저장 전에 보여 줄 것 (무엇이 바뀌나)

def frame(rows, cols=("a", "b")):
    return pd.DataFrame(rows, columns=list(cols), dtype=object)


def kinds(before, after):
    rows, total = im.row_changes(before, after)
    return [(r["kind"], r["row"]) for r in rows], total


def test_an_edited_row_is_called_edited():
    before = frame([["1", "x"], ["2", "y"]])
    after = frame([["1", "x"], ["2", "바뀜"]])
    assert kinds(before, after) == ([("수정", 2)], 1)


def test_a_row_added_at_the_end_is_called_new():
    before = frame([["1", "x"]])
    after = frame([["1", "x"], ["2", "y"]])
    assert kinds(before, after) == ([("신규", 2)], 1)


def test_a_row_inserted_in_the_middle_does_not_mark_the_rest_as_edited():
    """자리만 맞춰 비교하면 끼워 넣은 줄 아래가 전부 '수정' 으로 나온다.

    8000줄짜리에서 줄 하나 넣고 저장하면 '7999줄 수정' 이라고 뜨는 셈이라,
    사람이 그 창을 안 읽게 된다 -- 읽으라고 띄우는 창인데.
    """
    before = frame([["1", "x"], ["2", "y"], ["3", "z"]])
    after = frame([["1", "x"], ["새", "줄"], ["2", "y"], ["3", "z"]])
    assert kinds(before, after) == ([("신규", 2)], 1)


def test_a_deleted_row_is_called_deleted():
    before = frame([["1", "x"], ["2", "y"], ["3", "z"]])
    after = frame([["1", "x"], ["3", "z"]])
    assert kinds(before, after) == ([("삭제", 2)], 1)


def test_nothing_changed_means_nothing_to_show():
    before = frame([["1", "x"], ["2", "y"]])
    assert kinds(before, before.copy()) == ([], 0)


def test_a_blank_row_left_over_is_not_a_change():
    """'행 아래' 를 눌렀다 안 채우고 저장하는 일이 흔하다."""
    before = frame([["1", "x"]])
    after = frame([["1", "x"], ["", ""]])
    assert kinds(before, after) == ([], 0)


def test_the_changed_row_carries_its_whole_row():
    before = frame([["1", "x"]])
    after = frame([["1", "바뀜"]])
    rows, _ = im.row_changes(before, after)
    assert rows[0]["values"] == {"a": "1", "b": "바뀜"}


def test_too_many_changes_are_counted_but_not_all_listed():
    before = frame([[str(i), "x"] for i in range(500)])
    after = frame([[str(i), "y"] for i in range(500)])
    rows, total = im.row_changes(before, after, limit=10)
    assert total == 500 and len(rows) == 10


def test_workbook_changes_leaves_out_untouched_sheets():
    before = {"A": frame([["1", "x"]]), "B": frame([["1", "x"]])}
    after = {"A": frame([["1", "바뀜"]]), "B": frame([["1", "x"]])}
    got = im.workbook_changes(before, after)
    assert list(got) == ["A"]
    assert got["A"]["total"] == 1


def test_workbook_changes_notices_a_new_column():
    before = {"A": frame([["1", "x"]])}
    after = {"A": frame([["1", "x", ""]], cols=("a", "b", "c"))}
    assert "칸 추가: c" in im.workbook_changes(before, after)["A"]["note"]


def test_workbook_changes_notices_a_dropped_sheet():
    got = im.workbook_changes({"A": frame([["1", "x"]]), "B": frame([["1", "x"]])},
                              {"A": frame([["1", "x"]])})
    assert got["B"]["note"] == "시트 삭제"


# ------------------------------------------------------------ REV_INFO

def rev_book():
    return {"STEP": frame([["1", "x"]]),
            "REV_INFO": pd.DataFrame(
                [{"Date": "2026-09-01", "Remark": "최초", "user": "hong", "관련": ""}],
                dtype=object)}


def test_rev_columns_are_found():
    assert im.rev_columns(rev_book()) == ["Date", "Remark", "user", "관련"]


def test_no_rev_sheet_means_no_reason_is_demanded():
    assert im.rev_columns({"STEP": frame([["1", "x"]])}) is None


def test_the_reason_is_appended_at_the_bottom():
    """위에 끼워 넣으면 다음에 열었을 때 그 시트의 줄 번호가 전부 밀린다."""
    book = rev_book()
    out = im.append_rev_info(book, "2026-09-21", "오탈자", "김철수", "JIRA-1")
    assert out["REV_INFO"].values.tolist() == [
        ["2026-09-01", "최초", "hong", ""],
        ["2026-09-21", "오탈자", "김철수", "JIRA-1"],
    ]


def test_appending_does_not_touch_what_was_passed_in():
    book = rev_book()
    im.append_rev_info(book, "2026-09-21", "사유", "나", "")
    assert len(book["REV_INFO"]) == 1


def test_the_optional_field_may_be_empty():
    out = im.append_rev_info(rev_book(), "2026-09-21", "사유", "나", "")
    assert out["REV_INFO"].iloc[-1]["관련"] == ""


def test_columns_are_matched_ignoring_case_and_spaces():
    book = {"REV_INFO": pd.DataFrame(columns=[" DATE ", "remark", "USER", "관련"],
                                     dtype=object)}
    out = im.append_rev_info(book, "2026-09-21", "사유", "나", "세부")
    assert out["REV_INFO"].iloc[-1].tolist() == ["2026-09-21", "사유", "나", "세부"]


def test_a_rev_sheet_with_other_columns_leaves_them_blank():
    book = {"REV_INFO": pd.DataFrame(columns=["Date", "Remark", "user", "관련", "기타"],
                                     dtype=object)}
    out = im.append_rev_info(book, "2026-09-21", "사유", "나", "")
    assert out["REV_INFO"].iloc[-1]["기타"] == ""


def test_the_saved_file_carries_the_new_rev_row():
    book = rev_book()
    with_rev = im.append_rev_info(book, "2026-09-21", "오탈자", "김철수", "")
    im.save_workbook("A", with_rev, "김철수")
    back = im.load_workbook("A")[0]["REV_INFO"]
    assert back.iloc[-1]["Remark"] == "오탈자"
    assert back.iloc[-1]["Date"] == "2026-09-21"
