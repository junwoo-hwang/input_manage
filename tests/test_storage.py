"""기준 정보가 조용히 사라지거나 덮어써지지 않는가 (S3 는 가짜로 대신한다)."""
import io
import sys
import zipfile
from pathlib import Path

import openpyxl
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


# ------------------------------------- 이력은 REV_INFO 시트 한 곳에만

def test_saving_leaves_no_files_beside_the_workbook():
    """옆에 _history/ 나 _audit.csv 를 만들지 않는다.

    기준 정보를 받아 보는 사람이 파일 하나만 열면 이력까지 같이 보게
    하려고 그렇게 정했다. 옆 파일이 생기면 그 둘이 갈리기 시작한다.
    """
    im.save_workbook("A", sheets(S=[{"a": 1}]), "hong")
    im.save_workbook("A", sheets(S=[{"a": 2}]), "kim")
    assert list(fake_s3.STORE) == ["2GAPU/input/A.xlsx"], list(fake_s3.STORE)


def test_the_change_log_lands_in_the_workbook_itself():
    book = rev_book()
    after = {k: v.copy() for k, v in book.items()}
    after["STEP"].loc[0, "b"] = "바뀜"
    changes = im.workbook_changes(book, after)
    body = im.append_rev_info(after, "2026-09-21", "사유", "나", "",
                              im.changes_text(changes))
    im.save_workbook("A", body, "나")

    back = im.load_workbook("A")[0]["REV_INFO"]
    assert back.iloc[-1]["Remark"] == "사유"
    assert "수정 1행" in back.iloc[-1]["관련"], back.iloc[-1]["관련"]


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


# ------------------------------------------------------------ 수식 지키기

def book_with_formula():
    """C 칸에 VLOOKUP 이 걸린 시트."""
    sheets = {"STEP": pd.DataFrame(
        [{"step": "0010", "ppid": "P-01", "찾은값": "가"},
         {"step": "0020", "ppid": "P-02", "찾은값": "나"}], dtype=object)}
    formulas = {"STEP": {(0, "찾은값"): "VLOOKUP(B2,MAP!$A$1:$B$99,2,0)",
                         (1, "찾은값"): "VLOOKUP(B3,MAP!$A$1:$B$99,2,0)"}}
    return sheets, formulas


def test_a_formula_survives_a_round_trip():
    """이게 안 되면 VLOOKUP 이 마지막 계산값으로 굳어 버린다."""
    sheets, formulas = book_with_formula()
    im.save_workbook("A", sheets, "hong", formulas=formulas)
    got: dict = {}
    im.load_workbook("A", got)
    assert got["STEP"][(0, "찾은값")] == "VLOOKUP(B2,MAP!$A$1:$B$99,2,0)"


def test_the_saved_file_really_holds_a_formula_not_a_value():
    sheets, formulas = book_with_formula()
    im.save_workbook("A", sheets, "hong", formulas=formulas)
    raw = fake_s3.STORE["2GAPU/input/A.xlsx"]
    book = openpyxl.load_workbook(io.BytesIO(raw))
    assert book["STEP"]["C2"].value == "=VLOOKUP(B2,MAP!$A$1:$B$99,2,0)"


def test_editing_another_column_keeps_the_formula():
    sheets, formulas = book_with_formula()
    after = {"STEP": sheets["STEP"].copy()}
    after["STEP"].loc[0, "ppid"] = "P-99"
    kept, lost = im.surviving_formulas(sheets, after, formulas)
    assert len(kept["STEP"]) == 2 and not lost


def test_typing_over_a_formula_cell_wins():
    """사람이 그 칸에 직접 값을 적었으면 그 값이 이긴다."""
    sheets, formulas = book_with_formula()
    after = {"STEP": sheets["STEP"].copy()}
    after["STEP"].loc[0, "찾은값"] = "손으로 적음"
    kept, lost = im.surviving_formulas(sheets, after, formulas)
    assert (0, "찾은값") not in kept.get("STEP", {})
    assert (1, "찾은값") in kept["STEP"]
    assert lost == {"STEP": 1}


def test_a_row_inserted_above_drops_the_formulas_below():
    """5번 줄의 =VLOOKUP(B5,..) 는 6번 줄로 밀리면 B6 을 봐야 맞다.

    자리를 따라 고쳐 주지는 못하므로, 틀린 수식을 남기는 대신 버린다.
    """
    sheets, formulas = book_with_formula()
    grown = pd.concat([
        pd.DataFrame([{"step": "0005", "ppid": "P-00", "찾은값": ""}], dtype=object),
        sheets["STEP"]], ignore_index=True)
    kept, lost = im.surviving_formulas(sheets, {"STEP": grown}, formulas)
    assert not kept
    assert lost == {"STEP": 2}


def test_a_row_added_at_the_end_keeps_the_formulas_above():
    sheets, formulas = book_with_formula()
    grown = pd.concat([
        sheets["STEP"],
        pd.DataFrame([{"step": "0030", "ppid": "P-03", "찾은값": ""}], dtype=object)],
        ignore_index=True)
    kept, lost = im.surviving_formulas(sheets, {"STEP": grown}, formulas)
    assert len(kept["STEP"]) == 2 and not lost


def test_a_dropped_sheet_drops_its_formulas():
    sheets, formulas = book_with_formula()
    kept, lost = im.surviving_formulas(sheets, {}, formulas)
    assert not kept and lost == {"STEP": 2}


def test_a_column_inserted_to_the_left_does_not_move_the_formula():
    """칸을 번호가 아니라 이름으로 붙들어 두는 자리다."""
    sheets, formulas = book_with_formula()
    after = sheets["STEP"].copy()
    after.insert(0, "새칸", ["", ""])
    kept, _lost = im.surviving_formulas(sheets, {"STEP": after}, formulas)
    im.save_workbook("A", {"STEP": after}, "hong", formulas=kept)
    raw = fake_s3.STORE["2GAPU/input/A.xlsx"]
    book = openpyxl.load_workbook(io.BytesIO(raw))
    assert book["STEP"]["A1"].value == "새칸"
    assert str(book["STEP"]["D2"].value).startswith("=VLOOKUP")


def test_a_blank_row_removed_on_save_pulls_the_formula_up():
    """_clean 이 빈 줄을 빼면 아래 줄이 당겨진다. 수식도 같이 당겨야 한다."""
    sheets = {"STEP": pd.DataFrame(
        [{"a": "", "f": ""}, {"a": "x", "f": "1"}], dtype=object)}
    formulas = {"STEP": {(1, "f"): "SUM(A3:A3)"}}
    im.save_workbook("A", sheets, "hong", formulas=formulas)
    raw = fake_s3.STORE["2GAPU/input/A.xlsx"]
    book = openpyxl.load_workbook(io.BytesIO(raw))
    assert book["STEP"]["B2"].value == "=SUM(A3:A3)", \
        [c.value for c in book["STEP"]["B"]]


def test_excel_is_told_to_recompute_on_open():
    """적어 둔 값은 낡았을 수 있다. 엑셀이 열 때 다시 계산해야 한다."""
    sheets, formulas = book_with_formula()
    im.save_workbook("A", sheets, "hong", formulas=formulas)
    raw = fake_s3.STORE["2GAPU/input/A.xlsx"]
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        assert b'fullCalcOnLoad="1"' in zf.read("xl/workbook.xml")


# ------------------------------------------------------------- 이력 내용



def test_the_change_text_reads_like_the_popup():
    """창에서 본 것과 같은 것이 파일에 남아야 한다 -- 나중에 되짚는 사람은
    그 창을 못 보고 이 칸만 본다."""
    before = {"STEP": frame([["1", "x"], ["2", "y"]])}
    after = {"STEP": frame([["1", "바뀜"], ["2", "y"], ["3", "새"]])}
    text = im.changes_text(im.workbook_changes(before, after))
    assert "[STEP]" in text
    assert "수정 1행: 1 | 바뀜" in text, text
    assert "신규 3행: 3 | 새" in text, text


def test_the_change_text_names_added_columns():
    before = {"STEP": frame([["1", "x"]])}
    after = {"STEP": frame([["1", "x", ""]], cols=("a", "b", "c"))}
    assert "칸 추가: c" in im.changes_text(im.workbook_changes(before, after))


def test_the_change_text_cannot_overflow_an_excel_cell():
    """엑셀 칸 하나는 32,767자까지다. 넘기면 파일이 안 열린다.

    줄 수는 300개로 막아 두지만 칸이 넓고 값이 길면 그것만으로 넘길 수 있다.
    """
    wide = [[("가" * 200) for _ in range(20)] for _ in range(300)]
    before = {"STEP": frame(wide, cols=[f"c{i}" for i in range(20)])}
    after = {"STEP": frame([[v + "!" for v in row] for row in wide],
                           cols=[f"c{i}" for i in range(20)])}
    text = im.changes_text(im.workbook_changes(before, after))
    assert len(text) <= im.CELL_MAX
    assert "잘림" in text, "잘렸다는 말이 없으면 뒤가 없는 건지 알 수 없다"


def test_a_truncated_change_text_says_so():
    before = {"STEP": frame([["1", "x"], ["2", "y"]])}
    after = {"STEP": frame([["1", "바뀜"], ["2", "또바뀜"]])}
    text = im.changes_text(im.workbook_changes(before, after), limit=40)
    assert len(text) <= 40 and text.endswith("잘림)"), repr(text)


def test_what_the_person_typed_comes_before_the_auto_part():
    out = im.append_rev_info(rev_book(), "2026-09-21", "사유", "나",
                             "JIRA-1", "[STEP] ...\n수정 1행: x")
    related = out["REV_INFO"].iloc[-1]["관련"]
    assert related.startswith("JIRA-1"), related
    assert "수정 1행" in related


def test_the_auto_part_alone_is_fine_when_nothing_was_typed():
    out = im.append_rev_info(rev_book(), "2026-09-21", "사유", "나", "",
                             "[STEP] ...")
    assert out["REV_INFO"].iloc[-1]["관련"] == "[STEP] ..."
