"""격자를 진짜 브라우저에서 두들겨 본다.

끌어서 선택 / 복사 / 붙여넣기 / 행·열 넣고 빼기는 전부 index.html 안의
자바스크립트에 있어서 파이썬 테스트로는 닿지 않는다. 그래서 여기만
브라우저를 띄운다 -- playwright 가 없으면 통째로 건너뛰므로 나머지
테스트는 브라우저 없이도 돈다.

    pip install playwright && playwright install chromium
"""
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

sync_playwright = pytest.importorskip(
    "playwright.sync_api", reason="playwright 가 없으면 브라우저 검사는 건너뛴다"
).sync_playwright


def _free_port():
    with socket.socket() as sock:
        sock.bind(("", 0))
        return sock.getsockname()[1]


@pytest.fixture(scope="module")
def server():
    port = _free_port()
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", str(ROOT / "app_local.py"),
         "--server.port", str(port), "--server.headless", "true",
         "--browser.gatherUsageStats", "false"],
        cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    url = f"http://localhost:{port}/"
    for _ in range(120):
        try:
            with socket.create_connection(("localhost", port), timeout=0.5):
                break
        except OSError:
            time.sleep(0.5)
    else:
        proc.terminate()
        pytest.fail("streamlit 이 안 떴습니다")
    time.sleep(3)
    yield url
    proc.terminate()
    proc.wait(timeout=20)


@pytest.fixture(scope="module")
def page(server):
    exe = os.environ.get("PLAYWRIGHT_CHROMIUM", "/opt/pw-browsers/chromium")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            executable_path=exe if os.path.exists(exe) else None)
        ctx = browser.new_context(viewport={"width": 1400, "height": 950})
        ctx.grant_permissions(["clipboard-read", "clipboard-write"])
        pg = ctx.new_page()
        errors = []
        pg.on("pageerror", lambda e: errors.append(str(e)))
        pg.goto(server)
        pg.wait_for_timeout(7000)
        pg.errors = errors
        yield pg
        browser.close()


def grid(page):
    """격자가 든 iframe."""
    return page.frame_locator("iframe[title*='sheet_grid'], "
                              "iframe[src*='sheet_grid']").first


def cells(page):
    return grid(page).locator(".tbl .row .cell:not(.rowhead)")


def table(page):
    """격자의 지금 내용 (헤더, 줄들)."""
    return grid(page).locator(".tbl").first.evaluate("""(t) => ({
      cols: [...t.querySelectorAll('.hrow .colhead')].map(e => e.textContent),
      rows: [...t.querySelectorAll('.row')].map(
              tr => [...tr.querySelectorAll('.cell:not(.rowhead)')].map(e => e.textContent)),
    })""")


def click_cell(page, r, c):
    """칸을 누른다. 겸사겸사 iframe 에 초점이 가서 자판/클립보드가 거기로 간다."""
    grid(page).locator(f".cell[data-r='{r}'][data-c='{c}']").click()
    page.wait_for_timeout(250)


def wait_dirty(page):
    """고친 것이 파이썬까지 올라가 '저장하지 않은 수정' 이 뜰 때까지.

    격자는 고친 것을 한 박자 모았다가 올리고, 그러면 streamlit 이 스크립트를
    다시 돈다. 그 왕복 시간은 시트 수와 줄 수에 따라 들쭉날쭉하다.
    """
    page.wait_for_function(
        "() => document.body.innerText.includes('저장하지 않은 수정')", timeout=30000)


def open_review(page, table=True):
    """저장을 눌러 '변경내용' 창을 띄운다.

    바뀐 줄을 담은 표는 창보다 한 박자 늦게 도착한다 (streamlit 이 화면을
    조각내어 보낸다). 표를 볼 검사는 그것까지 기다려야 한다.
    """
    page.get_by_role("button", name="저장", exact=True).first.click()
    page.wait_for_function(
        "() => document.body.innerText.includes('변경내용')", timeout=60000)
    if table:
        page.wait_for_selector("[role='dialog'] [data-testid='stTable'], "
                               "[data-testid='stTable']", timeout=60000)


def review_text(page):
    """'변경내용' 창 안의 글자만. 뒤에 깔린 화면 글자와 안 섞이게."""
    box = page.locator("[role='dialog']")
    return box.inner_text() if box.count() else page.inner_text("body")


def confirm_save(page, remark="사유"):
    """저장 -> 사유 적기 -> 진짜 저장. '저장했습니다' 가 뜰 때까지 기다린다.

    적자마자 바로 누른다. 사람이 그렇게 하기 때문이다 -- 칸을 떠나라고
    Tab 을 눌러 주거나 한 박자 기다려 주지 않는다.

    user 는 로그인한 사람으로 고정이라 못 고친다.
    """
    open_review(page, table=False)
    page.get_by_label("Remark — 사유").fill(remark)
    page.get_by_role("button", name="저장", exact=True).last.click()
    page.wait_for_function(
        "() => document.body.innerText.includes('저장했습니다')", timeout=60000)


def paste(page, text):
    """엑셀에서 긁어온 것처럼 클립보드에 넣고 붙여넣는다.

    locator.press 가 아니라 page.keyboard 를 쓴다 -- 전자는 그 요소로만
    키를 보내서 브라우저가 진짜 paste 이벤트를 만들지 않는다.
    """
    page.evaluate("(t) => navigator.clipboard.writeText(t)", text)
    page.keyboard.press("Control+v")
    page.wait_for_timeout(1500)


BOOK = "FAB_INPUT_ULY_r0"        # 시트가 넷이라 검사할 거리가 있는 쪽


def settle(page):
    """격자가 다 그려지고 '고친 것 없음' 이 뜰 때까지 기다린다."""
    page.wait_for_function(
        "() => document.body.innerText.includes('고친 것 없음')", timeout=30000)
    grid(page).locator(".cell[data-r='0'][data-c='0']").wait_for(timeout=30000)
    page.wait_for_timeout(300)


def reset(page):
    """페이지를 새로 열고 검사할 파일을 고른다.

    한 페이지를 여러 검사가 나눠 쓰면 앞 검사가 남긴 것(고르던 시트, 저장
    안 한 수정, 늘어난 칸)을 다음 검사가 그대로 물고 시작한다. 실제로 그래서
    따로 돌리면 통과하고 같이 돌리면 깨지는 검사가 여럿 나왔다. 새로 열면
    streamlit 세션이 새로 생겨 session_state 까지 깨끗해진다.
    """
    page.goto(page.url)
    page.wait_for_timeout(4000)
    page.get_by_role("combobox").click()
    page.wait_for_timeout(600)
    page.get_by_text(BOOK, exact=True).click()
    settle(page)


# ------------------------------------------------------------ 보이는가

def test_the_whole_sheet_is_shown_with_its_columns(page):
    reset(page)
    got = table(page)
    assert got["cols"], "칸 이름 줄이 없습니다"
    assert got["rows"], "값 줄이 없습니다"
    assert len(got["rows"][0]) == len(got["cols"]), "칸 수와 값 수가 안 맞습니다"


def test_no_javascript_errors(page):
    reset(page)
    assert page.errors == []


def test_ids_keep_their_leading_zeros(page):
    """'0010' 이 10 으로 바뀌면 기준 정보로 못 쓴다."""
    reset(page)
    flat = [v for row in table(page)["rows"] for v in row]
    assert any(v.startswith("0") and len(v) > 1 for v in flat), flat


# ---------------------------------------------------- 끌어서 범위 선택

def test_dragging_selects_a_rectangle(page):
    reset(page)
    a = grid(page).locator(".cell[data-r='0'][data-c='0']")
    b = grid(page).locator(".cell[data-r='0'][data-c='2']")
    a.hover(); page.mouse.down()
    b.hover(); page.mouse.up()
    page.wait_for_timeout(400)
    assert grid(page).locator(".cell.sel").count() == 3
    assert "선택 1x3" in grid(page).locator(".sheetbar .count").inner_text()


def test_clicking_a_column_header_selects_the_whole_column(page):
    reset(page)
    grid(page).locator(".colhead[data-c='1']").click()
    page.wait_for_timeout(400)
    rows = len(table(page)["rows"])
    assert grid(page).locator(".cell.sel").count() == rows


# -------------------------------------------------- 복사 / 붙여넣기

def test_pasting_a_range_from_excel_spreads_across_cells(page):
    """엑셀은 클립보드에 탭/줄바꿈으로 나눈 글자를 넣는다. 그대로 퍼져야 한다."""
    reset(page)
    before = table(page)
    click_cell(page, 0, 0)
    paste(page, "X1\tX2\nY1\tY2")
    got = table(page)
    assert got["rows"][0][0] == "X1" and got["rows"][0][1] == "X2"
    assert got["rows"][1][0] == "Y1" and got["rows"][1][1] == "Y2"
    assert len(got["rows"]) >= len(before["rows"]), "줄이 사라졌습니다"


def test_pasting_more_columns_than_exist_grows_the_table(page):
    reset(page)
    n = len(table(page)["cols"])
    click_cell(page, 0, n - 1)
    paste(page, "a\tb\tc")
    assert len(table(page)["cols"]) == n + 2


def test_copying_a_range_puts_tab_separated_text_on_the_clipboard(page):
    reset(page)
    a = grid(page).locator(".cell[data-r='0'][data-c='0']")
    b = grid(page).locator(".cell[data-r='0'][data-c='1']")
    a.hover(); page.mouse.down(); b.hover(); page.mouse.up()
    want = table(page)["rows"][0][:2]
    page.keyboard.press("Control+c")
    page.wait_for_timeout(900)
    text = page.evaluate("navigator.clipboard.readText()")
    assert text == "\t".join(want), (text, want)


def test_delete_clears_the_selected_range(page):
    reset(page)
    a = grid(page).locator(".cell[data-r='0'][data-c='0']")
    b = grid(page).locator(".cell[data-r='0'][data-c='1']")
    a.hover(); page.mouse.down(); b.hover(); page.mouse.up()
    page.keyboard.press("Delete")
    page.wait_for_timeout(900)
    assert table(page)["rows"][0][:2] == ["", ""]


# ---------------------------------------------------- 행 / 열 넣고 빼기

def test_inserting_and_deleting_a_row(page):
    reset(page)
    n = len(table(page)["rows"])
    click_cell(page, 0, 0)
    grid(page).locator("#bar [data-act='row-below']").click()
    page.wait_for_timeout(600)
    assert len(table(page)["rows"]) == n + 1
    grid(page).locator("#bar [data-act='row-del']").click()
    page.wait_for_timeout(600)
    assert len(table(page)["rows"]) == n


def test_inserting_and_deleting_a_column(page):
    """st.data_editor 로는 안 되는 것. 이게 이 격자를 직접 만든 이유다."""
    reset(page)
    cols = table(page)["cols"]
    click_cell(page, 0, 0)
    grid(page).locator("#bar [data-act='col-right']").click()
    page.wait_for_timeout(600)
    grown = table(page)["cols"]
    assert len(grown) == len(cols) + 1
    assert grown[1] not in cols, "새 칸 이름이 기존 것과 겹칩니다"

    click_cell(page, 0, 1)
    grid(page).locator("#bar [data-act='col-del']").click()
    page.wait_for_timeout(600)
    assert table(page)["cols"] == cols


def test_undo_puts_back_what_a_delete_removed(page):
    reset(page)
    before = table(page)
    click_cell(page, 0, 0)
    grid(page).locator("#bar [data-act='row-del']").click()
    page.wait_for_timeout(600)
    assert table(page)["rows"] != before["rows"]
    grid(page).locator("button[data-act='undo']").click()
    page.wait_for_timeout(600)
    assert table(page)["rows"] == before["rows"]


# ------------------------------------------------------- 고치고 저장

@pytest.mark.parametrize("typed", ["바뀐값", "NEW1"])
def test_typing_into_a_cell_and_saving_changes_what_is_stored(page, typed):
    """한글도 되어야 한다. 한글은 keydown 이 아니라 조합으로 들어와서,
    격자에 바로 자판을 받으면 한 글자도 안 들어온다."""
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type(typed)
    page.keyboard.press("Enter")
    wait_dirty(page)
    save = page.get_by_role("button", name="저장", exact=True)
    assert save.is_enabled(), "저장 버튼이 안 켜졌습니다"
    confirm_save(page, remark="검사")
    reset(page)
    assert table(page)["rows"][0][2] == typed, "저장한 값이 안 남았습니다"


# ------------------------------------------- 아래 시트 탭 (엑셀과 같은 자리)

def test_sheet_tabs_sit_below_the_grid(page):
    """엑셀처럼 시트가 표 아래에 있어야 한다."""
    reset(page)
    tabs = grid(page).locator(".sheetbar .tab")
    assert tabs.count() >= 2, "시트 탭이 안 보입니다"
    box = grid(page).locator(".sheetbar").bounding_box()
    grid_box = grid(page).locator(".scroll").bounding_box()
    assert box["y"] > grid_box["y"], "시트 탭이 표 위에 있습니다"


def test_switching_sheets_shows_that_sheet(page):
    reset(page)
    tabs = grid(page).locator(".sheetbar .tab")
    first = table(page)["cols"]
    names = [tabs.nth(i).inner_text() for i in range(tabs.count())]
    tabs.nth(1).click()
    page.wait_for_timeout(700)
    second = table(page)["cols"]
    assert second != first, f"{names[0]} -> {names[1]} 인데 표가 그대로입니다"
    assert "on" in (tabs.nth(1).get_attribute("class") or "")


def test_an_edit_on_one_sheet_survives_a_trip_to_another(page):
    """시트를 옮겼다 돌아오면 고치던 게 남아 있어야 한다.

    시트마다 iframe 을 따로 두면 이게 깨진다 -- 그래서 시트 전체를 컴포넌트
    하나가 들고 있다.
    """
    reset(page)
    click_cell(page, 0, 0)
    page.keyboard.type("남아라")
    page.keyboard.press("Enter")
    wait_dirty(page)

    tabs = grid(page).locator(".sheetbar .tab")
    tabs.nth(1).click(); page.wait_for_timeout(800)
    tabs.nth(0).click(); page.wait_for_timeout(800)
    assert table(page)["rows"][0][0] == "남아라"


def test_adding_a_sheet(page):
    reset(page)
    n = grid(page).locator(".sheetbar .tab").count()
    grid(page).locator("#addSheet").click()
    page.wait_for_timeout(600)
    assert grid(page).locator(".sheetbar .tab").count() == n + 1
    wait_dirty(page)


# ------------------------------------------- 주소 / 수식줄 / 찾기 / 내려받기

def test_the_address_box_uses_excel_notation(page):
    """A1, E1493 처럼 적어야 한다 -- 사람끼리 자리를 말할 때 쓰는 말이 그거다."""
    reset(page)
    click_cell(page, 0, 0)
    assert grid(page).locator("#addr").inner_text() == "A2"
    click_cell(page, 2, 4)
    assert grid(page).locator("#addr").inner_text() == "E4"


def test_the_address_row_matches_what_excel_would_show(page):
    """화면 1행은 엑셀에서 2행이다 (엑셀 1행은 칸 이름 줄).

    여기가 어긋나면 "A5 고쳐줘" 하고 엑셀을 열었을 때 한 줄씩 밀린다.
    """
    reset(page)
    click_cell(page, 1, 1)
    assert grid(page).locator("#addr").inner_text() == "B3"


def test_the_formula_bar_shows_the_selected_value(page):
    reset(page)
    click_cell(page, 1, 2)
    want = table(page)["rows"][1][2]
    assert grid(page).locator("#val").input_value() == want


def test_editing_in_the_formula_bar_changes_the_cell(page):
    reset(page)
    click_cell(page, 0, 3)
    box = grid(page).locator("#val")
    box.click()
    box.fill("수식줄에서")
    box.press("Enter")
    wait_dirty(page)
    assert table(page)["rows"][0][3] == "수식줄에서"


def test_find_counts_and_jumps(page):
    reset(page)
    click_cell(page, 0, 0)          # 격자에 초점을 준다 (Ctrl+F 가 거기로 가게)
    page.keyboard.press("Control+f")
    page.wait_for_timeout(500)
    assert "on" in (grid(page).locator("#find").get_attribute("class") or "")
    grid(page).locator("#findInput").fill("AA94")
    page.wait_for_timeout(700)
    assert grid(page).locator("#findHits").inner_text() == "1 / 3"
    assert grid(page).locator(".cell.hit-now").count() == 1
    grid(page).locator("#findNext").click()
    page.wait_for_timeout(400)
    assert grid(page).locator("#findHits").inner_text() == "2 / 3"
    # 찾은 자리로 선택이 따라간다
    assert grid(page).locator("#addr").inner_text() == "B3"


def test_find_says_so_when_there_is_nothing(page):
    reset(page)
    click_cell(page, 0, 0)
    page.keyboard.press("Control+f")
    page.wait_for_timeout(400)
    grid(page).locator("#findInput").fill("그런값없음")
    page.wait_for_timeout(600)
    assert grid(page).locator("#findHits").inner_text() == "없음"


def test_downloading_gives_the_saved_file_not_the_screen(page):
    """'저장 안 한 수정 미포함' 이다. 그래서 우리가 다시 만들지 않고 S3 에서
    받아 온 바이트를 그대로 내준다 -- 다시 만들면 저장된 판과 한 글자라도
    다를 수 있고, 내려받아 고쳐 올릴 사람에게는 그게 곧 사고다.
    """
    reset(page)
    with page.expect_download() as got:
        page.get_by_role("button", name="엑셀 다운로드", exact=True).click()
    assert got.value.suggested_filename == f"{BOOK}.xlsx"


def test_the_four_buttons_are_the_same_size_and_line_up(page):
    """고르개에만 이름이 붙으면 그 칸만 내려앉아 밑줄이 안 맞는다."""
    reset(page)
    boxes = [page.locator("div[data-testid='stSelectbox']").first.bounding_box()]
    for name in ("초기화", "엑셀 다운로드", "엑셀 업로드", "저장"):
        boxes.append(page.get_by_role("button", name=name, exact=True)
                     .first.bounding_box())
    ys = {round(b["y"]) for b in boxes}
    hs = {round(b["height"]) for b in boxes}
    ws = {round(b["width"]) for b in boxes}
    assert len(ys) == 1, f"높이가 안 맞습니다: {ys}"
    assert len(hs) == 1, f"키가 안 맞습니다: {hs}"
    assert len(ws) == 1, f"너비가 안 맞습니다: {ws}"


def open_menu(page, r, c):
    grid(page).locator(f".cell[data-r='{r}'][data-c='{c}']").click(button="right")
    page.wait_for_timeout(500)
    return grid(page).locator("#menu")


def test_right_click_opens_the_menu_not_the_browser_one(page):
    reset(page)
    menu = open_menu(page, 1, 1)
    assert "on" in (menu.get_attribute("class") or "")
    items = menu.locator(".item").all_inner_texts()
    assert any("복사" in i for i in items)
    assert any("위에 행 삽입" in i for i in items)
    assert any("열 삭제" in i for i in items)


def test_the_menu_closes_when_you_click_away(page):
    reset(page)
    open_menu(page, 1, 1)
    grid(page).locator(".cell[data-r='0'][data-c='0']").click()
    page.wait_for_timeout(400)
    assert "on" not in (grid(page).locator("#menu").get_attribute("class") or "")


def test_right_clicking_outside_the_selection_moves_it(page):
    """엑셀과 같게. 안 그러면 엉뚱한 자리에 행이 들어간다."""
    reset(page)
    grid(page).locator(".cell[data-r='0'][data-c='0']").click()
    page.wait_for_timeout(300)
    open_menu(page, 2, 3)
    assert grid(page).locator("#addr").inner_text() == "D4"


def test_right_clicking_inside_the_selection_keeps_it(page):
    reset(page)
    a = grid(page).locator(".cell[data-r='0'][data-c='0']")
    b = grid(page).locator(".cell[data-r='2'][data-c='2']")
    a.hover(); page.mouse.down(); b.hover(); page.mouse.up()
    page.wait_for_timeout(400)
    open_menu(page, 1, 1)
    assert grid(page).locator(".cell.sel").count() == 9, "범위 선택이 풀렸습니다"


def test_inserting_a_row_from_the_menu(page):
    reset(page)
    n = len(table(page)["rows"])
    menu = open_menu(page, 0, 0)
    menu.locator("[data-act='row-below']").click()
    wait_dirty(page)
    assert len(table(page)["rows"]) == n + 1


def test_inserting_a_column_from_the_menu(page):
    reset(page)
    n = len(table(page)["cols"])
    menu = open_menu(page, 0, 0)
    menu.locator("[data-act='col-right']").click()
    wait_dirty(page)
    assert len(table(page)["cols"]) == n + 1


def test_clearing_from_the_menu(page):
    reset(page)
    menu = open_menu(page, 0, 1)
    menu.locator("[data-act='clear']").click()
    wait_dirty(page)
    assert table(page)["rows"][0][1] == ""


def test_copying_from_the_menu(page):
    reset(page)
    want = table(page)["rows"][1][2]
    menu = open_menu(page, 1, 2)
    menu.locator("[data-act='copy']").click()
    page.wait_for_timeout(700)
    assert page.evaluate("navigator.clipboard.readText()") == want


def test_cutting_from_the_menu_copies_and_clears(page):
    reset(page)
    want = table(page)["rows"][1][3]
    menu = open_menu(page, 1, 3)
    menu.locator("[data-act='cut']").click()
    wait_dirty(page)
    assert page.evaluate("navigator.clipboard.readText()") == want
    assert table(page)["rows"][1][3] == ""


# ------------------------------------------------------- 많은 줄 버티기

def test_the_last_short_block_promises_its_own_height(page):
    """시트 끝 묶음은 50줄이 안 찬다. 그 묶음이 미리 알려 주는 높이는 제
    줄 수만큼이어야 한다 -- 1250px 로 두면 끝에 빈 자리가 생긴다."""
    reset(page)
    got = grid(page).locator(".tbl .blk").last.evaluate("""(b) => ({
      height: b.getBoundingClientRect().height,
      promised: getComputedStyle(b).containIntrinsicSize,
      rows: b.querySelectorAll('.row').length,
    })""")
    assert got["rows"] < 50, got
    assert got["height"] == got["rows"] * 25, got
    assert f"{got['rows'] * 25}px" in got["promised"], got


def test_undo_after_two_edits_in_one_row_goes_back_one_step_at_a_time(page):
    """되돌리기 판은 줄을 나눠 쓴다. 줄을 그 자리에서 고치면 지난 판까지 같이
    바뀌어서, 되돌려도 안 돌아가거나 두 걸음이 한꺼번에 사라진다."""
    reset(page)
    first = table(page)["rows"][0][:]
    click_cell(page, 0, 1)
    page.keyboard.type("하나")
    page.keyboard.press("Enter")
    page.wait_for_timeout(300)
    click_cell(page, 0, 2)
    page.keyboard.type("둘")
    page.keyboard.press("Enter")
    page.wait_for_timeout(300)
    row = table(page)["rows"][0]
    assert row[1] == "하나" and row[2] == "둘", row

    undo = grid(page).locator("button[data-act='undo']")
    undo.click()
    page.wait_for_timeout(400)
    row = table(page)["rows"][0]
    assert row[1] == "하나" and row[2] == first[2], row
    undo.click()
    page.wait_for_timeout(400)
    assert table(page)["rows"][0] == first

    grid(page).locator("button[data-act='redo']").click()
    page.wait_for_timeout(400)
    row = table(page)["rows"][0]
    assert row[1] == "하나" and row[2] == first[2], row


def test_the_grid_is_not_a_table_element(page):
    """<div> 로 그려야 화면 밖 줄의 배치를 건너뛸 수 있다.

    CSS 의 크기 가둠은 표의 행에는 적용되지 않게 정해져 있다. <table> 로
    되돌리면 content-visibility 가 있어도 소용이 없어지고, 15,000줄에서
    여는 데 걸리는 시간이 1.8초에서 8초대로 돌아간다.
    """
    assert grid(page).locator("table").count() == 0
    assert grid(page).locator(".tbl .row .cell").count() > 0


# ------------------------------------------------- 저장 전에 보여주는 창

def test_the_review_says_which_rows_changed_and_how(page):
    """고친 줄은 '수정', 새로 만든 줄은 '신규' 로 나와야 한다."""
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("고침")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page)
    body = review_text(page)
    assert "sheet : STEP" in body, body[:300]
    assert "수정" in body, body[:600]
    assert "고침" in body, f"바뀐 줄의 값이 안 보입니다: {body[:600]}"
    page.get_by_role("button", name="취소").click()
    page.wait_for_timeout(600)


def test_a_new_row_is_marked_new_not_edited(page):
    reset(page)
    click_cell(page, 0, 0)
    grid(page).locator("#bar [data-act='row-below']").click()
    page.wait_for_timeout(300)
    click_cell(page, 1, 0)
    page.keyboard.type("새줄")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page)
    body = review_text(page)
    assert "신규" in body, body[:600]
    assert "새줄" in body, body[:600]
    page.get_by_role("button", name="취소").click()
    page.wait_for_timeout(600)


def test_saving_needs_a_reason(page):
    """Remark 를 안 적으면 저장이 안 되고, 왜 안 되는지 말해 준다."""
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("사유없이")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page)
    page.get_by_role("button", name="저장", exact=True).last.click()
    page.wait_for_timeout(2000)
    assert "Remark 를 적어야" in review_text(page), review_text(page)
    assert "저장했습니다" not in page.inner_text("body"), "사유 없이 저장됐습니다"
    page.get_by_role("button", name="취소").click()
    page.wait_for_timeout(600)


def test_one_click_saves_right_after_typing_the_reason(page):
    """적자마자 누른다 -- 칸을 떠나 주기를 기다려 주는 사람은 없다.

    예전에는 그 누름이 '칸을 떠났다' 로 먼저 처리되고, 그 판에서 저장
    단추는 사유가 아직 빈 줄 알고 꺼져 있었다. 그래서 한 번 눌러서는
    저장이 안 됐다.
    """
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("한번에저장")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page, table=False)
    page.get_by_label("Remark — 사유").fill("한 번에")
    page.get_by_role("button", name="저장", exact=True).last.click()   # 한 번만
    page.wait_for_function(
        "() => document.body.innerText.includes('저장했습니다')", timeout=60000)
    settle(page)
    assert "한번에저장" in [v for row in table(page)["rows"] for v in row]


def test_esc_does_not_close_the_save_popup(page):
    """Esc 로 닫히면 적던 사유가 통째로 날아간다."""
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("esc확인")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page, table=False)
    page.get_by_label("Remark — 사유").fill("적던 중")
    page.keyboard.press("Escape")
    page.wait_for_timeout(1500)
    assert "변경내용" in page.inner_text("body"), "Esc 에 창이 닫혔습니다"
    assert page.get_by_label("Remark — 사유").input_value() == "적던 중"
    page.get_by_role("button", name="취소").click()     # 닫는 길은 그대로 있다
    page.wait_for_timeout(1200)
    assert "변경내용" not in page.inner_text("body"), "취소로도 안 닫힙니다"


def test_typing_the_reason_does_not_wake_python(page):
    """사유를 칠 때마다 파이썬이 돌면 15,000행짜리 격자를 그때마다 다시
    내려보내느라 창이 굼떠진다. 그래서 단추를 누를 때 한 번만 돈다.
    """
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("굼뜸확인")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page, table=False)
    before = grid(page).locator("body").evaluate("() => window.__rendered || 0")
    page.get_by_label("Remark — 사유").fill("사유를 적는다")
    page.keyboard.press("Tab")
    page.get_by_label("관련 — 세부 내용 (필수X)").fill("세부 내용도 적는다")
    page.keyboard.press("Tab")
    page.wait_for_timeout(2500)
    after = grid(page).locator("body").evaluate("() => window.__rendered || 0")
    assert after == before, f"칸 두 개 적는 동안 파이썬이 {after - before}번 돌았습니다"
    page.get_by_role("button", name="취소").click()
    page.wait_for_timeout(600)


def test_the_date_cannot_be_typed_over(page):
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("날짜확인")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page)
    box = page.get_by_label("Date")
    assert box.is_disabled(), "날짜 칸을 고칠 수 있으면 안 됩니다"
    import datetime as _dt
    today = _dt.datetime.now(_dt.timezone(_dt.timedelta(hours=9))).strftime("%Y-%m-%d")
    assert box.input_value() == today, box.input_value()
    page.get_by_role("button", name="취소").click()
    page.wait_for_timeout(600)


def test_the_reason_lands_in_the_rev_info_sheet(page):
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("기록남기기")
    page.keyboard.press("Enter")
    wait_dirty(page)
    confirm_save(page, remark="오탈자 고침")

    reset(page)
    page.wait_for_timeout(400)
    grid(page).locator(".sheetbar .tab", has_text="REV_INFO").click()
    page.wait_for_timeout(600)
    last = table(page)["rows"][-1]
    assert "오탈자 고침" in last, last
    import datetime as _dt
    today = _dt.datetime.now(_dt.timezone(_dt.timedelta(hours=9))).strftime("%Y-%m-%d")
    assert today in last, last


def test_the_save_button_sits_above_the_grid(page):
    """단추가 표 아래에 있으면 15,000행짜리 표에 밀려 화면 밖으로 나간다.

    실제로 배포하고 나서 '저장 버튼이 어디 있냐' 는 말을 들은 자리다.
    """
    reset(page)
    save = page.get_by_role("button", name="저장", exact=True).first
    frame = page.locator("iframe[title*='sheet_grid'], iframe[src*='sheet_grid']").first
    assert save.bounding_box()["y"] < frame.bounding_box()["y"], \
        "저장 단추가 표보다 아래에 있습니다"


def test_the_save_button_is_visible_without_scrolling(page):
    reset(page)
    save = page.get_by_role("button", name="저장", exact=True).first
    assert save.is_visible()
    box, view = save.bounding_box(), page.viewport_size
    assert box["y"] + box["height"] <= view["height"], \
        f"저장 단추가 첫 화면 밖입니다: y={box['y']}, 화면높이={view['height']}"


def test_editing_many_cells_reruns_python_only_once(page):
    """칸마다 파이썬을 깨우면 고칠 때마다 화면에 로딩이 번쩍인다.

    편집 중에 파이썬이 알아야 하는 것은 '고친 게 있다' 하나뿐이고, 그건
    한 번 참이 되면 저장하거나 초기화할 때까지 계속 참이다.
    """
    reset(page)
    for r, c in ((0, 2), (1, 2), (2, 2), (0, 3)):
        click_cell(page, r, c)
        page.keyboard.type(f"값{r}{c}")
        page.keyboard.press("Enter")
        page.wait_for_timeout(700)
    wait_dirty(page)
    sent = grid(page).locator("body").evaluate("() => window.__sent || 0")
    assert sent == 1, f"칸 4개 고치는데 {sent}번 올렸습니다"


def test_the_reset_button_is_called_초기화(page):
    reset(page)
    assert page.get_by_role("button", name="초기화").count() == 1


# --------------------------------------------- 줄이 많을 때 나눠 그리기

@pytest.fixture(scope="module")
def big_server():
    """줄이 3000개인 시트로 따로 띄운다 (기본 씨앗은 3줄이라 안 걸린다)."""
    port = _free_port()
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", str(ROOT / "app_local.py"),
         "--server.port", str(port), "--server.headless", "true",
         "--browser.gatherUsageStats", "false"],
        cwd=ROOT, env=dict(os.environ, IM_LOCAL_ROWS="3000"),
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(120):
        try:
            with socket.create_connection(("localhost", port), timeout=0.5):
                break
        except OSError:
            time.sleep(0.5)
    else:
        proc.terminate()
        pytest.fail("streamlit 이 안 떴습니다")
    time.sleep(3)
    yield f"http://localhost:{port}/"
    proc.terminate()
    proc.wait(timeout=20)


@pytest.fixture(scope="module")
def big_page(big_server, page):
    pg = page.context.new_page()
    pg.goto(big_server)
    pg.wait_for_timeout(6000)
    pg.get_by_role("combobox").click()
    pg.wait_for_timeout(600)
    pg.get_by_text(BOOK, exact=True).click()
    pg.wait_for_timeout(9000)
    yield pg
    pg.close()


def test_every_row_ends_up_drawn(big_page):
    """나눠 그리더라도 결국 한 줄도 빠지지 않아야 한다.

    보이는 만큼만 그리고 마는 것이 아니다 -- 그러면 브라우저 Ctrl+F 가
    값을 못 찾는다. 첫 화면을 먼저 내놓고 나머지를 이어 붙일 뿐이다.
    """
    g = big_page.frame_locator(
        "iframe[title*='sheet_grid'], iframe[src*='sheet_grid']").first
    assert g.locator(".tbl .row").count() == 3000
    assert g.locator(".tbl .filler").count() == 0, "다 그리고 나면 빈 자리는 없어야 한다"


def test_the_scrollbar_is_right_while_it_is_still_filling(big_page):
    """아직 안 그린 만큼을 자리로 잡아 두지 않으면, 줄이 붙을 때마다 막대가
    자라서 잡고 있던 자리가 밀린다."""
    g = big_page.frame_locator(
        "iframe[title*='sheet_grid'], iframe[src*='sheet_grid']").first
    got = g.locator(".tbl").evaluate("t => t.getBoundingClientRect().height")
    assert abs(got - 3001 * 25) <= 2, got      # 줄 3000 + 머리글 1


# ------------------------------------------------------- 엑셀 업로드

def make_xlsx(tmp_path, rows):
    import openpyxl
    b = openpyxl.Workbook(); ws = b.active; ws.title = "STEP"
    for i, v in enumerate(["step_seq", "step_id", "step_desc", "ppid", "사용"],
                          start=1):
        ws.cell(1, i, v)
    for r in rows:
        ws.append(r)
    path = tmp_path / "올릴것.xlsx"
    b.save(path)
    return str(path)


def upload(page, path):
    """엑셀 업로드 -> 파일 고르기 -> 화면에 넣기."""
    page.get_by_role("button", name="엑셀 업로드", exact=True).click()
    page.wait_for_function(
        "() => document.body.innerText.includes('올릴 엑셀 파일')", timeout=30000)
    page.locator("input[type='file']").set_input_files(path)
    page.wait_for_function(
        "() => document.body.innerText.includes('화면에 넣기')", timeout=30000)
    page.get_by_role("button", name="화면에 넣기").click()
    page.wait_for_function(
        "() => document.body.innerText.includes('올린 엑셀의 내용')", timeout=30000)
    page.wait_for_timeout(1500)


def test_an_uploaded_file_shows_up_in_the_grid(tmp_path, page):
    reset(page)
    upload(page, make_xlsx(tmp_path, [
        ["0010", "BB100001TR01", "올린값", "P-TTS-01", "Y"],
        ["0020", "BB100002TR01", "새줄", "P-TTS-02", "N"]]))
    got = table(page)
    assert len(got["rows"]) == 2, got
    assert "올린값" in got["rows"][0], got["rows"][0]


def test_uploading_does_not_save_by_itself(tmp_path, page):
    """올린다고 S3 가 바뀌면 안 된다. 저장을 눌러야 들어간다."""
    reset(page)
    upload(page, make_xlsx(tmp_path, [["0010", "x", "안저장됨", "p", "Y"]]))
    assert page.get_by_role("button", name="저장", exact=True).first.is_enabled()
    reset(page)                                   # 저장 안 하고 다시 불러오면
    assert "안저장됨" not in str(table(page)["rows"])


def test_saving_an_upload_records_what_changed_like_any_other_edit(tmp_path, page):
    reset(page)
    upload(page, make_xlsx(tmp_path, [
        ["0010", "BB100001TR01", "업로드로바꿈", "P-TTS-01", "Y"]]))
    open_review(page)
    body = review_text(page)
    assert "sheet : STEP" in body, body[:400]
    assert "업로드로바꿈" in body, body[:400]
    page.get_by_role("button", name="취소").click()
    page.wait_for_timeout(800)


def test_a_file_that_is_not_an_excel_is_refused(tmp_path, page):
    reset(page)
    bad = tmp_path / "아님.xlsx"
    bad.write_bytes(b"this is not a zip")
    page.get_by_role("button", name="엑셀 업로드", exact=True).click()
    page.wait_for_function(
        "() => document.body.innerText.includes('올릴 엑셀 파일')", timeout=30000)
    page.locator("input[type='file']").set_input_files(str(bad))
    page.wait_for_function(
        "() => document.body.innerText.includes('엑셀로 읽지 못했습니다')",
        timeout=30000)
    assert not page.get_by_role("button", name="화면에 넣기").count()
    page.get_by_role("button", name="닫기").click()
    page.wait_for_timeout(800)


def test_a_block_of_rows_is_as_tall_as_it_promises(big_page):
    """줄은 50줄씩 묶어 화면 밖 묶음의 배치를 미룬다. 미루는 동안에는 미리
    알려 준 높이(1250px)로 자리를 잡으므로, 그 값이 실제 높이와 다르면 스크롤
    막대 길이가 틀어져 끝까지 내렸는데 줄이 더 남아 있거나 빈 자리가 생긴다.
    CSS 의 padding 하나만 건드려도 어긋난다."""
    g = big_page.frame_locator(
        "iframe[title*='sheet_grid'], iframe[src*='sheet_grid']").first
    got = g.locator(".tbl .blk").first.evaluate("""(b) => ({
      height: b.getBoundingClientRect().height,
      promised: getComputedStyle(b).containIntrinsicSize,
      skipping: getComputedStyle(b).contentVisibility,
      rows: b.querySelectorAll('.row').length,
      rowHeight: b.querySelector('.row').getBoundingClientRect().height,
    })""")
    assert got["rowHeight"] == 25, got
    assert got["rows"] == 50 and got["height"] == 50 * 25, got
    assert "1250px" in got["promised"], got
    assert got["skipping"] == "auto", got


def test_typing_while_scrolled_away_brings_the_cell_back(big_page):
    """엑셀처럼, 고른 칸이 화면 밖일 때 글자를 치면 그 칸으로 데려간다.
    안 그러면 보이지 않는 칸에 글자가 들어간다."""
    g = big_page.frame_locator(
        "iframe[title*='sheet_grid'], iframe[src*='sheet_grid']").first
    g.locator(".cell[data-r='0'][data-c='1']").click()
    big_page.wait_for_timeout(200)
    g.locator("#scroll").evaluate("el => { el.scrollTop = 40000; }")
    big_page.wait_for_timeout(300)
    big_page.keyboard.type("x")
    big_page.wait_for_timeout(300)
    top = g.locator("#scroll").evaluate("el => el.scrollTop")
    big_page.keyboard.press("Escape")               # 고친 것은 버린다
    big_page.wait_for_timeout(300)
    assert top < 1000, f"고치는 칸(맨 위)이 화면 밖에 있습니다: scrollTop={top}"


def test_saving_sends_only_the_sheet_you_touched(big_server, page):
    """3,000줄짜리 STEP 은 그대로 두고 작은 시트 하나만 고쳤으면, 저장할 때
    오가는 것도 그 한 장이어야 한다. 예전에는 저장을 누를 때마다 파일 전체가
    브라우저로 한 번 내려오고(누른 표시가 바뀌었다고) 또 전체가 올라갔다."""
    pg = page.context.new_page()
    sent, got = [], []
    pg.on("websocket", lambda ws: (
        ws.on("framesent", lambda p: sent.append(len(p))),
        ws.on("framereceived", lambda p: got.append(len(p)))))
    try:
        pg.goto(big_server)
        pg.wait_for_timeout(5000)
        pg.get_by_role("combobox").click()
        pg.wait_for_timeout(600)
        pg.get_by_text(BOOK, exact=True).click()
        pg.wait_for_function(
            "() => document.body.innerText.includes('고친 것 없음')", timeout=60000)
        pg.wait_for_timeout(3000)
        assert max(got) > 60_000, "처음에는 표가 통째로 내려와야 한다"

        g = pg.frame_locator("iframe[title*='sheet_grid'], iframe[src*='sheet_grid']").first
        g.locator("#sheetbar .tab", has_text="ITEM").click()
        pg.wait_for_timeout(500)
        g.locator(".cell[data-r='0'][data-c='1']").click()
        pg.keyboard.type("작은시트만")
        pg.keyboard.press("Enter")
        pg.wait_for_function(
            "() => document.body.innerText.includes('저장하지 않은 수정')", timeout=30000)
        pg.wait_for_timeout(1000)

        sent.clear()
        got.clear()
        pg.get_by_role("button", name="저장", exact=True).first.click()
        pg.wait_for_selector("[role='dialog'] [data-testid='stTable']", timeout=60000)
        pg.wait_for_timeout(1000)
        assert "sheet : ITEM" in pg.locator("[role='dialog']").inner_text()
        assert max(sent) < 20_000, f"안 고친 시트까지 올라갔습니다: {max(sent)} bytes"
        assert max(got) < 60_000, f"표가 다시 내려왔습니다: {max(got)} bytes"
        pg.get_by_role("button", name="취소").click()
        pg.wait_for_timeout(600)
    finally:
        pg.close()
