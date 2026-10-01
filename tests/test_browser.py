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
        cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        # 임시 저장은 끈다: 앞 검사가 고쳐 둔 채로 다음 검사가 페이지를 새로 열면
        # '복구하시겠습니까?' 창이 떠서 격자를 가린다. 임시 저장은 backup_server 에서 본다.
        env=dict(os.environ, IM_LOCAL_BACKUP="0"))
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
    """저장 -> 사유 적기 -> 진짜 저장. '저장 완료!' 창이 뜨면 닫는다.

    적자마자 바로 누른다. 사람이 그렇게 하기 때문이다 -- 칸을 떠나라고
    Tab 을 눌러 주거나 한 박자 기다려 주지 않는다.

    user 는 로그인한 사람으로 고정이라 못 고친다.
    """
    open_review(page, table=False)
    page.get_by_label("Remark — 사유").fill(remark)
    page.get_by_role("button", name="저장", exact=True).last.click()
    wait_saved(page)


def wait_saved(page):
    """'저장 완료!' 창이 뜰 때까지 기다렸다가 확인을 눌러 닫는다.

    창을 닫아 두어야 뒤이은 검사가 격자를 누를 수 있다 (창이 가린다).
    """
    page.wait_for_function(
        "() => document.body.innerText.includes('저장 완료!')", timeout=60000)
    page.get_by_role("button", name="확인").click()
    page.wait_for_function(
        "() => !document.body.innerText.includes('저장 완료!')", timeout=30000)


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


def _forget_edits(page):
    """격자가 '고친 것이 있다' 고 들고 있는 것을 지운다.

    고친 채로 페이지를 옮기면 격자가 '벗어나겠습니까?' 를 묻는다(브라우저의
    beforeunload). 검사는 그 창에 답하지 않으므로 페이지가 안 넘어간다.
    """
    page.evaluate("""() => { for (const f of document.querySelectorAll('iframe')) {
      try { f.contentWindow.eval('try { dirty = false; unsavedFromPython = false } catch (e) {}') }
      catch (e) {} } }""")


def reset(page):
    """페이지를 새로 열고 검사할 파일을 고른다.

    한 페이지를 여러 검사가 나눠 쓰면 앞 검사가 남긴 것(고르던 시트, 저장
    안 한 수정, 늘어난 칸)을 다음 검사가 그대로 물고 시작한다. 실제로 그래서
    따로 돌리면 통과하고 같이 돌리면 깨지는 검사가 여럿 나왔다. 새로 열면
    streamlit 세션이 새로 생겨 session_state 까지 깨끗해진다.
    """
    _forget_edits(page)
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
    below = table(page)["rows"][1][3]
    box.fill("수식줄에서")
    box.press("Enter")
    wait_dirty(page)
    got = table(page)["rows"]
    assert got[0][3] == "수식줄에서"
    # Enter 로 내려간 뒤 입력줄에서 초점이 빠질 때(blur) 같은 글자를 아래
    # 칸에 한 번 더 써 넣었다
    assert got[1][3] == below, "아래 칸까지 덮어썼습니다"


# ------------------------------------------- 한글 입력 (IME) 과 Enter
#
# 한글은 마지막 글자가 '조합 중' 인 채로 Enter 가 눌린다. 그때 바로 칸을
# 닫고 내려가면, IME 가 그 글자를 새로 초점이 간 아래 칸에 확정해 넣어
# 아래 칸 값이 덮였다. 크롬의 IME 흉내(CDP)로 실제 순서를 그대로 보낸다:
# 조합 -> Enter(keyCode 229, 조합 중) -> 확정 [-> 브라우저에 따라 Enter 한 번 더].

def _ime_type_then_enter(page, echo):
    cdp = page.context.new_cdp_session(page)

    def key(vk, kind="rawKeyDown"):
        cdp.send("Input.dispatchKeyEvent", {"type": kind, "key": "Enter", "code": "Enter",
                                            "windowsVirtualKeyCode": vk,
                                            "nativeVirtualKeyCode": vk})
    for t in ["ㅎ", "하", "한"]:
        cdp.send("Input.imeSetComposition", {"text": t, "selectionStart": 1, "selectionEnd": 1})
    cdp.send("Input.insertText", {"text": "한"})
    for t in ["ㄱ", "그", "글"]:
        cdp.send("Input.imeSetComposition", {"text": t, "selectionStart": 1, "selectionEnd": 1})
    page.wait_for_timeout(100)
    key(229)                                   # 조합 중에 누른 Enter
    cdp.send("Input.insertText", {"text": "글"})  # IME 가 마지막 글자를 확정
    if echo:
        key(13)                                # 일부 브라우저: Enter 가 한 번 더
    key(13, "keyUp")
    cdp.detach()
    page.wait_for_timeout(600)


@pytest.mark.parametrize("echo", [False, True], ids=["enter-once", "enter-twice"])
def test_korean_typed_in_a_cell_then_enter_stays_in_that_cell(page, echo):
    reset(page)
    before = table(page)["rows"]
    click_cell(page, 0, 2)
    page.keyboard.press("Enter")               # 칸 고치기 시작
    grid(page).locator(".cell.editing").evaluate("""(td) => {
      const r = document.createRange(); r.selectNodeContents(td);
      const s = getSelection(); s.removeAllRanges(); s.addRange(r); }""")
    _ime_type_then_enter(page, echo)
    wait_dirty(page)
    got = table(page)["rows"]
    assert got[0][2] == "한글"
    assert got[1][2] == before[1][2], "아래 칸에 글자가 들어갔습니다"
    assert got[2][2] == before[2][2]
    # 한 칸만 내려가 있고, 아래 칸을 고치는 중이 아니다
    assert grid(page).locator(".cell.anchor").get_attribute("data-r") == "1"
    assert grid(page).locator(".cell.editing").count() == 0


@pytest.mark.parametrize("echo", [False, True], ids=["enter-once", "enter-twice"])
def test_korean_typed_in_the_formula_bar_then_enter_stays_in_that_cell(page, echo):
    reset(page)
    before = table(page)["rows"]
    click_cell(page, 0, 2)
    box = grid(page).locator("#val")
    box.click()
    box.evaluate("el => el.select()")
    _ime_type_then_enter(page, echo)
    wait_dirty(page)
    got = table(page)["rows"]
    assert got[0][2] == "한글"
    assert got[1][2] == before[1][2], "아래 칸에 글자가 들어갔습니다"
    assert grid(page).locator(".cell.anchor").get_attribute("data-r") == "1"
    assert grid(page).locator(".cell.editing").count() == 0


# ------------------------------------------- 검색칸에서 친 키는 검색칸의 것

def _open_find(page):
    click_cell(page, 0, 0)
    page.keyboard.press("Control+f")
    page.wait_for_timeout(400)
    return grid(page).locator("#findInput")


def test_backspace_in_the_find_box_erases_the_search_not_the_cell(page):
    reset(page)
    before = table(page)["rows"]
    box = _open_find(page)
    page.keyboard.type("AA94")
    page.wait_for_timeout(500)
    page.keyboard.press("Backspace")
    page.keyboard.press("Backspace")
    page.keyboard.press("Delete")
    page.wait_for_timeout(500)
    assert box.input_value() == "AA"
    assert table(page)["rows"] == before, "검색칸에서 지웠는데 칸이 지워졌습니다"


def test_enter_in_the_find_box_moves_to_the_next_hit_without_editing(page):
    reset(page)
    _open_find(page)
    page.keyboard.type("AA94")
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(300)
    assert grid(page).locator("#findHits").inner_text() == "2 / 3"
    assert grid(page).locator(".cell.editing").count() == 0, "칸 고치기가 시작됐습니다"
    page.keyboard.type("x")
    page.wait_for_timeout(300)
    assert grid(page).locator("#findInput").input_value() == "AA94x"


def test_pasting_into_the_find_box_does_not_paste_into_the_grid(page):
    reset(page)
    before = table(page)["rows"]
    box = _open_find(page)
    paste(page, "AA94")
    assert box.input_value() == "AA94"
    assert table(page)["rows"] == before, "검색칸에 붙여넣었는데 격자에 들어갔습니다"


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
    assert "저장 완료!" not in page.inner_text("body"), "사유 없이 저장됐습니다"
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
    wait_saved(page)
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
        cwd=ROOT, env=dict(os.environ, IM_LOCAL_ROWS="3000", IM_LOCAL_BACKUP="0"),
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


# ---------------------------------------------- 저장하는 동안 단추를 끈다

@pytest.fixture(scope="module")
def slow_server():
    """저장에 3초가 걸리는 서버. 가짜 저장소는 순식간이라 '저장하는 동안'
    이 없어서, 그동안 단추가 어떻게 보이는지 볼 수가 없다."""
    port = _free_port()
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", str(ROOT / "app_local.py"),
         "--server.port", str(port), "--server.headless", "true",
         "--browser.gatherUsageStats", "false"],
        cwd=ROOT, env=dict(os.environ, IM_LOCAL_SLOW_SAVE="3", IM_LOCAL_BACKUP="0"),
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


def _open_on(page, url):
    """다른 서버에 새 탭을 열고 검사할 파일을 골라 한 칸을 고쳐 둔다."""
    import re
    pg = page.context.new_page()
    pg.goto(url)
    pg.wait_for_timeout(5000)
    pg.get_by_role("combobox").click()
    pg.wait_for_timeout(600)
    pg.get_by_text(BOOK, exact=True).click()
    settle(pg)
    shown = re.search(r"REV_INFO · (\d+)줄", pg.inner_text("body"))
    click_cell(pg, 0, 2)
    pg.keyboard.type(f"느린저장{time.time():.0f}")
    pg.keyboard.press("Enter")
    wait_dirty(pg)
    open_review(pg, table=False)
    pg.get_by_label("Remark — 사유").fill("단추 끄기")
    return pg, int(shown.group(1))


def _rev_lines(pg):
    import re
    pg.wait_for_function("() => /REV_INFO · \\d+줄/.test(document.body.innerText)",
                         timeout=30000)
    return int(re.search(r"REV_INFO · (\d+)줄", pg.inner_text("body")).group(1))


def test_the_save_button_greys_out_the_moment_you_press_it(slow_server, page):
    """저장이 도는 몇 초 동안 단추가 켜져 있으면 한 번 더 누르게 된다."""
    pg, _before = _open_on(page, slow_server)
    try:
        pg.get_by_role("button", name="저장", exact=True).last.click()
        busy = pg.get_by_role("button", name="저장 중…")
        busy.wait_for(timeout=2500)              # 저장(3초)이 끝나기 전에
        assert busy.is_disabled(), "저장하는 동안 단추가 켜져 있습니다"
        assert pg.get_by_role("button", name="취소").is_disabled()
        assert pg.get_by_label("Remark — 사유").is_disabled()
        wait_saved(pg)
    finally:
        pg.close()


def test_pressing_save_twice_quickly_saves_once(slow_server, page):
    """두 번째 누름이 단추가 꺼지기 전에 들어가도 저장은 한 번이어야 한다.
    두 번 저장되면 REV_INFO 에 같은 줄이 두 번 쌓이고 이력 사본도 두 벌이다."""
    pg, before = _open_on(page, slow_server)
    try:
        pg.get_by_role("button", name="저장", exact=True).last.dblclick()
        wait_saved(pg)
        pg.wait_for_timeout(4000)                # 두 번째 누름이 뭔가 한다면 이 사이에
        assert _rev_lines(pg) == before + 1
    finally:
        pg.close()


# ------------------------------------------------------- 저장 완료 창

def test_saving_says_save_complete(page):
    """저장이 끝나면 '저장 완료!' 창이 뜬다."""
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("완료창")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page, table=False)
    page.get_by_label("Remark — 사유").fill("완료")
    page.get_by_role("button", name="저장", exact=True).last.click()
    page.wait_for_function(
        "() => document.body.innerText.includes('저장 완료!')", timeout=60000)
    box = page.locator("[role='dialog']").inner_text()
    assert "저장 완료!" in box and BOOK in box, box
    assert "raw data" not in box, "반영 안내는 뺐다"
    page.get_by_role("button", name="확인").click()
    page.wait_for_function(
        "() => !document.body.innerText.includes('저장 완료!')", timeout=30000)


def test_the_save_complete_popup_stays_until_closed_and_then_stays_closed(page):
    """띄우면서 치우면 격자가 새 판을 받았다고 알려 오는 다음 판에서 창이
    저절로 닫힌다 (눈 깜짝할 새 번쩍였다 사라진다). 닫은 뒤에는 칸을
    고쳐서 다시 그려도 되살아나면 안 된다."""
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.type("창유지")
    page.keyboard.press("Enter")
    wait_dirty(page)
    open_review(page, table=False)
    page.get_by_label("Remark — 사유").fill("창")
    page.get_by_role("button", name="저장", exact=True).last.click()
    page.wait_for_function(
        "() => document.body.innerText.includes('저장 완료!')", timeout=60000)
    page.wait_for_timeout(4000)
    assert "저장 완료!" in page.inner_text("body"), "창이 저절로 닫혔습니다"

    page.keyboard.press("Escape")                 # X 로 닫는 것과 같다
    page.wait_for_function(
        "() => !document.body.innerText.includes('저장 완료!')", timeout=30000)
    settle(page)
    click_cell(page, 1, 2)
    page.keyboard.type("다시그림")
    page.keyboard.press("Enter")
    wait_dirty(page)
    page.wait_for_timeout(1500)
    assert "저장 완료!" not in page.inner_text("body"), "닫은 창이 되살아났습니다"


# ------------------------------------------- REV_INFO 는 화면에서 못 고친다

def open_rev_info(page):
    grid(page).locator(".sheetbar .tab", has_text="REV_INFO").click()
    page.wait_for_timeout(600)


def test_rev_info_cannot_be_edited_in_the_grid(page):
    """REV_INFO 는 저장할 때 자동으로 한 줄씩 적히는 기록이다. 칸 고치기,
    지우기, 붙여넣기, 줄 넣기가 전부 먹히지 않아야 하고, 그래서 '고친 것'
    도 생기지 않아야 한다."""
    reset(page)
    open_rev_info(page)
    assert grid(page).locator("#locknote").is_visible(), "잠겼다는 안내가 없습니다"
    before = table(page)

    click_cell(page, 0, 1)
    page.keyboard.type("몰래고침")
    page.keyboard.press("Enter")
    page.wait_for_timeout(300)
    click_cell(page, 0, 1)
    page.keyboard.press("Delete")
    paste(page, "X\tY")
    grid(page).locator("#bar [data-act='row-above']").click()
    page.wait_for_timeout(600)

    assert table(page) == before, "잠긴 시트가 바뀌었습니다"
    assert grid(page).locator("#val").evaluate("el => el.readOnly"), \
        "수식 입력줄로 고칠 수 있습니다"
    page.wait_for_timeout(800)
    assert "저장하지 않은 수정" not in page.inner_text("body")


def test_rev_info_can_still_be_selected_and_copied(page):
    reset(page)
    open_rev_info(page)
    want = table(page)["rows"][0][1]
    click_cell(page, 0, 1)
    page.keyboard.press("Control+c")
    page.wait_for_timeout(400)
    assert page.evaluate("() => navigator.clipboard.readText()") == want


def test_rev_info_tab_cannot_be_renamed(page):
    """이름이 바뀌면 저장할 때 기록을 붙여 적을 시트를 못 찾는다."""
    reset(page)
    # 먼저 그 탭을 연다. 안 연 탭을 두 번 누르면 첫 누름에 탭 줄이 새로
    # 그려져서, 두 번째 누름이 다른 요소에 떨어져 두 번 누른 것으로 안 잡힐
    # 때가 있다 (검사가 가끔 깨지던 까닭).
    open_rev_info(page)
    said = []
    # 알림창은 받자마자 닫는다. 안 닫으면 두 번 누르기가 그 창이 닫히기를
    # 기다리며 멈춘다.
    page.once("dialog", lambda d: (said.append(d.message), d.accept()))
    grid(page).locator(".sheetbar .tab", has_text="REV_INFO").dblclick()
    for _ in range(40):
        if said:
            break
        page.wait_for_timeout(100)
    assert said and "이름을 바꿀 수 없습니다" in said[0], said
    assert grid(page).locator(".sheetbar .tab", has_text="REV_INFO").count() == 1


def test_other_sheets_stay_editable_after_looking_at_rev_info(page):
    reset(page)
    open_rev_info(page)
    grid(page).locator(".sheetbar .tab", has_text="STEP").click()
    page.wait_for_timeout(600)
    assert not grid(page).locator("#locknote").is_visible()
    click_cell(page, 0, 2)
    page.keyboard.type("다시고침")
    page.keyboard.press("Enter")
    wait_dirty(page)
    assert table(page)["rows"][0][2] == "다시고침"


def test_an_uploaded_rev_info_is_ignored_and_the_log_grows_by_one(tmp_path, page):
    """고친 REV_INFO 가 든 엑셀을 올려도 기록은 S3 것 그대로에 한 줄만 붙는다."""
    import openpyxl
    reset(page)
    open_rev_info(page)
    stored = table(page)["rows"]

    b = openpyxl.Workbook()
    ws = b.active
    ws.title = "STEP"
    ws.append(["step_seq", "step_id", "step_desc", "ppid", "사용"])
    ws.append(["0010", "AA941234TR01", "업로드기록시험", "P-ULY-01", "Y"])
    rev = b.create_sheet("REV_INFO")
    rev.append(["Date", "Remark", "user", "관련"])
    rev.append(["1999-01-01", "위조된 기록", "누군가", ""])
    path = tmp_path / "기록위조.xlsx"
    b.save(path)

    reset(page)
    page.get_by_role("button", name="엑셀 업로드", exact=True).click()
    page.wait_for_function(
        "() => document.body.innerText.includes('올릴 엑셀 파일')", timeout=30000)
    page.locator("input[type='file']").set_input_files(str(path))
    page.wait_for_function(
        "() => document.body.innerText.includes('화면에 넣기')", timeout=30000)
    box = page.locator("[role='dialog']").inner_text()
    assert "올린 파일의 것을 쓰지 않습니다" in box, box
    page.get_by_role("button", name="화면에 넣기").click()
    page.wait_for_function(
        "() => document.body.innerText.includes('올린 엑셀의 내용')", timeout=30000)
    page.wait_for_timeout(1500)
    confirm_save(page, remark="업로드 뒤 기록")

    reset(page)
    open_rev_info(page)
    now = table(page)["rows"]
    assert now[:len(stored)] == stored, "S3 의 기록이 바뀌었습니다"
    assert len(now) == len(stored) + 1, now
    assert "업로드 뒤 기록" in now[-1]
    assert not any("위조된 기록" in r for r in now)


# ------------------------------------------- 엑셀과 복사·붙여넣기 (여러 줄 칸)

def test_pasting_a_multi_line_cell_from_excel_keeps_it_in_one_cell(page):
    """엑셀은 Alt+Enter 로 줄을 바꾼 칸을 "..." 로 감싸 클립보드에 넣는다.
    줄바꿈마다 자르면 그 칸이 두 줄로 쪼개져 아래 줄이 전부 밀린다."""
    reset(page)
    click_cell(page, 0, 0)
    paste(page, 'A1\t"첫줄\n둘째줄"\r\nB1\tB2\r\n')
    got = table(page)["rows"]
    assert got[0][0] == "A1" and got[0][1] == "첫줄\n둘째줄", got[:2]
    assert got[1][0] == "B1" and got[1][1] == "B2", got[:2]


def test_copying_a_multi_line_cell_quotes_it_like_excel(page):
    reset(page)
    click_cell(page, 0, 1)
    paste(page, '"한\n칸 ""따옴표"""')
    click_cell(page, 0, 1)
    page.keyboard.press("Control+c")
    page.wait_for_timeout(400)
    assert page.evaluate("() => navigator.clipboard.readText()") == \
        '"한\n칸 ""따옴표"""'


def test_the_file_picker_is_locked_while_there_are_unsaved_edits(page):
    """파일을 바꾸면 고친 것을 들고 가지 않는다. 한 번 잘못 누른 것으로
    말없이 사라지지 않게, 저장하거나 초기화하기 전에는 못 바꾼다."""
    reset(page)
    assert page.get_by_role("combobox").is_enabled()
    click_cell(page, 0, 2)
    page.keyboard.type("잠금")
    page.keyboard.press("Enter")
    wait_dirty(page)
    page.wait_for_timeout(800)
    assert page.get_by_role("combobox").is_disabled(), "고친 게 있는데 파일을 바꿀 수 있다"
    page.get_by_role("button", name="초기화").click()
    settle(page)
    page.wait_for_timeout(800)
    assert page.get_by_role("combobox").is_enabled(), "초기화한 뒤에도 잠겨 있다"


# ------------------------------------------------------------- 수정 잠금
#
# 누가 고치기 시작해서 저장 단추가 켜진 동안에는 다른 사람이 그 파일을 못
# 고친다. 다른 사람에게는 "hong님이 수정중입니다." 창이 뜬다. 저장하거나
# 초기화하면 풀리고, 기다리던 사람 화면은 새 판을 읽어 온다.

@pytest.fixture(scope="module")
def lock_server():
    """잠금 heartbeat 를 2초로 줄인 서버. 1분을 기다릴 수는 없다."""
    port = _free_port()
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", str(ROOT / "app_local.py"),
         "--server.port", str(port), "--server.headless", "true",
         "--browser.gatherUsageStats", "false"],
        cwd=ROOT, env=dict(os.environ, IM_LOCAL_BEAT="2", IM_LOCAL_BACKUP="0"),
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


def _open_as(page, url, user, menu=""):
    """그 사람으로 새 탭을 열고 검사할 파일을 고른다 (탭마다 세션이 따로다)."""
    pg = page.context.new_page()
    pg.goto(f"{url}?user={user}" + (f"&menu={menu}" if menu else ""))
    pg.wait_for_timeout(4000)
    pg.get_by_role("combobox").click()
    pg.wait_for_timeout(600)
    pg.get_by_text(BOOK, exact=True).click()
    grid(pg).locator(".cell[data-r='0'][data-c='0']").wait_for(timeout=30000)
    pg.wait_for_timeout(800)
    return pg


def _says(pg, text, timeout=30000):
    pg.wait_for_function(f"() => document.body.innerText.includes({text!r})", timeout=timeout)


def _close_lock_popup(pg):
    if pg.locator("[role='dialog']").count():
        pg.get_by_role("button", name="확인").click()
        pg.wait_for_timeout(800)


def _edit(pg, r, c, text):
    click_cell(pg, r, c)
    pg.keyboard.type(text)
    pg.keyboard.press("Enter")


def test_while_someone_edits_others_are_told_and_cannot_edit(lock_server, page):
    hong = _open_as(page, lock_server, "hong")
    kim = None
    try:
        settle(hong)
        _edit(hong, 0, 2, "잠금시험")
        wait_dirty(hong)

        kim = _open_as(page, lock_server, "kim")
        _says(kim, "hong님이 수정중입니다.")
        box = kim.locator("[role='dialog']")
        assert "hong님이 수정중입니다." in box.inner_text()
        kim.get_by_role("button", name="확인").click()
        kim.wait_for_function("() => !document.querySelector(\"[role='dialog']\")", timeout=15000)
        kim.wait_for_timeout(3000)              # heartbeat 가 몇 번 돌아도
        assert kim.locator("[role='dialog']").count() == 0, "닫은 창이 다시 떴습니다"

        before = table(kim)["rows"]
        _edit(kim, 0, 3, "몰래고침")
        kim.wait_for_timeout(800)
        assert table(kim)["rows"] == before, "남이 고치는 중인 파일을 고쳤습니다"
        assert kim.get_by_role("button", name="저장", exact=True).first.is_disabled()
        assert "hong님이 수정중입니다" in grid(kim).locator("#locknote").inner_text()

        # hong 의 화면은 잠기지 않는다 (제 잠금이다)
        assert grid(hong).locator("#locknote").is_hidden()

        # hong 이 저장하면 풀리고, kim 은 hong 이 저장한 판을 받는다
        confirm_save(hong)
        _says(kim, "hong님이 저장한 최신 내용을 불러왔습니다.", timeout=30000)
        kim.wait_for_timeout(1500)
        assert table(kim)["rows"][0][2] == "잠금시험"
        _edit(kim, 0, 3, "이제고침")
        wait_dirty(kim)
        assert table(kim)["rows"][0][3] == "이제고침"

        # 이번엔 kim 이 고치는 중이다 -- hong 쪽이 잠긴다
        _says(hong, "kim님이 수정중입니다.", timeout=30000)
    finally:
        hong.close()
        if kim:
            kim.close()


def test_resetting_releases_the_lock(lock_server, page):
    hong = _open_as(page, lock_server, "hong")
    kim = None
    try:
        # 앞 검사가 남긴 잠금(탭을 닫아 버린 사람 것)은 그 사람이 다시 들어와
        # 초기화를 누르면 풀린다 -- 몇 분을 기다릴 것 없이
        kim = _open_as(page, lock_server, "kim")
        for pg in (hong, kim):
            _close_lock_popup(pg)
            pg.get_by_role("button", name="초기화").click()
            pg.wait_for_timeout(1500)
        for pg in (hong, kim):
            _close_lock_popup(pg)
            pg.get_by_role("button", name="초기화").click()
            settle(pg)

        _edit(hong, 1, 2, "초기화시험")
        wait_dirty(hong)
        _says(kim, "hong님이 수정중입니다.", timeout=30000)
        _close_lock_popup(kim)      # 창이 떠 있는 동안은 화면이 쉰다

        hong.get_by_role("button", name="초기화").click()
        settle(hong)
        kim.wait_for_function(
            "() => !document.body.innerText.includes('hong님이 수정중입니다')", timeout=30000)
        assert grid(kim).locator("#locknote").is_hidden()
    finally:
        hong.close()
        if kim:
            kim.close()


# ---------------------------------------- 고친 채로 다른 메뉴로 가려 하면 묻는다

LEAVE_MSG = "페이지를 벗어나면 수정 사항이 사라집니다. 벗어나겠습니까?"


def _open_with_menu(page):
    _forget_edits(page)
    page.goto(page.url.split("?")[0] + "?menu=1")
    page.wait_for_timeout(4000)
    page.get_by_role("combobox").click()
    page.wait_for_timeout(600)
    page.get_by_text(BOOK, exact=True).click()
    settle(page)


def _menu_frame(page):
    return page.frame_locator("[data-testid='stSidebar'] iframe").first


def test_leaving_with_unsaved_edits_asks_first_and_stays_if_no(page):
    _open_with_menu(page)
    _edit(page, 0, 2, "떠나기전")
    wait_dirty(page)

    asked = []
    page.once("dialog", lambda d: (asked.append(d.message), d.dismiss()))
    page.get_by_text("Home", exact=True).click()
    page.wait_for_timeout(2500)
    assert asked == [LEAVE_MSG]
    assert "홈 화면" not in page.inner_text("body"), "아니오를 눌렀는데 벗어났습니다"
    assert table(page)["rows"][0][2] == "떠나기전", "고친 것이 사라졌습니다"

    # 메뉴 컴포넌트(따로 뜨는 틀) 안의 메뉴도 같다
    asked.clear()
    page.once("dialog", lambda d: (asked.append(d.message), d.dismiss()))
    _menu_frame(page).locator("#menu-link").click()
    page.wait_for_timeout(800)
    assert asked == [LEAVE_MSG]
    assert _menu_frame(page).locator("body").get_attribute("data-clicked") is None, \
        "아니오를 눌렀는데 메뉴가 눌렸습니다"

    # 예를 누르면 간다
    page.once("dialog", lambda d: d.accept())
    page.get_by_text("Home", exact=True).click()
    page.wait_for_function("() => document.body.innerText.includes('홈 화면')", timeout=30000)


def test_leaving_without_edits_does_not_ask(page):
    _open_with_menu(page)
    asked = []
    handler = lambda d: (asked.append(d.message), d.dismiss())
    page.on("dialog", handler)
    try:
        page.get_by_text("Home", exact=True).click()
        page.wait_for_function("() => document.body.innerText.includes('홈 화면')", timeout=30000)
        assert asked == []
    finally:
        page.remove_listener("dialog", handler)


def test_clicks_on_the_page_itself_do_not_ask(page):
    """저장 단추나 변경 이력 같은 이 화면 안의 것은 벗어나는 것이 아니다."""
    _open_with_menu(page)
    click_cell(page, 0, 0)
    paste(page, "a\nb")                          # 줄이 둘은 있게
    _edit(page, 0, 2, "안묻기")
    wait_dirty(page)
    asked = []
    handler = lambda d: (asked.append(d.message), d.dismiss())
    page.on("dialog", handler)
    try:
        page.get_by_text("변경 이력").click()
        page.wait_for_timeout(800)
        click_cell(page, 1, 1)
        assert asked == []
    finally:
        page.remove_listener("dialog", handler)


def test_closing_the_tab_with_unsaved_edits_asks(page):
    """창을 닫거나 새로 고칠 때는 브라우저가 제 문구로 묻는다."""
    pg = page.context.new_page()
    try:
        pg.goto(page.url.split("?")[0])
        pg.wait_for_timeout(4000)
        pg.get_by_role("combobox").click()
        pg.wait_for_timeout(600)
        pg.get_by_text(BOOK, exact=True).click()
        settle(pg)
        _edit(pg, 0, 2, "닫기전")
        wait_dirty(pg)
        kinds = []
        pg.on("dialog", lambda d: (kinds.append(d.type), d.dismiss()))
        pg.close(run_before_unload=True)
        pg.wait_for_timeout(1500)
        assert kinds == ["beforeunload"]
        assert not pg.is_closed(), "아니오를 눌렀는데 닫혔습니다"
    finally:
        if not pg.is_closed():
            _forget_edits(pg)
            pg.close()


# ---------------------------------------------- 복사·붙여넣기 모양과 글자만

def _three_rows(page):
    """앞 검사들이 저장해 둔 것과 상관없이 3줄 x 5칸을 깔고 시작한다.

    검사 서버의 저장소는 검사들이 나눠 쓴다. 앞에서 1줄짜리로 저장해 두면 줄
    수를 기대한 검사가 엉뚱하게 깨진다.
    """
    reset(page)
    click_cell(page, 0, 0)
    paste(page, "\n".join("\t".join(f"{c}{r}" for c in "abcde") for r in (1, 2, 3)))
    wait_dirty(page)
    return table(page)["rows"]


def _drag(page, r0, c0, r1, c1):
    a = grid(page).locator(f".cell[data-r='{r0}'][data-c='{c0}']")
    b = grid(page).locator(f".cell[data-r='{r1}'][data-c='{c1}']")
    a.hover(); page.mouse.down(); b.hover(); page.mouse.up()
    page.wait_for_timeout(200)


@pytest.mark.parametrize("shape", ["세로", "가로"])
def test_a_copied_range_pastes_back_in_the_same_shape(page, shape):
    """세로로 복사한 것은 세로로, 가로로 복사한 것은 가로로."""
    before = _three_rows(page)
    if shape == "세로":
        _drag(page, 0, 0, 1, 0)
        want = [before[0][0], before[1][0]]
        target = (0, 3)
    else:
        _drag(page, 0, 0, 0, 1)
        want = before[0][:2]
        target = (2, 2)
    page.keyboard.press("Control+c")
    page.wait_for_timeout(300)
    click_cell(page, *target)
    page.keyboard.press("Control+v")
    page.wait_for_timeout(1200)
    got = table(page)["rows"]
    if shape == "세로":
        assert [got[0][3], got[1][3]] == want
        assert got[2][3] == before[2][3], "아래 칸까지 번졌습니다"
        assert got[0][4] == before[0][4], "옆 칸까지 번졌습니다"
    else:
        assert got[2][2:4] == want
        assert got[2][4] == before[2][4], "옆 칸까지 번졌습니다"
        assert got[1][2] == before[1][2], "위 칸이 바뀌었습니다"
    assert len(got) == len(before), "줄 수가 바뀌었습니다"


def test_copy_works_without_the_clipboard_api(page):
    """사내 포털은 http 라 navigator.clipboard 가 없다. 그래도 Ctrl+C 가 돼야 한다."""
    _three_rows(page)
    grid(page).locator("body").evaluate(
        "() => Object.defineProperty(navigator, 'clipboard', {value: undefined, configurable: true})")
    page.evaluate("navigator.clipboard.writeText('예전 것')")
    _drag(page, 0, 0, 1, 0)
    want = "\n".join(r[0] for r in table(page)["rows"][:2])
    page.keyboard.press("Control+c")
    page.wait_for_timeout(500)
    assert page.evaluate("navigator.clipboard.readText()") == want


def test_pasting_cells_while_editing_a_cell_spreads_them(page):
    """칸을 고치는 중에 엑셀 여러 칸을 붙이면 한 칸에 몰리지 않고 펼쳐진다."""
    before = _three_rows(page)
    click_cell(page, 0, 2)
    page.keyboard.press("Enter")                  # 고치기 시작
    page.wait_for_timeout(200)
    paste(page, "가\n나\n다\n")
    got = table(page)["rows"]
    assert [got[r][2] for r in range(3)] == ["가", "나", "다"]
    assert len(got) == len(before), "끝 줄바꿈 때문에 빈 줄이 붙었습니다"
    assert got[0][1] == before[0][1] and got[0][3] == before[0][3], "옆 칸까지 번졌습니다"
    assert grid(page).locator(".cell.editing").count() == 0


def test_pasting_one_value_while_editing_inserts_plain_text(page):
    reset(page)
    click_cell(page, 0, 2)
    page.keyboard.press("Enter")
    page.wait_for_timeout(200)
    page.keyboard.press("Control+a")
    page.evaluate("""async () => {
      const html = new Blob(['<b style="color:red">굵게</b>'], {type: 'text/html'});
      const text = new Blob(['굵게'], {type: 'text/plain'});
      await navigator.clipboard.write([new ClipboardItem({'text/html': html, 'text/plain': text})]);
    }""")
    page.keyboard.press("Control+v")
    page.wait_for_timeout(300)
    td = grid(page).locator(".cell.editing")
    assert td.inner_html() == "굵게", "서식이 같이 들어왔습니다"
    page.keyboard.press("Enter")
    wait_dirty(page)
    assert table(page)["rows"][0][2] == "굵게"


def test_backspace_in_the_formula_bar_edits_the_text(page):
    """수식 입력줄의 Backspace 를 격자가 가로채 고른 칸을 지우던 것."""
    reset(page)
    click_cell(page, 0, 2)
    before = table(page)["rows"]
    box = grid(page).locator("#val")
    box.click()
    page.keyboard.press("End")
    page.keyboard.press("Backspace")
    page.keyboard.press("Backspace")
    page.wait_for_timeout(300)
    assert box.input_value() == before[0][2][:-2]
    assert table(page)["rows"] == before, "칸이 지워졌습니다"


# ------------------------------------------------------------- 임시 저장
#
# 고치는 동안 몇 초마다 S3 에 임시로 적어 둔다. 화면이 날아간 뒤 같은 사람이
# 다시 들어오면 '임시 저장된 내용이 있습니다. 복구하시겠습니까?' 를 묻는다.
# 고친 채로 다른 메뉴로 가겠다고 하면 5분(검사에서는 18초) 동안만 들고 있고,
# 그동안은 잠금도 그대로다. 그 뒤로는 둘 다 사라져 다른 사람이 고칠 수 있다.

RESTORE_Q = "임시 저장된 내용이 있습니다. 복구하시겠습니까?"
KEEP_SECONDS = 18


@pytest.fixture(scope="module")
def backup_server():
    port = _free_port()
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", str(ROOT / "app_local.py"),
         "--server.port", str(port), "--server.headless", "true",
         "--browser.gatherUsageStats", "false"],
        cwd=ROOT, env=dict(os.environ, IM_LOCAL_BEAT="2", IM_LOCAL_BACKUP="1",
                           IM_LOCAL_LEAVE_KEEP=str(KEEP_SECONDS / 60)),
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


def _answer(pg, button):
    """창의 단추를 누른다. 창이 막 뜬 참에는 파일을 바꿔 읽는 몇 판이 창을 다시
    그려서 단추가 자리를 못 잡는다 -- 한 박자 기다렸다 누른다."""
    _says(pg, RESTORE_Q)
    pg.wait_for_timeout(1500)
    try:
        pg.get_by_role("button", name=button).click(timeout=15000)
    except Exception:
        pg.screenshot(path="/tmp/claude-0/-home-user-dc-ocap-code/2a01401e-db5d-5432-aecf-9b0e63c99f54/scratchpad/answer_fail.png")
        print("BODY:", pg.inner_text("body")[:1500])
        raise


def _drop_tab(pg):
    """창이 날아간 것처럼 닫는다 (묻지 않고)."""
    _forget_edits(pg)
    pg.close()


def _edit_and_wait_backup(pg, r, c, text):
    _edit(pg, r, c, text)
    wait_dirty(pg)
    pg.wait_for_timeout(4000)                 # 1초 모았다가 올리고, S3 에 적는다


def test_after_losing_the_page_the_same_person_can_restore(backup_server, page):
    hong = _open_as(page, backup_server, "hong")
    settle(hong)
    _edit_and_wait_backup(hong, 0, 2, "날아갈뻔")
    _drop_tab(hong)

    # 다른 사람에게는 안 묻는다
    kim = _open_as(page, backup_server, "kim")
    kim.wait_for_timeout(1500)
    assert RESTORE_Q not in kim.inner_text("body")
    kim.close()

    hong = _open_as(page, backup_server, "hong")
    try:
        _answer(hong, "복구")
        wait_dirty(hong)
        hong.wait_for_timeout(800)
        assert table(hong)["rows"][0][2] == "날아갈뻔"
        # 되살린 것도 여느 때처럼 저장된다. 저장하면 임시 저장은 사라진다.
        confirm_save(hong)
        assert table(hong)["rows"][0][2] == "날아갈뻔"
    finally:
        hong.close()
    hong = _open_as(page, backup_server, "hong")
    try:
        hong.wait_for_timeout(2000)
        assert RESTORE_Q not in hong.inner_text("body"), "저장했는데 또 묻습니다"
    finally:
        hong.close()


def test_discarding_a_temporary_save_forgets_it(backup_server, page):
    hong = _open_as(page, backup_server, "hong")
    settle(hong)
    before = table(hong)["rows"]
    _edit_and_wait_backup(hong, 1, 2, "버릴것")
    _drop_tab(hong)
    hong = _open_as(page, backup_server, "hong")
    try:
        _answer(hong, "버리기")
        settle(hong)
        assert table(hong)["rows"] == before
    finally:
        hong.close()
    hong = _open_as(page, backup_server, "hong")
    try:
        hong.wait_for_timeout(2000)
        assert RESTORE_Q not in hong.inner_text("body")
    finally:
        hong.close()


def _leave_to_home(pg):
    """메뉴의 Home 을 누르고 '벗어나겠습니까?' 에 확인. 한 번에 넘어가야 한다."""
    asked = []
    pg.once("dialog", lambda d: (asked.append(d.message), d.accept()))
    pg.get_by_text("Home", exact=True).click()
    pg.wait_for_function("() => document.body.innerText.includes('홈 화면')", timeout=15000)
    assert asked == [LEAVE_MSG]


def test_leaving_keeps_the_edits_and_the_lock_for_a_while(backup_server, page):
    hong = _open_as(page, backup_server, "hong", menu="1")
    kim = None
    try:
        settle(hong)
        _edit(hong, 0, 3, "실수로나감")
        wait_dirty(hong)
        _leave_to_home(hong)
        assert hong.locator("[role='dialog']").count() == 0

        # 그동안은 다른 사람이 못 고친다
        kim = _open_as(page, backup_server, "kim")
        _says(kim, "hong님이 수정중입니다.")
        kim.close()
        kim = None

        # 돌아오면 묻는다 -- 실수로 확인을 눌렀어도 되살릴 수 있다
        hong.get_by_text("기준 정보 관리", exact=True).click()
        _answer(hong, "복구")
        wait_dirty(hong)
        hong.wait_for_timeout(800)
        assert table(hong)["rows"][0][3] == "실수로나감"
        hong.get_by_role("button", name="초기화").click()   # 뒷정리: 잠금·임시 저장 풀기
        settle(hong)
    finally:
        _forget_edits(hong)
        hong.close()
        if kim:
            kim.close()


def test_after_the_wait_the_edits_are_gone_and_others_can_edit(backup_server, page):
    hong = _open_as(page, backup_server, "hong", menu="1")
    try:
        settle(hong)
        before = table(hong)["rows"]
        _edit(hong, 0, 4, "곧사라짐")
        wait_dirty(hong)
        _leave_to_home(hong)
        hong.wait_for_timeout((KEEP_SECONDS + 3) * 1000)

        kim = _open_as(page, backup_server, "kim")
        try:
            kim.wait_for_timeout(1500)
            assert "수정중입니다" not in kim.inner_text("body"), "시간이 지났는데 잠겨 있습니다"
            settle(kim)
        finally:
            kim.close()

        hong.get_by_text("기준 정보 관리", exact=True).click()
        settle(hong)
        hong.wait_for_timeout(1500)
        assert RESTORE_Q not in hong.inner_text("body"), "시간이 지났는데 또 묻습니다"
        assert table(hong)["rows"] == before
    finally:
        hong.close()


@pytest.mark.parametrize("menu", ["sac", "option"])
def test_leaving_through_the_portal_menu_takes_one_confirm(backup_server, page, menu):
    """포털이 쓰는 메뉴 컴포넌트 그대로. 확인 한 번에 넘어가야 한다."""
    pytest.importorskip({"sac": "streamlit_antd_components",
                         "option": "streamlit_option_menu"}[menu])
    hong = _open_as(page, backup_server, "hong", menu=menu)
    try:
        settle(hong)
        _edit(hong, 1, 3, "메뉴시험")
        wait_dirty(hong)
        asked = []
        hong.once("dialog", lambda d: (asked.append(d.message), d.accept()))
        hong.frame_locator("[data-testid='stSidebar'] iframe").first \
            .get_by_text("Home", exact=True).click()
        hong.wait_for_function("() => document.body.innerText.includes('홈 화면')",
                               timeout=15000)
        assert asked == [LEAVE_MSG]
    finally:
        hong.close()
    # 뒷정리: 남겨 둔 잠금과 임시 저장을 푼다
    hong = _open_as(page, backup_server, "hong")
    try:
        _answer(hong, "버리기")
        settle(hong)
    finally:
        hong.close()
