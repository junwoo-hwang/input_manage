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
    return grid(page).locator("tbody td:not(.rowhead)")


def table(page):
    """격자의 지금 내용 (헤더, 줄들)."""
    return grid(page).locator("table").first.evaluate("""(t) => ({
      cols: [...t.querySelectorAll('thead th.colhead')].map(e => e.textContent),
      rows: [...t.querySelectorAll('tbody tr')].map(
              tr => [...tr.querySelectorAll('td:not(.rowhead)')].map(e => e.textContent)),
    })""")


def click_cell(page, r, c):
    """칸을 누른다. 겸사겸사 iframe 에 초점이 가서 자판/클립보드가 거기로 간다."""
    grid(page).locator(f"td[data-r='{r}'][data-c='{c}']").click()
    page.wait_for_timeout(250)


def paste(page, text):
    """엑셀에서 긁어온 것처럼 클립보드에 넣고 붙여넣는다.

    locator.press 가 아니라 page.keyboard 를 쓴다 -- 전자는 그 요소로만
    키를 보내서 브라우저가 진짜 paste 이벤트를 만들지 않는다.
    """
    page.evaluate("(t) => navigator.clipboard.writeText(t)", text)
    page.keyboard.press("Control+v")
    page.wait_for_timeout(1500)


def reset(page):
    """항상 같은 자리에서 시작한다 (S3 의 지금 값을 다시 읽는다)."""
    page.get_by_role("button", name="다시 불러오기").click()
    page.wait_for_timeout(3000)


# ------------------------------------------------------------ 보이는가

def test_the_whole_sheet_is_shown_with_its_columns(page):
    got = table(page)
    assert got["cols"], "칸 이름 줄이 없습니다"
    assert got["rows"], "값 줄이 없습니다"
    assert len(got["rows"][0]) == len(got["cols"]), "칸 수와 값 수가 안 맞습니다"


def test_no_javascript_errors(page):
    assert page.errors == []


def test_ids_keep_their_leading_zeros(page):
    """'0010' 이 10 으로 바뀌면 기준 정보로 못 쓴다."""
    flat = [v for row in table(page)["rows"] for v in row]
    assert any(v.startswith("0") and len(v) > 1 for v in flat), flat


# ---------------------------------------------------- 끌어서 범위 선택

def test_dragging_selects_a_rectangle(page):
    reset(page)
    a = grid(page).locator("td[data-r='0'][data-c='0']")
    b = grid(page).locator("td[data-r='0'][data-c='2']")
    a.hover(); page.mouse.down()
    b.hover(); page.mouse.up()
    page.wait_for_timeout(400)
    assert grid(page).locator("td.sel").count() == 3
    assert "선택 1x3" in grid(page).locator("#where").inner_text()


def test_clicking_a_column_header_selects_the_whole_column(page):
    reset(page)
    grid(page).locator("th.colhead[data-c='1']").click()
    page.wait_for_timeout(400)
    rows = len(table(page)["rows"])
    assert grid(page).locator("td.sel").count() == rows


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
    a = grid(page).locator("td[data-r='0'][data-c='0']")
    b = grid(page).locator("td[data-r='0'][data-c='1']")
    a.hover(); page.mouse.down(); b.hover(); page.mouse.up()
    want = table(page)["rows"][0][:2]
    page.keyboard.press("Control+c")
    page.wait_for_timeout(900)
    text = page.evaluate("navigator.clipboard.readText()")
    assert text == "\t".join(want), (text, want)


def test_delete_clears_the_selected_range(page):
    reset(page)
    a = grid(page).locator("td[data-r='0'][data-c='0']")
    b = grid(page).locator("td[data-r='0'][data-c='1']")
    a.hover(); page.mouse.down(); b.hover(); page.mouse.up()
    page.keyboard.press("Delete")
    page.wait_for_timeout(900)
    assert table(page)["rows"][0][:2] == ["", ""]


# ---------------------------------------------------- 행 / 열 넣고 빼기

def test_inserting_and_deleting_a_row(page):
    reset(page)
    n = len(table(page)["rows"])
    click_cell(page, 0, 0)
    grid(page).get_by_text("행 아래", exact=True).click()
    page.wait_for_timeout(600)
    assert len(table(page)["rows"]) == n + 1
    grid(page).get_by_text("행 삭제", exact=True).click()
    page.wait_for_timeout(600)
    assert len(table(page)["rows"]) == n


def test_inserting_and_deleting_a_column(page):
    """st.data_editor 로는 안 되는 것. 이게 이 격자를 직접 만든 이유다."""
    reset(page)
    cols = table(page)["cols"]
    click_cell(page, 0, 0)
    grid(page).get_by_text("열 오른쪽", exact=True).click()
    page.wait_for_timeout(600)
    grown = table(page)["cols"]
    assert len(grown) == len(cols) + 1
    assert grown[1] not in cols, "새 칸 이름이 기존 것과 겹칩니다"

    click_cell(page, 0, 1)
    grid(page).get_by_text("열 삭제", exact=True).click()
    page.wait_for_timeout(600)
    assert table(page)["cols"] == cols


def test_undo_puts_back_what_a_delete_removed(page):
    reset(page)
    before = table(page)
    click_cell(page, 0, 0)
    grid(page).get_by_text("행 삭제", exact=True).click()
    page.wait_for_timeout(600)
    assert table(page)["rows"] != before["rows"]
    grid(page).get_by_text("되돌리기", exact=True).click()
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
    # 격자는 고친 값을 한 박자(350ms) 모았다가 올리고, 그러면 streamlit 이
    # 스크립트를 다시 돈다. 그 왕복을 기다린다.
    page.wait_for_timeout(3000)

    assert "고친 칸 1개" in page.inner_text("body"), page.inner_text("body")[:400]
    save = page.get_by_role("button", name="저장")
    assert save.is_enabled(), "저장 버튼이 안 켜졌습니다"
    save.click()
    page.wait_for_timeout(3000)

    assert "저장했습니다" in page.inner_text("body")
    reset(page)
    assert table(page)["rows"][0][2] == typed, "저장한 값이 안 남았습니다"
