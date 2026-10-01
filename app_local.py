"""로컬에서 화면만 확인할 때 쓴다. 진짜 S3 대신 메모리 가짜 저장소로 돈다.

    streamlit run app_local.py

포털에 붙는 것은 src/input_manage/input_manage.py 의 show_input_manage()
하나뿐이고, 이 파일은 저장소에만 있다 (포털은 이 파일을 안 읽는다).
"""
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent))
from src.input_manage import input_manage as im
from tests import fake_s3

im.s3 = fake_s3
im.FOLDER_PATH = "2GAPU/input"

# 진짜 S3 처럼 저장에 시간이 걸리게 할 때 (초):  IM_LOCAL_SLOW_SAVE=2
# 가짜 저장소는 순식간이라 '저장하는 동안 단추가 꺼져 있나' 를 볼 수가 없다.
SLOW = float(__import__("os").environ.get("IM_LOCAL_SLOW_SAVE", "0"))
if SLOW and not getattr(fake_s3.put_object, "_slow", False):
    _put = fake_s3.put_object

    def _slow_put(key, data):
        __import__("time").sleep(SLOW / 2)       # 이력 한 번, 본 파일 한 번
        return _put(key, data)
    _slow_put._slow = True
    fake_s3.put_object = _slow_put


BIG_ROWS = int(__import__("os").environ.get("IM_LOCAL_ROWS", "0"))


def big_step(n):
    """실제 파일 크기(8000행 넘음)에서 어떻게 도는지 보려고 부풀린 시트."""
    return pd.DataFrame({
        "step_seq": [f"{(i + 1) * 10:04d}" for i in range(n)],
        "step_id": [f"AA{940000 + i}TR01" for i in range(n)],
        "step_desc": ["PRE", "MAIN", "POST"][:1] * n,
        "ppid": [f"P-ULY-{i % 50:02d}" for i in range(n)],
        "사용": ["Y" if i % 7 else "N" for i in range(n)],
        "비고": [""] * n,
    })


def seed():
    books = {
        "FAB_INPUT_ULY_r0": {
            "STEP": big_step(BIG_ROWS) if BIG_ROWS else pd.DataFrame([
                {"step_seq": "0010", "step_id": "AA941234TR01", "step_desc": "PRE",
                 "ppid": "P-ULY-01", "사용": "Y"},
                {"step_seq": "0020", "step_id": "AA941235TR01", "step_desc": "MAIN",
                 "ppid": "P-ULY-02", "사용": "Y"},
                {"step_seq": "0030", "step_id": "AA941236TR01", "step_desc": "POST",
                 "ppid": "P-ULY-03", "사용": "N"},
            ]),
            "ITEM": pd.DataFrame([
                {"item_id": "item1", "unit": "mV", "owner": "홍길동", "비고": ""},
                {"item_id": "item3", "unit": "uA", "owner": "김철수", "비고": "관리 강화"},
            ]),
            "PROBE CARD": pd.DataFrame([
                {"card": "PC001", "상태": "사용", "교체주기": "90"},
                {"card": "PC002", "상태": "점검", "교체주기": "60"},
            ]),
            "EQP": pd.DataFrame([
                {"eqp": "PRB01", "line": "L1", "사용": "Y"},
            ]),
            # 저장할 때 사유를 받아 한 줄씩 쌓는 시트. 실제 파일에 있는 것과
            # 칸 이름을 맞춰 둔다.
            "REV_INFO": pd.DataFrame([
                {"Date": "2026-09-01", "Remark": "최초 등록",
                 "user": "홍길동", "관련": ""},
            ]),
        },
        "FAB_INPUT_TTS_r0": {
            "STEP": pd.DataFrame([
                {"step_seq": "0010", "step_id": "BB100001TR01", "step_desc": "PRE",
                 "ppid": "P-TTS-01", "사용": "Y"},
            ]),
        },
    }
    for name, sheets in books.items():
        fake_s3.put_object(f"2GAPU/input/{name}.xlsx", im.to_xlsx(sheets))


if not fake_s3.STORE:
    seed()
    # 진짜 크기의 엑셀로 보고 싶을 때:  IM_LOCAL_XLSX=경로 streamlit run app_local.py
    real = __import__("os").environ.get("IM_LOCAL_XLSX")
    if real:
        fake_s3.put_object("2GAPU/input/AAA_REAL.xlsx", Path(real).read_bytes())

from src.input_manage.input_manage import show_input_manage

st.set_page_config(page_title="기준 정보 관리", layout="wide")
# 두 사람을 흉내 낼 때:  http://localhost:8501/?user=kim
st.session_state.setdefault("user_id", st.query_params.get("user", "hong"))
# 잠금 heartbeat 를 검사에서 빨리 돌릴 때 (초):  IM_LOCAL_BEAT=2
_beat = __import__("os").environ.get("IM_LOCAL_BEAT")
if _beat:
    im.LOCK_BEAT_SECONDS = float(_beat)
    im.LOCK_PEEK_SECONDS = min(im.LOCK_PEEK_SECONDS, float(_beat))
# 임시 저장 간격(초)과, 고친 채 떠난 것을 들고 있는 시간(분)을 검사에서 줄일 때
_bk = __import__("os").environ.get("IM_LOCAL_BACKUP")
if _bk:
    im.BACKUP_SECONDS = float(_bk)
_keep = __import__("os").environ.get("IM_LOCAL_LEAVE_KEEP")
if _keep:
    im.INPUT_LEAVE_KEEP_MINUTES = float(_keep)
# 포털 메뉴 흉내 (?menu=1). 사이드바의 streamlit 위젯과, 메뉴 컴포넌트처럼
# 따로 뜨는 틀 안의 링크 둘 다. 고치다가 메뉴를 누르면 묻는지 본다.
# ?menu=sac / ?menu=option 이면 포털이 쓰는 그 메뉴 컴포넌트로 (깔려 있을 때만).
_menu = st.query_params.get("menu")
if _menu:
    import streamlit.components.v1 as components
    with st.sidebar:
        if _menu == "sac":
            import streamlit_antd_components as sac
            pick = sac.menu([sac.MenuItem("기준 정보 관리"), sac.MenuItem("Home")],
                            index=0, key="menu_sac", size="sm", variant="subtle")
            where = pick or "기준 정보 관리"
        elif _menu == "option":
            from streamlit_option_menu import option_menu
            where = option_menu("메뉴", ["기준 정보 관리", "Home"], key="menu_opt")
        else:
            where = st.radio("메뉴", ["기준 정보 관리", "Home"], key="menu_pick")
            components.html('<a href="#" id="menu-link" '
                            'onclick="document.body.dataset.clicked=1">다른 메뉴</a>', height=40)
    if where == "Home":
        st.write("홈 화면")
        st.stop()
show_input_manage()
