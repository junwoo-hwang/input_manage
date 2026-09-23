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
st.session_state.setdefault("user_id", "hong")
show_input_manage()
