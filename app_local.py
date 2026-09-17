"""로컬에서 화면만 확인할 때 쓴다. 진짜 S3 대신 메모리 가짜 저장소로 돈다.

    streamlit run app_local.py

포털에 붙는 것은 src/input_manage/input_manage.py 의 show_input_manage()
하나뿐이고, 이 파일은 저장소에만 있다 (포털은 이 파일을 안 읽는다).
"""
import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent))
from src.input_manage import storage
from tests import fake_s3

storage.s3io = fake_s3
storage.PREFIX = "2GAPU/input"


def seed():
    books = {
        "FAB_INPUT_ULY_r0": {
            "STEP": pd.DataFrame([
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
        },
        "FAB_INPUT_TTS_r0": {
            "STEP": pd.DataFrame([
                {"step_seq": "0010", "step_id": "BB100001TR01", "step_desc": "PRE",
                 "ppid": "P-TTS-01", "사용": "Y"},
            ]),
        },
    }
    for name, sheets in books.items():
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for sheet, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        fake_s3.put_object(f"2GAPU/input/{name}.xlsx", buf.getvalue())


if not fake_s3.STORE:
    seed()

from src.input_manage.input_manage import show_input_manage

st.set_page_config(page_title="기준 정보 관리", layout="wide")
st.session_state.setdefault("user_id", "hong")
show_input_manage()
