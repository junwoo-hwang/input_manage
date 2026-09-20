"""기준 정보 관리 — 포털 메뉴 하나.

S3 drive 에 있는 기준 정보 엑셀을 브라우저에서 엑셀처럼 고치고 바로
저장한다. 내려받았다 고쳐 다시 올리는 왕복이 없다.

    G-DVC / 2GAPU/input /
        FAB_INPUT_ULY_r0.xlsx   <- 이 폴더의 .xlsx 가 곧 편집 대상 목록
        FAB_INPUT_TTS_r0.xlsx
        _history/FAB_INPUT_ULY_r0/20260918_1041_a3f1_hong.xlsx
        _audit.csv

_history 와 _audit.csv 는 우리가 만드는 것이라 이름 앞에 _ 를 붙였다.
목록에서 빼는 기준도 그거다 -- 엑셀이 늘어나도 코드를 안 고치고, 우리
파일은 편집 대상으로 잡히지 않는다.

포털에서는 show_input_manage() 하나만 부르면 된다.

화면의 격자는 sheet_grid/frontend/index.html 이다. 거기만 따로인 이유는
streamlit 컴포넌트가 iframe 에 띄우는 방식이라 파일이 나뉠 수밖에 없어서다.
"""
from __future__ import annotations

import csv
import hashlib
import inspect
import io
import os
import threading
from datetime import datetime, timedelta, timezone
from pathlib import Path

import boto3
import botocore.exceptions
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

KST = timezone(timedelta(hours=9))


# ======================================================================
# 1. S3. 사내 환경에 닿는 자리는 여기뿐이다.
#
# 설정과 클라이언트 만드는 법은 포털의 src/dc_ocap/dc_ocap.py 와 같게 맞췄다.
# 다른 점은 목적뿐이다 -- 거기는 .html 을 내려받아 보여주기만 하고, 여기는
# .xlsx 를 읽고 다시 올린다. 자격증명은 환경변수에서만 읽는다.
# ======================================================================
AWS_ACCESS_KEY = os.getenv("AWS_ACCESS_KEY")
AWS_SECRET_KEY = os.getenv("AWS_SECRET_KEY")

BUCKET_NAME = os.getenv("INPUT_S3_BUCKET", "G-DVC")
FOLDER_PATH = os.getenv("INPUT_S3_PREFIX", "2GAPU/input").strip("/")
S3_ENDPOINT = os.getenv("INPUT_S3_ENDPOINT", "http://s3.dataplatform.samsungds.net:9020")

_client_lock = threading.Lock()
_client = None


def _s3_client():
    """클라이언트 하나를 만들어 재사용한다.

    streamlit 은 사람이 버튼 한 번 누를 때마다 스크립트를 처음부터 다시
    돌린다. 그때마다 새로 만들면 매번 연결을 다시 맺느라 저장이 눈에 띄게
    느려진다.
    """
    global _client
    with _client_lock:
        if _client is None:
            _client = boto3.client(
                service_name="s3",
                aws_access_key_id=AWS_ACCESS_KEY,
                aws_secret_access_key=AWS_SECRET_KEY,
                endpoint_url=S3_ENDPOINT,
            )
        return _client


def get_object(key: str) -> tuple[bytes, str] | None:
    """(내용, 버전표). 없으면 None.

    버전표는 S3 가 주는 ETag(내용 해시)를 그대로 쓴다. '내가 화면에 띄운 뒤
    다른 사람이 먼저 저장했는가' 를 가리는 데 쓰는데, 그 판단 하나 하자고
    파일을 통째로 다시 받아 해시를 뜰 필요가 없다는 것이 ETag 의 쓸모다.
    """
    try:
        got = _s3_client().get_object(Bucket=BUCKET_NAME, Key=key)
    except botocore.exceptions.ClientError as err:
        if err.response["Error"]["Code"] in ("NoSuchKey", "404", "NoSuchBucket"):
            return None
        raise
    return got["Body"].read(), got.get("ETag", "").strip('"')


def head_etag(key: str) -> str:
    """지금 올라가 있는 것의 버전표. 없으면 빈 글자."""
    try:
        got = _s3_client().head_object(Bucket=BUCKET_NAME, Key=key)
    except botocore.exceptions.ClientError:
        return ""
    return got.get("ETag", "").strip('"')


def put_object(key: str, data: bytes) -> str:
    _s3_client().put_object(Bucket=BUCKET_NAME, Key=key, Body=data)
    return head_etag(key)


def list_keys(prefix: str) -> list[str]:
    out: list[str] = []
    paginator = _s3_client().get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=BUCKET_NAME, Prefix=prefix):
        for obj in page.get("Contents", []):
            if not obj["Key"].endswith("/"):
                out.append(obj["Key"])
    return sorted(out)


def delete_object(key: str) -> None:
    _s3_client().delete_object(Bucket=BUCKET_NAME, Key=key)


# 위 다섯 개를 한 묶음으로 들고 다닌다. 테스트는 이것만 가짜로 갈아끼워서
# S3 없이 돌고, 사내 헬퍼로 바꿔 끼울 때도 여기만 손대면 된다.
class _S3:
    get_object = staticmethod(get_object)
    head_etag = staticmethod(head_etag)
    put_object = staticmethod(put_object)
    list_keys = staticmethod(list_keys)
    delete_object = staticmethod(delete_object)


s3 = _S3()


# ======================================================================
# 2. 엑셀 읽고 쓰기. 화면은 여기 안 들어온다.
#
# 기준 정보는 여러 사람이 같이 고치는 값이라, 조용히 덮어써지거나 저장이
# 반쯤 되다 마는 일이 생기면 무엇이 맞는 값인지 아무도 모르게 된다. 이
# 구역은 브라우저 없이 검사할 수 있게 화면과 떼어 두었다.
# ======================================================================
HISTORY_KEEP = 100          # 파일 하나당 남길 이력 개수
AUDIT_COLS = ["saved_at", "user_id", "workbook", "sheet",
              "changed_cells", "rows_before", "rows_after"]


class ConcurrentEdit(Exception):
    """내가 화면에 띄운 뒤 다른 사람이 먼저 저장했다.

    그냥 덮어쓰면 그 사람이 고친 값이 소리 없이 사라진다. 누가 맞는지는
    코드가 정할 수 없으므로 여기서 멈추고 사람에게 넘긴다.
    """


def _key(*parts: str) -> str:
    return "/".join([FOLDER_PATH, *parts])


def audit_key() -> str:
    return _key("_audit.csv")


def list_workbooks() -> list[str]:
    """고칠 수 있는 엑셀 파일 이름들 (확장자 뺀 것)."""
    names = []
    for key in s3.list_keys(FOLDER_PATH + "/"):
        name = key[len(FOLDER_PATH) + 1:]
        if "/" in name or not name.lower().endswith(".xlsx") or name.startswith("_"):
            continue                       # _history/ 안의 것과 내부 파일은 뺀다
        names.append(name[:-len(".xlsx")])
    return names


def load_workbook(book: str) -> tuple[dict[str, pd.DataFrame], str]:
    """{시트이름: DataFrame} 과 그 시점의 버전표."""
    got = s3.get_object(_key(f"{book}.xlsx"))
    if got is None:
        return {}, ""
    data, etag = got
    sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, dtype=object)
    # 엑셀의 빈 칸은 NaN 으로 읽힌다. 그대로 두면 사람이 손도 안 댄 칸이
    # 나중에 "nan" 이라는 글자로 저장된다.
    return {name: df.where(pd.notna(df), None) for name, df in sheets.items()}, etag


def save_workbook(book: str, sheets: dict[str, pd.DataFrame], user_id: str,
                  base_stamp: str | None = None) -> str:
    """고친 값을 올리고 새 버전표를 돌려준다.

    base_stamp 를 주면 그 사이에 다른 사람이 올렸는지 보고, 그랬으면
    아무것도 쓰지 않고 ConcurrentEdit 를 던진다.
    """
    key = _key(f"{book}.xlsx")
    if base_stamp is not None and s3.head_etag(key) != base_stamp:
        raise ConcurrentEdit(
            f"'{book}' 을(를) 화면에 띄운 뒤 다른 사람이 먼저 저장했습니다. "
            f"덮어쓰지 않았습니다 -- 다시 불러와서 고친 내용을 옮겨 주세요."
        )

    before = {}
    got = s3.get_object(key)
    if got is not None:
        before = pd.read_excel(io.BytesIO(got[0]), sheet_name=None, dtype=object)

    # 통째로 만들어 한 번에 올린다. S3 의 put 은 그 자체로 원자적이라,
    # 올리다 끊겨도 옛 파일이 반쯤 덮어써지는 일은 없다.
    body = to_xlsx(sheets)

    now = datetime.now(KST)
    # 이력을 먼저 올린다. 순서가 반대면, 본 파일은 바뀌었는데 이력이 없는
    # 순간이 생긴다 -- 되돌릴 것을 찾을 때 하필 그 판이 없는 쪽이 더 아프다.
    s3.put_object(_key("_history", book, _history_name(now, user_id, body)), body)
    etag = s3.put_object(key, body)
    _trim_history(book)
    _append_audit(now, user_id, book, before, sheets)
    return etag


def to_xlsx(sheets: dict[str, pd.DataFrame]) -> bytes:
    """시트들을 엑셀 파일 한 벌로. 저장과 내려받기가 같은 것을 쓴다."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            _clean(df).to_excel(writer, sheet_name=_sheet_name(name), index=False)
    return buf.getvalue()


def _history_name(now: datetime, user_id: str, body: bytes) -> str:
    """이력 파일 이름. 시각 + 내용 네 글자 + 누가.

    내용 네 글자를 끼우는 이유는, 같은 초에 두 번 저장되면(두 사람이 거의
    동시에 눌렀을 때) 이름이 같아져 앞 이력이 덮어써지기 때문이다. 되돌릴
    판이 하나 사라지는 셈이라 조용히 넘길 일이 아니다. 내용이 정말 같으면
    이름도 같은데, 그때는 덮어써도 잃는 것이 없다.
    """
    short = hashlib.md5(body).hexdigest()[:4]
    return f"{now:%Y%m%d_%H%M%S}_{short}_{_safe(user_id)}.xlsx"


def _sheet_name(name: str) -> str:
    """엑셀 시트 이름 규칙에 맞춘다 (31자, : \\ / ? * [ ] 못 씀)."""
    clean = "".join(" " if c in ':\\/?*[]' else c for c in str(name))
    return clean[:31] or "Sheet1"


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    """엑셀로 나가기 전에 다듬는다.

    - 통째로 빈 줄은 뺀다. 격자에서 '행 아래' 를 눌렀다 안 채우고 저장하면
      빈 줄이 그대로 쌓이는데, 그걸 읽는 쪽에서는 결측 한 줄이 된다.
    - None/NaN 은 빈 칸으로 쓴다 ("nan" 이라는 글자로 저장되지 않게).
    """
    out = df.copy().where(pd.notna(df), None)
    if not len(out):
        return out
    blank = out.apply(lambda row: all(
        v is None or str(v).strip() == "" for v in row), axis=1)
    return out[~blank]


def _safe(text: str) -> str:
    """S3 키에 넣어도 되는 꼴로. 빈 값이면 'unknown'."""
    kept = "".join(c for c in str(text or "") if c.isalnum() or c in "-_.")
    return kept[:40] or "unknown"


def history_keys(book: str, limit: int = 50) -> list[str]:
    return sorted(s3.list_keys(_key("_history", book) + "/"), reverse=True)[:limit]


def _trim_history(book: str) -> None:
    keys = sorted(s3.list_keys(_key("_history", book) + "/"))
    for old in keys[:-HISTORY_KEEP]:
        s3.delete_object(old)


def _as_text(df: pd.DataFrame, cols: list[str], rows: int) -> pd.DataFrame:
    """값을 글자로 눕혀 같은 모양으로 맞춘다 (없는 칸은 빈 글자).

    astype(object) 를 먼저 하는 것이 중요하다. 줄 수가 다른 두 표를 맞추려면
    reindex 로 빈 줄을 채우는데, 정수 칸에 NaN 이 들어가면 pandas 가 그 칸을
    통째로 실수로 올려서 1 이 1.0 이 된다. 그러면 손도 안 댄 칸까지
    '바뀌었다' 로 세어져, 저장 전에 보여주는 숫자가 사람이 고친 칸 수와
    안 맞는다.
    """
    out = df.set_axis([str(c) for c in df.columns], axis=1).astype(object)
    out = out.where(pd.notna(out), "")
    out = out.reindex(index=range(rows), columns=cols, fill_value="")
    return out.astype(str).apply(lambda s: s.str.strip())


def changed_cells(before: pd.DataFrame | None, after: pd.DataFrame) -> int:
    """두 표 사이에 값이 다른 칸이 몇 개인가.

    줄이나 칸이 늘고 준 것도 센다 -- 저장 전에 사람이 확인하려는 숫자라서.
    빈 칸과 '없는 칸' 은 같게 본다 (격자에서 행을 늘렸다 비워둔 것은 고친
    것이 아니다).
    """
    after = _clean(after)
    before = _clean(before) if before is not None else after.iloc[:0]
    cols = list(dict.fromkeys([*map(str, before.columns), *map(str, after.columns)]))
    rows = max(len(before), len(after))
    if not cols or rows == 0:
        return 0
    a, b = _as_text(before, cols, rows), _as_text(after, cols, rows)
    n = int((a != b).to_numpy().sum())

    # 값만 견주면 '빈 칸을 새로 넣은 것' 이 0 으로 나온다 -- 없는 칸도 빈
    # 글자로 채워 맞추기 때문이다. 그러면 열을 하나 넣고 저장을 누를 수가
    # 없다. 내용이 있는 칸은 이미 위에서 세어졌으므로, 비어 있는 채로
    # 생기거나 없어진 칸만 한 개씩 더한다.
    before_cols = set(map(str, before.columns))
    after_cols = set(map(str, after.columns))
    for name in (after_cols - before_cols) | (before_cols - after_cols):
        if not (b[name] if name in after_cols else a[name]).str.strip().any():
            n += 1
    return n


def _append_audit(now: datetime, user_id: str, book: str,
                  before: dict[str, pd.DataFrame],
                  after: dict[str, pd.DataFrame]) -> None:
    """누가 언제 어느 시트를 몇 칸 고쳤는지 한 줄씩 덧붙인다.

    기준 정보는 '언제부터 이 값이었나' 를 되짚을 일이 반드시 생긴다.
    엑셀 파일만 남기면 그 답을 못 한다.
    """
    rows = []
    for name, df in after.items():
        old = before.get(name)
        n = changed_cells(old, df)
        if n == 0 and old is not None:
            continue                          # 안 바뀐 시트는 적지 않는다
        rows.append([f"{now:%Y-%m-%d %H:%M:%S}", user_id or "unknown", book, name,
                     n, 0 if old is None else len(old), len(_clean(df))])
    if not rows:
        return

    got = s3.get_object(audit_key())
    buf = io.StringIO()
    if got is None:
        csv.writer(buf).writerow(AUDIT_COLS)
    else:
        buf.write(got[0].decode("utf-8-sig"))
        if not buf.getvalue().endswith("\n"):
            buf.write("\n")
    writer = csv.writer(buf, lineterminator="\n")
    for row in rows:
        writer.writerow(row)
    s3.put_object(audit_key(), buf.getvalue().encode("utf-8-sig"))


def read_audit(book: str | None = None, limit: int = 200) -> pd.DataFrame:
    got = s3.get_object(audit_key())
    if got is None:
        return pd.DataFrame(columns=AUDIT_COLS)
    df = pd.read_csv(io.BytesIO(got[0]), dtype=str, encoding="utf-8-sig")
    if book is not None and "workbook" in df.columns:
        df = df[df["workbook"] == book]
    return df.tail(limit).iloc[::-1].reset_index(drop=True)      # 최근 것이 위


# ======================================================================
# 3. 격자 (streamlit 컴포넌트).
#
# st.data_editor 를 안 쓰는 이유는 하나다: 열을 못 넣고 못 뺀다. 칸 구성이
# DataFrame 스키마로 고정돼서, 칸 하나 추가하려면 엑셀을 내려받아 고쳐 다시
# 올려야 한다 -- 그게 하기 싫어서 만든 화면이라 그걸로는 끝이 안 난다.
#
# 시트 전체를 컴포넌트 하나에 넘긴다. st.tabs 로 시트를 나누면 탭이 표 위에
# 붙어 엑셀과 다르게 보이고, 시트마다 iframe 이 하나씩 생겨 시트를 오갈 때마다
# 화면이 끊긴다. 지금은 격자가 제 아래에 엑셀처럼 시트 탭을 그린다.
# ======================================================================
_FRONTEND = Path(__file__).parent / "sheet_grid" / "frontend"
_grid = components.declare_component("input_manage_sheet_grid", path=str(_FRONTEND))


def sheet_grid(sheets: dict[str, pd.DataFrame], version: str, key: str,
               max_height: int = 520) -> dict[str, pd.DataFrame]:
    """격자를 그리고, 사람이 고친 시트들을 돌려준다.

    version 은 '이 데이터가 갈렸다' 를 알리는 표다. 격자는 이 값이 바뀔
    때만 제 상태를 갈아엎는다 -- 값을 올려보낼 때마다 streamlit 이 스크립트를
    다시 돌리면서 같은 데이터가 되돌아오는데, 그때마다 새로 그리면 방금 고친
    칸과 고른 자리, 보고 있던 시트가 날아간다.
    """
    payload = [{
        "name": str(name),
        "cols": [str(c) for c in df.columns],
        "rows": [["" if pd.isna(v) else str(v) for v in row]
                 for row in df.itertuples(index=False, name=None)],
    } for name, df in sheets.items()]

    got = _grid(sheets=payload, version=version, max_height=max_height,
                key=key, default=None)
    if not got:
        return sheets
    return to_frames(got)


def to_frames(payload: dict) -> dict[str, pd.DataFrame]:
    """격자가 올려준 것을 {시트이름: DataFrame} 으로."""
    out: dict[str, pd.DataFrame] = {}
    for i, sheet in enumerate(payload.get("sheets", []) or []):
        name = str(sheet.get("name") or f"Sheet{i + 1}")
        while name in out:                      # 시트 이름도 겹치면 안 된다
            name += "_"
        out[name] = _to_frame(sheet)
    return out


def _to_frame(sheet: dict) -> pd.DataFrame:
    """칸 이름이 겹치면 뒤엣것에 번호를 붙인다.

    겹친 채로 두면 pandas 에서 df["a"] 가 Series 가 아니라 DataFrame 이 되고,
    엑셀로 내보낼 때 값 대신 칸 이름이 실린 파일이 오류 없이 만들어진다.
    """
    cols = [str(c) for c in sheet.get("cols", [])]
    seen: dict[str, int] = {}
    unique = []
    for name in cols:
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        unique.append(name)
    rows = sheet.get("rows", []) or []
    fixed = [(list(r) + [""] * len(unique))[:len(unique)] for r in rows]
    return pd.DataFrame(fixed, columns=unique, dtype=object)


# ======================================================================
# 4. 화면. 포털이 부르는 것은 show_input_manage() 하나다.
# ======================================================================
S_BOOK = "_im_book"        # 지금 고르고 있는 엑셀 파일
S_SHEETS = "_im_sheets"    # 그 파일을 띄웠을 때의 원본
S_STAMP = "_im_stamp"      # 그 원본이 어느 판이었는지 (S3 ETag)
# 몇 번째로 불러온 것인지. 격자에 넘기는 판 번호에 섞는다.
#
# ETag 만으로는 모자란다: '다시 불러오기' 는 아무도 저장하지 않았으면 같은
# 파일을 다시 읽으므로 ETag 가 그대로고, 그러면 격자가 '갈린 게 없다' 며
# 제 상태를 그대로 둔다 -- 버리려고 누른 수정이 화면에 그대로 남는다.
S_NONCE = "_im_nonce"
# 저장 직후에 보여줄 한 줄. 저장하고 바로 st.rerun() 을 하는데, rerun 은
# 스크립트를 처음부터 다시 돌리므로 그 전에 그린 st.success 는 화면에 남지
# 않는다. 그래서 문구를 여기 맡겨 두고 다음 판에서 그린다.
S_TOAST = "_im_toast"
# 내려받을 엑셀. 만들어 둔 뒤에야 st.download_button 을 그릴 수 있다.
#
# 버튼에 바로 못 붙이는 이유: st.download_button 은 파일 내용을 미리 받아야
# 하는데, 8000행짜리 엑셀을 만드는 데 1초 넘게 걸린다. 그걸 화면 그릴 때마다
# 하면 칸 하나 고칠 때마다 그 값을 치르게 된다. 그래서 누를 때만 만든다.
S_DOWNLOAD = "_im_download"


# 칸 너비를 꽉 채우라고 말하는 법이 streamlit 버전마다 다르다. 새 버전은
# width="stretch", 예전 버전은 use_container_width=True 다. 포털이 어느
# 버전인지 모르는 채로 한쪽만 쓰면 화면이 아예 안 뜬다 (TypeError).
_WIDE = ({"width": "stretch"}
         if "width" in inspect.signature(st.button).parameters
         else {"use_container_width": True})


def _load(book: str) -> None:
    sheets, stamp = load_workbook(book)
    st.session_state[S_BOOK] = book
    st.session_state[S_SHEETS] = sheets
    st.session_state[S_STAMP] = stamp
    st.session_state[S_NONCE] = st.session_state.get(S_NONCE, 0) + 1


def show_input_manage() -> None:
    st.markdown('<div class="pretendard-area"><h2>기준 정보 관리</h2></div>',
                unsafe_allow_html=True)

    toast = st.session_state.pop(S_TOAST, None)
    if toast:
        st.success(toast)

    try:
        books = list_workbooks()
    except Exception as err:
        st.error(f"S3 에서 기준 정보 목록을 읽지 못했습니다: {err}")
        st.caption(f"버킷 `{BUCKET_NAME}` / 폴더 `{FOLDER_PATH}` · "
                   f"AWS_ACCESS_KEY, AWS_SECRET_KEY 가 설정돼 있는지 확인하세요.")
        return
    if not books:
        st.warning(f"`{BUCKET_NAME}/{FOLDER_PATH}/` 아래에 .xlsx 가 없습니다.")
        return

    # 고르개는 파일 이름만 들어가면 되므로 좁게 둔다. 화면 폭을 다 쓰면
    # 정작 넓어야 할 표가 그만큼 아래로 밀린다.
    top, refresh, _rest = st.columns([2, 1.2, 5.8])
    with top:
        book = st.selectbox("관리할 파일", books, key="im_book_pick")
    with refresh:
        st.write("")
        reload_now = st.button("다시 불러오기", **_WIDE,
                               help="저장하지 않은 수정을 버리고 S3 의 지금 값을 다시 읽습니다")

    # 파일을 바꿔 고르면 그 파일을 새로 읽는다. 이전 파일의 미저장 수정은
    # 들고 가지 않는다 -- 시트 이름이 겹칠 때 엉뚱한 표에 얹히기 때문이다.
    if reload_now or st.session_state.get(S_BOOK) != book:
        # 만들어 둔 내려받기 파일은 버린다. 안 그러면 파일을 바꿔 골랐는데
        # 이전 파일 내용이 담긴 버튼이 그대로 남는다.
        st.session_state.pop(S_DOWNLOAD, None)
        _load(book)
        if reload_now:
            st.rerun()

    sheets: dict[str, pd.DataFrame] = st.session_state[S_SHEETS]
    if not sheets:
        st.warning(f"'{book}' 에 시트가 없습니다.")
        return

    user_id = st.session_state.get("user_id") or "unknown"
    edited = sheet_grid(
        sheets,
        version=f"{book}|{st.session_state[S_STAMP]}|{st.session_state[S_NONCE]}",
        key="im_grid",
    )

    # 칸 값만 세면 안 된다. 시트를 새로 만들거나 이름을 바꾸거나 지운 것도
    # '고친 것' 인데, 빈 시트를 하나 더한 경우 칸 기준으로는 0 이 나와서
    # 저장 버튼이 안 켜진다 (실제로 그랬다).
    counts: dict[str, int] = {}
    for name, df in edited.items():
        if name in sheets:
            counts[name] = changed_cells(sheets[name], df)
        else:
            counts[f"{name} (새 시트)"] = max(changed_cells(None, df), 1)
    for name in sheets:
        if name not in edited:
            counts[f"{name} (지움)"] = max(len(sheets[name]), 1)
    total = sum(counts.values())

    save_col, make_col, get_col, _gap = st.columns([1, 1.2, 1.6, 3])
    with save_col:
        if st.button("저장", type="primary", disabled=total == 0, **_WIDE):
            _save(book, edited, user_id)
    with make_col:
        if st.button("엑셀 만들기", **_WIDE,
                     help="지금 화면의 값(저장 안 한 수정 포함)으로 엑셀 파일을 만듭니다"):
            st.session_state[S_DOWNLOAD] = (f"{book}.xlsx", to_xlsx(edited))
    with get_col:
        ready = st.session_state.get(S_DOWNLOAD)
        if ready:
            st.download_button(f"⬇ {ready[0]}", ready[1], file_name=ready[0],
                               **_WIDE,
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet")

    if total:
        changed = ", ".join(f"{n}({c})" for n, c in counts.items() if c)
        st.info(f"저장하지 않은 수정 {total}칸 — {changed}")
    else:
        st.caption("고친 것 없음")

    _show_history(book)


def _save(book: str, edited: dict[str, pd.DataFrame], user_id: str) -> None:
    try:
        stamp = save_workbook(book, edited, user_id,
                              base_stamp=st.session_state[S_STAMP])
    except ConcurrentEdit as err:
        # 덮어쓰지 않는다. 누구 값이 맞는지는 코드가 못 정한다.
        st.error(str(err))
        return
    except Exception as err:
        st.error(f"저장하지 못했습니다: {err}\n\n"
                 f"S3 의 값은 그대로입니다. 고친 내용은 화면에 남아 있습니다.")
        return
    # 새 판을 원본으로 삼는다. 판 번호가 바뀌므로 격자도 이 값으로 다시
    # 그려지고, '고친 칸' 은 0 으로 돌아간다.
    st.session_state[S_SHEETS] = {k: v.copy() for k, v in edited.items()}
    st.session_state[S_STAMP] = stamp
    st.session_state.pop(S_DOWNLOAD, None)
    st.session_state[S_TOAST] = f"'{book}' 저장했습니다 ({user_id})."
    st.rerun()


def _show_history(book: str) -> None:
    with st.expander("변경 이력"):
        try:
            log = read_audit(book)
        except Exception as err:
            st.caption(f"이력을 읽지 못했습니다: {err}")
            return
        if log.empty:
            st.caption("아직 저장된 적이 없습니다.")
        else:
            st.dataframe(log, hide_index=True, **_WIDE)
