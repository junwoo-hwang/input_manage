"""기준 정보 관리 화면.

위 드롭다운에서 파일을 고르면 그 엑셀이 통째로 뜬다. 시트가 여럿이면
탭으로 나뉘고, 각 탭은 엑셀과 같은 격자다 -- 칸 이름 줄이 위에 있고 값이
그 아래로 쭉 붙는다. 셀을 끌어서 범위를 고르고, 복사/붙여넣기가 엑셀과
오가고, 행과 열을 넣고 뺄 수 있다.

내려받았다 올리는 왕복은 없다. 저장을 누르면 S3 의 그 엑셀이 바로 바뀐다.

읽고 쓰는 일은 storage.py, 격자는 sheet_grid/ 가 한다. 여기서는 둘을
잇고 저장 버튼이 무엇을 하는지만 정한다.
"""
from __future__ import annotations

import pandas as pd
import streamlit as st

from . import storage
from .sheet_grid import sheet_grid

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


def _grid_key(book: str, sheet: str) -> str:
    return f"im_grid::{book}::{sheet}"


def _load(book: str) -> None:
    sheets, stamp = storage.load_workbook(book)
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
        books = storage.list_workbooks()
    except Exception as err:
        st.error(f"S3 에서 기준 정보 목록을 읽지 못했습니다: {err}")
        st.caption(f"버킷 `{storage.s3io.BUCKET_NAME}` / 폴더 `{storage.PREFIX}` · "
                   f"AWS_ACCESS_KEY, AWS_SECRET_KEY 가 설정돼 있는지 확인하세요.")
        return
    if not books:
        st.warning(f"`{storage.s3io.BUCKET_NAME}/{storage.PREFIX}/` 아래에 .xlsx 가 없습니다.")
        return

    top, refresh = st.columns([5, 1])
    with top:
        book = st.selectbox("관리할 파일", books, key="im_book_pick")
    with refresh:
        st.write("")
        reload_now = st.button("다시 불러오기", width="stretch",
                               help="저장하지 않은 수정을 버리고 S3 의 지금 값을 다시 읽습니다")

    # 파일을 바꿔 고르면 그 파일을 새로 읽는다. 이전 파일의 미저장 수정은
    # 들고 가지 않는다 -- 시트 이름이 겹칠 때 엉뚱한 표에 얹히기 때문이다.
    if reload_now or st.session_state.get(S_BOOK) != book:
        _load(book)
        if reload_now:
            st.rerun()

    sheets: dict[str, pd.DataFrame] = st.session_state[S_SHEETS]
    if not sheets:
        st.warning(f"'{book}' 에 시트가 없습니다.")
        return

    user_id = st.session_state.get("user_id") or "unknown"
    stamp = st.session_state[S_STAMP]
    nonce = st.session_state[S_NONCE]

    # 탭은 한 번에 하나만 보이지만 격자는 안 보이는 탭 것도 값을 들고 있다.
    # 저장은 늘 전체를 함께 쓴다 -- 보이는 탭만 쓰면 나머지 시트가 통째로
    # 사라진 엑셀이 올라간다.
    edited: dict[str, pd.DataFrame] = {}
    counts: dict[str, int] = {}
    names = list(sheets)
    for tab, name in zip(st.tabs(names), names):
        with tab:
            edited[name] = sheet_grid(
                sheets[name],
                # 판 번호에 시트 이름과 불러온 횟수까지 넣는다. 시트 이름이
                # 없으면 한 파일 안의 시트들이 같은 ETag 를 써서 탭을 오갈 때
                # 격자가 옆 시트 표를 자기 것으로 착각하고, 횟수가 없으면
                # '다시 불러오기' 가 아무것도 안 되돌린다 (위 S_NONCE 참고).
                version=f"{book}|{name}|{stamp}|{nonce}",
                key=_grid_key(book, name),
            )
            counts[name] = storage.changed_cells(sheets[name], edited[name])
            st.caption(f"{len(edited[name])}행 x {len(edited[name].columns)}열"
                       + (f" · 고친 칸 {counts[name]}개" if counts[name]
                          else " · 고친 것 없음"))

    total = sum(counts.values())
    left, _gap = st.columns([1, 5])
    with left:
        if st.button("저장", type="primary", disabled=total == 0, width="stretch"):
            _save(book, edited, user_id)

    if total:
        changed = ", ".join(f"{n}({c})" for n, c in counts.items() if c)
        st.info(f"저장하지 않은 수정 {total}칸 — {changed}")

    _show_history(book)


def _save(book: str, edited: dict[str, pd.DataFrame], user_id: str) -> None:
    try:
        stamp = storage.save_workbook(book, edited, user_id,
                                      base_stamp=st.session_state[S_STAMP])
    except storage.ConcurrentEdit as err:
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
    st.session_state[S_TOAST] = f"'{book}' 저장했습니다 ({user_id})."
    st.rerun()


def _show_history(book: str) -> None:
    with st.expander("변경 이력"):
        try:
            log = storage.read_audit(book)
        except Exception as err:
            st.caption(f"이력을 읽지 못했습니다: {err}")
            return
        if log.empty:
            st.caption("아직 저장된 적이 없습니다.")
        else:
            st.dataframe(log, width="stretch", hide_index=True)
