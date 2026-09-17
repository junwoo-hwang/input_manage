"""엑셀처럼 고치는 격자. streamlit 컴포넌트 하나.

st.data_editor 를 안 쓰는 이유는 하나다: 열을 못 넣고 못 뺀다. 칸 구성이
DataFrame 스키마로 고정돼서, 기준 정보에 칸 하나 추가하려면 엑셀을
내려받아 고쳐 다시 올려야 한다. 그게 하기 싫어서 만든 화면이므로 그걸로는
끝이 안 난다.

그래서 격자를 직접 만들었다. frontend/index.html 하나가 전부이고, npm
빌드 단계가 없다 -- iframe 과 부모 사이의 postMessage 규약 세 가지
(componentReady / render / setComponentValue)만 직접 지킨다. 사내망에서
npm 의존성을 하나 더 들이는 것보다 이쪽이 싸다.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit.components.v1 as components

_FRONTEND = Path(__file__).parent / "frontend"
_component = components.declare_component("input_manage_sheet_grid", path=str(_FRONTEND))


def sheet_grid(df: pd.DataFrame, version: str, key: str,
               max_height: int = 520) -> pd.DataFrame:
    """격자를 그리고, 사람이 고친 표를 돌려준다.

    version 은 '이 데이터가 갈렸다' 를 알리는 표다. 격자는 이 값이 바뀔
    때만 제 상태를 갈아엎는다 -- 값을 올려보낼 때마다 streamlit 이 스크립트를
    다시 돌리면서 같은 데이터가 되돌아오는데, 그때마다 새로 그리면 방금 고친
    칸과 고른 자리가 날아간다.
    """
    cols = [str(c) for c in df.columns]
    rows = [["" if pd.isna(v) else str(v) for v in row]
            for row in df.itertuples(index=False, name=None)]

    got = _component(cols=cols, rows=rows, version=version,
                     max_height=max_height, key=key, default=None)
    if not got:
        return df
    return to_frame(got)


def to_frame(payload: dict) -> pd.DataFrame:
    """격자가 올려준 것을 DataFrame 으로.

    칸 이름이 겹치면 뒤엣것에 번호를 붙인다. 겹친 채로 두면 pandas 에서
    df["a"] 가 Series 가 아니라 DataFrame 이 되고, 엑셀로 내보낼 때 값 대신
    칸 이름이 실린 파일이 오류 없이 만들어진다.
    """
    cols = [str(c) for c in payload.get("cols", [])]
    rows = payload.get("rows", []) or []
    seen: dict[str, int] = {}
    unique = []
    for name in cols:
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        unique.append(name)
    fixed = [list(r) + [""] * (len(unique) - len(r)) for r in rows]
    fixed = [r[:len(unique)] for r in fixed]
    return pd.DataFrame(fixed, columns=unique, dtype=object)
