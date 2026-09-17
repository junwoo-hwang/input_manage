"""기준 정보 엑셀(S3 drive)을 읽고 쓴다. 화면(streamlit)은 여기 안 들어온다.

화면 없이 도는 부분만 따로 둔 이유는 테스트 때문이다. 기준 정보는 여러
사람이 같이 고치는 값이라, 조용히 덮어써지거나 저장이 반쯤 되다 마는 일이
생기면 무엇이 맞는 값인지 아무도 모르게 된다. 그 부분은 브라우저 없이
검사할 수 있어야 한다.

S3 에 닿는 일은 전부 s3io.py 가 한다.

    G-DVC/2GAPU/input/
        FAB_INPUT_ULY_r0.xlsx     지금 값. 엑셀로 직접 열어도 되는 그 파일이다
        _history/FAB_INPUT_ULY_r0/20260917_191817_a3f1_hong.xlsx
        _audit.csv                누가 언제 어느 시트를 몇 칸 고쳤나

_history 와 _audit.csv 는 우리가 만드는 것이라 이름 앞에 _ 를 붙였다.
목록에서 빼는 기준도 그거다 -- 기준 정보 엑셀이 늘어나도 코드를 안 고치고,
우리 파일은 편집 대상으로 잡히지 않는다.
"""
from __future__ import annotations

import csv
import hashlib
import io
from datetime import datetime, timedelta, timezone

import pandas as pd

from . import s3io

KST = timezone(timedelta(hours=9))

# 기준 정보 엑셀이 있는 자리. 파일 이름을 여기 적지 않는 이유는, 파일이
# 늘 수 있고 그때마다 배포하고 싶지 않아서다 -- 이 폴더에 있는 .xlsx 를
# 전부 기준 정보로 본다.
PREFIX = s3io.FOLDER_PATH
HISTORY_KEEP = 100          # 파일 하나당 남길 이력 개수
AUDIT_COLS = ["saved_at", "user_id", "workbook", "sheet",
              "changed_cells", "rows_before", "rows_after"]


class ConcurrentEdit(Exception):
    """내가 화면에 띄운 뒤 다른 사람이 먼저 저장했다.

    그냥 덮어쓰면 그 사람이 고친 값이 소리 없이 사라진다. 누가 맞는지는
    코드가 정할 수 없으므로 여기서 멈추고 사람에게 넘긴다.
    """


def _key(*parts: str) -> str:
    return "/".join([PREFIX, *parts])


def audit_key() -> str:
    return _key("_audit.csv")


def list_workbooks() -> list[str]:
    """고칠 수 있는 엑셀 파일 이름들 (확장자 뺀 것)."""
    names = []
    for key in s3io.list_keys(PREFIX + "/"):
        name = key[len(PREFIX) + 1:]
        if "/" in name or not name.lower().endswith(".xlsx") or name.startswith("_"):
            continue                       # _history/ 안의 것과 내부 파일은 뺀다
        names.append(name[:-len(".xlsx")])
    return names


def load_workbook(book: str) -> tuple[dict[str, pd.DataFrame], str]:
    """{시트이름: DataFrame} 과 그 시점의 버전표."""
    got = s3io.get_object(_key(f"{book}.xlsx"))
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
    if base_stamp is not None and s3io.head_etag(key) != base_stamp:
        raise ConcurrentEdit(
            f"'{book}' 을(를) 화면에 띄운 뒤 다른 사람이 먼저 저장했습니다. "
            f"덮어쓰지 않았습니다 -- 다시 불러와서 고친 내용을 옮겨 주세요."
        )

    before = {}
    got = s3io.get_object(key)
    if got is not None:
        before = pd.read_excel(io.BytesIO(got[0]), sheet_name=None, dtype=object)

    # 통째로 만들어 한 번에 올린다. S3 의 put 은 그 자체로 원자적이라,
    # 올리다 끊겨도 옛 파일이 반쯤 덮어써지는 일은 없다.
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            _clean(df).to_excel(writer, sheet_name=_sheet_name(name), index=False)
    body = buf.getvalue()

    now = datetime.now(KST)
    # 이력을 먼저 올린다. 순서가 반대면, 본 파일은 바뀌었는데 이력이 없는
    # 순간이 생긴다 -- 되돌릴 것을 찾을 때 하필 그 판이 없는 쪽이 더 아프다.
    s3io.put_object(_key("_history", book, _history_name(now, user_id, body)), body)
    etag = s3io.put_object(key, body)
    _trim_history(book)
    _append_audit(now, user_id, book, before, sheets)
    return etag


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

    - 통째로 빈 줄은 뺀다. 격자에서 '행 추가' 를 눌렀다 안 채우고 저장하면
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
    return sorted(s3io.list_keys(_key("_history", book) + "/"), reverse=True)[:limit]


def _trim_history(book: str) -> None:
    keys = sorted(s3io.list_keys(_key("_history", book) + "/"))
    for old in keys[:-HISTORY_KEEP]:
        s3io.delete_object(old)


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
    return int((a != b).to_numpy().sum())


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

    got = s3io.get_object(audit_key())
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
    s3io.put_object(audit_key(), buf.getvalue().encode("utf-8-sig"))


def read_audit(book: str | None = None, limit: int = 200) -> pd.DataFrame:
    got = s3io.get_object(audit_key())
    if got is None:
        return pd.DataFrame(columns=AUDIT_COLS)
    df = pd.read_csv(io.BytesIO(got[0]), dtype=str, encoding="utf-8-sig")
    if book is not None and "workbook" in df.columns:
        df = df[df["workbook"] == book]
    return df.tail(limit).iloc[::-1].reset_index(drop=True)      # 최근 것이 위
