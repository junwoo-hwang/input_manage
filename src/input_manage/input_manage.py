"""기준 정보 관리 — 포털 메뉴 하나.

S3 drive 에 있는 기준 정보 엑셀을 브라우저에서 엑셀처럼 고치고 바로
저장한다. 내려받았다 고쳐 다시 올리는 왕복이 없다.

    G-DVC / 2GAPU/input /
        FAB_INPUT_ULY_r0.xlsx   <- 이 폴더의 .xlsx 가 곧 편집 대상 목록
        FAB_INPUT_TTS_r0.xlsx

옆에 남기는 파일은 없다. 누가 언제 무엇을 바꿨는지는 그 엑셀 안의
REV_INFO 시트에 한 줄씩 쌓인다 -- 기준 정보를 받아 보는 사람이 파일
하나만 열면 이력까지 같이 보는 것이 맞다. 이름이 _ 로 시작하는 파일은
목록에서 뺀다 (누군가 임시로 올려 둔 것을 편집 대상으로 잡지 않게).

포털에서는 show_input_manage() 하나만 부르면 된다.

화면의 격자는 sheet_grid/frontend/index.html 이다. 거기만 따로인 이유는
streamlit 컴포넌트가 iframe 에 띄우는 방식이라 파일이 나뉠 수밖에 없어서다.
"""
from __future__ import annotations

import difflib
import functools
import gc
import inspect
import json
import io
import os
import re
import threading
import zipfile
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from typing import NamedTuple
from xml.etree import ElementTree as ET

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
# 저장할 때마다 사본을 쌓아 두는 폴더 (기준 정보 폴더 바로 아래)
HISTORY_DIR = os.getenv("INPUT_S3_HISTORY_DIR", "이력")

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
    """지금 올라가 있는 것의 버전표. 없으면 빈 글자.

    '없다' 는 404 일 때뿐이다. 권한이 없거나 S3 가 잠깐 탈이 난 것까지
    '없다' 로 삼키면, 저장 직전 확인에서 '다른 사람이 먼저 저장했다' 는
    엉뚱한 말이 뜨거나 버전표가 빈 채로 남아 다음 저장이 까닭 없이 막힌다.
    """
    try:
        got = _s3_client().head_object(Bucket=BUCKET_NAME, Key=key)
    except botocore.exceptions.ClientError as err:
        if err.response.get("Error", {}).get("Code") in ("404", "NoSuchKey", "NotFound"):
            return ""
        raise
    return got.get("ETag", "").strip('"')


def put_object(key: str, data: bytes) -> str:
    """올리고 새 버전표를 돌려준다. 버전표는 올린 대답에 들어 있다 -- 확인하려고
    HEAD 를 한 번 더 보낼 것 없다 (저장 한 번에 두 번씩 헛걸음이었다)."""
    got = _s3_client().put_object(Bucket=BUCKET_NAME, Key=key, Body=data)
    return got.get("ETag", "").strip('"') or head_etag(key)


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
# 2. .xlsx 를 파이썬 기본 기능만으로 읽고 쓴다.
#
# openpyxl 을 안 쓰는 이유는 하나다: 사내 pypi 미러에 없다. 그것 하나 때문에
# 화면 전체를 못 올리는 것보다, 우리가 쓰는 만큼만 직접 다루는 쪽이 낫다.
# 여기서 필요한 것은 값뿐이고(서식은 안 다루기로 했다) xlsx 는 XML 몇 장을
# zip 으로 묶은 것이라 그 정도는 기본 기능으로 된다.
#
# 읽을 때 감당하는 것: sharedStrings 에 모인 글자(엑셀이 저장한 파일), 칸
# 안에 그대로 있는 글자(inlineStr), 수식 칸(마지막 계산값), 날짜(엑셀은
# 날짜를 수로 저장하고 서식으로만 구분한다), 참/거짓, 빈 칸, 건너뛴 칸.
# 쓸 때는 글자와 수만 쓴다.
# ======================================================================
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# 엑셀이 미리 정해 둔 날짜/시간 서식 번호. 사용자가 만든 서식은 아래에서
# 서식 문자열을 보고 가린다.
BUILTIN_DATE_FMTS = set(range(14, 23)) | {45, 46, 47}
# 엑셀의 날짜 0 일. 1900 년 윤년 버그 때문에 1899-12-30 에서 센다.
EPOCH = datetime(1899, 12, 30)
MIDNIGHT = time(0, 0)


class BadWorkbook(Exception):
    """엑셀 파일로 읽을 수 없다."""


# ----------------------------------------------------------------------
# 큰 표를 다루는 동안 파이썬의 순환 쓰레기 수거(GC)를 잠시 멈춘다.
#
# 8MB 짜리 엑셀 하나를 읽으면 XML 칸 객체가 300만 개 가까이 생긴다. GC 는
# 객체가 일정 수 늘 때마다 깨어나 '지금까지 만든 것 전부' 를 훑는데, 만드는
# 중에 그게 수십 번 일어나서 읽는 시간의 절반 이상이 거기 들었다 (실측 7.2초
# 중 4초). 멈추고 읽으면 같은 파서로 3.3초다.
#
# 멈춰도 메모리가 새지는 않는다. 파이썬은 대부분의 객체를 참조 수로 바로
# 치우고, GC 는 서로 물고 있는 고리만 치운다 -- 그것도 다시 켜면 치운다.
# 여러 사람이 동시에 읽을 수 있으므로 몇 명이 들어와 있는지 세어서, 마지막
# 사람이 나갈 때 원래대로 켠다. 원래 꺼져 있었으면 켜지 않는다.
# ----------------------------------------------------------------------
_gc_lock = threading.Lock()
_gc_inside = 0
_gc_was_on = True


def _no_gc(fn):
    @functools.wraps(fn)
    def run(*args, **kwargs):
        global _gc_inside, _gc_was_on
        with _gc_lock:
            if _gc_inside == 0:
                _gc_was_on = gc.isenabled()
                gc.disable()
            _gc_inside += 1
        try:
            return fn(*args, **kwargs)
        finally:
            with _gc_lock:
                _gc_inside -= 1
                if _gc_inside == 0 and _gc_was_on:
                    gc.enable()
    return run


def xlsx_read(data: bytes, formulas: dict | None = None) -> dict[str, list[list]]:
    """{시트이름: [[값, ...], ...]}. 첫 줄도 값으로 그대로 돌려준다.

    formulas 를 주면 거기에 {시트이름: {(줄, 칸): '수식'}} 을 채운다. 줄과
    칸은 0 부터 세는 자리이고 첫 줄(머리글)도 0 번이다. 값만 읽고 수식을
    버리면, 저장할 때 VLOOKUP 이 걸려 있던 칸이 마지막으로 계산된 값으로
    굳어 버린다.
    """
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as err:
        raise BadWorkbook(f"엑셀 파일이 아닙니다: {err}") from err

    with zf:
        names = set(zf.namelist())
        shared = _read_shared(zf) if "xl/sharedStrings.xml" in names else []
        date_styles = _read_date_styles(zf) if "xl/styles.xml" in names else set()
        out: dict[str, list[list]] = {}
        for name, path in _sheet_paths(zf).items():
            if path not in names:
                continue
            found: dict[tuple[int, int], str] = {}
            out[name] = _read_sheet(zf.read(path), shared, date_styles, found)
            if formulas is not None and found:
                formulas[name] = found
        if not out:
            raise BadWorkbook("시트를 하나도 찾지 못했습니다.")
        return out


def _read_shared(zf: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return [_rich_text(si) for si in root.findall(NS + "si")]


def _rich_text(node) -> str:
    """<si> 나 <is> 안의 글자. 글자마다 서식이 다르면 <r><t>..</t></r> 여럿으로
    나뉘어 있어 이어 붙인다. 발음 표기(<rPh>, 일본어 후리가나)는 칸의 글자가
    아니라서 뺀다."""
    out = []
    for child in node:
        if child.tag == _T:
            out.append(child.text or "")
        elif child.tag == _R:
            t = child.find(_T)
            if t is not None:
                out.append(t.text or "")
    return "".join(out)


def _read_date_styles(zf: zipfile.ZipFile) -> set[int]:
    """날짜로 보이는 서식을 쓰는 칸 스타일 번호들.

    엑셀은 날짜를 그냥 수로 저장하고 '이 칸은 날짜 서식' 이라고만 적어 둔다.
    그래서 서식을 안 보면 2026-09-20 이 46285 라는 수로 읽힌다.
    """
    root = ET.fromstring(zf.read("xl/styles.xml"))
    custom = set()
    for fmt in root.iter(NS + "numFmt"):
        if _is_date_format(fmt.get("formatCode", "")):
            custom.add(int(fmt.get("numFmtId")))

    out = set()
    xfs = root.find(NS + "cellXfs")
    if xfs is None:
        return out
    for i, xf in enumerate(xfs.findall(NS + "xf")):
        fid = int(xf.get("numFmtId", "0") or 0)
        if fid in BUILTIN_DATE_FMTS or fid in custom:
            out.add(i)
    return out


def _is_date_format(code: str) -> bool:
    """서식 문자열이 날짜/시각 서식인가.

    날짜 기호(y m d h s)만 남기고 본다. 빼는 것:
    - "…" 안의 글자 (0"m" 은 단위 m 을 붙이는 수 서식)
    - \\x 와 _x, *x (그대로 찍는 글자, 자리 채우기)
    - [Red] [$-412] [>100] 같은 꺾쇠. 회계 서식 #,##0;[Red]-#,##0 은 Red 의 d
      때문에 날짜로 잘못 읽혀, 그 서식의 수가 전부 날짜로 바뀌었다. 단 [h]
      [mm] [ss] 는 '지난 시간' 서식이라 남긴다.
    """
    bare = re.sub(r'"[^"]*"', "", code)
    bare = re.sub(r"\\.|[_*].", "", bare)
    bare = re.sub(r"\[(?![hHmMsS]+\])[^\]]*\]", "", bare)
    return bool(re.search(r"[yYmMdDhHsS]", bare))


def _sheet_paths(zf: zipfile.ZipFile) -> dict[str, str]:
    """{시트이름: zip 안의 경로}. 통합문서에 적힌 순서를 지킨다."""
    rels = {}
    if "xl/_rels/workbook.xml.rels" in zf.namelist():
        for rel in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels")):
            target = rel.get("Target", "")
            # 대부분은 workbook.xml 이 있는 xl/ 에서 센 길이지만(worksheets/
            # sheet1.xml), 앞에 / 가 붙으면 꾸러미 뿌리에서 센 길이다
            # (/xl/worksheets/sheet1.xml -- openpyxl 이 이렇게 쓴다).
            rels[rel.get("Id")] = (target[1:] if target.startswith("/")
                                   else "xl/" + target.removeprefix("./"))

    out = {}
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    for i, sheet in enumerate(root.iter(NS + "sheet"), start=1):
        name = sheet.get("name") or f"Sheet{i}"
        rid = sheet.get(NS_R + "id") or sheet.get("id")
        out[name] = rels.get(rid) or f"xl/worksheets/sheet{i}.xml"
    return out


def col_index(ref: str) -> int:
    """'A1' -> 0, 'E12' -> 4. 칸이 중간에 비어 건너뛰어도 자리를 맞추려고 쓴다."""
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def col_letter(index: int) -> str:
    """0 -> 'A', 26 -> 'AA'."""
    out = ""
    n = index
    while True:
        out = chr(65 + n % 26) + out
        n = n // 26 - 1
        if n < 0:
            return out


# ----------------------------------------------------------------------
# 수식을 다른 칸으로 옮겨 적기 (엑셀의 '채우기' 와 같은 셈)
#
# 엑셀은 수식을 아래로 끌어 채우면 그 수식을 파일에 한 번만 적는다. 맨 윗칸
# 에만 수식이 있고(<f t="shared" ref="C2:C9" si="0">VLOOKUP(E2,..)</f>),
# 나머지 칸은 '0번 공유 수식을 내 자리로 옮겨 써라'(<f t="shared" si="0"/>)
# 뿐이다. 그 칸들의 수식은 맨 윗칸 수식의 상대 참조(E2)를 줄·칸 차이만큼
# 옮겨서(E3, E4, ...) 우리가 만들어야 한다. 안 그러면 저장할 때 맨 윗칸만
# 수식으로 남고 나머지는 값으로 굳는다.
#
# 옮기는 것은 $ 가 안 붙은 쪽뿐이다 ($A$1 은 그대로, $A1 은 줄만, A$1 은
# 칸만). 글자열("A2")과 따옴표 친 시트 이름('A 1')은 건드리지 않는다.
# ----------------------------------------------------------------------
_F_SKIP = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'')
_F_REF = re.compile(
    r"(?<![\w.$\]])"                     # 이름·수의 한가운데가 아니다
    r"(?:"
    r"(?P<c1d>\$?)(?P<c1>[A-Za-z]{1,3}):(?P<c2d>\$?)(?P<c2>[A-Za-z]{1,3})"   # A:C
    r"|(?P<r1d>\$?)(?P<r1>\d+):(?P<r2d>\$?)(?P<r2>\d+)"                     # 2:5
    r"|(?P<cd>\$?)(?P<col>[A-Za-z]{1,3})(?P<rd>\$?)(?P<row>\d+)"            # E2
    r")(?![\w(\[!.])")                   # 함수 이름(LOG10()·시트 이름(Q1!)이 아니다
_MAX_COL = 16383                         # XFD
_MAX_ROW = 1048576


class _BadShift(ValueError):
    """옮긴 참조가 시트 밖으로 나간다 (엑셀이면 #REF!)."""


def _shift_col(letters: str, fixed: str, dc: int) -> str:
    if fixed:
        return fixed + letters.upper()
    c = col_index(letters.upper()) + dc
    if not 0 <= c <= _MAX_COL:
        raise _BadShift(letters)
    return col_letter(c)


def _shift_row(digits: str, fixed: str, dr: int) -> str:
    if fixed:
        return fixed + digits
    r = int(digits) + dr
    if not 1 <= r <= _MAX_ROW:
        raise _BadShift(digits)
    return str(r)


def _shift_formula(text: str, dr: int, dc: int) -> str:
    """수식을 dr 줄, dc 칸 떨어진 칸으로 옮겨 적은 꼴. 못 옮기면 _BadShift."""
    def one(m: re.Match) -> str:
        if m.group("col") is not None:
            return (_shift_col(m.group("col"), m.group("cd"), dc)
                    + _shift_row(m.group("row"), m.group("rd"), dr))
        if m.group("c1") is not None:
            return (_shift_col(m.group("c1"), m.group("c1d"), dc) + ":"
                    + _shift_col(m.group("c2"), m.group("c2d"), dc))
        return (_shift_row(m.group("r1"), m.group("r1d"), dr) + ":"
                + _shift_row(m.group("r2"), m.group("r2d"), dr))

    if not dr and not dc:
        return text
    out, pos = [], 0
    for skip in _F_SKIP.finditer(text):
        out.append(_F_REF.sub(one, text[pos:skip.start()]))
        out.append(skip.group(0))
        pos = skip.end()
    out.append(_F_REF.sub(one, text[pos:]))
    return "".join(out)


_A1 = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)")


class ArrayFormula(str):
    """배열 수식 ({=...}). 글자로는 보통 수식과 같고, 차지하는 칸 수를 들고
    다닌다 -- 쓸 때 다시 배열 수식으로 적어야 엑셀이 같게 계산한다."""
    span: tuple[int, int] = (1, 1)       # (줄 수, 칸 수)

    @classmethod
    def of(cls, text: str, ref: str | None) -> "ArrayFormula":
        out = cls(text)
        if ref and ":" in ref:
            a, b = (_A1.fullmatch(p) for p in ref.split(":", 1))
            if a and b:
                out.span = (abs(int(b[2]) - int(a[2])) + 1,
                            abs(col_index(b[1].upper()) - col_index(a[1].upper())) + 1)
        return out


def _read_sheet(raw: bytes, shared: list[str], date_styles: set[int],
                formulas: dict[tuple[int, int], str] | None = None) -> list[list]:
    rows: list[list] = []
    root = ET.fromstring(raw)
    cell_tag, f_tag = NS + "c", NS + "f"
    # 칸 이름에서 자리를 따는 것은 칸마다 한 번씩 일어난다. 15,000행 x 10칸
    # 이면 15만 번이라, 같은 칸 이름('A','B',...)의 답을 적어 두고 쓴다.
    seen: dict[str, int] = {}
    # 공유 수식 번호(si) -> (맨 윗칸 수식, 그 칸의 줄, 칸)
    shared_f: dict[str, tuple[str, int, int]] = {}
    for row in root.iter(NS + "row"):
        # 줄 번호가 건너뛰었으면 그만큼 빈 줄을 채운다
        at_row = int(row.get("r") or len(rows) + 1) - 1
        while len(rows) < at_row:
            rows.append([])
        values: list = []
        add = values.append
        for cell in row:
            if cell.tag != cell_tag:
                continue
            ref = cell.get("r")
            if ref is None:
                at = len(values)
            else:
                letters = ref.rstrip("0123456789")
                at = seen.get(letters)
                if at is None:
                    at = seen[letters] = col_index(letters)
            while len(values) < at:
                add(None)                    # 건너뛴 칸은 빈 칸이다
            add(_cell_value(cell, shared, date_styles))
            if formulas is not None:
                f = cell.find(f_tag)
                if f is not None:
                    _read_formula(f, len(rows), at, formulas, shared_f)
        rows.append(values)
    return rows


def _read_formula(f, r: int, c: int, formulas: dict, shared_f: dict) -> None:
    """수식 칸 하나. 끌어 채운 수식(공유 수식)은 제 자리의 수식으로 펼친다."""
    kind = f.get("t")
    text = f.text or ""
    if kind == "shared":
        si = f.get("si")
        if text.strip():                          # 맨 윗칸: 수식이 여기 있다
            shared_f[si] = (text, r, c)
            formulas[(r, c)] = text
        elif si in shared_f:                      # 나머지: 맨 윗칸 것을 옮겨 쓴다
            base, r0, c0 = shared_f[si]
            try:
                formulas[(r, c)] = _shift_formula(base, r - r0, c - c0)
            except _BadShift:
                pass                              # 엑셀이면 #REF! -- 값으로 둔다
    elif kind == "array":
        if text.strip():
            formulas[(r, c)] = ArrayFormula.of(text, f.get("ref"))
    elif kind == "dataTable":
        pass            # '데이터 표' 는 수식이 아니다 -- 계산된 값만 둔다
    elif text.strip():
        formulas[(r, c)] = text


_V, _IS, _T, _R = NS + "v", NS + "is", NS + "t", NS + "r"


def _cell_value(cell, shared: list[str], date_styles: set[int]):
    kind = cell.get("t", "n")
    if kind == "inlineStr":
        node = cell.find(_IS)
        if node is None:
            return None
        # 글자 하나짜리가 거의 전부다 (칸 안에서 서식이 갈리지 않는 한).
        # 그 경우를 먼저 쳐내면 15만 번의 join 과 generator 를 아낀다.
        if len(node) == 1 and node[0].tag == _T:
            return node[0].text or ""
        return _rich_text(node)
    if kind == "s":                                   # sharedStrings 색인
        v = cell.find(_V)
        if v is None or v.text is None:
            return None
        i = int(v.text)
        return shared[i] if 0 <= i < len(shared) else None
    if kind in ("str", "e"):                          # 수식 결과 / 오류
        v = cell.find(_V)
        return v.text if v is not None else None

    v = cell.find(_V)
    if v is None or not v.text:
        return None
    if kind == "b":
        return v.text not in ("0", "false", "FALSE")
    try:
        number = float(v.text)
    except ValueError:
        return v.text
    style = int(cell.get("s", "0") or 0)
    if style in date_styles:
        return _to_datetime(number)
    return int(number) if number.is_integer() else number


def _to_datetime(serial: float):
    """엑셀의 날짜 수를 날짜로. 범위를 벗어나면 수 그대로 둔다.

    시각이 0시 0분이면 date 로 돌려준다. 화면에는 글자로 찍히는데
    '2026-09-20 00:00:00' 보다 '2026-09-20' 이 읽기 낫고, 기준 정보에
    들어 있는 날짜는 죄다 날짜뿐이라 시각이 의미가 없다.
    """
    try:
        at = EPOCH + timedelta(days=float(serial))
    except (OverflowError, ValueError):
        return serial
    return at.date() if at.time() == MIDNIGHT else at


# 손봐야 할 글자가 하나라도 있는가. 대부분의 칸은 하나도 없어서, 이것
# 하나로 걸러내면 나머지 치환을 통째로 건너뛴다.
_XML_SPECIAL = re.compile('[&<>"\x00-\x08\x0b\x0c\x0e-\x1f]')
# 엑셀이 못 읽는 제어문자 (탭 \t, 줄바꿈 \n \r 은 된다)
_XML_CTRL = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _esc(text) -> str:
    out = str(text)
    if not _XML_SPECIAL.search(out):
        return out
    out = out.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    out = out.replace('"', "&quot;")
    # 엑셀이 못 읽는 제어문자는 뺀다. 붙여넣기로 섞여 들어오면 파일 자체가
    # 안 열리는데, 그게 제일 알아채기 어려운 고장이다.
    return _XML_CTRL.sub("", out)


# 엑셀이 수로 읽는 꼴. 파이썬의 float() 는 'nan', 'inf', '1_000' 도 받아 주는데,
# 그걸 수로 적으면 엑셀이 파일을 열 때 '복구' 창을 띄운다.
_EXCEL_NUM = re.compile(r"^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")


def _looks_numeric(text: str) -> bool:
    return bool(_EXCEL_NUM.match(text))


def _sheet_xml(rows: list[list],
               formulas: dict[tuple[int, int], str] | None = None,
               strings: dict[str, int] | None = None) -> bytes:
    """시트 한 장의 XML.

    글자는 엑셀처럼 통합문서 한 곳(sharedStrings)에 모으고 칸에는 그 번호만
    적는다. strings 가 그 모음이고, 시트끼리 나눠 쓴다. 칸마다 글자를 그대로
    넣던 때(inlineStr)보다 XML 이 절반 가까이 줄어서 쓰기도, 나중에 다시
    읽기도 그만큼 빠르다. 같은 글자('Y', 'PRE', ...)는 한 번만 적힌다.
    """
    formulas = formulas or {}
    strings = {} if strings is None else strings
    f_rows = {r for r, _c in formulas}
    width = max(map(len, rows), default=0)
    letters = [col_letter(c) for c in range(width)]   # 칸마다 새로 셀 것 없다
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/'
             'spreadsheetml/2006/main"><sheetData>']
    add = parts.append
    for r0, row in enumerate(rows):
        r = r0 + 1
        has_f = r0 in f_rows
        if not has_f and not any(v is not None and v != "" for v in row):
            continue                                  # 빈 줄은 안 쓴다
        add(f'<row r="{r}">')
        for c, value in enumerate(row):
            if has_f:
                formula = formulas.get((r0, c))
                if formula is not None:
                    # 수식은 수식으로 쓴다. 마지막으로 계산된 값도 같이 적어
                    # 둬야 엑셀로 열기 전에(우리 화면에서) 빈 칸으로 보이지
                    # 않는다. 진짜 값은 엑셀이 열 때 다시 계산한다(아래
                    # fullCalcOnLoad).
                    cached = ("" if value is None else str(value))
                    kind = "" if _looks_numeric(cached) else ' t="str"'
                    body = f"<v>{_esc(cached)}</v>" if cached else ""
                    ftag = "<f>"
                    if isinstance(formula, ArrayFormula):
                        nr, nc = formula.span
                        end = (f":{col_letter(c + nc - 1)}{r + nr - 1}"
                               if (nr, nc) != (1, 1) else "")
                        ftag = f'<f t="array" ref="{letters[c]}{r}{end}">'
                    add(f'<c r="{letters[c]}{r}"{kind}>{ftag}{_esc(formula)}</f>'
                        f'{body}</c>')
                    continue
            if value is None or value == "":
                continue
            kind = type(value)
            if kind is str:
                at = strings.get(value)
                if at is None:
                    at = strings[value] = len(strings)
                add(f'<c r="{letters[c]}{r}" t="s"><v>{at}</v></c>')
            elif kind is int or kind is float:
                add(f'<c r="{letters[c]}{r}"><v>{value}</v></c>')
            elif isinstance(value, bool):
                add(f'<c r="{letters[c]}{r}" t="b"><v>{1 if value else 0}</v></c>')
            elif isinstance(value, datetime):
                # 엑셀의 날짜는 날 수 + 날짜 서식이다. 글자로 적으면 그 칸이
                # 날짜가 아니게 되어, pandas 로 읽는 쪽에서 형이 바뀐다.
                add(f'<c r="{letters[c]}{r}" s="2"><v>{_serial(value)!r}</v></c>')
            elif isinstance(value, date):
                add(f'<c r="{letters[c]}{r}" s="1"><v>{int(_serial(value))}</v></c>')
            elif isinstance(value, (int, float)):
                add(f'<c r="{letters[c]}{r}"><v>{value}</v></c>')
            else:
                text = str(value)
                if text == "":
                    continue
                at = strings.get(text)
                if at is None:
                    at = strings[text] = len(strings)
                add(f'<c r="{letters[c]}{r}" t="s"><v>{at}</v></c>')
        add("</row>")
    add("</sheetData></worksheet>")
    return "".join(parts).encode("utf-8")


def _shared_xml(strings: dict[str, int]) -> bytes:
    """모은 글자들을 sharedStrings.xml 로. 번호 순서대로 적는다."""
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
             f' uniqueCount="{len(strings)}">']
    add = parts.append
    for text in strings:                      # dict 는 넣은 순서 = 번호 순서
        # 앞뒤 빈칸은 이렇게 적어 둬야 엑셀이 안 떼어 먹는다
        if text[:1].isspace() or text[-1:].isspace():
            add(f'<si><t xml:space="preserve">{_esc(text)}</t></si>')
        else:
            add(f"<si><t>{_esc(text)}</t></si>")
    add("</sst>")
    return "".join(parts).encode("utf-8")


def xlsx_write(sheets: dict[str, list[list]],
               formulas: dict[str, dict] | None = None) -> bytes:
    """{시트이름: [[값, ...], ...]} -> .xlsx 바이트.

    formulas 는 {시트이름: {(줄, 칸): '수식'}}. 그 칸은 값 대신 수식으로
    쓰고, 마지막으로 계산된 값을 함께 적어 둔다.
    """
    formulas = formulas or {}
    names = list(sheets) or ["Sheet1"]
    # 시트를 먼저 만든다 -- 글자 모음(sharedStrings)은 시트를 다 훑어야 나온다
    strings: dict[str, int] = {}
    sheet_xml = [_sheet_xml(sheets.get(name, []), formulas.get(name), strings)
                 for name in names]
    buf = io.BytesIO()
    # 압축 단계는 기본(6) 그대로 둔다. 1 로 낮추면 8MB 짜리에서 0.8초 빨라지지만
    # 파일이 2MB 커지는데, 그 파일은 저장 한 번에 두 번(이력, 본 파일) 올라가고
    # 열 때마다 내려온다.
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
            'package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + "".join(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType='
                f'"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for i in range(1, len(names) + 1))
            + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.'
              'openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
              '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
              'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
              '</Types>')

        zf.writestr("_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats'
            '.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')

        zf.writestr("xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<sheets>"
            + "".join(f'<sheet name="{_esc(n)}" sheetId="{i}" r:id="rId{i}"/>'
                      for i, n in enumerate(names, start=1))
            + "</sheets>"
            # 우리는 VLOOKUP 을 계산하지 못한다. 적어 둔 값은 마지막으로
            # 엑셀이 계산한 것이라 낡았을 수 있으므로, 열 때 다시 계산하라고
            # 적어 둔다.
            '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')

        zf.writestr("xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships">'
            + "".join(
                f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/'
                f'officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{i}.xml"/>'
                for i in range(1, len(names) + 1))
            + f'<Relationship Id="rId{len(names) + 1}" Type="http://schemas.'
              f'openxmlformats.org/officeDocument/2006/relationships/styles" '
              f'Target="styles.xml"/>'
            + f'<Relationship Id="rId{len(names) + 2}" Type="http://schemas.'
              f'openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
              f'Target="sharedStrings.xml"/></Relationships>')

        # 서식은 날짜 둘만 쓴다 (칸 서식 1 = 날짜, 2 = 날짜+시각). 날짜를 날 수로
        # 적으므로 이게 없으면 엑셀에서 46285 같은 수로 보인다.
        zf.writestr("xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>'
            '<numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
            '</cellStyleXfs><cellXfs count="3">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            '</cellXfs></styleSheet>')

        zf.writestr("xl/sharedStrings.xml", _shared_xml(strings))
        for i, body in enumerate(sheet_xml, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", body)
    return buf.getvalue()


# ======================================================================
# 3. 기준 정보를 읽고 쓴다. 화면은 여기 안 들어온다.
#
# 기준 정보는 여러 사람이 같이 고치는 값이라, 조용히 덮어써지거나 저장이
# 반쯤 되다 마는 일이 생기면 무엇이 맞는 값인지 아무도 모르게 된다. 이
# 구역은 브라우저 없이 검사할 수 있게 화면과 떼어 두었다.
# ======================================================================
class ConcurrentEdit(Exception):
    """내가 화면에 띄운 뒤 다른 사람이 먼저 저장했다.

    그냥 덮어쓰면 그 사람이 고친 값이 소리 없이 사라진다. 누가 맞는지는
    코드가 정할 수 없으므로 여기서 멈추고 사람에게 넘긴다.
    """


def _key(*parts: str) -> str:
    return "/".join([FOLDER_PATH, *parts])


def list_workbooks() -> list[str]:
    """고칠 수 있는 엑셀 파일 이름들 (확장자 뺀 것)."""
    names = []
    for key in s3.list_keys(FOLDER_PATH + "/"):
        name = key[len(FOLDER_PATH) + 1:]
        if "/" in name or not name.lower().endswith(".xlsx") or name.startswith("_"):
            continue                       # 이력 폴더 안의 것과 임시 파일은 뺀다
        names.append(name[:-len(".xlsx")])
    return names


def load_workbook(book: str,
                  formulas: dict | None = None) -> tuple[dict[str, pd.DataFrame], str]:
    """{시트이름: DataFrame} 과 그 시점의 버전표.

    formulas 를 주면 {시트이름: {(줄, 칸이름): '수식'}} 을 거기 채운다.
    줄은 머리글을 뺀 0 부터다 (DataFrame 의 줄 번호와 같다).
    """
    got = s3.get_object(_key(f"{book}.xlsx"))
    if got is None:
        return {}, ""
    data, etag = got
    return read_xlsx(data, formulas), etag


# ----------------------------------------------------------------------
# 읽은 것을 판(ETag)마다 한 번만 읽는다.
#
# 8MB 짜리 엑셀을 읽는 데 몇 초가 든다. 그런데 같은 판을 여러 번 읽는 일이
# 많다 -- 여럿이 같은 파일을 열 때, 파일을 바꿔 골랐다 돌아올 때, 아무도
# 저장하지 않았는데 '초기화' 를 누를 때. 판이 같으면 읽은 결과도 같으므로
# 한 번 읽은 것을 서버 안에서 나눠 쓴다. 판을 가리는 것은 S3 의 ETag(내용
# 해시)라서, 누가 저장하면 ETag 가 바뀌고 그때는 새로 읽는다.
#
# 먼저 HEAD 로 ETag 만 물어본다. 가진 판과 같으면 파일을 내려받지도 않는다.
#
# 여기 든 표와 수식은 여러 사람이 같이 본다 -- 고치면 안 된다. 화면은
# _seed 에서 사본을 떠서 쓴다. 파일마다 가장 최근 판 하나만 들고 있는다.
# ----------------------------------------------------------------------
_BOOKS: dict[str, tuple[str, bytes, dict, dict]] = {}
_BOOKS_LOCK = threading.Lock()
# 같은 파일을 둘이 동시에 처음 열면 둘 다 읽느라 몇 초씩 쓴다. 파일마다
# 자물쇠를 두어 뒷사람은 앞사람이 다 읽을 때까지 기다렸다가 그걸 받는다.
_BOOK_LOCKS: dict[str, threading.Lock] = {}


def fetch_workbook(book: str) -> tuple[bytes, dict[str, pd.DataFrame], dict, str]:
    """(파일 바이트, {시트: DataFrame}, 수식, 버전표). 없는 파일이면 비어 있다.

    돌려주는 표와 수식은 나눠 쓰는 것이다. 읽기만 한다.
    """
    key = _key(f"{book}.xlsx")
    with _BOOKS_LOCK:
        lock = _BOOK_LOCKS.setdefault(key, threading.Lock())
    with lock:
        stamp = s3.head_etag(key)
        with _BOOKS_LOCK:
            hit = _BOOKS.get(key)
        if hit is not None and stamp and hit[0] == stamp:
            _stamp, raw, sheets, formulas = hit
            return raw, sheets, formulas, stamp

        got = s3.get_object(key)
        if got is None:
            with _BOOKS_LOCK:
                _BOOKS.pop(key, None)
            return b"", {}, {}, ""
        raw, stamp = got
        formulas: dict = {}
        sheets = read_xlsx(raw, formulas) if raw else {}
        with _BOOKS_LOCK:
            _BOOKS[key] = (stamp, raw, sheets, formulas)
        return raw, sheets, formulas, stamp


@_no_gc
def read_xlsx(data: bytes, formulas: dict | None = None) -> dict[str, pd.DataFrame]:
    """엑셀 바이트 -> {시트이름: DataFrame}. 첫 줄이 칸 이름이다."""
    raw: dict[str, dict] = {}
    out = {name: _frame(rows)
           for name, rows in xlsx_read(data, raw).items()}
    if formulas is None:
        return out
    for name, found in raw.items():
        df = out.get(name)
        if df is None:
            continue
        cols = [str(c) for c in df.columns]
        # 칸을 번호가 아니라 이름으로 붙들어 둔다. 왼쪽에 칸을 하나
        # 끼워 넣어도 수식이 엉뚱한 칸으로 옮겨가지 않게.
        got_f = {(r - 1, cols[c]): f
                 for (r, c), f in found.items() if r >= 1 and c < len(cols)}
        if got_f:
            formulas[name] = got_f
    return out


def _frame(rows: list[list]) -> pd.DataFrame:
    """[[값...]...] 의 첫 줄을 칸 이름으로 삼아 DataFrame 으로."""
    if not rows:
        return pd.DataFrame()
    head, body = rows[0], rows[1:]
    # 줄마다 칸 수가 다를 수 있다 (엑셀은 오른쪽 빈 칸을 아예 안 적는다).
    width = max([len(head)] + [len(r) for r in body] or [0])
    cols, seen = [], {}
    for i in range(width):
        name = head[i] if i < len(head) else None
        name = f"Unnamed: {i}" if name is None or str(name) == "" else str(name)
        # 칸 이름이 겹치면 pandas 에서 df[이름] 이 Series 가 아니라 DataFrame 이
        # 되고, 그때부터 값 대신 표가 실려 나간다. 뒤엣것에 번호를 붙인다.
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        cols.append(name)
    fixed = [(list(r) + [None] * width)[:width] for r in body]
    return pd.DataFrame(fixed, columns=cols, dtype=object)


class Saved(NamedTuple):
    """저장하고 나서 알게 되는 것들.

    body 와 sheets 를 같이 돌려주는 이유: 저장한 뒤 화면을 맞추려면 방금
    올린 판이 필요한데, 그걸 S3 에서 도로 내려받아 다시 읽으면 8MB 짜리
    파일에서 몇 초가 그냥 간다. 방금 우리가 올린 것이 곧 S3 에 있는 것이다.
    """
    stamp: str                          # 새 버전표 (S3 ETag)
    body: bytes                         # 올린 파일 그대로
    sheets: dict[str, pd.DataFrame]     # 그 파일에 실제로 담긴 값


def save_workbook(book: str, sheets: dict[str, pd.DataFrame], user_id: str,
                  base_stamp: str | None = None,
                  formulas: dict | None = None) -> Saved:
    """고친 값을 올리고 새 버전표를 돌려준다.

    base_stamp 를 주면 그 사이에 다른 사람이 올렸는지 보고, 그랬으면
    아무것도 쓰지 않고 ConcurrentEdit 를 던진다.
    """
    key = _key(f"{book}.xlsx")

    def check() -> None:
        if base_stamp is not None and s3.head_etag(key) != base_stamp:
            raise ConcurrentEdit(
                f"'{book}' 을(를) 화면에 띄운 뒤 다른 사람이 먼저 저장했습니다. "
                f"덮어쓰지 않았습니다 -- 다시 불러와서 고친 내용을 옮겨 주세요."
            )

    check()                  # 먼저 한 번 -- 몇 초씩 파일을 만들기 전에 알린다

    # 통째로 만들어 한 번에 올린다. S3 의 put 은 그 자체로 원자적이라,
    # 올리다 끊겨도 옛 파일이 반쯤 덮어써지는 일은 없다.
    body, written = build_xlsx(sheets, formulas)

    # 이력 폴더에 한 벌 먼저 넣는다. 본 파일을 먼저 덮어쓰고 나면, 그 뒤에
    # 이력 넣기가 실패했을 때 되돌릴 것이 없는 채로 끝난다. 순서를 이렇게
    # 두면 '사본을 못 남기면 덮어쓰지도 않는다' 가 된다.
    s3.put_object(_history_key(book, user_id), body)

    # 본 파일을 올리기 바로 앞에서 한 번 더 본다. 파일을 만들고 이력을 올리는
    # 몇 초 사이에 누가 저장했으면, 처음 확인만으로는 그걸 덮어쓴다.
    check()
    return Saved(s3.put_object(key, body), body, written)


def _history_key(book: str, user_id: str, now: datetime | None = None) -> str:
    """이력 폴더에 넣을 이름. 260921_원래이름_junwoo.hwang.xlsx

    날짜가 앞에 오므로 폴더를 이름순으로 보면 그대로 시간순이 된다.

    같은 사람이 같은 날 두 번 저장하면 이름이 겹친다. 그대로 두면 앞의 것이
    조용히 덮어써지는데, 되돌릴 판이 하나 사라지는 셈이라 그럴 수 없다.
    겹치면 뒤에 번호를 붙인다.
    """
    now = now or datetime.now(KST)
    base = f"{now:%y%m%d}_{book}_{_safe(user_id)}"
    # 겹칠 수 있는 것은 이름이 이렇게 시작하는 것뿐이다. 폴더를 통째로 훑으면
    # 저장할수록 이력이 쌓여 점점 느려진다.
    taken = {k.rsplit("/", 1)[-1] for k in s3.list_keys(_key(HISTORY_DIR, base))}
    name = f"{base}.xlsx"
    n = 2
    while name in taken:
        name = f"{base}_{n}.xlsx"
        n += 1
    return _key(HISTORY_DIR, name)


def _safe(text: str) -> str:
    """파일 이름에 넣어도 되는 꼴로. 빈 값이면 'unknown'.

    사번이나 이름이 그대로 들어오므로 / 나 .. 가 섞이면 이력이 폴더 밖에
    떨어진다.
    """
    kept = "".join(c for c in str(text or "") if c.isalnum() or c in "-_.")
    return kept.strip(".")[:40] or "unknown"


def to_xlsx(sheets: dict[str, pd.DataFrame],
            formulas: dict | None = None) -> bytes:
    """시트들을 엑셀 파일 한 벌로."""
    return build_xlsx(sheets, formulas)[0]


@_no_gc
def build_xlsx(sheets: dict[str, pd.DataFrame],
               formulas: dict | None = None) -> tuple[bytes, dict]:
    """엑셀 파일 한 벌과, 그 안에 실제로 담긴 값.

    formulas 는 {시트이름: {(줄, 칸이름): '수식'}}. 그 칸은 값 대신 수식으로
    나간다 -- 안 그러면 VLOOKUP 이 걸려 있던 칸이 마지막으로 계산된 값으로
    굳어 버린다.

    같이 적어 두는 캐시 값(=엑셀이 열 때까지 화면에 보여줄 값이자, 이 파일을
    pandas 등으로 직접 읽는 쪽이 곧이곧대로 가져가는 값)은 정확매칭 VLOOKUP
    에 한해 지금 값으로 다시 계산한다 -- refresh_formula_cache 참고.
    """
    formulas = formulas or {}
    sheets = refresh_formula_cache(sheets, formulas)
    out: dict[str, list[list]] = {}
    out_f: dict[str, dict] = {}
    for name, df in sheets.items():
        key = _sheet_name(name)
        while key in out:                    # 31자로 자르다 보면 겹칠 수 있다
            key = key[:30] + "_"
        want = formulas.get(name) or {}
        clean = _clean(df, keep={row for row, _col in want})
        cols = [str(c) for c in clean.columns]
        # 이름이 없던 머리글은 읽을 때 'Unnamed: 5' 로 부른다. 쓸 때는 도로
        # 빈칸으로 -- 안 그러면 원래 비어 있던 머리글에 그 글자가 박힌다.
        head = ["" if c == f"Unnamed: {i}" else c for i, c in enumerate(cols)]
        # 칸이 146만 개라 칸마다 부르는 함수 한 겹도 1초 가까이 된다. 격자가
        # 올려준 값은 거의 다 글자이고 나머지는 빈 칸이라, 그 둘은 바로 처리한다.
        out[key] = ([head] + [
            [_number(v) if type(v) is str else (None if v is None else _cell(v))
             for v in row]
            for row in clean.to_numpy(dtype=object).tolist()])

        if not want:
            continue
        # 줄 이름 -> 파일 안의 몇 번째 줄. 끝의 빈 줄만 빠지므로 자리는 그대로다.
        moved = {old: new for new, old in enumerate(clean.index)}
        at = {col: i for i, col in enumerate(cols)}
        placed = {}
        for (row, col), text in want.items():
            if row in moved and col in at:
                placed[(moved[row] + 1, at[col])] = text   # +1 은 머리글 줄
        if placed:
            out_f[key] = placed
    return xlsx_write(out, out_f), sheets


def _cell(value):
    """엑셀 칸에 넣을 꼴로. 수는 수로, 나머지는 글자로 둔다."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        # 격자는 모든 값을 글자로 올려보낸다. 수로 적힌 것은 수로 되돌려야
        # 엑셀에서 정렬과 합계가 되고, 다시 읽었을 때 값이 그대로다.
        return _number(value)
    if isinstance(value, (datetime, date)):
        return value                     # 날짜로 적는다 (_sheet_xml)
    return str(value)


def _number(text: str):
    """'12' -> 12, '1.5' -> 1.5, 아니면 글자 그대로.

    수로 바꿨다 되돌렸을 때 글자가 한 자도 다르지 않을 때만 바꾼다. 그래서
    '0010' 은 10 이 되지 않고(앞의 0 은 품번에서 뜻이 있다) '1.50' 도 1.5 가
    되지 않는다(사람이 적은 자릿수다). 사람이 친 글자를 조용히 고치는 것은
    저장이 할 일이 아니다.
    """
    if not text or text[0] not in "-0123456789":
        return text
    try:
        number = int(text) if "." not in text else float(text)
    except ValueError:
        return text
    return number if str(number) == text else text


def _sheet_name(name: str) -> str:
    """엑셀 시트 이름 규칙에 맞춘다 (31자, : \\ / ? * [ ] 못 씀)."""
    clean = "".join(" " if c in ':\\/?*[]' else c for c in str(name))
    return clean[:31] or "Sheet1"


def _blank(row) -> bool:
    return not any(v is not None and (v.strip() if type(v) is str else str(v).strip())
                   for v in row)


def _clean(df: pd.DataFrame, keep=frozenset()) -> pd.DataFrame:
    """엑셀로 나가기 전에 다듬는다.

    - None/NaN 은 빈 칸으로 쓴다 ("nan" 이라는 글자로 저장되지 않게).
    - **맨 끝의** 통째로 빈 줄만 뺀다 ('행 아래' 를 눌렀다 안 채운 것). 단 keep
      에 든 줄(수식이 있는 줄 -- 수식 결과가 빈 글자면 빈 줄처럼 보인다)은
      빼지 않는다.

    가운데의 빈 줄은 두지 않고 빼던 때가 있었는데, 그러면 그 아래 줄이 한 줄씩
    당겨지면서 수식 안의 줄 번호(=VLOOKUP(E10220,..))는 따라 바뀌지 않아, 엑셀로
    열어 다시 계산하면 한 줄씩 어긋난 값을 끌어왔다. 원래 파일에 있던 빈 줄은
    그 자리에 그대로 있어야 한다.
    """
    out = df.copy().where(pd.notna(df), None)
    if not len(out):
        return out
    rows = out.to_numpy(dtype=object).tolist()
    labels = list(out.index)
    end = len(rows)
    while end and labels[end - 1] not in keep and _blank(rows[end - 1]):
        end -= 1
    return out.iloc[:end]


def _text_rows(df: pd.DataFrame | None, cols: list[str],
               rows: int | None = None) -> list[tuple[str, ...]]:
    """칸 값을 글자로 눕혀(빈 칸은 빈 글자, 앞뒤 공백은 뗀다) 줄마다 튜플로.

    같은 일을 pandas 로 하면 칸 하나 꺼낼 때마다 pandas 를 한 겹씩 거친다.
    52,000줄 x 28칸이면 그 칸이 146만 개라, '무엇이 바뀌었나' 하나 세는 데
    10초가 걸렸다. 여기서는 칸 하나씩 꺼내 쓰는 쪽이 전부 파이썬 튜플이다.

    칸은 cols 순서로 맞추고 df 에 없는 칸은 빈 글자다. 자리는 줄 번호가
    아니라 몇 번째 줄인지로 센다.
    """
    n_have = 0 if df is None else len(df)
    n = n_have if rows is None else rows
    if not cols:
        return [()] * n
    if df is None or not n_have:
        return [("",) * len(cols)] * n
    where = {str(c): i for i, c in enumerate(df.columns)}
    arr = df.to_numpy(dtype=object)
    blank = pd.isna(arr)
    columns = []
    for name in cols:
        i = where.get(name)
        if i is None:
            columns.append([""] * n)
            continue
        col = ["" if gone else (v.strip() if type(v) is str else str(v).strip())
               for v, gone in zip(arr[:n, i].tolist(), blank[:n, i].tolist())]
        if len(col) < n:
            col.extend([""] * (n - len(col)))
        columns.append(col)
    return list(zip(*columns))


# ----------------------------------------------------------------------
# 무엇이 바뀌었나 -- 저장 전에 사람에게 보여 줄 것
# ----------------------------------------------------------------------
REV_SHEET = "REV_INFO"
# 그 시트에 적을 칸들. 없는 칸은 건너뛰고, 있는 칸만 채운다.
REV_DATE, REV_REMARK, REV_USER, REV_LINK = "Date", "Remark", "user", "관련"


def _trim(rows: list[tuple]) -> list[tuple]:
    """맨 끝의 통째로 빈 줄을 뗀다."""
    end = len(rows)
    while end and not any(rows[end - 1]):
        end -= 1
    return rows[:end]


def row_changes(before: pd.DataFrame | None, after: pd.DataFrame,
                limit: int = 300) -> tuple[list[dict], int]:
    """어느 줄이 어떻게 바뀌었는지. ([{kind, row, values}], 전체 개수)

    kind 는 '신규' / '수정' / '삭제' 다. 자리만 비교하면(첫 줄부터 차례로)
    가운데에 줄 하나를 끼워 넣었을 때 그 아래가 전부 '수정' 으로 나온다.
    그래서 difflib 으로 '무엇이 그대로이고 무엇이 끼어들었나' 를 먼저 맞춘다.

    limit 은 팝업에 띄울 개수다. 몇 천 줄을 붙여넣고 저장하는 일이 있는데,
    그걸 다 그리면 팝업이 안 뜬다. 넘치는 것은 개수로만 알린다.
    """
    cols = list(dict.fromkeys([*map(str, (before.columns if before is not None else [])),
                               *map(str, after.columns)]))
    # 맨 끝의 빈 줄은 저장할 때 빠진다 ('행 아래' 를 눌렀다 안 채운 것). 그건
    # 바뀐 것이 아니므로 양쪽에서 떼고 견준다. 가운데의 빈 줄은 그대로 저장
    # 되므로(_clean) 새로 끼운 빈 줄도 '신규' 로 센다 -- 아래 줄이 밀린다.
    old_rows = _trim(_text_rows(before, cols)) if before is not None else []
    new_rows = _trim(_text_rows(after, cols))

    out: list[dict] = []
    total = 0
    if old_rows == new_rows:
        # 대개는 여기서 끝난다. 시트가 다섯 장이어도 고친 것은 한두 장이라,
        # 나머지는 줄 맞추기(difflib)를 돌릴 것도 없이 같은지만 보면 된다.
        return out, total

    def add(kind: str, n: int, values: tuple):
        nonlocal total
        total += 1
        if len(out) < limit:
            out.append({"kind": kind, "row": n,
                        "values": dict(zip(cols, values))})

    matcher = difflib.SequenceMatcher(None, old_rows, new_rows, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag in ("replace", "insert"):
            for j in range(j1, j2):
                add("수정" if tag == "replace" else "신규", j + 1, new_rows[j])
        if tag in ("replace", "delete"):
            for i in range(i1, i2):
                if tag == "replace" and i - i1 < j2 - j1:
                    continue            # 같은 자리의 '수정' 으로 이미 셌다
                add("삭제", i + 1, old_rows[i])
    return out, total


# ----------------------------------------------------------------------
# 정확매칭 VLOOKUP 을 우리가 대신 계산한다
#
# 우리는 엑셀이 아니라서 수식을 계산하지 않는다. 그래서 저장할 때 같이
# 적어 두는 캐시 값은 '마지막으로 엑셀이 계산했을 때' 값 그대로인데, 같은
# 파일 안의 다른 시트를 고친 뒤에도 그 값이 그대로면 엑셀로 열지 않고
# pandas 등으로 직접 읽는 쪽은 낡은 값을 그대로 가져간다.
#
# VLOOKUP(키, 시트!범위, 열번호, 0) 꼴의 정확매칭만 계산한다. 그 이유:
#   - 마지막 인자가 0/FALSE 인 정확매칭은 '이 값과 완전히 같은 줄을 찾는다'
#     는 뜻이라 계산이 명확하다. 근사매칭(정렬돼 있다고 가정하는 것)은
#     엑셀의 이분 탐색을 그대로 흉내 내야 해서 잘못 계산할 위험이 크다 --
#     안 하느니만 못하다.
#   - 다른 함수가 섞였거나 이 꼴에 안 맞으면 계산하지 않는다. 그 칸은
#     여전히 낡은 채로 남는다 (지금까지와 같다 -- 더 나빠지지 않는다).
# ----------------------------------------------------------------------
_NOT_EVALUATED = object()

# 끌어 채운 VLOOKUP 은 찾을 칸(E2, E3, ...)만 다르고 뒤(,ET추출여부!$A:$C,3,0))는
# 같다. 그래서 둘로 나눠 뒤쪽은 한 번만 풀어 두고 나눠 쓴다 (_Grids.spec) --
# 수식 32,000개를 하나하나 통째로 풀면 그것만 0.4초였다.
_VL_HEAD = re.compile(r"^\s*VLOOKUP\(\s*([^,]+?)\s*,(.*)$", re.IGNORECASE | re.DOTALL)
_VL_TAIL = re.compile(
    r"^\s*(?:(?P<sheet>'(?:[^']|'')+'|[^!,']+)!)?"         # 시트 이름 (없으면 제 시트)
    r"\$?(?P<c1>[A-Za-z]{1,3})\$?(?P<r1>\d*)"               # $A 또는 $A$2
    r":\$?(?P<c2>[A-Za-z]{1,3})\$?(?P<r2>\d*)"               # $C 또는 $C$500
    r"\s*,\s*(?P<idx>\d+)\s*,\s*(?:0|FALSE)\s*\)\s*$",     # 정확매칭만
    re.IGNORECASE)
# 첫 인자가 수로 적힌 것 (VLOOKUP(1001, ...))
_NUM_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
# 첫 인자가 같은 시트의 칸 자리를 가리키는 경우 (E10220 처럼). $ 는 있어도
# 없어도 된다 -- 엑셀에서 상대/절대 참조 표기 차이일 뿐 우리에겐 같다.
_CELL_REF_RE = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+)$')


def refresh_formula_cache(sheets: dict[str, pd.DataFrame],
                          formulas: dict) -> dict[str, pd.DataFrame]:
    """VLOOKUP 정확매칭이 걸린 칸의 값을 지금 시트 값으로 다시 계산한다.

    formulas 는 {시트이름: {(줄, 칸이름): '수식'}} (surviving_formulas 가
    돌려준 것과 같은 꼴). 계산에 성공한 칸만 그 시트의 사본에서 값을
    바꾸고, 나머지 시트는 원래 것을 그대로 돌려준다 -- 계산 못 한 칸은
    그대로 낡은 값이 남는다.
    """
    if not formulas:
        return sheets
    out = dict(sheets)
    grids = _Grids(out)
    for name, cells in formulas.items():
        df = out.get(name)
        if df is None or not cells:
            continue
        # 수식의 첫 인자(찾을 값)는 고치기 전의 이 시트에서 읽는다
        own = grids.values(name)
        where = {c: i for i, c in enumerate(df.columns)}
        rows_at = _positions(df.index)
        touched = None
        for (row, col), text in cells.items():
            ci, ri = where.get(col), rows_at(row)
            if ci is None or ri is None:
                continue
            got = _eval_vlookup(text, own, name, grids)
            if got is _NOT_EVALUATED:
                continue
            if touched is None:
                touched = own.copy()
            touched[ri, ci] = got
        if touched is not None:
            out[name] = pd.DataFrame(touched, index=df.index, columns=df.columns,
                                     dtype=object)
            # 이 시트의 값이 갈렸으니 들고 있던 것도 버린다. 다른 시트의
            # 수식이 이 시트를 찾아본다면 방금 고친 값으로 찾아야 맞다.
            grids.forget(name)
    return out


def _positions(index: pd.Index):
    """줄 이름 -> 몇 번째 줄인지. 없는 이름이면 None."""
    if isinstance(index, pd.RangeIndex) and index.start == 0 and index.step == 1:
        n = len(index)

        def at(row):
            try:
                i = row.__index__()           # int 이든 numpy 정수든
            except AttributeError:
                return None
            return i if 0 <= i < n else None
        return at
    found = {label: i for i, label in enumerate(index)}
    return found.get


class _Grids:
    """VLOOKUP 을 계산하는 동안 시트들을 들고 있는 곳.

    시트마다 한 번만 값 배열로 꺼내 두고, 찾을 표는 한 번만 뒤집어 둔다.
    수식 하나를 계산할 때마다 pandas 에서 칸을 하나씩 꺼내면(.iloc, .at)
    그 한 번이 수십 마이크로초라, 줄마다 수식이 걸린 16,000줄 시트에서는
    그것만 몇 초였다. 찾을 표를 수식마다 처음부터 훑으면 수식 개수 x 표 줄수가
    되어 저장이 분 단위로 걸렸다.
    """

    def __init__(self, sheets: dict[str, pd.DataFrame]):
        self.sheets = sheets
        self._values: dict[str, object] = {}
        self._index: dict[tuple[str, int], dict[str, int]] = {}
        self._specs: dict[tuple[str, str], tuple | None] = {}

    def values(self, name: str):
        got = self._values.get(name)
        if got is None:
            got = self._values[name] = self.sheets[name].to_numpy(dtype=object)
        return got

    def lookup(self, name: str, col: int, lo: int = 0,
               hi: int | None = None) -> dict:
        """찾을 값 -> 그 값이 처음 나온 줄. 표 하나를 한 번만 훑는다.

        lo..hi 는 수식에 적힌 줄 범위($A$2:$C$500)를 DataFrame 자리로 옮긴
        것이다 (없으면 끝까지). 처음 나온 줄만 담는 것은 엑셀과 같다 -- 같은
        키가 여러 줄이면 VLOOKUP 은 맨 위엣것을 준다. 빈 칸은 무엇과도 안 맞는다.
        """
        got = self._index.get((name, col, lo, hi))
        if got is None:
            got = {}
            column = self.values(name)[:, col].tolist()
            stop = len(column) if hi is None else min(hi + 1, len(column))
            for row in range(max(lo, 0), stop):
                key = _lookup_norm(column[row])
                if key is not None and key not in got:
                    got[key] = row
            self._index[(name, col, lo, hi)] = got
        return got

    def spec(self, own_name: str, tail: str):
        """VLOOKUP 의 뒤쪽을 푼 것: (찾을 시트, 키 칸, 값 칸, 첫 줄, 끝 줄).
        정확매칭 꼴이 아니거나 범위가 시트 밖이면 None."""
        key = (own_name, tail)
        if key in self._specs:
            return self._specs[key]
        out = None
        m = _VL_TAIL.match(tail)
        if m:
            sheet_name = _unquote_sheet(m["sheet"]) if m["sheet"] else own_name
            target = self.sheets.get(sheet_name)
            if target is not None:
                width = len(target.columns)
                start, end = col_index(m["c1"].upper()), col_index(m["c2"].upper())
                pos = start + int(m["idx"]) - 1
                if start <= pos <= end and pos < width:
                    # 엑셀 줄 번호 N 은 DataFrame 의 N-2 번째 줄이다 (머리글이 1행)
                    lo = int(m["r1"]) - 2 if m["r1"] else 0
                    hi = int(m["r2"]) - 2 if m["r2"] else None
                    out = (sheet_name, start, pos, lo, hi)
        self._specs[key] = out
        return out

    def forget(self, name: str) -> None:
        self._values.pop(name, None)
        for key in [k for k in self._index if k[0] == name]:
            del self._index[key]


def _lookup_norm(value):
    """엑셀의 정확매칭 VLOOKUP 이 같다고 보는 값끼리 같아지는 꼴. 빈 칸은 None.

    - 글자는 대소문자를 가리지 않는다 ('ab12' 와 'AB12' 는 같다). 앞뒤 빈칸은
      떼지 않는다 ('K1 ' 과 'K1' 은 다르다).
    - 수와 글자는 다르다 (1001 과 '0010'). 격자가 올려준 '1001' 은 저장하면 수로
      적히므로(_number) 수로 본다.
    - 날짜는 엑셀처럼 수(날 수)로 본다.
    """
    if value is None or (isinstance(value, float) and value != value):
        return None
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, (int, float)):
        return ("n", float(value))
    if isinstance(value, (datetime, date)):
        return ("n", _serial(value))
    text = value if isinstance(value, str) else str(value)
    if text == "":
        return None
    number = _number(text)
    if not isinstance(number, str):
        return ("n", float(number))
    return ("s", text.casefold())


def _serial(value) -> float:
    """날짜를 엑셀의 날 수로."""
    if not isinstance(value, datetime):
        value = datetime(value.year, value.month, value.day)
    return (value.replace(tzinfo=None) - EPOCH).total_seconds() / 86400


def _unquote_sheet(name: str) -> str:
    name = name.strip()
    if len(name) >= 2 and name[0] == name[-1] == "'":
        return name[1:-1].replace("''", "'")
    return name


def _eval_vlookup(formula: str, own, own_name: str, grids: _Grids):
    """수식 하나를 지금 값으로. 못 하면 _NOT_EVALUATED.

    엑셀이 저장된 파일을 열어 다시 계산했을 때와 같은 값이어야 한다 --
    pandas 로 읽는 쪽은 이 값을 그대로 가져간다. 그래서 엑셀과 다르게 될 수
    있는 것(찾을 값이 식인 것, 와일드카드)은 계산하지 않고 둔다.

    own 은 이 수식이 들어 있는 시트의 값 배열, own_name 은 그 시트 이름이다
    (범위에 시트 이름이 없으면 제 시트에서 찾는다).
    """
    m = _VL_HEAD.match(formula)
    if not m:
        return _NOT_EVALUATED
    spec = grids.spec(own_name, m.group(2))
    if spec is None:
        return _NOT_EVALUATED
    sheet_name, start, pos, lo, hi = spec
    key = _resolve_ref(m.group(1), own)
    if key is _NOT_EVALUATED:
        return _NOT_EVALUATED
    if isinstance(key, str) and any(ch in key for ch in "*?~"):
        return _NOT_EVALUATED            # 엑셀은 와일드카드로 본다
    norm = _lookup_norm(key)
    if norm is None:
        return "#N/A"                    # 빈 칸을 찾으면 엑셀도 #N/A

    row = grids.lookup(sheet_name, start, lo, hi).get(norm)
    if row is None:
        return "#N/A"                    # 엑셀도 못 찾으면 이렇게 보여준다
    found = grids.values(sheet_name)[row, pos]
    # 찾은 칸이 비어 있으면 엑셀은 0 을 준다
    if found is None or found == "" or (isinstance(found, float) and found != found):
        return 0
    return found


def _resolve_ref(expr: str, own):
    """VLOOKUP 의 첫 인자를 값으로. 못 하면 _NOT_EVALUATED.

    받는 것은 세 가지뿐이다: 같은 시트의 칸 자리(E10220), 따옴표 친 글자
    ("K1"), 수(1001). 그 밖의 것(TRIM(E5), E5&F5, 다른시트!E5)은 계산하지
    않는다 -- 예전에는 그런 식을 글자 그대로 찾아서 엉뚱하게 #N/A 를 적었다.
    """
    expr = expr.strip()
    m = _CELL_REF_RE.match(expr)
    if m:
        letters, excel_row = m.groups()
        at_row = int(excel_row) - 2      # 머리글이 엑셀 1행이므로 -2
        c = col_index(letters.upper())
        rows, cols = own.shape
        if at_row < 0 or at_row >= rows or c >= cols:
            return _NOT_EVALUATED
        return own[at_row, c]
    if len(expr) >= 2 and expr[0] == expr[-1] == '"' and '"' not in expr[1:-1].replace('""', ""):
        return expr[1:-1].replace('""', '"')
    if _NUM_RE.match(expr):
        return float(expr)
    return _NOT_EVALUATED


@_no_gc
def surviving_formulas(before: dict[str, pd.DataFrame],
                       after: dict[str, pd.DataFrame],
                       formulas: dict) -> tuple[dict, dict[str, int]]:
    """아직 믿을 수 있는 수식만 남긴다. (남은 것, {시트: 버린 개수})

    수식은 제가 앉은 자리를 기준으로 옆 칸을 가리킨다 -- 5번 줄의
    =VLOOKUP(B5,...) 는 6번 줄로 밀리면 B6 을 봐야 맞다. 엑셀은 줄을
    끼워 넣을 때 그걸 알아서 고쳐 주지만 우리는 못 한다. 그래서 줄이
    제자리에 그대로 있을 때만 수식을 남기고, 밀린 것은 버린다.

    틀린 수식을 남겨 두는 것보다 버리는 쪽이 낫다: 남기면 조용히 엉뚱한
    값을 끌어오지만, 버리면 마지막으로 계산된 값이 그대로 보이고 저장
    전에 몇 개를 버리는지 알려 줄 수 있다.
    """
    kept: dict[str, dict] = {}
    lost: dict[str, int] = {}
    for name, want in formulas.items():
        old, new = before.get(name), after.get(name)
        if old is None or new is None:
            lost[name] = len(want)
            continue
        if old is new:
            kept[name] = dict(want)  # 손대지 않은 시트: 수식도 전부 제자리다
            continue
        cols = list(dict.fromkeys([*map(str, old.columns), *map(str, new.columns)]))
        old_rows = _text_rows(old, cols)
        new_rows = _text_rows(new, cols)
        same = (range(len(old_rows)) if old_rows == new_rows
                else _rows_in_place(old_rows, new_rows))
        at = {c: i for i, c in enumerate(cols)}
        new_cols = set(map(str, new.columns))
        live = {}
        for (row, col), text in want.items():
            if row not in same or col not in new_cols:
                continue
            # 그 칸을 사람이 직접 고쳤으면 사람이 적은 값이 이긴다
            i = at[col]
            if (row < len(old_rows) and row < len(new_rows)
                    and old_rows[row][i] != new_rows[row][i]):
                continue
            live[(row, col)] = text
        if live:
            kept[name] = live
        if len(live) < len(want):
            lost[name] = len(want) - len(live)
    return kept, lost


def _rows_in_place(old: list, new: list) -> set[int]:
    """줄 번호가 그대로인 줄들.

    '내용이 같은 줄' 이 아니라 '자리가 그대로인 줄' 이다. 같은 줄의 다른
    칸을 고친 것은 자리를 옮긴 것이 아니므로 그 줄의 수식은 멀쩡하다.
    위에 줄이 끼거나 빠져서 번호가 밀린 줄만 걸러낸다.
    """
    out = set()
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(
            None, old, new, autojunk=False).get_opcodes():
        if tag not in ("equal", "replace") or i1 != j1:
            continue
        out.update(range(i1, min(i2, i1 + (j2 - j1))))
    return out


@_no_gc
def workbook_changes(before: dict[str, pd.DataFrame],
                     after: dict[str, pd.DataFrame]) -> dict:
    """{시트이름: {"rows": [...], "total": n, "note": "..."}}. 안 바뀐 시트는 뺀다."""
    out: dict[str, dict] = {}
    for name, df in after.items():
        old = before.get(name)
        if old is df:
            continue                 # 격자가 손대지 않은 시트는 같은 표 그대로 온다
        rows, total = row_changes(old, df)
        notes = []
        if old is None:
            notes.append("새 시트")
        else:
            gone = [c for c in map(str, old.columns) if c not in map(str, df.columns)]
            new_cols = [c for c in map(str, df.columns) if c not in map(str, old.columns)]
            if new_cols:
                notes.append("칸 추가: " + ", ".join(new_cols))
            if gone:
                notes.append("칸 삭제: " + ", ".join(gone))
        if total or notes:
            out[name] = {"rows": rows, "total": total, "note": " · ".join(notes)}
    for name in before:
        if name not in after:
            out[name] = {"rows": [], "total": 0, "note": "시트 삭제"}
    return out


# 엑셀 칸 하나에 들어가는 글자 수 한계. 넘기면 엑셀이 파일을 못 연다.
CELL_MAX = 32767


def changes_text(changes: dict, limit: int = CELL_MAX) -> str:
    """'변경내용' 창에 뜬 것을 REV_INFO 칸 하나에 담을 글로.

    창에서 본 것과 같은 것이 파일에 남아야 한다 -- 나중에 '이때 뭘 바꿨지'
    를 되짚는 사람은 그 창을 못 보고 이 칸만 보기 때문이다.
    """
    out: list[str] = []
    for name, info in changes.items():
        head = f"[{name}]"
        if info["rows"]:
            head += " " + " | ".join(info["rows"][0]["values"])
        if info["note"]:
            head += f"  ({info['note']})"
        out.append(head)
        for row in info["rows"]:
            out.append(f"{row['kind']} {row['row']}행: "
                       + " | ".join(str(v) for v in row["values"].values()))
        if info["total"] > len(info["rows"]):
            out.append(f"…외 {info['total'] - len(info['rows'])}줄")
    text = "\n".join(out)
    if len(text) > limit:
        # 잘렸다는 것을 안 적으면, 뒤가 없는 것인지 잘린 것인지 알 수 없다
        tail = "\n…(너무 길어 잘림)"
        text = text[:limit - len(tail)] + tail
    return text


def pin_rev_info(sheets: dict[str, pd.DataFrame],
                 stored: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """REV_INFO 시트는 늘 S3 에 있는 그대로(stored) 둔다.

    그 시트는 저장할 때마다 우리가 한 줄씩 붙여 적는 기록이다. 화면에서
    고치거나, 고친 엑셀을 올려서 바꾸거나, 그 시트가 빠진 엑셀을 올려서
    지워지면 기록이 기록이 아니게 된다. 격자에서도 막지만(잠긴 시트) 저장은
    여기를 거치므로 여기서 한 번 더 막는다.

    그 시트가 이미 있으면 그 자리에, 없어졌으면 맨 뒤에 되살린다. S3 에 원래
    없던 파일이면 건드리지 않는다.
    """
    keep = stored.get(REV_SHEET)
    if keep is None:
        return sheets
    out = dict(sheets)
    out[REV_SHEET] = keep
    return out


def rev_columns(sheets: dict[str, pd.DataFrame]) -> list[str] | None:
    """REV_INFO 시트의 칸 이름들. 그 시트가 없으면 None."""
    df = sheets.get(REV_SHEET)
    return None if df is None else [str(c) for c in df.columns]


def append_rev_info(sheets: dict[str, pd.DataFrame], when: str, remark: str,
                    user: str, link: str,
                    detail: str = "") -> dict[str, pd.DataFrame]:
    """REV_INFO 시트 맨 아래에 이번 변경 한 줄을 붙인 사본을 돌려준다.

    맨 아래에 붙이는 이유는 그래야 위의 줄 번호가 그대로이기 때문이다.
    위에 끼워 넣으면 다음에 열었을 때 '무엇이 바뀌었나' 가 전부 한 칸씩
    밀려 보인다.
    """
    df = sheets.get(REV_SHEET)
    if df is None:
        return sheets
    # 사람이 적은 것 먼저, 그 아래에 무엇이 바뀌었는지를 붙인다
    related = "\n\n".join(x for x in (link, detail) if x)
    want = {REV_DATE: when, REV_REMARK: remark, REV_USER: user, REV_LINK: related}
    row = {}
    for col in df.columns:
        key = str(col).strip().lower()
        row[col] = next((v for k, v in want.items() if k.lower() == key), "")
    out = dict(sheets)
    out[REV_SHEET] = pd.concat(
        [df, pd.DataFrame([row], columns=df.columns)], ignore_index=True)
    return out


# ======================================================================
# 4. 격자 (streamlit 컴포넌트).
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
               max_height: int = 520, want_full: str = "") -> dict:
    """격자를 그리고, 격자가 올려준 것을 그대로 돌려준다.

    돌려주는 것: {"rev": n, "dirty": bool} 이고, 표를 달라고 했을 때만
    {"full": True, "token": ..., "sheets": [...]} 가 붙는다.

    version 은 '이 데이터가 갈렸다' 를 알리는 표다. 격자는 이 값이 바뀔
    때만 제 상태를 갈아엎는다 -- 값을 올려보낼 때마다 streamlit 이 스크립트를
    다시 돌리면서 같은 데이터가 되돌아오는데, 그때마다 새로 그리면 방금 고친
    칸과 고른 자리, 보고 있던 시트가 날아간다.

    want_full 은 '지금 표를 통째로 올려달라' 는 표다. 평소에는 빈 글자다 --
    칸 하나 고칠 때마다 15,000행을 통째로 주고받으면 한 번에 2.6초가 걸린다.
    저장할 때처럼 진짜로 값이 필요할 때만 표를 하나 들려 보낸다.

    표는 격자가 이 판을 아직 안 가졌을 때만 내려보낸다. 격자는 올려보낼
    때마다 자기가 가진 판(have)을 같이 알린다. streamlit 은 컴포넌트에 넘기는
    값이 하나라도 바뀌면 전부를 다시 보내는데, 저장을 누르면 want_full 이
    바뀌므로 그때마다 14MB 가 다시 내려갔다. 격자 틀이 새로 만들어져 표를
    잃었으면 격자가 have 를 비워 올리고, 그러면 다음 판에서 다시 보낸다.
    """
    prev = st.session_state.get(key)
    have = prev.get("have") if isinstance(prev, dict) else None
    data = None if have == version else _grid_payload(sheets, version, key)

    got = _grid(sheets=data, version=version, max_height=max_height,
                want_full=want_full, key=key, default=None)
    return got or {}


def _grid_payload(sheets: dict[str, pd.DataFrame], version: str,
                  key: str) -> list[dict]:
    """격자에 내려보낼 표. 판마다 한 번만 만든다.

    시트마다 줄들을 JSON 글자 하나로 싸 보낸다. 줄과 칸을 그대로 넘기면
    브라우저가 받는 쪽에서 146만 개의 글자를 하나씩 만들고, 격자 틀로 옮길
    때 또 하나씩 복사한다. 글자 하나로 넘기면 옮기는 것은 한 번의 복사이고,
    격자는 지금 보는 시트만 풀면 된다 (나머지는 그 시트를 열 때 푼다).
    """
    held = f"_im_grid_payload_{key}"
    got = st.session_state.get(held)
    if got is not None and got[0] == version:
        return got[1]
    payload = []
    for name, df in sheets.items():
        cols = [str(c) for c in df.columns]
        arr = df.to_numpy(dtype=object)
        blank = pd.isna(arr)
        rows = [["" if gone else (v if type(v) is str else str(v))
                 for v, gone in zip(row, holes)]
                for row, holes in zip(arr.tolist(), blank.tolist())]
        payload.append({"name": str(name), "cols": cols, "n": len(rows),
                        "locked": str(name) == REV_SHEET,
                        "rows_json": json.dumps(rows, ensure_ascii=False,
                                                separators=(",", ":"))})
    st.session_state[held] = (version, payload)
    return payload


@_no_gc
def to_frames(payload: dict,
              kept: dict[str, pd.DataFrame] | None = None) -> dict[str, pd.DataFrame]:
    """격자가 올려준 것을 {시트이름: DataFrame} 으로.

    격자는 손대지 않은 시트는 내용 없이 {"name", "keep": 원래이름} 만 올린다.
    그 시트는 kept(격자에 내려보냈던 표)에서 그대로 꺼낸다 -- 같은 객체를
    그대로 쓰므로, 뒤에서 '무엇이 바뀌었나' 를 셀 때 한눈에 안 바뀐 줄 안다.
    """
    kept = kept or {}
    out: dict[str, pd.DataFrame] = {}
    for i, sheet in enumerate(payload.get("sheets", []) or []):
        name = str(sheet.get("name") or f"Sheet{i + 1}")
        while name in out:                      # 시트 이름도 겹치면 안 된다
            name += "_"
        if "keep" in sheet:
            orig = str(sheet["keep"])
            if orig not in kept:
                # 격자와 파이썬이 서로 다른 판을 보고 있다. 없는 시트를
                # 빈 시트로 채워 저장하면 그 시트가 통째로 지워지므로 멈춘다.
                raise ValueError(f"'{orig}' 시트의 내용을 찾지 못했습니다.")
            out[name] = kept[orig]
        else:
            base = kept.get(str(sheet.get("orig") or name))
            out[name] = _restore_types(_to_frame(sheet), base)
    return out


_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_ISO_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(\.\d+)?$")


def _restore_types(df: pd.DataFrame, base: pd.DataFrame | None) -> pd.DataFrame:
    """격자는 모든 값을 글자로 돌려준다. 원래 날짜·참거짓이던 칸은 그 꼴로.

    안 그러면 한 번 저장할 때마다 날짜 칸이 글자 칸이 되어, 그 파일을
    pandas 로 읽는 쪽에서 형이 바뀐다. 원래 표(base)에서 날짜가 하나라도
    있던 칸만, 날짜 꼴('2026-09-20', '2026-09-20 13:45:00')인 글자만 되돌린다
    -- 날짜가 없던 칸에 사람이 '2026-09-20' 이라고 친 것은 글자로 둔다.
    """
    if base is None or not len(df):
        return df
    base_cols = {str(c): c for c in base.columns}
    out = None
    for col in df.columns:
        src = base_cols.get(str(col))
        if src is None:
            continue
        seen = {type(v) for v in base[src].tolist()
                if v is not None and not (isinstance(v, float) and v != v)}
        has_date = any(issubclass(t, date) for t in seen)
        has_bool = bool in seen
        if not (has_date or has_bool):
            continue
        values = df[col].tolist()
        fixed = [_typed(v, has_date, has_bool) for v in values]
        if fixed != values:
            if out is None:
                out = df.copy()
            out[col] = pd.Series(fixed, index=df.index, dtype=object)
    return df if out is None else out


def _typed(v, dates: bool, bools: bool):
    if type(v) is not str:
        return v
    if dates:
        try:
            if _ISO_DATE.match(v):
                return date.fromisoformat(v)
            if _ISO_DATETIME.match(v):
                return datetime.fromisoformat(v)
        except ValueError:
            return v
    if bools and v in ("True", "False"):
        return v == "True"
    return v


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
# 5. 화면. 포털이 부르는 것은 show_input_manage() 하나다.
# ======================================================================
S_BOOK = "_im_book"        # 지금 고르고 있는 엑셀 파일
S_SHEETS = "_im_sheets"    # 그 파일을 띄웠을 때의 원본
S_STAMP = "_im_stamp"      # 그 원본이 어느 판이었는지 (S3 ETag)
# 몇 번째로 불러온 것인지. 격자에 넘기는 판 번호에 섞는다.
#
# ETag 만으로는 모자란다: '초기화' 는 아무도 저장하지 않았으면 같은
# 파일을 다시 읽으므로 ETag 가 그대로고, 그러면 격자가 '갈린 게 없다' 며
# 제 상태를 그대로 둔다 -- 버리려고 누른 수정이 화면에 그대로 남는다.
S_NONCE = "_im_nonce"
# 저장 직후에 띄울 '저장 완료!' 창의 내용. 저장하고 바로 st.rerun() 을 하는데,
# rerun 은 스크립트를 처음부터 다시 돌리므로 그 전에 그린 것은 화면에 남지
# 않는다. 그래서 여기 맡겨 두고 다음 판에서 띄운다.
S_TOAST = "_im_toast"
# S3 에서 받아 온 파일 그대로의 바이트. '엑셀 다운로드' 가 이걸 그대로
# 내준다 -- 저장된 판을 그대로 주는 것이라 우리가 다시 만들 필요가 없고,
# 서식이든 뭐든 저장돼 있는 그대로 나간다.
S_RAW = "_im_raw"
# 격자에 띄울 값. 보통은 S_SHEETS 와 같지만, 엑셀을 업로드하면 그 내용이
# 여기 들어온다. S_SHEETS 는 S3 에 저장돼 있는 판 그대로 두어야 '무엇이
# 바뀌나' 를 견줄 기준이 남는다.
S_SHOWN = "_im_shown"
S_UPLOAD = "_im_upload"      # 업로드 창이 열려 있는가
S_UPLOADED = "_im_uploaded"  # 업로드한 내용이 아직 저장 안 됐는가
# 격자에 '표를 통째로 올려달라' 고 하면서 들려 보낸 표. 격자가 그 표를 달고
# 올려주면 그때가 우리가 요청한 그 값이다.
#
# 이 왕복이 필요한 이유: 평소에 격자는 '고친 게 있다' 만 올린다. 칸 하나
# 고칠 때마다 15,000행을 통째로 주고받으면 한 번에 2.6초가 걸리기 때문이다.
# 그래서 저장처럼 값이 진짜 필요한 순간에만 달라고 한다.
S_WANT = "_im_want"
S_PENDING = "_im_pending"    # 표를 받으면 할 일: 지금은 "save" 뿐
S_EDITED = "_im_edited"      # 격자가 올려준 지금 값
S_REVIEW = "_im_review"      # 저장 팝업에 띄울 변경 내역
# 그때 같이 셈해 둔 (살아남는 수식, 시트별로 버리는 개수). 창이 떠 있는
# 동안에는 값이 안 바뀌므로 다시 셀 일이 없다 -- 15,000행에서 한 번 세는 데
# 0.14초라, 사유를 한 글자 칠 때마다 다시 세면 창이 그만큼씩 굼떠진다.
S_KEPT = "_im_kept"
# 저장 창에서 저장을 눌렀다: {"remark", "link"}. 누른 판에서는 이것만 적고
# 곧바로 다시 그린다 -- 그래야 저장이 도는 몇 초 동안 단추가 회색으로 보인다.
# 누른 판에서 바로 저장하면 그동안 화면은 누르기 전 그대로라 단추가 켜져
# 있고, 한 번 더 누를 수 있다.
S_SAVING = "_im_saving"
S_SAVE_ERR = "_im_save_err"  # 저장이 실패한 까닭. 창을 다시 켜서 보여 준다
# 파일을 띄웠을 때 그 안에 있던 수식들. 격자는 값만 다루므로, 이걸 안 들고
# 있으면 저장할 때 VLOOKUP 이 걸려 있던 칸이 마지막 계산값으로 굳어 버린다.
S_FORMULAS = "_im_formulas"


# 칸 너비를 꽉 채우라고 말하는 법이 streamlit 버전마다 다르다. 새 버전은
# width="stretch", 예전 버전은 use_container_width=True 다. 포털이 어느
# 버전인지 모르는 채로 한쪽만 쓰면 화면이 아예 안 뜬다 (TypeError).
_WIDE = ({"width": "stretch"}
         if "width" in inspect.signature(st.button).parameters
         else {"use_container_width": True})

# 이름을 우리가 직접 그리므로 위젯이 제 이름을 또 그리지 않게 한다.
# label_visibility 는 예전 streamlit 에 없어서, 없으면 그냥 두 번 나온다.
_NO_LABEL = ({"label_visibility": "collapsed"}
             if "label_visibility" in inspect.signature(st.selectbox).parameters
             else {})

# 폼 단추를 끄는 법(disabled)도 예전 streamlit 에는 없다. 없으면 그냥 켜진
# 채로 둔다 -- 두 번 눌러도 두 번 저장되지는 않는다 (저장 중 표시를 먼저 본다).
def _off(on: bool) -> dict:
    return ({"disabled": on}
            if "disabled" in inspect.signature(st.form_submit_button).parameters
            else {})


# st.dialog 는 예전 streamlit 에 없다. 없으면 팝업 대신 화면 안에 펼쳐서
# 보여준다 -- 보기는 덜 좋아도 저장 전에 확인하는 절차는 그대로 지킨다.
_HAS_DIALOG = hasattr(st, "dialog")


def _load(book: str) -> None:
    """S3 에서 그 파일을 다시 읽어 화면 상태를 처음으로 되돌린다.

    받아 온 바이트를 그대로 들고 있는다 -- '엑셀 다운로드' 가 그걸 그대로
    내주기 때문이다. 우리가 다시 만들어 주면 저장된 판과 한 글자라도 다를
    수 있는데, 내려받아 고쳐서 다시 올릴 사람에게는 그게 곧 사고다.
    """
    raw, sheets, formulas, stamp = fetch_workbook(book)
    _seed(book, raw, sheets, formulas, stamp)


def _seed(book: str, raw: bytes, sheets: dict[str, pd.DataFrame],
          formulas: dict, stamp: str) -> None:
    """'이 파일의 지금 판은 이것' 으로 화면 상태를 통째로 갈아끼운다."""
    st.session_state[S_RAW] = raw
    st.session_state[S_FORMULAS] = formulas
    st.session_state[S_BOOK] = book
    st.session_state[S_SHEETS] = sheets
    # 사본을 뜨지 않고 같은 표를 가리킨다. 어느 쪽도 그 자리에서 고치지 않고
    # (고칠 때는 늘 새 표를 만든다), 같은 객체여야 저장할 때 격자가 손대지
    # 않은 시트를 '안 바뀜' 으로 바로 알아본다.
    st.session_state[S_SHOWN] = dict(sheets)
    st.session_state[S_STAMP] = stamp
    st.session_state[S_NONCE] = st.session_state.get(S_NONCE, 0) + 1
    # 다른 파일의 값과 진행 중이던 저장·업로드는 들고 가지 않는다
    for key in (S_EDITED, S_REVIEW, S_KEPT, S_WANT, S_PENDING,
                S_UPLOAD, S_UPLOADED, S_SAVING, S_SAVE_ERR):
        st.session_state.pop(key, None)


def show_input_manage() -> None:
    st.markdown('<div class="pretendard-area"><h2>기준 정보 관리</h2></div>',
                unsafe_allow_html=True)

    # 저장 직후 띄운다. 닫을 때(확인, X) 치운다 -- 띄우면서 치우면 바로 다음
    # 판(격자가 새 판을 받았다고 알려 오는 판)에서 창이 저절로 닫혀 버린다.
    done = st.session_state.get(S_TOAST)
    if done:
        _saved(done)

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

    # 고르개와 단추 넷을 같은 크기, 같은 높이로 한 줄에 둔다. 칸마다 이름
    # 줄을 하나씩 두어야 높이가 맞는다 -- 고르개만 이름이 붙으면 그만큼
    # 혼자 내려앉는다. 그래서 단추 쪽에는 빈 이름 줄을 같은 꼴로 넣는다.
    #
    # 칸은 여기서 한꺼번에 만들지만 저장 쪽은 아래에서 채운다. 저장을 켜고
    # 끄려면 격자가 '고친 게 있다' 를 알려 줘야 하는데 그건 격자를 그린
    # 뒤에야 안다 -- 그렇다고 단추를 표 아래에 두면 15,000행짜리 표에 밀려
    # 화면 밖으로 나가서, 저장하려고 스크롤을 해야 한다.
    c_book, c_reset, c_down, c_up, c_save, _gap = st.columns(
        [1, 1, 1, 1, 1, 1.6])
    # 고친 것이 있는 동안은 파일을 못 바꾼다. 파일을 바꾸면 고친 것을 들고 가지
    # 않는데(시트 이름이 겹치면 엉뚱한 표에 얹힌다), 고르개를 한 번 잘못 누른
    # 것만으로 고친 것이 말없이 사라졌다. 격자가 마지막에 알려 온 것을 본다.
    last = st.session_state.get("im_grid")
    unsaved = (bool(st.session_state.get(S_UPLOADED))
               or (isinstance(last, dict) and bool(last.get("dirty"))
                   and st.session_state.get(S_BOOK) is not None))
    with c_book:
        _row_label("관리 파일")
        book = st.selectbox(
            "관리 파일", books, key="im_book_pick", disabled=unsaved, **_NO_LABEL,
            help="저장하지 않은 수정이 있습니다 — 저장하거나 초기화한 뒤 바꿀 수 있습니다"
            if unsaved else None)
    with c_reset:
        _row_label()
        reload_now = st.button("초기화", **_WIDE,
                               help="저장하지 않은 수정을 버리고 S3 의 지금 값을 다시 읽습니다")

    # 파일을 바꿔 고르면 그 파일을 새로 읽는다.
    if reload_now or st.session_state.get(S_BOOK) != book:
        try:
            _load(book)
        except Exception as err:
            st.error(f"'{book}' 을(를) S3 에서 읽지 못했습니다: {err}")
            st.caption("잠시 뒤 초기화를 눌러 다시 해 보세요. 계속되면 S3 연결"
                       "(AWS_ACCESS_KEY, AWS_SECRET_KEY, 권한)을 확인하세요.")
            return
        if reload_now:
            st.rerun()

    shown: dict[str, pd.DataFrame] = st.session_state[S_SHOWN]
    if not shown:
        st.warning(f"'{book}' 에 시트가 없습니다.")
        return

    status = st.container()

    user_id = st.session_state.get("user_id") or "unknown"
    got = sheet_grid(
        shown,
        version=f"{book}|{st.session_state[S_STAMP]}|{st.session_state[S_NONCE]}",
        key="im_grid",
        want_full=st.session_state.get(S_WANT, ""),
    )
    # 업로드한 내용도 '아직 저장 안 한 수정' 이다. 격자는 새로 받은 판을
    # 깨끗한 것으로 치므로 그것만 보면 저장 단추가 안 켜진다.
    uploaded = bool(st.session_state.get(S_UPLOADED))
    dirty = bool(got.get("dirty")) or uploaded
    _take_full(got, book)

    with c_down:
        _row_label()
        st.download_button(
            "엑셀 다운로드", st.session_state.get(S_RAW) or b"",
            file_name=f"{book}.xlsx", **_WIDE,
            help="엑셀 파일을 내려받습니다.(저장 하지 않은 수정 내용 미포함)",
            mime="application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.sheet")
    with c_up:
        _row_label()
        if st.button("엑셀 업로드", **_WIDE,
                     help="많은 내용을 한 번에 바꿀 때. 내려받아 고친 엑셀을 "
                          "올리면 화면이 그 내용으로 바뀝니다. 저장을 눌러야 "
                          "S3 에 들어갑니다"):
            st.session_state[S_UPLOAD] = True
            st.rerun()
    with c_save:
        _row_label()
        if st.button("저장", type="primary", disabled=not dirty, **_WIDE,
                     help=("S3 의 이 엑셀을 지금 화면의 값으로 바꿉니다"
                           if dirty else "고친 것이 있어야 켜집니다")):
            _ask_full("save")

    with status:
        if st.session_state.get(S_WANT):
            st.caption("표를 받아오는 중입니다...")
        elif uploaded:
            st.info("올린 엑셀의 내용이 화면에 들어왔습니다. 아직 저장 전입니다 "
                    "— 저장을 누르면 무엇이 바뀌는지 먼저 보여 드립니다.")
        elif dirty:
            st.info("저장하지 않은 수정이 있습니다. "
                    "저장을 누르면 무엇이 바뀌는지 먼저 보여 드립니다.")
        else:
            st.caption("고친 것 없음 — 칸을 고치면 저장 단추가 켜집니다")

    if st.session_state.get(S_UPLOAD):
        _upload(book)
    if st.session_state.get(S_REVIEW):
        _review(book, user_id)
    else:
        # 창이 떠 있는 동안에는 안 그린다. 창에 가려 안 보이는데 그리느라
        # 창 안에서 뭘 할 때마다 그만큼씩 기다리게 된다.
        _show_history(book)


def _row_label(text: str = "") -> None:
    """한 줄에 놓인 칸들의 높이를 맞추는 이름 줄.

    글자가 없어도 자리는 차지해야 한다 -- 고르개에만 이름이 붙으면 그 칸만
    이름 높이만큼 내려앉아 단추들과 밑줄이 안 맞는다.
    """
    st.markdown(f"**{text}**" if text else "&nbsp;", unsafe_allow_html=True)


def _ask_full(what: str) -> None:
    """격자에게 '지금 표를 통째로 올려달라' 고 한다. 다음 판에서 받는다."""
    st.session_state[S_PENDING] = what
    st.session_state[S_WANT] = f"{what}-{datetime.now(KST):%H%M%S%f}"
    st.rerun()


def _take_full(got: dict, book: str) -> None:
    """격자가 올려준 표를 받아 두고, 달라고 한 이유대로 처리한다.

    토큰을 맞춰 보는 이유는 streamlit 이 컴포넌트가 마지막에 올린 값을
    계속 되돌려주기 때문이다. 그것만 보고 일하면 저장이 끝난 뒤에도 매 판마다
    또 저장하려 든다.
    """
    want = st.session_state.get(S_WANT)
    if not want or not got.get("full") or got.get("token") != want:
        return
    st.session_state.pop(S_WANT, None)
    try:
        edited = pin_rev_info(to_frames(got, st.session_state.get(S_SHOWN)),
                              st.session_state[S_SHEETS])
    except ValueError as err:
        st.session_state.pop(S_PENDING, None)
        st.error(f"화면의 표를 받지 못했습니다: {err} 초기화한 뒤 다시 해 주세요.")
        return
    st.session_state[S_EDITED] = edited
    if st.session_state.pop(S_PENDING, None) == "save":
        # 여기서 한 번만 센다. 창이 떠 있는 동안 streamlit 이 스크립트를
        # 다시 돌 때마다 또 세면, 15,000행에서는 그때마다 0.4초씩 얼어붙는다.
        st.session_state[S_REVIEW] = workbook_changes(
            st.session_state[S_SHEETS], edited)
        # 수식의 자리는 '화면에 띄운 판' 기준이다 (엑셀을 올렸으면 그 판).
        # 무엇이 바뀌었나는 'S3 에 있는 판' 기준이고, 둘은 다를 수 있다.
        st.session_state[S_KEPT] = surviving_formulas(
            st.session_state[S_SHOWN], edited,
            st.session_state.get(S_FORMULAS, {}))
    st.rerun()


def _upload(book: str) -> None:
    """내려받아 고친 엑셀을 올려 화면 값으로 삼는다."""
    if _HAS_DIALOG:
        kw = ({"width": "large"}
              if "width" in inspect.signature(st.dialog).parameters else {})
        st.dialog("엑셀 업로드", **kw)(_upload_body)(book)
    else:
        with st.container(border=True):
            _upload_body(book)


def _upload_body(book: str) -> None:
    st.markdown(
        "많은 내용을 한 번에 바꿀 때 씁니다. **엑셀 다운로드**로 받아 엑셀에서"
        " 고친 뒤 여기에 올리면, 그 내용이 화면에 그대로 들어옵니다.\n\n"
        "올린다고 저장되는 것은 아닙니다. **저장**을 눌러야 S3 에 들어가고,"
        " 그때 무엇이 바뀌는지 여느 때처럼 먼저 보여 드립니다.")

    got = st.file_uploader("올릴 엑셀 파일", type=["xlsx"], key="im_upload_file")
    if got is not None:
        formulas: dict = {}
        try:
            sheets = read_xlsx(got.getvalue(), formulas)
        except BadWorkbook as err:
            st.error(f"엑셀로 읽지 못했습니다: {err}")
            sheets = None
        except Exception as err:
            st.error(f"읽는 중에 문제가 생겼습니다: {err}")
            sheets = None

        if sheets:
            st.caption(" · ".join(f"{n} {len(df)}행 x {len(df.columns)}열"
                                  for n, df in sheets.items()))
            here = set(st.session_state[S_SHEETS])
            # REV_INFO 는 올린 파일의 것을 쓰지 않으므로 견줄 것이 없다
            there = set(pin_rev_info(sheets, st.session_state[S_SHEETS]))
            if REV_SHEET in here:
                st.caption(f"`{REV_SHEET}` 시트는 올린 파일의 것을 쓰지 않습니다 — "
                           "저장할 때 자동으로 한 줄씩 적히는 기록이라 그대로 둡니다.")
            # 시트 구성이 다르면 통째로 갈아엎는 셈이다. 막지는 않는다 --
            # 일부러 시트를 더하거나 뺄 수도 있다 -- 대신 눈에 띄게 알린다.
            if here != there:
                gone, fresh = sorted(here - there), sorted(there - here)
                said = []
                if gone:
                    said.append("없어지는 시트: " + ", ".join(gone))
                if fresh:
                    said.append("새로 생기는 시트: " + ", ".join(fresh))
                st.warning("지금 파일과 시트 구성이 다릅니다 — " + " / ".join(said))

            ok, cancel, _gap = st.columns([1, 1, 3])
            with ok:
                if st.button("화면에 넣기", type="primary", **_WIDE):
                    st.session_state[S_SHOWN] = pin_rev_info(
                        sheets, st.session_state[S_SHEETS])
                    st.session_state[S_FORMULAS] = formulas
                    st.session_state[S_UPLOADED] = True
                    # 판 번호를 올려야 격자가 새 값으로 다시 그려진다
                    st.session_state[S_NONCE] = st.session_state.get(S_NONCE, 0) + 1
                    for key in (S_UPLOAD, S_EDITED, "im_upload_file"):
                        st.session_state.pop(key, None)
                    st.rerun()
            with cancel:
                if st.button("취소", **_WIDE):
                    for key in (S_UPLOAD, "im_upload_file"):
                        st.session_state.pop(key, None)
                    st.rerun()
            return

    if st.button("닫기", **_WIDE):
        for key in (S_UPLOAD, "im_upload_file"):
            st.session_state.pop(key, None)
        st.rerun()


def _review(book: str, user_id: str) -> None:
    """저장 전에 '무엇이 바뀌는가' 를 보여주고 REV_INFO 를 받는다."""
    if _HAS_DIALOG:
        kw = ({"width": "large"}
              if "width" in inspect.signature(st.dialog).parameters else {})
        st.dialog("변경내용", **kw)(_review_body)(book, user_id)
    else:
        with st.container(border=True):
            _review_body(book, user_id)


def _keep_open() -> None:
    """이 창이 떠 있는 동안 Esc 로 닫히지 않게 한다.

    streamlit 의 창은 Esc 를 누르면 닫히고, 그러면 적던 사유가 통째로
    사라진다. 끄는 설정이 따로 없어서 Esc 를 창 밖으로 못 나가게 막는다.

    막는 것을 그만둘 때를 이 틀(iframe)이 알려 준다 -- 창이 닫히면 이 틀도
    같이 화면에서 빠지므로, 그때 손을 뗀다. 창을 닫는 길은 그대로 있다
    (취소, 오른쪽 위 X).
    """
    components.html("""
<script>
(function () {
  var me = window.frameElement;
  var doc = window.parent.document;
  if (!me || !doc || me.dataset.imEsc) return;
  me.dataset.imEsc = "1";
  function block(e) {
    if (!me.isConnected) {              // 창이 닫혔다 -- 이제 남 일이다
      doc.removeEventListener("keydown", block, true);
      return;
    }
    if (e.key === "Escape" || e.keyCode === 27) {
      e.stopImmediatePropagation();
      e.preventDefault();
    }
  }
  doc.addEventListener("keydown", block, true);
})();
</script>""", height=0)


def _review_body(book: str, user_id: str) -> None:
    changes: dict = st.session_state[S_REVIEW]
    edited: dict[str, pd.DataFrame] = st.session_state[S_EDITED]
    _keep_open()

    if not any(v["total"] or v["note"] for v in changes.values()):
        st.info("바뀐 것이 없습니다. 저장할 것이 없어요.")
        if st.button("닫기", **_WIDE):
            st.session_state.pop(S_REVIEW, None)
            st.rerun()
        return

    for name, info in changes.items():
        head = f"**sheet : {name}** — {info['total']}줄"
        if info["note"]:
            head += f" · {info['note']}"
        st.markdown(head)
        if info["rows"]:
            # st.dataframe 이 아니라 st.table 이다. dataframe 은 캔버스로
            # 그려서 눈으로는 보이지만 글자로는 안 잡히고, 스크롤을 따로
            # 해야 한다. 여기는 '읽고 판단하는' 표라 통째로 펼쳐 두는 쪽이 낫다.
            st.table(pd.DataFrame(
                [{"": r["kind"], "행": r["row"], **r["values"]}
                 for r in info["rows"]]).set_index(""))
        if info["total"] > len(info["rows"]):
            st.caption(f"…외 {info['total'] - len(info['rows'])}줄은 접었습니다.")

    kept, lost = st.session_state.get(S_KEPT) or ({}, {})
    if lost:
        st.warning(
            "**수식이 사라집니다** — "
            + ", ".join(f"{n} {c}개" for n, c in lost.items())
            + "\n\n수식은 제가 앉은 자리를 기준으로 옆 칸을 가리킵니다"
              " (5번 줄의 `=VLOOKUP(B5,...)`). 줄을 넣거나 빼서 자리가"
              " 밀리면 그 수식은 더 이상 맞지 않는데, 엑셀처럼 자리를 따라"
              " 고쳐 주지는 못합니다. 그래서 틀린 수식을 남기는 대신"
              " 마지막으로 계산된 값으로 굳힙니다."
              " 수식을 지키려면 취소하고, 줄을 넣고 빼는 것은 엑셀에서 하세요.")

    st.divider()

    cols = rev_columns(edited)
    # 사유와 단추를 st.form 으로 묶는다. 묶지 않으면 한 글자 칠 때마다,
    # 칸을 떠날 때마다 streamlit 이 스크립트를 처음부터 다시 도는데, 그때마다
    # 15,000행짜리 격자를 다시 내려보내느라 창이 굼떠진다.
    #
    # 굼뜬 것보다 나빴던 것은 따로 있다. 칸에 적은 값은 칸을 떠나야 파이썬에
    # 닿는데, 사람은 적자마자 저장을 누른다 -- 그 누름은 '칸을 떠났다' 로
    # 먼저 처리되고, 그 판에서 저장 단추는 아직 사유가 빈 줄 알고 꺼져 있다.
    # 그래서 한 번 눌러서는 저장이 안 되고 두 번 눌러야 했다. form 안에서는
    # 누름 한 번에 칸 값이 같이 실려 온다.
    # 저장을 눌러 지금 저장하는 중이면 칸과 단추를 전부 끈다 (회색). 두 번
    # 눌러 두 번 저장되는 일이 없게.
    pending = st.session_state.get(S_SAVING)
    busy = pending is not None
    failed = st.session_state.pop(S_SAVE_ERR, None)
    if failed:
        st.error(failed)

    with st.form("im_rev_form", clear_on_submit=False):
        if cols is None:
            st.caption(f"이 파일에는 `{REV_SHEET}` 시트가 없어 변경 사유는 안 받습니다.")
            remark = link = ""
        else:
            st.markdown(f"**{REV_SHEET} 에 남길 기록**")
            today = f"{datetime.now(KST):%Y-%m-%d}"
            # 날짜와 사람은 사람이 못 바꾼다. 언제 누가 바꿨는지는 기록이지
            # 입력이 아니다 -- 고칠 수 있으면 남의 이름으로 적을 수도 있다.
            c1, c2 = st.columns(2)
            with c1:
                st.text_input(REV_DATE, value=today, disabled=True,
                              key="im_rev_date")
            with c2:
                st.text_input(REV_USER, value=user_id, disabled=True,
                              key="im_rev_user")
            remark = st.text_input(f"{REV_REMARK} — 사유", key="im_rev_remark",
                                   disabled=busy)
            link = st.text_input(f"{REV_LINK} — 세부 내용 (필수X)", key="im_rev_link",
                                 disabled=busy)
            st.caption(f"{REV_REMARK} 를 적어야 저장할 수 있습니다.")
            if user_id == "unknown":
                st.caption("로그인한 사람을 못 읽어 'unknown' 으로 남습니다. "
                           "포털이 st.session_state['user_id'] 를 채우는지 봐 주세요.")

        go, cancel, _gap = st.columns([1, 1, 3])
        with go:
            saving = st.form_submit_button("저장 중…" if busy else "저장",
                                           type="primary", **_off(busy), **_WIDE)
        with cancel:
            quit_now = st.form_submit_button("취소", **_off(busy), **_WIDE)

    who = user_id
    if busy:
        # 두 번째 판: 단추는 위에서 이미 회색으로 나갔다. 이제 진짜 저장한다.
        # 표시는 저장하기 전에 치운다 -- 저장 도중 무슨 일로 이 판이 다시
        # 돌더라도 두 번 저장하지 않게.
        st.session_state.pop(S_SAVING, None)
        body = edited
        if cols is not None:
            body = append_rev_info(
                edited, f"{datetime.now(KST):%Y-%m-%d}",
                pending["remark"], who.strip(), pending["link"],
                changes_text(changes))
        with st.spinner("저장하는 중입니다…"):
            _save(book, body, who.strip() or user_id, kept)
        return
    if quit_now:
        st.session_state.pop(S_REVIEW, None)
        st.rerun()
    if saving:
        if cols is not None and not remark.strip():
            st.error(f"{REV_REMARK} 를 적어야 저장할 수 있습니다.")
            return
        # 첫 판: 눌렀다는 것만 적고 곧바로 다시 그린다 (단추가 회색이 된다)
        st.session_state[S_SAVING] = {"remark": remark.strip(),
                                      "link": link.strip()}
        st.rerun()


def _save(book: str, edited: dict[str, pd.DataFrame], user_id: str,
          formulas: dict | None = None) -> None:
    try:
        done = save_workbook(book, edited, user_id,
                             base_stamp=st.session_state[S_STAMP],
                             formulas=formulas)
    except ConcurrentEdit as err:
        # 덮어쓰지 않는다. 누구 값이 맞는지는 코드가 못 정한다.
        st.session_state[S_SAVE_ERR] = str(err)
        st.rerun()
    except Exception as err:
        # 창을 다시 켜서(단추도 다시 켜진다) 까닭을 보여 준다. 적어 둔
        # 사유는 그대로 있으니 다시 누르면 된다.
        st.session_state[S_SAVE_ERR] = (
            f"저장하지 못했습니다: {err}\n\n"
            f"S3 의 값은 그대로입니다. 고친 내용은 화면에 남아 있습니다.")
        st.rerun()
    # 방금 올린 판으로 화면을 맞춘다. S3 에서 도로 내려받아 다시 읽으면
    # 확실하기야 하겠지만, 8MB 짜리 파일에서 그 왕복만 몇 초다 -- 그리고
    # 방금 우리가 올린 바이트가 곧 지금 S3 에 있는 바이트다. 버전표까지
    # 그 put 이 돌려준 것이라, 이어서 또 저장할 때도 맞는 판을 짚는다.
    _seed(book, done.body, done.sheets, formulas or {}, done.stamp)
    # 다음 저장 때 지난번 사유가 그대로 남아 있으면, 그걸 못 보고 그대로
    # 눌러 버린다. 사유는 매번 새로 받는 것이 맞다.
    for key in ("im_rev_remark", "im_rev_link"):
        st.session_state.pop(key, None)
    st.session_state[S_TOAST] = {"book": book, "user": user_id}
    st.rerun()


def _saved(done: dict) -> None:
    """'저장 완료!' 창. 창이 없는 예전 streamlit 에서는 화면 위에 적는다."""
    if _HAS_DIALOG:
        kw = {}
        # X 나 Esc 로 닫았을 때도 치운다. 이걸 못 받는 예전 streamlit 에서는
        # 확인을 누를 때까지 다시 그릴 때마다 창이 또 뜬다.
        if "on_dismiss" in inspect.signature(st.dialog).parameters:
            kw["on_dismiss"] = _forget_saved
        st.dialog("저장 완료!", **kw)(_saved_body)(done)
    else:
        st.session_state.pop(S_TOAST, None)       # 한 번 적고 만다
        with st.container(border=True):
            st.success("저장 완료!")
            _saved_body(done)


def _forget_saved() -> None:
    st.session_state.pop(S_TOAST, None)


def _saved_body(done: dict) -> None:
    st.caption(f"{done['book']} · {done['user']}")
    if _HAS_DIALOG and st.button("확인", type="primary", **_WIDE):
        _forget_saved()
        st.rerun()


def _show_history(book: str) -> None:
    """변경 이력. 따로 모아 둔 것이 아니라 이 파일의 REV_INFO 시트다."""
    log = st.session_state.get(S_SHEETS, {}).get(REV_SHEET)
    if log is None:
        return                       # 그 시트가 없는 파일은 이력도 없다
    with st.expander(f"변경 이력 ({REV_SHEET} · {len(log)}줄)"):
        if log.empty:
            st.caption("아직 저장된 적이 없습니다.")
        else:
            # 최근 것이 위로. 맨 아래에 쌓으니 뒤집어서 보여 준다.
            st.dataframe(log.iloc[::-1], hide_index=True, **_WIDE)
