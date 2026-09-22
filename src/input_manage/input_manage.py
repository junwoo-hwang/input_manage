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
import inspect
import io
import os
import re
import threading
import zipfile
from datetime import datetime, time, timedelta, timezone
from pathlib import Path
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
    # <si> 안에 <t> 가 여럿일 수 있다 (글자마다 서식이 다른 경우). 이어 붙인다.
    return ["".join(t.text or "" for t in si.iter(NS + "t"))
            for si in root.findall(NS + "si")]


def _read_date_styles(zf: zipfile.ZipFile) -> set[int]:
    """날짜로 보이는 서식을 쓰는 칸 스타일 번호들.

    엑셀은 날짜를 그냥 수로 저장하고 '이 칸은 날짜 서식' 이라고만 적어 둔다.
    그래서 서식을 안 보면 2026-09-20 이 46285 라는 수로 읽힌다.
    """
    root = ET.fromstring(zf.read("xl/styles.xml"))
    custom = set()
    for fmt in root.iter(NS + "numFmt"):
        code = fmt.get("formatCode", "")
        # 따옴표 안의 글자는 서식이 아니라 그대로 찍는 글자다
        bare = re.sub(r'"[^"]*"', "", code)
        if re.search(r"[yYmMdDhHsS]", bare):
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


def _read_sheet(raw: bytes, shared: list[str], date_styles: set[int],
                formulas: dict[tuple[int, int], str] | None = None) -> list[list]:
    rows: list[list] = []
    root = ET.fromstring(raw)
    cell_tag, f_tag = NS + "c", NS + "f"
    # 칸 이름에서 자리를 따는 것은 칸마다 한 번씩 일어난다. 15,000행 x 10칸
    # 이면 15만 번이라, 같은 칸 이름('A','B',...)의 답을 적어 두고 쓴다.
    seen: dict[str, int] = {}
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
                # 배열 수식의 나머지 칸(t="shared" 이면서 내용이 빈 것)은
                # 본체가 따로 있어서 여기 적을 것이 없다
                if f is not None and (f.text or "").strip():
                    formulas[(len(rows), at)] = f.text
        rows.append(values)
    return rows


_V, _IS, _T = NS + "v", NS + "is", NS + "t"


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
        return "".join(t.text or "" for t in node.iter(_T))
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


def _esc(text) -> str:
    out = str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    out = out.replace('"', "&quot;")
    # 엑셀이 못 읽는 제어문자는 뺀다. 붙여넣기로 섞여 들어오면 파일 자체가
    # 안 열리는데, 그게 제일 알아채기 어려운 고장이다.
    return "".join(c for c in out if c in "\t\n\r" or ord(c) >= 32)


def _looks_numeric(text: str) -> bool:
    try:
        float(text)
        return True
    except (TypeError, ValueError):
        return False


def _sheet_xml(rows: list[list],
               formulas: dict[tuple[int, int], str] | None = None) -> bytes:
    formulas = formulas or {}
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/'
             'spreadsheetml/2006/main"><sheetData>']
    for r, row in enumerate(rows, start=1):
        has_f = any((r - 1, c) in formulas for c in range(len(row)))
        if not has_f and not any(v is not None and str(v) != "" for v in row):
            continue                                  # 빈 줄은 안 쓴다
        parts.append(f'<row r="{r}">')
        for c, value in enumerate(row):
            ref = f"{col_letter(c)}{r}"
            formula = formulas.get((r - 1, c))
            if formula is not None:
                # 수식은 수식으로 쓴다. 마지막으로 계산된 값도 같이 적어 둬야
                # 엑셀로 열기 전에(우리 화면에서) 빈 칸으로 보이지 않는다.
                # 진짜 값은 엑셀이 열 때 다시 계산한다(아래 fullCalcOnLoad).
                cached = ("" if value is None else str(value))
                kind = "" if _looks_numeric(cached) else ' t="str"'
                body = f"<v>{_esc(cached)}</v>" if cached else ""
                parts.append(f'<c r="{ref}"{kind}><f>{_esc(formula)}</f>{body}</c>')
                continue
            if value is None or str(value) == "":
                continue
            if isinstance(value, bool):
                parts.append(f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>')
            elif isinstance(value, (int, float)):
                parts.append(f'<c r="{ref}"><v>{value}</v></c>')
            else:
                # 글자는 칸 안에 그대로 넣는다(inlineStr). sharedStrings 를
                # 쓰면 파일이 조금 작아지지만 표를 하나 더 만들어야 하고,
                # 기준 정보는 같은 글자가 반복되는 표가 아니라 이득이 적다.
                parts.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space='
                             f'"preserve">{_esc(value)}</t></is></c>')
        parts.append("</row>")
    parts.append("</sheetData></worksheet>")
    return "".join(parts).encode("utf-8")


def xlsx_write(sheets: dict[str, list[list]],
               formulas: dict[str, dict] | None = None) -> bytes:
    """{시트이름: [[값, ...], ...]} -> .xlsx 바이트.

    formulas 는 {시트이름: {(줄, 칸): '수식'}}. 그 칸은 값 대신 수식으로
    쓰고, 마지막으로 계산된 값을 함께 적어 둔다.
    """
    formulas = formulas or {}
    names = list(sheets) or ["Sheet1"]
    buf = io.BytesIO()
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
              'openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')

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
              f'Target="styles.xml"/></Relationships>')

        # 서식은 안 쓰지만 styles.xml 자체는 있어야 엑셀이 연다
        zf.writestr("xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
            '</cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" '
            'borderId="0" xfId="0"/></cellXfs></styleSheet>')

        for i, name in enumerate(names, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml",
                        _sheet_xml(sheets.get(name, []), formulas.get(name)))
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


def save_workbook(book: str, sheets: dict[str, pd.DataFrame], user_id: str,
                  base_stamp: str | None = None,
                  formulas: dict | None = None) -> str:
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

    # 통째로 만들어 한 번에 올린다. S3 의 put 은 그 자체로 원자적이라,
    # 올리다 끊겨도 옛 파일이 반쯤 덮어써지는 일은 없다.
    body = to_xlsx(sheets, formulas)

    # 이력 폴더에 한 벌 먼저 넣는다. 본 파일을 먼저 덮어쓰고 나면, 그 뒤에
    # 이력 넣기가 실패했을 때 되돌릴 것이 없는 채로 끝난다. 순서를 이렇게
    # 두면 '사본을 못 남기면 덮어쓰지도 않는다' 가 된다.
    s3.put_object(_history_key(book, user_id), body)
    return s3.put_object(key, body)


def _history_key(book: str, user_id: str, now: datetime | None = None) -> str:
    """이력 폴더에 넣을 이름. 260921_원래이름_junwoo.hwang.xlsx

    날짜가 앞에 오므로 폴더를 이름순으로 보면 그대로 시간순이 된다.

    같은 사람이 같은 날 두 번 저장하면 이름이 겹친다. 그대로 두면 앞의 것이
    조용히 덮어써지는데, 되돌릴 판이 하나 사라지는 셈이라 그럴 수 없다.
    겹치면 뒤에 번호를 붙인다.
    """
    now = now or datetime.now(KST)
    base = f"{now:%y%m%d}_{book}_{_safe(user_id)}"
    taken = {k.rsplit("/", 1)[-1] for k in s3.list_keys(_key(HISTORY_DIR) + "/")}
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
    """시트들을 엑셀 파일 한 벌로. 저장과 내려받기가 같은 것을 쓴다.

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
        clean = _clean(df)
        cols = [str(c) for c in clean.columns]
        out[key] = ([cols] + [[_cell(v) for v in row]
                              for row in clean.itertuples(index=False, name=None)])

        want = formulas.get(name)
        if not want:
            continue
        # _clean 이 빈 줄을 빼면 그 아래 줄이 위로 당겨진다. 수식의 자리도
        # 같이 당겨 줘야 엉뚱한 줄에 붙지 않는다.
        moved = {old: new for new, old in enumerate(clean.index)}
        at = {col: i for i, col in enumerate(cols)}
        placed = {}
        for (row, col), text in want.items():
            if row in moved and col in at:
                placed[(moved[row] + 1, at[col])] = text   # +1 은 머리글 줄
        if placed:
            out_f[key] = placed
    return xlsx_write(out, out_f)


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


# ----------------------------------------------------------------------
# 무엇이 바뀌었나 -- 저장 전에 사람에게 보여 줄 것
# ----------------------------------------------------------------------
REV_SHEET = "REV_INFO"
# 그 시트에 적을 칸들. 없는 칸은 건너뛰고, 있는 칸만 채운다.
REV_DATE, REV_REMARK, REV_USER, REV_LINK = "Date", "Remark", "user", "관련"


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
    old = _as_text(before, cols, len(before)) if before is not None else None
    new = _as_text(after, cols, len(after))
    old_rows = [tuple(r) for r in old.itertuples(index=False, name=None)] if old is not None else []
    new_rows = [tuple(r) for r in new.itertuples(index=False, name=None)]

    out: list[dict] = []
    total = 0

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
                # 빈 줄을 새로 만들어 두고 안 채운 것은 '바뀐 것' 이 아니다
                if tag == "insert" and not any(v.strip() for v in new_rows[j]):
                    continue
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

_VLOOKUP_RE = re.compile(
    r'^VLOOKUP\(\s*([^,]+?)\s*,\s*([^!,]+)!\$?([A-Za-z]{1,3})\$?\d*'
    r':\$?([A-Za-z]{1,3})\$?\d*\s*,\s*(\d+)\s*,\s*(?:0|FALSE)\s*\)$',
    re.IGNORECASE)
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
    for name, cells in formulas.items():
        df = out.get(name)
        if df is None or not cells:
            continue
        cols = list(df.columns)
        touched = None
        for (row, col), text in cells.items():
            if col not in cols or row not in df.index:
                continue
            got = _eval_vlookup(text, own=df, sheets=out)
            if got is _NOT_EVALUATED:
                continue
            if touched is None:
                touched = df.copy()
            touched.at[row, col] = got
        if touched is not None:
            out[name] = touched
    return out


def _eval_vlookup(formula: str, own: pd.DataFrame, sheets: dict[str, pd.DataFrame]):
    """수식 하나를 지금 값으로. 못 하면 _NOT_EVALUATED.

    own 은 이 수식이 들어 있는 시트다 -- 첫 인자(찾을 값)가 칸 자리를
    가리키면 그 시트에서 값을 가져와야 하므로, 수식을 어느 시트가 들고
    있는지가 따로 필요하다. sheets 는 VLOOKUP 이 찾아볼 대상 시트를 이름으로
    꺼내려고 쓴다.
    """
    m = _VLOOKUP_RE.match(formula.strip())
    if not m:
        return _NOT_EVALUATED
    lookup_expr, sheet_name, c1, c2, idx = m.groups()
    target = sheets.get(sheet_name.strip())
    if target is None:
        return _NOT_EVALUATED
    key = _resolve_ref(lookup_expr, own)
    if key is _NOT_EVALUATED:
        return _NOT_EVALUATED

    t_cols = list(target.columns)
    start, end = col_index(c1.upper()), col_index(c2.upper())
    idx = int(idx)
    if idx < 1 or idx - 1 > end - start or start >= len(t_cols):
        return _NOT_EVALUATED
    pos = start + idx - 1
    if pos >= len(t_cols):
        return _NOT_EVALUATED
    key_col, val_col = t_cols[start], t_cols[pos]

    key_text = str(key).strip()
    key_series = target[key_col].apply(
        lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v))
        else str(v).strip())
    matches = target.index[key_series == key_text]
    if len(matches) == 0:
        return "#N/A"                        # 엑셀도 못 찾으면 이렇게 보여준다
    found = target.loc[matches[0], val_col]
    return "" if found is None or (isinstance(found, float) and pd.isna(found)) else found


def _resolve_ref(expr: str, own: pd.DataFrame):
    """VLOOKUP 의 첫 인자를 값으로. 같은 시트의 칸 자리(E10220)면 own 에서
    그 값을 가져오고, 아니면 글자/수로 적은 값 그대로다.

    참조에 시트 이름이 안 붙어 있으므로(그냥 'E10220') 수식이 든 시트
    자신을 본다 -- VLOOKUP 의 찾을 값은 대개 자기 줄의 다른 칸이다.
    """
    expr = expr.strip()
    m = _CELL_REF_RE.match(expr)
    if not m:
        if expr.startswith('"') and expr.endswith('"') and len(expr) >= 2:
            return expr[1:-1]
        return expr
    letters, excel_row = m.groups()
    at_row = int(excel_row) - 2          # 머리글이 엑셀 1행이므로 -2
    cols = list(own.columns)
    c = col_index(letters.upper())
    if at_row < 0 or at_row >= len(own) or c >= len(cols):
        return _NOT_EVALUATED
    try:
        return own.iloc[at_row, c]
    except Exception:
        return _NOT_EVALUATED


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
        cols = list(dict.fromkeys([*map(str, old.columns), *map(str, new.columns)]))
        old_text = _as_text(old, cols, len(old))
        new_text = _as_text(new, cols, len(new))
        same = _rows_in_place(
            [tuple(r) for r in old_text.itertuples(index=False, name=None)],
            [tuple(r) for r in new_text.itertuples(index=False, name=None)])
        live = {}
        for (row, col), text in want.items():
            if row not in same or col not in map(str, new.columns):
                continue
            # 그 칸을 사람이 직접 고쳤으면 사람이 적은 값이 이긴다
            if (row < len(old_text) and row < len(new_text)
                    and col in cols
                    and old_text.iloc[row][col] != new_text.iloc[row][col]):
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


def workbook_changes(before: dict[str, pd.DataFrame],
                     after: dict[str, pd.DataFrame]) -> dict:
    """{시트이름: {"rows": [...], "total": n, "note": "..."}}. 안 바뀐 시트는 뺀다."""
    out: dict[str, dict] = {}
    for name, df in after.items():
        old = before.get(name)
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
    저장이나 엑셀 만들기처럼 진짜로 값이 필요할 때만 표를 하나 들려 보낸다.
    """
    payload = [{
        "name": str(name),
        "cols": [str(c) for c in df.columns],
        "rows": [["" if pd.isna(v) else str(v) for v in row]
                 for row in df.itertuples(index=False, name=None)],
    } for name, df in sheets.items()]

    got = _grid(sheets=payload, version=version, max_height=max_height,
                want_full=want_full, key=key, default=None)
    return got or {}


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
# 격자에 '표를 통째로 올려달라' 고 하면서 들려 보낸 표. 격자가 그 표를 달고
# 올려주면 그때가 우리가 요청한 그 값이다.
#
# 이 왕복이 필요한 이유: 평소에 격자는 '고친 게 있다' 만 올린다. 칸 하나
# 고칠 때마다 15,000행을 통째로 주고받으면 한 번에 2.6초가 걸리기 때문이다.
# 그래서 저장이나 엑셀 만들기처럼 값이 진짜 필요한 순간에만 달라고 한다.
S_WANT = "_im_want"
S_PENDING = "_im_pending"    # 표를 받으면 할 일: "save" | "download"
S_EDITED = "_im_edited"      # 격자가 올려준 지금 값
S_REVIEW = "_im_review"      # 저장 팝업에 띄울 변경 내역
# 파일을 띄웠을 때 그 안에 있던 수식들. 격자는 값만 다루므로, 이걸 안 들고
# 있으면 저장할 때 VLOOKUP 이 걸려 있던 칸이 마지막 계산값으로 굳어 버린다.
S_FORMULAS = "_im_formulas"


# 칸 너비를 꽉 채우라고 말하는 법이 streamlit 버전마다 다르다. 새 버전은
# width="stretch", 예전 버전은 use_container_width=True 다. 포털이 어느
# 버전인지 모르는 채로 한쪽만 쓰면 화면이 아예 안 뜬다 (TypeError).
_WIDE = ({"width": "stretch"}
         if "width" in inspect.signature(st.button).parameters
         else {"use_container_width": True})

# st.dialog 는 예전 streamlit 에 없다. 없으면 팝업 대신 화면 안에 펼쳐서
# 보여준다 -- 보기는 덜 좋아도 저장 전에 확인하는 절차는 그대로 지킨다.
_HAS_DIALOG = hasattr(st, "dialog")


def _load(book: str) -> None:
    formulas: dict = {}
    sheets, stamp = load_workbook(book, formulas)
    st.session_state[S_FORMULAS] = formulas
    st.session_state[S_BOOK] = book
    st.session_state[S_SHEETS] = sheets
    st.session_state[S_STAMP] = stamp
    st.session_state[S_NONCE] = st.session_state.get(S_NONCE, 0) + 1
    # 다른 파일의 값과 진행 중이던 저장은 들고 가지 않는다
    for key in (S_EDITED, S_REVIEW, S_WANT, S_PENDING):
        st.session_state.pop(key, None)


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
    #
    # 칸은 여기서 한꺼번에 만들지만 저장 쪽은 아래에서 채운다. 저장을 켜고
    # 끄려면 격자가 '고친 게 있다' 를 알려 줘야 하는데 그건 격자를 그린
    # 뒤에야 안다 -- 그렇다고 단추를 표 아래에 두면 15,000행짜리 표에 밀려
    # 화면 밖으로 나가서, 저장하려고 스크롤을 해야 한다.
    c_book, c_reset, c_save, c_make, c_get, _gap = st.columns(
        [2, 1, 1, 1.3, 1.7, 2])
    with c_book:
        book = st.selectbox("관리할 파일", books, key="im_book_pick")
    with c_reset:
        st.write("")
        reload_now = st.button("초기화", **_WIDE,
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

    status = st.container()

    user_id = st.session_state.get("user_id") or "unknown"
    got = sheet_grid(
        sheets,
        version=f"{book}|{st.session_state[S_STAMP]}|{st.session_state[S_NONCE]}",
        key="im_grid",
        want_full=st.session_state.get(S_WANT, ""),
    )
    dirty = bool(got.get("dirty"))
    _take_full(got, book)

    with c_save:
        st.write("")
        if st.button("저장", type="primary", disabled=not dirty, **_WIDE,
                     help=("S3 의 이 엑셀을 지금 화면의 값으로 바꿉니다"
                           if dirty else "고친 것이 있어야 켜집니다")):
            _ask_full("save")
    with c_make:
        st.write("")
        if st.button("엑셀 만들기", **_WIDE,
                     help="지금 화면의 값(저장 안 한 수정 포함)으로 엑셀 파일을 "
                          "만들어 내려받습니다. S3 는 안 바뀝니다"):
            _ask_full("download")
    with c_get:
        ready = st.session_state.get(S_DOWNLOAD)
        if ready:
            st.write("")
            st.download_button(f"⬇ {ready[0]}", ready[1], file_name=ready[0],
                               **_WIDE,
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet")

    with status:
        if st.session_state.get(S_WANT):
            st.caption("표를 받아오는 중입니다...")
        elif dirty:
            st.info("저장하지 않은 수정이 있습니다. "
                    "저장을 누르면 무엇이 바뀌는지 먼저 보여 드립니다.")
        else:
            st.caption("고친 것 없음 — 칸을 고치면 저장 단추가 켜집니다")

    if st.session_state.get(S_REVIEW):
        _review(book, user_id)

    _show_history(book)


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
    edited = to_frames(got)
    st.session_state[S_EDITED] = edited
    what = st.session_state.pop(S_PENDING, None)
    if what == "download":
        kept, _lost = surviving_formulas(
            st.session_state[S_SHEETS], edited,
            st.session_state.get(S_FORMULAS, {}))
        st.session_state[S_DOWNLOAD] = (f"{book}.xlsx", to_xlsx(edited, kept))
    elif what == "save":
        st.session_state[S_REVIEW] = workbook_changes(
            st.session_state[S_SHEETS], edited)
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


def _review_body(book: str, user_id: str) -> None:
    changes: dict = st.session_state[S_REVIEW]
    edited: dict[str, pd.DataFrame] = st.session_state[S_EDITED]

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
            st.caption(f"…외 {info['total'] - len(info['rows'])}줄은 접었습니다. "
                       f"전부 보려면 '엑셀 만들기' 로 받아서 비교하세요.")

    kept, lost = surviving_formulas(
        st.session_state[S_SHEETS], edited, st.session_state.get(S_FORMULAS, {}))
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

    cols = rev_columns(st.session_state[S_SHEETS])
    if cols is None:
        st.caption(f"이 파일에는 `{REV_SHEET}` 시트가 없어 변경 사유는 안 받습니다.")
        remark = link = ""
        who = user_id
        ok = True
    else:
        st.markdown(f"**{REV_SHEET} 에 남길 기록**")
        today = f"{datetime.now(KST):%Y-%m-%d}"
        # 날짜와 사람은 사람이 못 바꾼다. 언제 누가 바꿨는지는 기록이지
        # 입력이 아니다 -- 고칠 수 있으면 남의 이름으로 적을 수도 있다.
        c1, c2 = st.columns(2)
        with c1:
            st.text_input(REV_DATE, value=today, disabled=True, key="im_rev_date")
        with c2:
            st.text_input(REV_USER, value=user_id, disabled=True, key="im_rev_user")
        who = user_id
        remark = st.text_input(f"{REV_REMARK} — 사유", key="im_rev_remark")
        link = st.text_input(f"{REV_LINK} — 세부 내용 (필수X)", key="im_rev_link")
        ok = bool(remark.strip())
        if not ok:
            st.caption(f"{REV_REMARK} 를 적어야 저장할 수 있습니다.")
        if user_id == "unknown":
            st.caption("로그인한 사람을 못 읽어 'unknown' 으로 남습니다. "
                       "포털이 st.session_state['user_id'] 를 채우는지 봐 주세요.")

    go, cancel, _gap = st.columns([1, 1, 3])
    with go:
        if st.button("저장", type="primary", disabled=not ok, **_WIDE):
            body = edited
            if cols is not None:
                body = append_rev_info(
                    edited, f"{datetime.now(KST):%Y-%m-%d}",
                    remark.strip(), who.strip(), link.strip(),
                    changes_text(changes))
            st.session_state.pop(S_REVIEW, None)
            _save(book, body, who.strip() or user_id, kept)
    with cancel:
        if st.button("취소", **_WIDE):
            st.session_state.pop(S_REVIEW, None)
            st.rerun()


def _save(book: str, edited: dict[str, pd.DataFrame], user_id: str,
          formulas: dict | None = None) -> None:
    try:
        stamp = save_workbook(book, edited, user_id,
                              base_stamp=st.session_state[S_STAMP],
                              formulas=formulas)
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
    st.session_state[S_FORMULAS] = formulas or {}
    st.session_state[S_STAMP] = stamp
    st.session_state.pop(S_DOWNLOAD, None)
    st.session_state.pop(S_EDITED, None)
    # 다음 저장 때 지난번 사유가 그대로 남아 있으면, 그걸 못 보고 그대로
    # 눌러 버린다. 사유는 매번 새로 받는 것이 맞다.
    for key in ("im_rev_remark", "im_rev_link"):
        st.session_state.pop(key, None)
    st.session_state[S_TOAST] = f"'{book}' 저장했습니다 ({user_id})."
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
