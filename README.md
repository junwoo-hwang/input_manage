# input_manage — 기준 정보 관리

SF GDVC 포털의 메뉴 하나. S3 drive 에 있는 기준 정보 엑셀을 **브라우저에서
엑셀처럼 고치고 바로 저장**한다. 내려받았다 고쳐서 다시 올리는 왕복이 없다.

```
G-DVC / 2GAPU/input /
    FAB_INPUT_ULY_r0.xlsx     <- 이 폴더의 .xlsx 가 곧 편집 대상 목록
    FAB_INPUT_TTS_r0.xlsx
    _history/FAB_INPUT_ULY_r0/20260917_191817_a3f1_hong.xlsx
    _audit.csv
```

## 되는 것

- 위 드롭다운에서 파일을 고르면 그 엑셀이 통째로 뜬다 (시트는 표 아래 탭)
- **이름 상자 + 수식 입력줄** — 주소는 `A1`, `E1493` 엑셀 표기이고 거기서
  값을 바로 고칠 수 있다
- 셀을 **끌어서 범위 선택**, 열 머리글로 열 전체 선택
- **오른쪽 클릭 차림표** — 잘라내기 / 복사 / 붙여넣기 / 행·열 넣고 빼기 /
  내용 지우기
- **복사 / 붙여넣기가 엑셀과 오간다** (클립보드 TSV)
- **행 / 열 추가·삭제**, 열 이름 바꾸기(머리글 더블클릭)
- **찾기** (Ctrl+F) — 몇 개 중 몇 번째인지 세어 주고 그 칸으로 데려간다
- 되돌리기 / 다시 (Ctrl+Z / Ctrl+Y)
- **엑셀 내려받기** — 지금 화면의 값으로 파일을 만들어 받는다
- 저장하면 S3 의 그 엑셀이 바로 바뀐다. 이력 한 벌과 감사 기록이 남는다

### 안 되는 것

- **글꼴·색·굵게 같은 서식**은 없다. 값만 다룬다. 서식 단추를 달아 봐야
  저장되지 않아 다시 열면 사라지므로, 없는 편이 덜 헷갈린다.
- 수식(`=SUM`), 셀 병합, 틀 고정.

## 포털에 붙이기

`portal.py` 에 세 줄:

```python
from src.input_manage.input_manage import show_input_manage      # 임포트

MENU_CONFIG = {
    ...
    "기준 정보 관리": {"func": show_input_manage, "subs": None},
}

ICON_MAP = {
    ...
    "기준 정보 관리": "table",
}
```

## 환경변수 (p-dep)

```
AWS_ACCESS_KEY / AWS_SECRET_KEY        필수
INPUT_S3_BUCKET     기본 G-DVC
INPUT_S3_PREFIX     기본 2GAPU/input
INPUT_S3_ENDPOINT   기본 http://s3.dataplatform.samsungds.net:9020
```

## 구조

| 파일 | 하는 일 |
|---|---|
| `src/input_manage/input_manage.py` | 전부. S3 · 엑셀 읽기쓰기 · 화면. 포털이 부르는 것은 `show_input_manage()` 하나 |
| `src/input_manage/sheet_grid/frontend/index.html` | 격자. streamlit 컴포넌트라 이 파일만 따로일 수밖에 없다 |

`input_manage.py` 안에서 S3 에 닿는 것은 위쪽 다섯 함수(`s3` 묶음)뿐이다.
테스트는 그것만 가짜로 갈아끼워서 S3 없이 돌고, 사내 헬퍼로 바꿔 끼울 때도
거기만 손대면 된다.

## 필요한 것: streamlit, pandas, boto3. 끝.

**openpyxl 은 안 쓴다.** 사내 pypi 미러에 없어서(`No matching distribution
found for openpyxl`) 그것 하나 때문에 화면 전체를 못 올릴 판이었다. .xlsx 는
XML 몇 장을 zip 으로 묶은 것이고 우리가 다루는 것은 값뿐이라(서식은 애초에
안 다룬다) `zipfile` 과 `xml.etree` 로 직접 읽고 쓴다 — `input_manage.py`
2번 구역이다.

읽을 때 감당하는 것: 엑셀 본프로그램이 쓰는 sharedStrings, openpyxl 이 쓰는
inlineStr, 수식 칸(마지막 계산값), 날짜(엑셀은 날짜를 수로 저장하고 서식으로만
구분한다), 참/거짓, 중간이 비어 건너뛴 칸과 줄.

직접 만든 것이라 '엑셀이 진짜로 읽고 쓰는 꼴' 과 맞는지는 남이 봐 줘야 한다.
`tests/test_xlsx.py` 가 그걸 openpyxl 에 맡긴다 — 개발 환경에만 깔고
(`requirements-dev.txt`) 양쪽으로 맞춰 본다. 돌아가는 코드는 안 쓴다.

## 15,000행에서

가장 큰 시트가 15,000행이다. 실측(15,000행 x 6칸 = 9만 칸):

| | 지금 | 고치기 전 |
|---|---|---|
| 처음 다 그리기 | **1.8초** | 8.2초 |
| 한 번에 멈추는 시간 | **0.75초** | 6.9초 |
| 칸 하나 고치기 | 0.9초 | 2.6초 |
| 맨 끝까지 스크롤 | 0.5초 | 0.3초 |

줄은 **전부** 그린다. 보이는 만큼만 그리면 브라우저 Ctrl+F 가 값을 못 찾고,
스크롤 막대가 실제 줄과 어긋나 보인다 -- 눈으로 훑어 찾는 표라 그게 더 나쁘다.
대신 두 가지로 값을 치른다.

**하나. 표를 `<table>` 이 아니라 `<div>` 로 그린다.** 보기에는 같다. CSS 의
크기 가둠(containment)이 표의 행에는 적용되지 않게 정해져 있어서, `<tr>` 에는
`content-visibility` 를 줘도 브라우저가 배치를 건너뛰지 못한다. `<div>` 면
건너뛴다 -- 15,000줄을 그리고 높이를 재는 데 4,794ms 대 83ms 다. 줄은 9만 칸
그대로 DOM 에 있고, 브라우저는 화면에 보이는 줄만 배치한다.

여기서 조심할 것: 줄 높이(25px)가 `contain-intrinsic-size` 로 미리 알려 주는
값과 같아야 한다. 다르면 스크롤 막대 길이가 틀어진다. `padding` 하나만
건드려도 어긋나므로 검사로 묶어 뒀다.

**둘. 값이 바뀌면 그 칸 하나만, 고른 자리가 바뀌면 칠한 사각형 두 개만
손본다.** 표를 통째로 다시 만들던 때는 칸 하나 고치는 데 11.7초가 걸렸다.

## 왜 st.data_editor 가 아닌가

`st.data_editor` 는 **열을 못 넣고 못 뺀다.** 칸 구성이 DataFrame 스키마로
고정돼서, 기준 정보에 칸 하나 추가하려면 결국 엑셀을 내려받아 고쳐 다시
올려야 한다 — 그게 하기 싫어서 만든 화면이라 그걸로는 끝이 안 난다.

그래서 격자를 직접 만들었다. `frontend/index.html` 하나이고 npm 빌드 단계가
없다. iframe 과 부모 사이의 postMessage 규약 세 가지(componentReady /
render / setComponentValue)만 직접 지킨다.

## 테스트

```bash
pip install -r requirements.txt -r requirements-dev.txt
playwright install chromium
pytest -q            # 117개
```

- `tests/test_storage.py` — 저장이 조용히 덮어써지거나 반쯤 되다 말지 않는가
- `tests/test_xlsx.py` — .xlsx 를 직접 읽고 쓰는 부분 (openpyxl 을 자로)
- `tests/test_sheet_grid.py` — 격자가 올려준 것을 표로 되돌리는 부분
- `tests/test_browser.py` — 끌어서 선택 / 복사 / 붙여넣기 / 행·열 넣고 빼기를
  진짜 브라우저로 (playwright 가 없으면 통째로 건너뛴다)

로컬에서 화면만 보려면 (S3 없이, 가짜 저장소로):

```bash
streamlit run app_local.py
```
