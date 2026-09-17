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

- 위 드롭다운에서 파일을 고르면 그 엑셀이 통째로 뜬다 (시트는 탭)
- 셀을 **끌어서 범위 선택**, 열 머리글로 열 전체 선택
- **복사 / 붙여넣기가 엑셀과 오간다** (클립보드 TSV)
- **행 / 열 추가·삭제**, 열 이름 바꾸기(머리글 더블클릭)
- 되돌리기 / 다시 (Ctrl+Z / Ctrl+Y)
- 저장하면 S3 의 그 엑셀이 바로 바뀐다. 이력 한 벌과 감사 기록이 남는다

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
| `src/input_manage/input_manage.py` | 화면. 포털이 부르는 `show_input_manage()` 하나 |
| `src/input_manage/storage.py` | 엑셀 읽기/쓰기, 이력, 감사, 동시 편집 방어 |
| `src/input_manage/s3io.py` | S3 에 닿는 유일한 자리 |
| `src/input_manage/sheet_grid/` | 격자(컴포넌트). `frontend/index.html` 하나가 전부 |

`storage.py` 는 `s3io` 의 함수 다섯 개 말고는 S3 를 모른다. 그래서 테스트는
메모리 가짜 저장소로 S3 없이 돈다.

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
pytest -q            # 47개
```

- `tests/test_storage.py` — 저장이 조용히 덮어써지거나 반쯤 되다 말지 않는가
- `tests/test_sheet_grid.py` — 격자가 올려준 것을 표로 되돌리는 부분
- `tests/test_browser.py` — 끌어서 선택 / 복사 / 붙여넣기 / 행·열 넣고 빼기를
  진짜 브라우저로 (playwright 가 없으면 통째로 건너뛴다)

로컬에서 화면만 보려면 (S3 없이, 가짜 저장소로):

```bash
streamlit run app_local.py
```
