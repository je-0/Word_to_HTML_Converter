# Word to HTML Converter

Word(`.docx`) 문서를 **웹 CMS에 붙일 HTML**로 바꾸는 Python 스크립트 모음입니다.  
FreightWaves 재게시 기사, 주간 물류 브리핑, 일반 상세 페이지 등 **용도별 마크업**을 생성합니다.

**Repository:** [github.com/je-0/Word_to_HTML_Converter](https://github.com/je-0/Word_to_HTML_Converter)

---

## 요구 사항

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/)

```bash
pip install -r requirements.txt
```

가상환경 사용을 권장합니다.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

---

## 변환기 한눈에

| 스크립트 | 출력 형태 | 용도 |
|----------|-----------|------|
| `change-global-news.py` | HTML **조각** (`img-top`, `source-area`, `text`) | FreightWaves 등 글로벌 뉴스 1건 |
| `convert_docx_to_html_weekly_logistics.py` | HTML **조각** (`briefing-box` 반복) | 국가별 주간 물류 브리핑 |
| `convert_docx_to_html.py` | **완전한 HTML 문서** (`<html><body>`) | 폰트 크기 기반 상세 본문 |

모든 결과 파일은 **UTF-8**로 저장됩니다.

---

## 1. 글로벌 뉴스 — `change-global-news.py`

### Word 작성 규칙

단락(줄) 순서와 접두어를 맞춰야 제목·본문이 올바르게 잡힙니다.

1. (선택) `Original Article: https://...`
2. (선택) `Author: 이름`
3. (선택) `Date: Tuesday, April 15, 2025`
4. `(Photo: ...)` 가 포함된 줄 — 사진 캡션
5. **제목** — Photo 줄 다음 첫 번째 줄
6. **본문** — 그 이후 모든 줄

`(Photo:` 줄이 없으면 제목·본문이 비어 있을 수 있습니다.

### CLI

```bash
python change-global-news.py -i article.docx -o article.html
python change-global-news.py --author "Custom Author" --photo-desc "Caption text"
python change-global-news.py --image-src "/namoImgView.do?nfile=/images/..."
```

| 옵션 | 설명 |
|------|------|
| `-i`, `--input` | 입력 `.docx` (기본: `etc/global-news.docx`) |
| `-o`, `--output` | 출력 `.html` (기본: `etc/글로벌뉴스.html`) |
| `--image-src` | CMS 이미지 URL |
| `--author` | 문서에 Author가 없을 때 사용 |
| `--photo-desc` | Photo 캡션을 문서에서 못 찾을 때 사용 |

### Python에서 함수만 쓰기

파일명에 하이픈(`-`)이 있어 일반 `import` 대신 `importlib` 또는 CLI 실행을 사용합니다.

```python
import importlib.util
from pathlib import Path

root = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location(
    "change_global_news", root / "change-global-news.py"
)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

lines = mod.extract_text_from_docx("article.docx")
link, photo, author, date, title, body = mod.extract_metadata(lines)
html = mod.convert_news_to_html(link, photo, title, body, author, date)
mod.save_html_file(html, "out.html")
```

---

## 2. 주간 물류 브리핑 — `convert_docx_to_html_weekly_logistics.py`

### Word 작성 규칙

여러 기사를 한 문서에 넣을 때 아래 패턴을 반복합니다.

```
[나이지리아]5.11
제목 줄
본문 텍스트 (여러 줄 가능)
[인도]5.9
다음 기사 제목
...
```

- `[국가명]날짜` 형식으로 새 항목 시작 (`]` 뒤에 날짜)
- 다음 줄: 제목
- 그 다음 줄들: 본문 (공백으로 이어 붙임)

### 실행

CLI 인자는 없습니다. 다른 스크립트에서 함수를 import 하거나, 로컬 샘플 경로로 실행합니다.

```python
from convert_docx_to_html_weekly_logistics import (
    extract_text_from_docx,
    convert_news_to_html,
    save_html_file,
)

text = extract_text_from_docx("briefing.docx")
save_html_file(convert_news_to_html(text), "briefing.html")
```

`python convert_docx_to_html_weekly_logistics.py` 를 실행하면  
(로컬에 `etc/`가 있을 때) `etc/example.docx` → `etc/뉴스.html` 로 변환합니다.

---

## 3. 일반 Word → HTML — `convert_docx_to_html.py`

문단별 **폰트 크기(pt)** 로 CSS 클래스를 붙입니다.

| 폰트 크기 | class |
|-----------|--------|
| 24pt 이상 | `detail-large-tit` |
| 18pt 이상 | `detail-medium-tit` |
| 14pt 이상 | `detail-body-caption` |
| 그 외 / 미지정 | `detail-body-txt` |

```python
from convert_docx_to_html import docx_to_html

docx_to_html("input.docx", "output.html")
```

---

## 프로젝트 구조

```
Word_to_HTML_Converter/
├── change-global-news.py
├── convert_docx_to_html.py
├── convert_docx_to_html_weekly_logistics.py
├── requirements.txt
├── README.md
├── .gitignore
└── etc/                 # Git 제외 — 로컬 샘플·보조 도구
    ├── *.docx, *.html   # 테스트용 원본/결과
    ├── batch_converter.py
    ├── quick_convert.py
    └── …
```

`etc/`는 `.gitignore`에 포함되어 **원격 저장소에는 올라가지 않습니다.**  
샘플 Word 파일, 배치 변환, Excel 변환 등은 각자 `etc/`에 두고 사용하면 됩니다.

---

## CMS 반영 시 참고

- `change-global-news.py`, `convert_docx_to_html_weekly_logistics.py` 출력은 **본문 영역에 그대로 붙이는 조각**입니다.
- 이미지 URL은 Namo 계열 경로(`/namoImgView.do?nfile=...`) 형식을 기본으로 두었습니다. 사이트에 맞게 `--image-src` 또는 스크립트 상단 상수를 수정하세요.
- `convert_docx_to_html.py`는 `<html>` 전체 문서이므로, CMS가 조각만 받는 경우 `<body>` 안 내용만 추출해 사용하면 됩니다.

---

## 문제 해결

| 증상 | 확인 |
|------|------|
| `ModuleNotFoundError: docx` | `pip install -r requirements.txt` |
| 글로벌 뉴스 제목/본문이 비음 | Word에 `(Photo:` 줄과 그 다음 제목 줄이 있는지 확인 |
| `etc/...` 파일 없음 | `etc/` 폴더와 docx를 로컬에 직접 준비 (Git에는 없음) |

---

## 라이선스

내부·업무용 도구로 사용합니다. 필요 시 저장소 Settings에서 라이선스를 추가하세요.
