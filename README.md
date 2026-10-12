# Word → HTML 변환기

Cellosquare CMS용 Word(.docx) HTML 조각 변환 스크립트입니다.

## 스크립트

| 파일 | 용도 |
|------|------|
| `change-global-news.py` | FreightWaves 형식 글로벌 뉴스 기사 |
| `convert_docx_to_html_weekly_logistics.py` | `[국가]날짜` 형식 주간 물류 브리핑 |
| `convert_docx_to_html.py` | 폰트 크기 기반 일반 Word → HTML |

## 설치

```bash
pip install -r requirements.txt
```

## 사용 예

```bash
python change-global-news.py -i article.docx -o article.html
python convert_docx_to_html_weekly_logistics.py
python convert_docx_to_html.py
```

배치 변환·Excel 변환·샘플 docx/html 등은 로컬 `etc/` 폴더에 두며, Git에는 포함하지 않습니다.
