from pathlib import Path

from docx import Document
from html import escape

def docx_to_html(docx_path, html_path):
    doc = Document(docx_path)
    html_content = "<html>\n<head><meta charset='utf-8'></head>\n<body>\n"

    for para in doc.paragraphs:
        text = escape(para.text)
        
        # 각 문단의 스타일과 폰트 크기 확인
        style = para.style.name.lower()
        font_size = None
        
        # 폰트 크기 추출 (각 문단의 runs 중에서 폰트 크기 찾기)
        for run in para.runs:
            if run.font.size:
                font_size = run.font.size.pt  # 폰트 크기를 포인트로 가져옴
        
        # 폰트 크기에 맞는 class 설정
        if font_size:
            if font_size >= 24:
                class_name = "detail-large-tit"
            elif font_size >= 18:
                class_name = "detail-medium-tit"
            elif font_size >= 14:
                class_name = "detail-body-caption"
            else:
                class_name = "detail-body-txt"
        else:
            class_name = "detail-body-txt"  # 폰트 크기가 없으면 기본 class

        # 문단을 p 태그로 감싸고 class 추가
        html_content += f"<p class='{class_name}'>{text}</p>\n"

    html_content += "</body>\n</html>"

    # HTML 파일로 저장
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"변환 완료! HTML 파일 저장 위치: {html_path}")

if __name__ == "__main__":
    root = Path(__file__).resolve().parent
    docx_to_html(root / "etc" / "example.docx", root / "etc" / "output.html")
