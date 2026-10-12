from pathlib import Path

from docx import Document

FIXED_IMAGE_SRC = "/namoImgView.do?nfile=/images/000150/112d9c94-85cb-4efb-a017-56a52e7fa5c4.jpg"

def extract_text_from_docx(filepath):
    doc = Document(filepath)
    full_text = ''
    for para in doc.paragraphs:
        if para.text.strip():
            full_text += para.text.strip() + '\n'
    return full_text

def convert_news_to_html(text):
    lines = text.strip().split('\n')
    entries = []
    current = None

    for line in lines:
        line = line.strip()
        if line.startswith('['):
            if current:
                entries.append(current)
            country_part, date_part = line.split(']')
            current = {
                'country': country_part[1:],
                'date': date_part.strip(),
                'title': '',
                'text': ''
            }
        elif current and current['title'] == '':
            current['title'] = line
        elif current:
            current['text'] += line + ' '

    if current:
        entries.append(current)

    html_output = ''
    for entry in entries:
        html_output += f'''
<div class="briefing-box">
    <div class="text-area">
        <div class="title-bar">
            <b>[{entry["country"]}]</b><i>{entry["date"]}</i>
            <strong class="title">{entry["title"]}</strong>
        </div>
        <div class="text">
            {entry["text"].strip()}
        </div>
    </div>
    <div class="img">
        <img src="{FIXED_IMAGE_SRC}" alt="{entry["title"]}" style="vertical-align: baseline;">
    </div>
</div>
'''
    return html_output.strip()

def save_html_file(html_content, output_path='output.html'):
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

if __name__ == "__main__":
    root = Path(__file__).resolve().parent
    docx_path = root / "etc" / "example.docx"
    out_path = root / "etc" / "뉴스.html"
    text = extract_text_from_docx(str(docx_path))
    html = convert_news_to_html(text)
    save_html_file(html, str(out_path))
    print(f"✅ 변환 완료! → {out_path} 로 저장됨.")


# python convert_docx_to_html_weekly_logistics.py