import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

from docx import Document

_PROJECT_ROOT = Path(__file__).resolve().parent

# 기본 설정값
DEFAULT_IMAGE_SRC = "/namoImgView.do?nfile=/images/000148/1257a0e1-d6ee-4a4a-90c6-6ef4614cbfff.jpg"
DEFAULT_AUTHOR = "Stuart Chirls"
DEFAULT_SOURCE = "Articles Reproduced by Permission of FreightWaves"


def extract_text_from_docx(filepath):
    """
    Word 문서에서 텍스트를 추출합니다.
    
    Args:
        filepath: Word 문서 파일 경로
        
    Returns:
        list: 각 단락의 텍스트 리스트
        
    Raises:
        FileNotFoundError: 파일이 존재하지 않을 때
        Exception: 문서 읽기 실패 시
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"❌ 파일을 찾을 수 없습니다: {filepath}")
    
    try:
        doc = Document(filepath)
        lines = []
        for para in doc.paragraphs:
            if para.text.strip():
                lines.append(para.text.strip())
        return lines
    except Exception as e:
        raise Exception(f"❌ 문서 읽기 실패: {str(e)}")


def extract_metadata(lines, default_photo_desc=None):
    """
    텍스트 라인에서 메타데이터를 추출합니다.
    
    Args:
        lines: 텍스트 라인 리스트
        default_photo_desc: 기본 사진 설명 (없을 경우 사용)
        
    Returns:
        tuple: (article_link, photo_description, author, date, title_text, body_lines)
    """
    article_link = ''
    photo_description = default_photo_desc or ''
    author = DEFAULT_AUTHOR
    date_str = datetime.now().strftime("%A, %B %d, %Y")
    title_text = ''
    body_lines = []

    found_photo = False
    found_title = False

    for idx, line in enumerate(lines):
        # 원문 링크 추출
        if line.startswith("Original Article:"):
            article_link = line.split("Original Article:")[1].strip()
        
        # 작성자 추출 (예: "Author: John Doe" 형식)
        elif line.startswith("Author:"):
            author = line.split("Author:")[1].strip()
        
        # 날짜 추출 (예: "Date: Tuesday, April 15, 2025" 형식)
        elif line.startswith("Date:"):
            date_str = line.split("Date:")[1].strip()
        
        # 사진 설명 추출
        elif "(Photo:" in line:
            # "(Photo: ...)" 형식에서 설명 추출 시도
            if "(Photo:" in line:
                # 사진 설명이 같은 줄에 있는 경우
                photo_desc_match = line.split("(Photo:")[1].split(")")[0] if ")" in line else None
                if photo_desc_match:
                    photo_description = photo_desc_match.strip()
                else:
                    # 다음 줄이 사진 설명일 수 있음
                    if idx + 1 < len(lines):
                        photo_description = lines[idx + 1].strip()
                found_photo = True
        
        # 제목 추출 (사진 설명 다음 첫 번째 줄)
        elif found_photo and not found_title:
            title_text = line.strip()
            found_title = True
        
        # 본문 추출 (제목 이후 모든 줄)
        elif found_title:
            body_lines.append(line.strip())

    return article_link, photo_description, author, date_str, title_text, body_lines


def convert_news_to_html(article_link, photo_description, title_text, body_lines, 
                         author=None, date_str=None, image_src=None, source=None):
    """
    추출한 메타데이터를 HTML 형식으로 변환합니다.
    
    Args:
        article_link: 원문 링크
        photo_description: 사진 설명
        title_text: 제목
        body_lines: 본문 라인 리스트
        author: 작성자 (기본값: DEFAULT_AUTHOR)
        date_str: 날짜 문자열 (기본값: 현재 날짜)
        image_src: 이미지 경로 (기본값: DEFAULT_IMAGE_SRC)
        source: 출처 정보 (기본값: DEFAULT_SOURCE)
        
    Returns:
        str: HTML 문자열
    """
    author = author or DEFAULT_AUTHOR
    date_str = date_str or datetime.now().strftime("%A, %B %d, %Y")
    image_src = image_src or DEFAULT_IMAGE_SRC
    source = source or DEFAULT_SOURCE
    
    html_output = ''
    
    # 이미지 및 출처 영역
    html_output += f'''
<div class="img img-top" style="margin-bottom: 0px">
    <div class="source-area">
        <p>
            {author}, {date_str}<br />
            Original Article:
            <a href="{article_link}" class="link" target="_blank">{article_link}</a><br />
            {source}
        </p>
    </div>
    <img src="{image_src}" alt="Photo" style="vertical-align: baseline" />
    <span class="im_info align-right">{photo_description}</span>
</div>
'''

    # 제목 및 본문 영역
    html_output += f'''
<div class="text">
    <strong class="title">{title_text}</strong>
    {' '.join(body_lines)}
</div>
'''

    return html_output.strip()


def save_html_file(html_content, output_path='output.html'):
    """
    HTML 내용을 파일로 저장합니다.
    
    Args:
        html_content: 저장할 HTML 내용
        output_path: 출력 파일 경로
        
    Raises:
        Exception: 파일 저장 실패 시
    """
    try:
        # 출력 디렉토리가 없으면 생성
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
    except Exception as e:
        raise Exception(f"❌ HTML 파일 저장 실패: {str(e)}")


def main():
    """메인 실행 함수"""
    parser = argparse.ArgumentParser(
        description='Word 문서를 HTML 뉴스 형식으로 변환합니다.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
예제:
  python change-global-news.py
  python change-global-news.py -i input.docx -o output.html
  python change-global-news.py --input global-news.docx --output 글로벌뉴스.html
        '''
    )
    
    parser.add_argument(
        '-i', '--input',
        type=str,
        default=str(_PROJECT_ROOT / 'etc' / 'global-news.docx'),
        help='입력 Word 문서 파일 경로 (기본값: etc/global-news.docx)'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=str,
        default=str(_PROJECT_ROOT / 'etc' / '글로벌뉴스.html'),
        help='출력 HTML 파일 경로 (기본값: etc/글로벌뉴스.html)'
    )
    
    parser.add_argument(
        '--image-src',
        type=str,
        default=DEFAULT_IMAGE_SRC,
        help=f'이미지 경로 (기본값: {DEFAULT_IMAGE_SRC})'
    )
    
    parser.add_argument(
        '--author',
        type=str,
        default=None,
        help='작성자 이름 (문서에서 추출 실패 시 사용)'
    )
    
    parser.add_argument(
        '--photo-desc',
        type=str,
        default=None,
        help='사진 설명 (문서에서 추출 실패 시 사용)'
    )
    
    args = parser.parse_args()
    
    try:
        # 1. Word 문서에서 텍스트 추출
        print(f"📄 문서 읽는 중: {args.input}")
        lines = extract_text_from_docx(args.input)
        
        if not lines:
            print("⚠️ 경고: 문서에서 텍스트를 찾을 수 없습니다.")
            return
        
        # 2. 메타데이터 추출
        print("🔍 메타데이터 추출 중...")
        article_link, photo_description, author, date_str, title_text, body_lines = extract_metadata(
            lines, 
            default_photo_desc=args.photo_desc
        )
        
        # 명령줄 인자로 제공된 작성자 사용
        if args.author:
            author = args.author
        
        # 필수 필드 검증
        if not title_text:
            print("⚠️ 경고: 제목을 찾을 수 없습니다.")
        
        if not body_lines:
            print("⚠️ 경고: 본문을 찾을 수 없습니다.")
        
        # 3. HTML 변환
        print("🔄 HTML 변환 중...")
        html = convert_news_to_html(
            article_link=article_link,
            photo_description=photo_description,
            title_text=title_text,
            body_lines=body_lines,
            author=author,
            date_str=date_str,
            image_src=args.image_src
        )
        
        # 4. 파일 저장
        print(f"💾 파일 저장 중: {args.output}")
        save_html_file(html, args.output)
        
        print(f"✅ 변환 완료! → {args.output} 로 저장됨.")
        print(f"   - 제목: {title_text[:50]}{'...' if len(title_text) > 50 else ''}")
        print(f"   - 본문 라인 수: {len(body_lines)}")
        
    except FileNotFoundError as e:
        print(str(e))
        sys.exit(1)
    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")
        sys.exit(1)


if __name__ == '__main__':
    main()
