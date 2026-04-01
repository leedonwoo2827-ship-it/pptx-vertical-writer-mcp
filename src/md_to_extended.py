"""
일반 마크다운을 확장 MD 포맷으로 자동 변환.
섹션 구조, 본문 길이, 테이블 유무 등을 분석하여 최적 템플릿을 자동 배정.
"""
import re
import sys
import os


# 템플릿별 대표 참조 슬라이드 번호
TEMPLATE_REF_SLIDES = {
    'T0': [17, 39, 56, 94],   # 구분페이지
    'T1': [5, 6, 7, 8, 78],   # 카드형 다중
    'T3': [2],                  # 거버닝메시지+범위
    'T6': [22, 34, 35, 60],   # 순수 데이터테이블
    'T7': [20, 28, 33],       # 테이블+설명shape
    'T9': [16, 29, 47, 62],   # 핵심메시지/다이어그램
}

# T0 구분페이지용 ref_slide 순환 인덱스
t0_idx = 0
t1_idx = 0
t6_idx = 0
t7_idx = 0
t9_idx = 0


def get_ref_slide(template, idx=0):
    """템플릿에 맞는 참조 슬라이드 번호 반환 (순환)"""
    slides = TEMPLATE_REF_SLIDES.get(template, [17])
    return slides[idx % len(slides)]


def parse_sections(md_text):
    """MD를 # 헤딩 기준으로 섹션 분리"""
    lines = md_text.split('\n')
    sections = []
    current = None
    current_lines = []

    for line in lines:
        if line.startswith('# '):
            if current:
                sections.append({
                    'title': current,
                    'body': '\n'.join(current_lines).strip()
                })
            current = line[2:].strip()
            current_lines = []
        else:
            current_lines.append(line)

    if current:
        sections.append({
            'title': current,
            'body': '\n'.join(current_lines).strip()
        })

    return sections


def analyze_section(section):
    """섹션 내용을 분석하여 특성 파악"""
    body = section['body']
    lines = [l for l in body.split('\n') if l.strip()]

    has_table = any('|' in l and l.strip().startswith('|') for l in lines)
    table_lines = [l for l in lines if '|' in l and l.strip().startswith('|')]
    text_lines = [l for l in lines if not (l.strip().startswith('|') or l.strip().startswith('---'))]
    h2_lines = [l for l in lines if l.startswith('## ')]
    h3_lines = [l for l in lines if l.startswith('### ')]

    return {
        'has_table': has_table,
        'table_line_count': len(table_lines),
        'text_line_count': len(text_lines),
        'h2_count': len(h2_lines),
        'h3_count': len(h3_lines),
        'total_lines': len(lines),
        'body_length': len(body),
    }


def extract_table(body):
    """본문에서 마크다운 테이블 추출"""
    lines = body.split('\n')
    table_lines = []
    in_table = False

    for line in lines:
        if '|' in line and line.strip().startswith('|'):
            in_table = True
            table_lines.append(line.strip())
        elif in_table:
            break

    return '\n'.join(table_lines) if table_lines else ''


def extract_paragraphs(body):
    """## 헤딩 뒤의 본문 텍스트들을 추출"""
    lines = body.split('\n')
    paragraphs = []

    for line in lines:
        line = line.strip()
        if line.startswith('## ') and not line.startswith('### '):
            text = line[3:].strip()
            if text and len(text) > 10:
                paragraphs.append(text)
        elif line.startswith('### '):
            text = line[4:].strip()
            if text:
                paragraphs.append(text)

    return paragraphs


def section_to_slide(section, analysis, counters):
    """섹션을 확장 MD 슬라이드로 변환"""
    title = section['title']
    body = section['body']
    paragraphs = extract_paragraphs(body)

    # 섹션 깊이 판단
    depth = title.count('.')
    is_major = depth <= 1  # 1.2, 2.1 등
    is_sub = depth == 2    # 1.2.1, 2.2.1 등

    # 본문이 거의 없는 섹션 헤더 → T0 구분페이지
    if analysis['total_lines'] <= 2 and not analysis['has_table']:
        ref = get_ref_slide('T0', counters['t0'])
        counters['t0'] += 1
        slide = f"""---slide
template: T0
ref_slide: {ref}
---
@content_1: {title}"""
        return [slide]

    slides = []

    # 테이블이 있는 섹션
    if analysis['has_table']:
        table_text = extract_table(body)
        # 거버닝 메시지 = 첫 번째 ## 텍스트
        gov_msg = paragraphs[0][:200] if paragraphs else title

        # 테이블 크기에 따라 템플릿 선택
        if analysis['table_line_count'] > 10:
            ref = get_ref_slide('T6', counters['t6'])
            counters['t6'] += 1
        else:
            ref = get_ref_slide('T7', counters['t7'])
            counters['t7'] += 1

        slide = f"""---slide
template: T6
ref_slide: {ref}
---
@governing_message: {gov_msg}
@breadcrumb: {title}

{table_text}"""
        slides.append(slide)

        # 테이블 외 텍스트가 많으면 추가 슬라이드
        non_table_paras = [p for p in paragraphs if '|' not in p]
        if len(non_table_paras) > 3:
            extra_slide = _make_text_slide(title, non_table_paras, counters)
            slides.append(extra_slide)

    # 테이블 없이 텍스트만 있는 섹션
    else:
        if not paragraphs:
            # 본문이 없으면 구분페이지
            ref = get_ref_slide('T0', counters['t0'])
            counters['t0'] += 1
            slide = f"""---slide
template: T0
ref_slide: {ref}
---
@content_1: {title}"""
            slides.append(slide)
        elif len(paragraphs) <= 3:
            # 짧은 텍스트 → T9 핵심메시지
            slide = _make_text_slide(title, paragraphs, counters)
            slides.append(slide)
        elif analysis['h3_count'] >= 2:
            # ### 소제목이 여러 개 → 카드형으로 변환
            slide = _make_card_slide(title, body, counters)
            slides.append(slide)
        else:
            # 긴 텍스트 → 여러 슬라이드로 분할
            chunk_size = 4
            for i in range(0, len(paragraphs), chunk_size):
                chunk = paragraphs[i:i+chunk_size]
                slide = _make_text_slide(title, chunk, counters)
                slides.append(slide)

    return slides


def _make_text_slide(title, paragraphs, counters):
    """텍스트 중심 슬라이드 생성 (T9 또는 T3)"""
    ref = get_ref_slide('T9', counters['t9'])
    counters['t9'] += 1

    gov_msg = paragraphs[0][:200] if paragraphs else title
    fields = [f'@governing_message: {gov_msg}']
    fields.append(f'@breadcrumb: {title}')

    for i, para in enumerate(paragraphs[:6], 1):
        fields.append(f'@content_{i}: {para[:300]}')

    return f"""---slide
template: T9
ref_slide: {ref}
---
""" + '\n'.join(fields)


def _make_card_slide(title, body, counters):
    """### 소제목을 카드형으로 변환 (T1)"""
    ref = get_ref_slide('T1', counters['t1'])
    counters['t1'] += 1

    lines = body.split('\n')
    cards = []
    current_h3 = None
    current_text = []

    for line in lines:
        line_stripped = line.strip()
        if line_stripped.startswith('### '):
            if current_h3:
                cards.append((current_h3, ' '.join(current_text)[:300]))
            current_h3 = line_stripped[4:].strip()
            current_text = []
        elif line_stripped.startswith('## ') and current_h3:
            text = line_stripped[3:].strip()
            if text and len(text) > 10:
                current_text.append(text)

    if current_h3:
        cards.append((current_h3, ' '.join(current_text)[:300]))

    # 첫 ## 텍스트를 거버닝 메시지로
    first_para = ''
    for line in lines:
        if line.strip().startswith('## ') and not line.strip().startswith('### '):
            first_para = line.strip()[3:]
            break

    fields = [f'@governing_message: {first_para[:200]}' if first_para else f'@governing_message: {title}']
    fields.append(f'@breadcrumb: {title}')

    for i, (h3_title, h3_body) in enumerate(cards[:6], 1):
        fields.append(f'@카드{i}_제목: {h3_title}')
        fields.append(f'@카드{i}_내용: {h3_body}')

    return f"""---slide
template: T1
ref_slide: {ref}
---
""" + '\n'.join(fields)


def convert_md_to_extended(input_path, output_path, ref_pptx_path):
    """일반 MD를 확장 MD로 변환"""
    with open(input_path, 'r', encoding='utf-8') as f:
        md_text = f.read()

    sections = parse_sections(md_text)

    # III. 사업관리 부문은 제외 (별도 PPTX)
    sections = [s for s in sections if not s['title'].startswith('III.')]

    counters = {'t0': 0, 't1': 0, 't3': 0, 't6': 0, 't7': 0, 't9': 0}

    output_lines = [f"""---config
reference_pptx: {ref_pptx_path}
---
"""]

    total_slides = 0
    for section in sections:
        analysis = analyze_section(section)
        slides = section_to_slide(section, analysis, counters)
        for slide in slides:
            total_slides += 1
            # 슬라이드 번호 주석 삽입 (---slide 바로 뒤에)
            slide = slide.replace('---slide', f'---slide\n# [S{total_slides:03d}] {section["title"][:60]}', 1)
            output_lines.append(slide)

    result = '\n\n'.join(output_lines)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(result)

    print(f'변환 완료: {len(sections)}개 섹션 → {total_slides}개 슬라이드')
    print(f'출력: {output_path}')

    return total_slides


if __name__ == '__main__':
    input_path = sys.argv[1] if len(sys.argv) > 1 else 'input/proposal-body.md'
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'input/proposal-body-auto.md'
    ref_pptx = sys.argv[3] if len(sys.argv) > 3 else 'reference-ppt/우즈베키스탄_사이버대학 설립 PMC_Ⅱ.기술 부문_v2.pptx'

    convert_md_to_extended(input_path, output_path, ref_pptx)
