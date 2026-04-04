"""
확장 마크다운 파서.
---slide 구분자와 @필드명: 값 문법으로 슬라이드별 데이터를 파싱.
"""
import re
from typing import List, Dict, Any, Optional


def parse_md(md_text: str) -> Dict[str, Any]:
    """
    확장 MD 포맷을 파싱하여 config + slides 리스트 반환.

    포맷:
    ---config
    reference_pptx: path/to/ref.pptx
    master_template: path/to/template.pptx
    ---

    ---slide
    template: T1
    ref_slide: 5
    ---
    @governing_message: 텍스트...
    @section_title: 제목 텍스트
    @카드1_제목: 카드 제목
    @카드1_내용: 카드 내용
    """
    result = {
        'config': {},
        'slides': []
    }

    # config 블록 추출
    config_match = re.search(
        r'---config\s*\n(.*?)\n---',
        md_text, re.DOTALL
    )
    if config_match:
        config_text = config_match.group(1)
        for line in config_text.strip().split('\n'):
            line = line.strip()
            if ':' in line:
                key, val = line.split(':', 1)
                result['config'][key.strip()] = val.strip()

    # slide 블록들 추출 (# [SNNN] 주석 라인 포함 가능)
    slide_blocks = re.split(r'---slide\s*\n(?:#\s*\[S\d+\].*\n)?', md_text)

    for block in slide_blocks[1:]:  # 첫 번째는 config 이전 텍스트
        slide_data = parse_slide_block(block)
        if slide_data:
            result['slides'].append(slide_data)

    return result


def parse_slide_block(block: str) -> Optional[Dict[str, Any]]:
    """단일 슬라이드 블록 파싱"""
    # 헤더 (template, ref_slide 등) 추출
    header_match = re.match(r'(.*?)\n---\s*\n(.*)', block, re.DOTALL)

    if header_match:
        header_text = header_match.group(1)
        body_text = header_match.group(2)
    else:
        # 헤더 없이 바로 내용
        header_text = block
        body_text = ''

    slide = {
        'template': None,
        'ref_slide': None,
        'fields': {},
        'tables': [],
        'bullets': [],
    }

    # 헤더 파싱
    for line in header_text.strip().split('\n'):
        line = line.strip()
        if line.startswith('template:'):
            slide['template'] = line.split(':', 1)[1].strip()
        elif line.startswith('ref_slide:'):
            try:
                slide['ref_slide'] = int(line.split(':', 1)[1].strip())
            except ValueError:
                pass
        elif line.startswith('reference_pptx:'):
            slide['reference_pptx'] = line.split(':', 1)[1].strip()

    if not body_text:
        body_text = header_text if not header_match else ''

    # 본문 파싱
    if body_text:
        parse_body(body_text.strip(), slide)

    return slide if slide['template'] or slide['ref_slide'] else None


def parse_body(body: str, slide: Dict):
    """본문에서 @필드, 마크다운 테이블, 불릿 등을 파싱"""
    lines = body.split('\n')
    current_field = None
    current_value_lines = []
    current_table = []
    in_table = False

    def flush_field():
        nonlocal current_field, current_value_lines
        if current_field:
            value = '\n'.join(current_value_lines).strip()
            slide['fields'][current_field] = value
            current_field = None
            current_value_lines = []

    def flush_table():
        nonlocal current_table, in_table
        if current_table:
            parsed = parse_md_table(current_table)
            if parsed:
                slide['tables'].append(parsed)
            current_table = []
            in_table = False

    for line in lines:
        stripped = line.strip()

        # 마크다운 테이블 감지
        if '|' in stripped and stripped.startswith('|'):
            if not in_table:
                flush_field()
            in_table = True
            current_table.append(stripped)
            continue
        elif in_table:
            flush_table()

        # @필드: 값
        field_match = re.match(r'^@(\S+?):\s*(.*)', stripped)
        if field_match:
            flush_field()
            current_field = field_match.group(1)
            value = field_match.group(2)
            if value:
                current_value_lines = [value]
            else:
                current_value_lines = []
            continue

        # 현재 필드의 연속 값 (불릿 또는 일반 텍스트)
        if current_field and stripped:
            current_value_lines.append(stripped)
            continue

        # 비어있는 줄
        if not stripped and current_field:
            flush_field()

    # 남은 것 처리
    flush_field()
    flush_table()


def parse_md_table(table_lines: List[str]) -> Optional[Dict]:
    """마크다운 테이블을 파싱하여 2D 배열로 반환"""
    if len(table_lines) < 2:
        return None

    rows = []
    for line in table_lines:
        # 구분선(---|---) 건너뛰기
        cells = [c.strip() for c in line.strip('|').split('|')]
        if all(re.match(r'^[-:]+$', c) for c in cells if c):
            continue
        rows.append(cells)

    if not rows:
        return None

    return {
        'headers': rows[0] if rows else [],
        'rows': rows[1:] if len(rows) > 1 else [],
        'raw_rows': rows,
    }


def parse_bullets(text: str) -> List[Dict]:
    """불릿 텍스트를 파싱"""
    bullets = []
    for line in text.split('\n'):
        line = line.strip()
        match = re.match(r'^[-*+]\s+(.*)', line)
        if match:
            bullets.append({
                'level': 0,
                'text': match.group(1)
            })
        elif re.match(r'^\d+\.\s+(.*)', line):
            match = re.match(r'^\d+\.\s+(.*)', line)
            bullets.append({
                'level': 0,
                'text': match.group(1)
            })
        elif line:
            bullets.append({
                'level': 0,
                'text': line
            })
    return bullets


def split_slide_blocks(md_text: str) -> tuple:
    """
    확장 MD를 파싱하여 (config, slide_blocks) 반환.
    slide_blocks: [{index, ref_slide, template, slide_md(원본 텍스트)}, ...]
    """
    config = {}
    config_match = re.search(r'---config\s*\n(.*?)\n---', md_text, re.DOTALL)
    if config_match:
        for line in config_match.group(1).strip().split('\n'):
            line = line.strip()
            if ':' in line:
                key, val = line.split(':', 1)
                config[key.strip()] = val.strip()

    # ---slide 구분자로 분할하되 원본 텍스트 보존
    parts = re.split(r'(---slide\s*\n(?:#\s*\[S\d+\].*\n)?)', md_text)

    blocks = []
    idx = 0
    for i in range(len(parts)):
        if re.match(r'---slide\s*\n', parts[i]):
            raw_block = parts[i + 1] if i + 1 < len(parts) else ''
            slide_data = parse_slide_block(raw_block)
            if slide_data:
                blocks.append({
                    'index': idx,
                    'ref_slide': slide_data.get('ref_slide'),
                    'template': slide_data.get('template'),
                    'slide_md': '---slide\n' + raw_block,
                })
                idx += 1

    return config, blocks


if __name__ == '__main__':
    # 테스트
    test_md = """
---config
reference_pptx: reference-ppt/test.pptx
master_template: templates/master.pptx
---

---slide
template: T1
ref_slide: 5
---
@governing_message: 제도 기반 인프라 구축의 문제를 해결하기 위해 체계적 접근이 필요합니다.
@section_title: 현황 및 문제점
@카드1_제목: 제도적 과제
@카드1_내용: 수원기관은 국가 주도의 디지털 교육 전환 로드맵에 참여하고 있으나, 법적 근거가 미흡합니다.
@카드2_제목: 기술적 과제
@카드2_내용: 이러닝 전용 스튜디오, 편집실, 서버실 등 핵심 인프라가 부재합니다.

---slide
template: T6
ref_slide: 22
---
@governing_message: 성과관리 방법론을 적용하여 체계적으로 관리합니다.

| 구분 | 방법론 | 비고 |
|---|---|---|
| 정보수집 | 설문조사 | 분기별 |
| 분석 | 통계분석 | 연간 |
"""
    import json
    result = parse_md(test_md)
    print(json.dumps(result, ensure_ascii=False, indent=2))
