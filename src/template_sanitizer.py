"""
참조 PPTX의 모든 텍스트를 플레이스홀더(████ 또는 Lorem ipsum)로 치환하여
민감 정보 없는 깨끗한 템플릿 PPTX를 생성.

사용법:
    python src/template_sanitizer.py reference.pptx slide_index.json output.pptx
"""
import json
import os
import shutil
import sys
import time
import comtypes.client

# Lorem ipsum 더미 텍스트 (긴 본문용)
LOREM = (
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. "
    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris "
    "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in "
    "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla "
    "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in "
    "culpa qui officia deserunt mollit anim id est laborum. "
    "Curabitur pretium tincidunt lacus. Nulla gravida orci a odio. "
    "Nullam varius, turpis et commodo pharetra, est eros bibendum elit. "
)

# 치환하지 않을 역할 (장식, 구조 요소)
SKIP_ROLES = {
    'group_decoration', 'decoration', 'empty_shape', 'unknown',
    'image', 'page_number', 'chapter_label',
}

# 고정 텍스트로 유지할 역할 (범용 라벨)
KEEP_LABEL_ROLES = {
    'governing_label',  # "Governing Message" 라벨
    'number_circle',    # 번호 원
}


def make_block_text(length):
    """원본 길이에 비례하는 블록 문자(████) 생성"""
    if length <= 5:
        return '████'
    elif length <= 15:
        return '████ ████'
    elif length <= 30:
        return '████ ████████ ████'
    else:
        # 긴 텍스트는 블록으로 여러 줄
        blocks_per_line = 4
        num_lines = min(length // 40 + 1, 5)
        line = '████ ' * blocks_per_line
        return '\n'.join([line.strip()] * num_lines)


def make_lorem_text(length):
    """원본 길이에 맞춘 Lorem ipsum 더미텍스트"""
    if length <= 30:
        return LOREM[:length]
    elif length <= 100:
        return LOREM[:length]
    else:
        # 긴 텍스트는 Lorem 반복
        result = LOREM
        while len(result) < length:
            result += LOREM
        return result[:length]


def get_placeholder_text(role, original_text, original_length):
    """역할에 따라 적절한 플레이스홀더 텍스트 반환"""
    if role in SKIP_ROLES:
        return None  # 건드리지 않음

    if role in KEEP_LABEL_ROLES:
        return None  # 원본 유지

    if not original_text or original_length <= 2:
        return None

    # 짧은 텍스트 (라벨, 제목) → 블록문자
    if original_length <= 30:
        return make_block_text(original_length)

    # 긴 텍스트 (본문, 설명) → Lorem ipsum
    return make_lorem_text(original_length)


def sanitize_pptx(ref_pptx_path, slide_index_path, output_pptx_path):
    """참조 PPTX를 플레이스홀더로 치환한 템플릿 PPTX 생성"""

    # slide_index.json 로드
    with open(slide_index_path, 'r', encoding='utf-8') as f:
        slide_index = json.load(f)

    ref_pptx_path = os.path.abspath(ref_pptx_path)
    output_pptx_path = os.path.abspath(output_pptx_path)

    # 참조 PPTX 복사
    shutil.copy2(ref_pptx_path, output_pptx_path)

    # PowerPoint COM으로 열기
    pp = comtypes.client.CreateObject('Powerpoint.Application')
    pp.Visible = 1

    try:
        prs = pp.Presentations.Open(output_pptx_path, WithWindow=False)
        total_slides = prs.Slides.Count
        replaced_count = 0
        skipped_count = 0

        for slide_info in slide_index.get('slides', []):
            sn = slide_info['slide_number']
            if sn > total_slides:
                continue

            slide = prs.Slides(sn)
            shapes_meta = slide_info.get('shapes', [])
            shapes_com = slide.Shapes

            for si, meta in enumerate(shapes_meta):
                role = meta.get('role', 'unknown')
                original_text = meta.get('text', '')
                original_length = meta.get('text_length', len(original_text))

                placeholder = get_placeholder_text(role, original_text, original_length)

                if placeholder is None:
                    skipped_count += 1
                    continue

                com_idx = si + 1  # 1-based
                if com_idx > shapes_com.Count:
                    continue

                shape = shapes_com(com_idx)

                try:
                    if shape.HasTextFrame:
                        shape.TextFrame.TextRange.Text = placeholder
                        replaced_count += 1
                    elif shape.HasTable:
                        table = shape.Table
                        for ri in range(1, table.Rows.Count + 1):
                            for ci in range(1, table.Columns.Count + 1):
                                try:
                                    cell_text = table.Cell(ri, ci).Shape.TextFrame.TextRange.Text
                                    if cell_text and len(cell_text.strip()) > 2:
                                        table.Cell(ri, ci).Shape.TextFrame.TextRange.Text = make_block_text(len(cell_text))
                                        replaced_count += 1
                                except Exception:
                                    pass
                except Exception:
                    skipped_count += 1

            if sn % 20 == 0:
                print(f'  Processed {sn}/{total_slides}...')

        # 저장 (ppSaveAsOpenXMLPresentation = 24)
        prs.Save()
        prs.Close()

        print(f'\nSanitized: {output_pptx_path}')
        print(f'  Replaced: {replaced_count} shapes')
        print(f'  Skipped: {skipped_count} shapes')

    finally:
        try:
            pp.Quit()
        except Exception:
            pass


def sanitize_slide_index(slide_index_path, output_path=None):
    """slide_index.json의 텍스트 필드도 플레이스홀더로 치환"""
    with open(slide_index_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    for slide in data.get('slides', []):
        for shape in slide.get('shapes', []):
            role = shape.get('role', 'unknown')
            text = shape.get('text', '')
            length = shape.get('text_length', len(text))

            if role in SKIP_ROLES or role in KEEP_LABEL_ROLES:
                continue

            if text and length > 2:
                if length <= 30:
                    shape['text'] = make_block_text(length)
                else:
                    shape['text'] = make_lorem_text(min(length, 100))

            # table_preview도 치환
            if 'table_preview' in shape:
                shape['table_preview'] = [
                    make_block_text(len(cell)) if cell and len(cell) > 2 else cell
                    for cell in shape['table_preview']
                ]

    out = output_path or slide_index_path
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'Sanitized slide_index: {out}')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python template_sanitizer.py <ref_pptx> <slide_index.json> [output.pptx]')
        sys.exit(1)

    ref_pptx = sys.argv[1]
    slide_idx = sys.argv[2]
    output = sys.argv[3] if len(sys.argv) > 3 else 'templates/placeholder.pptx'

    # 1. PPTX 치환
    sanitize_pptx(ref_pptx, slide_idx, output)

    # 2. slide_index.json도 치환
    sanitize_slide_index(slide_idx)
