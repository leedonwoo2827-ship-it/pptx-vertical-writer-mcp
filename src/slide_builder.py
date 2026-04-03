"""
슬라이드 빌더: PowerPoint COM API를 사용하여 참조 PPTX에서 슬라이드를 복제하고 텍스트 교체.

PowerPoint COM 방식의 장점:
- 슬라이드 복제 시 레이아웃/마스터/이미지/서식 100% 보존
- PowerPoint가 직접 처리하므로 호환성 문제 없음
"""
import json
import os
import shutil
import comtypes.client


def load_slide_index(json_path: str) -> dict:
    """slide_index.json 로드"""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _get_powerpoint():
    """PowerPoint COM 객체 생성"""
    pp = comtypes.client.CreateObject('Powerpoint.Application')
    pp.Visible = 1
    return pp


def replace_shape_text_com(shape_com, new_text: str):
    """
    PowerPoint COM shape의 텍스트를 서식 보존하면서 교체.
    TextRange.Text를 설정하면 첫 Run의 서식이 유지됨.
    """
    try:
        if not shape_com.HasTextFrame:
            return False

        tr = shape_com.TextFrame.TextRange
        tr.Text = new_text
        return True
    except Exception as e:
        # 일부 shape은 텍스트 수정 불가
        return False


def replace_table_cell_com(table_com, row, col, new_text: str):
    """PowerPoint COM 테이블 셀 텍스트 교체 (1-based index)"""
    try:
        cell = table_com.Cell(row, col)
        cell.Shape.TextFrame.TextRange.Text = new_text
        return True
    except Exception:
        return False


MARKER = '★미교체★ '

# 표식 대상에서 제외할 역할 (장식, 라벨 등)
SKIP_MARK_ROLES = {
    'group_decoration', 'decoration', 'empty_shape', 'unknown',
    'governing_label', 'chapter_label', 'page_number', 'number_circle',
    'image',
}


def mark_unreplaced_shapes(slide_com, slide_info: dict, fields: dict):
    """
    교체되지 않은 텍스트 shape에 ★미교체★ 표식을 삽입.
    사업자가 수동으로 교체하거나 삭제할 수 있도록 시각적 표시.
    """
    role_map = slide_info.get('role_map', {})
    shapes_meta = slide_info.get('shapes', [])
    shapes_com = slide_com.Shapes

    # 어떤 shape 인덱스가 교체되었는지 추적
    replaced_indices = set()

    # 카드 테이블
    for si in role_map.get('card_table', []):
        for n, idx in enumerate(role_map.get('card_table', []), 1):
            if idx == si:
                if f'카드{n}_제목' in fields or f'카드{n}_내용' in fields:
                    replaced_indices.add(si)

    # 일반 역할
    role_groups = {
        'governing_message': role_map.get('governing_message', []),
        'breadcrumb': role_map.get('breadcrumb', []),
        'section_title': role_map.get('section_title', []),
        'content': sorted(role_map.get('content_box', []) + role_map.get('content_shape', [])),
        'heading': role_map.get('heading_box', []),
        'label': sorted(role_map.get('label_box', []) + role_map.get('label_shape', [])),
        'text': role_map.get('text_content', []),
    }

    for role_name, indices in role_groups.items():
        for n, si in enumerate(indices, 1):
            individual_key = f'{role_name}_{n}'
            if individual_key in fields or role_name in fields:
                replaced_indices.add(si)

    # 교체되지 않은 shape에 표식
    for si, meta in enumerate(shapes_meta):
        role = meta.get('role', 'unknown')
        if role in SKIP_MARK_ROLES:
            continue
        if si in replaced_indices:
            continue

        text = meta.get('text', '')
        if not text or len(text) <= 5:
            continue

        com_idx = si + 1
        if 1 <= com_idx <= shapes_com.Count:
            shape = shapes_com(com_idx)
            try:
                if shape.HasTextFrame:
                    current = shape.TextFrame.TextRange.Text
                    if current and not current.startswith(MARKER):
                        shape.TextFrame.TextRange.Text = MARKER + current
            except Exception:
                pass


def apply_fields_com(slide_com, slide_info: dict, fields: dict, tables: list = None):
    """
    PowerPoint COM 슬라이드에 필드 데이터를 적용.
    slide_info의 role_map을 기반으로 shape을 찾아 텍스트 교체.
    """
    role_map = slide_info.get('role_map', {})

    # COM shapes는 1-based index
    shapes_com = slide_com.Shapes

    def get_com_shape(meta_index):
        """slide_index.json의 0-based index를 COM 1-based index로 변환하여 shape 반환"""
        com_idx = meta_index + 1  # 1-based
        if 1 <= com_idx <= shapes_com.Count:
            return shapes_com(com_idx)
        return None

    # 역할별 shape 인덱스를 통합 정리
    role_groups = {
        'governing_message': role_map.get('governing_message', []),
        'breadcrumb': role_map.get('breadcrumb', []),
        'section_title': role_map.get('section_title', []),
        'content': sorted(role_map.get('content_box', []) + role_map.get('content_shape', [])),
        'heading': role_map.get('heading_box', []),
        'label': sorted(role_map.get('label_box', []) + role_map.get('label_shape', [])),
        'text': role_map.get('text_content', []),
    }

    # 일반 텍스트 shape 교체 (모든 역할에 대해 _N 인덱싱 지원)
    for role_name, indices in role_groups.items():
        for n, si in enumerate(indices, 1):
            # @role (전체 교체) 또는 @role_N (개별 교체) 지원
            individual_key = f'{role_name}_{n}'
            if individual_key in fields:
                shape = get_com_shape(si)
                if shape:
                    replace_shape_text_com(shape, fields[individual_key])
            elif role_name in fields:
                # _N 없이 role만 있으면 모든 해당 shape에 같은 텍스트
                shape = get_com_shape(si)
                if shape:
                    replace_shape_text_com(shape, fields[role_name])

    # 카드 테이블 교체 (카드N_제목, 카드N_내용)
    card_indices = role_map.get('card_table', [])
    for card_num, si in enumerate(card_indices, 1):
        title_key = f'카드{card_num}_제목'
        content_key = f'카드{card_num}_내용'
        shape = get_com_shape(si)
        if shape and shape.HasTable:
            table = shape.Table
            if title_key in fields and table.Rows.Count >= 1:
                replace_table_cell_com(table, 1, 1, fields[title_key])
            if content_key in fields and table.Rows.Count >= 2:
                replace_table_cell_com(table, 2, 1, fields[content_key])

    # text_content 교체 (하위 호환)
    for t_num, si in enumerate(role_map.get('text_content', []), 1):
        key = f'text_{t_num}'
        if key in fields:
            shape = get_com_shape(si)
            if shape:
                replace_shape_text_com(shape, fields[key])

    # 데이터 테이블 교체
    if tables:
        dt_indices = role_map.get('data_table', [])
        for ti, table_data in enumerate(tables):
            if ti < len(dt_indices):
                si = dt_indices[ti]
                shape = get_com_shape(si)
                if shape and shape.HasTable:
                    all_rows = table_data.get('raw_rows', [])
                    if not all_rows:
                        all_rows = [table_data.get('headers', [])] + table_data.get('rows', [])
                    table_com = shape.Table
                    for ri, row_data in enumerate(all_rows):
                        if ri >= table_com.Rows.Count:
                            break
                        for ci, cell_text in enumerate(row_data):
                            if ci >= table_com.Columns.Count:
                                break
                            replace_table_cell_com(table_com, ri + 1, ci + 1, str(cell_text))


def build_presentation(md_data: dict, slide_index: dict, output_path: str):
    """
    PowerPoint COM을 사용하여 PPTX 생성.

    전략:
    1. 참조 PPTX를 PowerPoint에서 열기
    2. 필요한 슬라이드만 새 프레젠테이션에 복사
    3. 텍스트 교체
    4. 저장
    """
    config = md_data.get('config', {})
    ref_pptx_path = os.path.abspath(config.get('reference_pptx', ''))

    if not os.path.exists(ref_pptx_path):
        raise FileNotFoundError(f'참조 PPTX 파일을 찾을 수 없습니다: {ref_pptx_path}')

    output_path = os.path.abspath(output_path)

    # 슬라이드 인덱스에서 참조 정보 조회
    slides_info = {s['slide_number']: s for s in slide_index.get('slides', [])}

    # PowerPoint COM 시작
    pp = _get_powerpoint()

    try:
        # 참조 PPTX 열기 (기본 + 슬라이드별 오버라이드 캐시)
        ref_prs_cache = {}  # path -> COM Presentation
        ref_prs = pp.Presentations.Open(ref_pptx_path, WithWindow=False)
        ref_prs_cache[ref_pptx_path] = ref_prs

        def get_ref_prs(pptx_path):
            """캐시된 참조 프레젠테이션 반환, 없으면 열기"""
            abs_path = os.path.abspath(pptx_path)
            if abs_path not in ref_prs_cache:
                if not os.path.exists(abs_path):
                    raise FileNotFoundError(f'슬라이드별 참조 PPTX 파일을 찾을 수 없습니다: {abs_path}')
                ref_prs_cache[abs_path] = pp.Presentations.Open(abs_path, WithWindow=False)
            return ref_prs_cache[abs_path]

        # 새 프레젠테이션 생성 (참조와 같은 크기)
        # 방법: 참조 파일을 복사해서 열고, 불필요한 슬라이드 제거
        temp_path = output_path + '.tmp.pptx'
        shutil.copy2(ref_pptx_path, temp_path)
        target_prs = pp.Presentations.Open(temp_path, WithWindow=False)

        # 원본의 모든 슬라이드 삭제 (뒤에서부터)
        for i in range(target_prs.Slides.Count, 0, -1):
            target_prs.Slides(i).Delete()

        # 필요한 슬라이드를 참조에서 복사
        md_slides = md_data.get('slides', [])
        for slide_data in md_slides:
            ref_slide_num = slide_data.get('ref_slide')
            if ref_slide_num is None:
                continue

            # 슬라이드별 reference_pptx 오버라이드 지원
            slide_ref_pptx = slide_data.get('reference_pptx')
            if slide_ref_pptx:
                current_ref_prs = get_ref_prs(slide_ref_pptx)
            else:
                current_ref_prs = ref_prs

            # slide_index에서 해당 슬라이드 정보 조회
            slide_info = slides_info.get(ref_slide_num, {})

            # source_slide: PPTX 내 실제 슬라이드 번호 (merge 시 가상 번호와 다를 수 있음)
            source_slide_num = slide_info.get('source_slide', ref_slide_num)

            if source_slide_num < 1 or source_slide_num > current_ref_prs.Slides.Count:
                print(f'  Warning: source_slide {source_slide_num} (ref_slide {ref_slide_num}) out of range, skipping')
                continue

            # 참조 슬라이드를 타겟에 복사 (클립보드 안정성을 위해 재시도)
            import time
            for attempt in range(3):
                try:
                    current_ref_prs.Slides(source_slide_num).Copy()
                    time.sleep(0.5)
                    target_prs.Slides.Paste()
                    break
                except Exception as e:
                    if attempt < 2:
                        time.sleep(1)
                    else:
                        raise e

            # 방금 붙여넣은 슬라이드 (마지막)
            new_slide_idx = target_prs.Slides.Count
            new_slide = target_prs.Slides(new_slide_idx)
            fields = slide_data.get('fields', {})
            tables = slide_data.get('tables', [])
            if fields or tables:
                apply_fields_com(new_slide, slide_info, fields, tables)

            # 미교체 텍스트에 ★미교체★ 표식 삽입
            mark_unreplaced_shapes(new_slide, slide_info, fields)

            print(f'  Built slide from ref #{ref_slide_num} (template: {slide_data.get("template", "?")})')

        # 저장
        # ppSaveAsOpenXMLPresentation = 24
        target_prs.SaveAs(output_path, 24)
        target_prs.Close()

        # 모든 캐시된 참조 프레젠테이션 닫기
        for cached_prs in ref_prs_cache.values():
            try:
                cached_prs.Close()
            except Exception:
                pass

        # 임시 파일 정리
        if os.path.exists(temp_path):
            os.remove(temp_path)

        print(f'\nSaved: {output_path}')
        print(f'Total slides: {len(md_slides)}')

    finally:
        try:
            pp.Quit()
        except Exception:
            pass
