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


def _build_name_map(shapes_com):
    """COM shapes에서 name → shape 매핑 딕셔너리 구축"""
    name_map = {}
    for i in range(1, shapes_com.Count + 1):
        try:
            shape = shapes_com(i)
            name_map[shape.Name] = shape
        except Exception:
            pass
    return name_map


def _get_shape_name(slide_info, meta_index):
    """slide_index.json의 shapes 배열에서 meta_index번째 shape의 이름 반환"""
    shapes_meta = slide_info.get('shapes', [])
    if 0 <= meta_index < len(shapes_meta):
        return shapes_meta[meta_index].get('name', '')
    return ''


def mark_unreplaced_shapes(slide_com, slide_info: dict, fields: dict):
    """
    교체되지 않은 텍스트 shape에 ★미교체★ 표식을 삽입.
    Shape name 기반 매칭으로 COM 복사 후에도 정확한 shape을 찾습니다.
    """
    role_map = slide_info.get('role_map', {})
    shapes_meta = slide_info.get('shapes', [])
    shapes_com = slide_com.Shapes
    name_map = _build_name_map(shapes_com)

    # 어떤 shape name이 교체되었는지 추적
    replaced_names = set()

    # 카드 테이블
    for n, si in enumerate(role_map.get('card_table', []), 1):
        if f'카드{n}_제목' in fields or f'카드{n}_내용' in fields:
            replaced_names.add(_get_shape_name(slide_info, si))

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
                replaced_names.add(_get_shape_name(slide_info, si))

    # 교체되지 않은 shape에 표식
    for si, meta in enumerate(shapes_meta):
        role = meta.get('role', 'unknown')
        if role in SKIP_MARK_ROLES:
            continue

        shape_name = meta.get('name', '')
        if shape_name in replaced_names:
            continue

        text = meta.get('text', '')
        if not text or len(text) <= 5:
            continue

        shape = name_map.get(shape_name)
        if shape:
            try:
                if shape.HasTextFrame:
                    current = shape.TextFrame.TextRange.Text
                    if current and not current.startswith(MARKER):
                        shape.TextFrame.TextRange.Text = MARKER + current
            except Exception:
                pass


def _collect_all_text_shapes(shapes_com):
    """COM 슬라이드의 모든 텍스트 가능 shape을 재귀 수집 (그룹 내부 포함)"""
    result = []
    for i in range(1, shapes_com.Count + 1):
        shape = shapes_com(i)
        try:
            if shape.Type == 6:  # msoGroup
                for gi in range(1, shape.GroupItems.Count + 1):
                    gshape = shape.GroupItems(gi)
                    try:
                        if gshape.HasTextFrame:
                            result.append(gshape)
                    except Exception:
                        pass
            elif shape.HasTextFrame:
                result.append(shape)
        except Exception:
            pass
    return result


def _find_shape_by_role_hint(text_shapes, role_hint):
    """역할 힌트로 shape을 추정 매칭 (Governing Message, breadcrumb 등)"""
    for shape in text_shapes:
        try:
            name = shape.Name.lower()

            if role_hint == 'governing_message':
                # 부제목 placeholder 또는 Governing Message 근처의 긴 텍스트
                if '부제목' in name or 'subtitle' in name.lower():
                    return shape
            elif role_hint == 'breadcrumb':
                # 제목 placeholder
                if '제목' in name and '부제목' not in name:
                    return shape
                if 'title' in name.lower() and 'subtitle' not in name.lower():
                    return shape
            elif role_hint == 'content_1' or role_hint == 'content':
                # content_box 역할이지만 role_map에 없는 경우
                if '둥근' in name or '양쪽' in name or '모서리' in name:
                    return shape
        except Exception:
            pass
    return None


def apply_fields_com(slide_com, slide_info: dict, fields: dict, tables: list = None):
    """
    PowerPoint COM 슬라이드에 필드 데이터를 적용.
    Shape name 기반 매칭 + 그룹 내부 탐색 fallback.
    """
    role_map = slide_info.get('role_map', {})
    shapes_com = slide_com.Shapes

    # COM shape들의 name → shape 매핑
    name_map = _build_name_map(shapes_com)

    # 그룹 내부 포함 모든 텍스트 shape 수집 (fallback용)
    all_text_shapes = _collect_all_text_shapes(shapes_com)

    # 교체 완료 추적
    applied_fields = set()

    def get_com_shape(meta_index):
        """slide_index.json의 0-based index에서 shape name을 조회 → COM shape 반환"""
        shape_name = _get_shape_name(slide_info, meta_index)
        if shape_name and shape_name in name_map:
            return name_map[shape_name]
        # fallback: 인덱스 기반 (name이 없거나 매칭 안 될 때)
        com_idx = meta_index + 1
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
            individual_key = f'{role_name}_{n}'
            if individual_key in fields:
                shape = get_com_shape(si)
                if shape:
                    if replace_shape_text_com(shape, fields[individual_key]):
                        applied_fields.add(individual_key)
            elif role_name in fields:
                shape = get_com_shape(si)
                if shape:
                    if replace_shape_text_com(shape, fields[role_name]):
                        applied_fields.add(role_name)

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

    # =====================================================================
    # Fallback: role_map에서 매칭 못 한 필드를 그룹 내부까지 탐색하여 교체
    # =====================================================================
    remaining = {k: v for k, v in fields.items() if k not in applied_fields}
    if remaining:
        # governing_message → 부제목/subtitle shape 탐색
        if 'governing_message' in remaining:
            shape = _find_shape_by_role_hint(all_text_shapes, 'governing_message')
            if shape:
                replace_shape_text_com(shape, remaining.pop('governing_message'))

        # breadcrumb → 제목/title shape 탐색
        if 'breadcrumb' in remaining:
            shape = _find_shape_by_role_hint(all_text_shapes, 'breadcrumb')
            if shape:
                replace_shape_text_com(shape, remaining.pop('breadcrumb'))

        # content_1 등 남은 content 필드 → 순서대로 텍스트 shape에 배치
        content_fields = sorted([k for k in remaining if k.startswith('content_')])
        if content_fields:
            # 텍스트가 있는(블록처리된) shape 중 아직 교체 안 된 것들
            available = []
            for shape in all_text_shapes:
                try:
                    t = shape.TextFrame.TextRange.Text
                    if t and '████' in t:
                        available.append(shape)
                except Exception:
                    pass

            for i, key in enumerate(content_fields):
                if i < len(available):
                    replace_shape_text_com(available[i], remaining[key])


def build_single_slide(slide_data: dict, slides_info: dict, output_path: str, slides_dir: str):
    """
    1장짜리 PPTX 생성: 템플릿 복사 → COM으로 텍스트 교체 → 저장.
    slide_data: parse_slide_block() 결과 (ref_slide, fields, tables 등)
    slides_info: {slide_number: slide_info, ...}
    """
    ref_slide_num = slide_data.get('ref_slide')
    if ref_slide_num is None:
        raise ValueError('ref_slide가 지정되지 않았습니다.')

    slide_file = os.path.join(slides_dir, f'S{ref_slide_num:04d}.pptx')
    if not os.path.exists(slide_file):
        raise FileNotFoundError(f'슬라이드 파일 없음: {slide_file}')

    # 템플릿 복사
    shutil.copy2(slide_file, output_path)

    fields = slide_data.get('fields', {})
    tables = slide_data.get('tables', [])
    if not fields and not tables:
        return output_path

    # COM으로 열어서 텍스트 교체
    pp = _get_powerpoint()
    try:
        prs = pp.Presentations.Open(os.path.abspath(output_path), WithWindow=False)
        slide_info = slides_info.get(ref_slide_num, {})
        apply_fields_com(prs.Slides(1), slide_info, fields, tables)
        prs.Save()
        prs.Close()
    finally:
        try:
            pp.Quit()
        except Exception:
            pass

    return output_path


def merge_pptx_files_safe(part_files: list, output_path: str, batch_size: int = 30):
    """여러 PPTX를 합치기. 파일 수가 많으면 2단계 merge."""
    if not part_files:
        raise ValueError('합칠 파일이 없습니다.')

    if len(part_files) == 1:
        shutil.copy2(part_files[0], output_path)
        return

    if len(part_files) <= batch_size:
        _merge_pptx_files(list(part_files), output_path)
        return

    # 2단계 merge: batch_size씩 중간파일 생성 → 중간파일 합치기
    intermediate_files = []
    for i in range(0, len(part_files), batch_size):
        batch = part_files[i:i + batch_size]
        if len(batch) == 1:
            intermediate_files.append(batch[0])
            continue
        intermediate = f'{output_path}.merge_{i}.pptx'
        _merge_pptx_files(list(batch), intermediate)
        intermediate_files.append(intermediate)

    _merge_pptx_files(intermediate_files, output_path)

    # 중간파일 정리
    for f in intermediate_files:
        if f != output_path and os.path.exists(f) and f not in part_files:
            try:
                os.remove(f)
            except Exception:
                pass


BATCH_SIZE = 20  # 배치당 슬라이드 수


def _build_batch(md_slides_batch, ref_pptx_path, slides_info, output_path, slides_dir=None):
    """배치 단위로 PPTX 생성. slides_dir이 있으면 1장짜리 파일에서 복사 (고속)."""
    import time

    pp = _get_powerpoint()
    try:
        built = 0

        if slides_dir and os.path.isdir(slides_dir):
            # === 분할 모드: 1장짜리 PPTX에서 복사 (고속) ===
            first_slide = md_slides_batch[0]
            first_ref = first_slide.get('ref_slide', 1)
            first_file = os.path.join(slides_dir, f'S{first_ref:04d}.pptx')

            if not os.path.exists(first_file):
                raise FileNotFoundError(f'슬라이드 파일 없음: {first_file}')

            # 첫 번째 슬라이드 파일을 복사해서 타겟으로 사용
            temp_path = output_path + '.tmp.pptx'
            shutil.copy2(first_file, temp_path)
            target_prs = pp.Presentations.Open(os.path.abspath(temp_path), WithWindow=False)

            # 첫 슬라이드 텍스트 교체
            slide_info = slides_info.get(first_ref, {})
            fields = first_slide.get('fields', {})
            tables = first_slide.get('tables', [])
            if fields or tables:
                apply_fields_com(target_prs.Slides(1), slide_info, fields, tables)
            built += 1

            # 나머지 슬라이드를 1장씩 열어서 복사
            for slide_data in md_slides_batch[1:]:
                ref_slide_num = slide_data.get('ref_slide')
                if ref_slide_num is None:
                    continue

                slide_file = os.path.join(slides_dir, f'S{ref_slide_num:04d}.pptx')
                if not os.path.exists(slide_file):
                    print(f'  Warning: {slide_file} not found, skipping')
                    continue

                src_prs = pp.Presentations.Open(os.path.abspath(slide_file), WithWindow=False)
                try:
                    src_prs.Slides(1).Copy()
                    time.sleep(0.3)
                    target_prs.Slides.Paste()
                except Exception as e:
                    print(f'  Error copying slide {ref_slide_num}: {e}')
                    src_prs.Close()
                    continue
                src_prs.Close()

                # 텍스트 교체
                new_slide = target_prs.Slides(target_prs.Slides.Count)
                slide_info = slides_info.get(ref_slide_num, {})
                fields = slide_data.get('fields', {})
                tables = slide_data.get('tables', [])
                if fields or tables:
                    apply_fields_com(new_slide, slide_info, fields, tables)
                built += 1

            target_prs.SaveAs(os.path.abspath(output_path), 24)
            target_prs.Close()

            if os.path.exists(temp_path):
                os.remove(temp_path)

        else:
            # === 기존 모드: 큰 placeholder에서 복사 (하위 호환) ===
            ref_prs_cache = {}
            ref_prs = pp.Presentations.Open(ref_pptx_path, WithWindow=False)
            ref_prs_cache[ref_pptx_path] = ref_prs

            def get_ref_prs(pptx_path):
                abs_path = os.path.abspath(pptx_path)
                if abs_path not in ref_prs_cache:
                    if not os.path.exists(abs_path):
                        raise FileNotFoundError(f'참조 PPTX: {abs_path}')
                    ref_prs_cache[abs_path] = pp.Presentations.Open(abs_path, WithWindow=False)
                return ref_prs_cache[abs_path]

            temp_path = output_path + '.tmp.pptx'
            shutil.copy2(ref_pptx_path, temp_path)
            target_prs = pp.Presentations.Open(temp_path, WithWindow=False)

            for i in range(target_prs.Slides.Count, 0, -1):
                target_prs.Slides(i).Delete()

            for slide_data in md_slides_batch:
                ref_slide_num = slide_data.get('ref_slide')
                if ref_slide_num is None:
                    continue

                slide_ref_pptx = slide_data.get('reference_pptx')
                current_ref_prs = get_ref_prs(slide_ref_pptx) if slide_ref_pptx else ref_prs

                slide_info = slides_info.get(ref_slide_num, {})
                source_slide_num = slide_info.get('source_slide', ref_slide_num)

                if source_slide_num < 1 or source_slide_num > current_ref_prs.Slides.Count:
                    print(f'  Warning: source_slide {source_slide_num} out of range, skipping')
                    continue

                for attempt in range(3):
                    try:
                        current_ref_prs.Slides(source_slide_num).Copy()
                        time.sleep(0.8)
                        target_prs.Slides.Paste()
                        break
                    except Exception as e:
                        if attempt < 2:
                            time.sleep(1)
                        else:
                            print(f'  Error copying slide {ref_slide_num}: {e}')

                new_slide = target_prs.Slides(target_prs.Slides.Count)
                fields = slide_data.get('fields', {})
                tables = slide_data.get('tables', [])
                if fields or tables:
                    apply_fields_com(new_slide, slide_info, fields, tables)
                built += 1

            target_prs.SaveAs(output_path, 24)
            target_prs.Close()

            for cached_prs in ref_prs_cache.values():
                try:
                    cached_prs.Close()
                except Exception:
                    pass

            if os.path.exists(temp_path):
                os.remove(temp_path)

        return built

    finally:
        try:
            pp.Quit()
        except Exception:
            pass


def _merge_pptx_files(part_files, output_path):
    """여러 PPTX 파일을 하나로 합치기"""
    import time

    if len(part_files) == 1:
        shutil.move(part_files[0], output_path)
        return

    pp = _get_powerpoint()
    try:
        # 첫 번째 파일을 기반으로
        target_prs = pp.Presentations.Open(os.path.abspath(part_files[0]), WithWindow=False)

        # 나머지 파일의 슬라이드를 추가
        for part_file in part_files[1:]:
            src_prs = pp.Presentations.Open(os.path.abspath(part_file), WithWindow=False)
            for si in range(1, src_prs.Slides.Count + 1):
                src_prs.Slides(si).Copy()
                time.sleep(0.5)
                target_prs.Slides.Paste()
            src_prs.Close()

        target_prs.SaveAs(os.path.abspath(output_path), 24)
        target_prs.Close()

    finally:
        try:
            pp.Quit()
        except Exception:
            pass

    # 임시 파일 삭제
    for f in part_files:
        try:
            os.remove(f)
        except Exception:
            pass


def build_presentation(md_data: dict, slide_index: dict, output_path: str, slides_dir: str = None):
    """
    PowerPoint COM을 사용하여 PPTX 생성.

    전략: 배치 단위(20장)로 분할 생성 후 합치기.
    - slides_dir이 있으면 1장짜리 파일에서 복사 (고속, 권장)
    - 없으면 큰 placeholder에서 복사 (하위 호환)
    """
    config = md_data.get('config', {})
    ref_pptx_path = os.path.abspath(config.get('reference_pptx', ''))

    if not os.path.exists(ref_pptx_path):
        raise FileNotFoundError(f'참조 PPTX 파일을 찾을 수 없습니다: {ref_pptx_path}')

    output_path = os.path.abspath(output_path)
    slides_info = {s['slide_number']: s for s in slide_index.get('slides', [])}

    md_slides = md_data.get('slides', [])
    total = len(md_slides)

    if total == 0:
        raise ValueError('슬라이드가 없습니다.')

    # 배치 분할
    batches = []
    for i in range(0, total, BATCH_SIZE):
        batches.append(md_slides[i:i + BATCH_SIZE])

    print(f'Total slides: {total}, Batches: {len(batches)} (batch size: {BATCH_SIZE})')

    # 각 배치를 개별 PPTX로 생성
    part_files = []
    for bi, batch in enumerate(batches):
        part_path = f'{output_path}.part{bi + 1}.pptx'
        print(f'\n--- Batch {bi + 1}/{len(batches)} ({len(batch)} slides) ---')
        built = _build_batch(batch, ref_pptx_path, slides_info, part_path, slides_dir=slides_dir)
        print(f'  Batch {bi + 1} done: {built} slides')
        if os.path.exists(part_path):
            part_files.append(part_path)

    # 합치기
    if part_files:
        print(f'\nMerging {len(part_files)} parts...')
        _merge_pptx_files(part_files, output_path)
        print(f'Saved: {output_path}')
        print(f'Total slides: {total}')
    else:
        raise RuntimeError('모든 배치 생성 실패')
