# -*- coding: utf-8 -*-
"""
md2ppt MCP Server

참조 PPTX 템플릿 기반으로 마크다운을 PowerPoint 프레젠테이션으로 변환하는 MCP 서버.
회사 표준 템플릿의 레이아웃/디자인을 100% 보존하면서 텍스트만 교체합니다.

실행: python server.py
"""

import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR / "src"))

from mcp.server.fastmcp import FastMCP
from md_parser import parse_md
from slide_builder import build_presentation, load_slide_index

mcp = FastMCP("pptx-vertical-writer")

DEFAULT_TEMPLATE_DIR = BASE_DIR / "templates"
DEFAULT_SLIDE_INDEX = DEFAULT_TEMPLATE_DIR / "slide_index.json"
DEFAULT_PLACEHOLDER_PPTX = DEFAULT_TEMPLATE_DIR / "placeholder.pptx"
TEMPLATE_IMAGES_DIR = BASE_DIR / "template-images" / "all-slides"


# ---------------------------------------------------------------------------
# 내부 유틸리티
# ---------------------------------------------------------------------------

def _resolve_output(project_dir: str, output_file: str, fallback: str = "output.pptx") -> Path:
    if project_dir:
        proj = Path(project_dir)
        out_dir = proj / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        name = Path(output_file).name if output_file else fallback
        return out_dir / name
    if output_file:
        p = Path(output_file)
        return p if p.is_absolute() else Path.home() / "Documents" / output_file
    return Path.home() / "Documents" / fallback


# ---------------------------------------------------------------------------
# MCP 도구
# ---------------------------------------------------------------------------

@mcp.tool()
def list_templates() -> str:
    """사용 가능한 슬라이드 템플릿 목록을 반환합니다.

    각 템플릿 타입(T0~T9)의 용도, 슬라이드 수, 대표 이미지 경로를 포함합니다.
    글쓰기 스킬에서 적합한 템플릿을 추천할 때 사용합니다.

    Returns:
        템플릿 카탈로그 (JSON)
    """
    if not DEFAULT_SLIDE_INDEX.exists():
        return "오류: slide_index.json이 없습니다. analyze_template()을 먼저 실행하세요."

    with open(DEFAULT_SLIDE_INDEX, 'r', encoding='utf-8') as f:
        idx = json.load(f)

    # 템플릿별 통계
    from collections import Counter
    tmpl_counts = Counter(s['template'] for s in idx.get('slides', []))

    TEMPLATE_NAMES = {
        'T0': '구분페이지 — 섹션 구분, 챕터 시작. 글: 제목 1줄',
        'T1': '카드형 다중 — 현황/문제점/개선, 전략/방법론. 글: 카드 2~6개(제목+내용)',
        'T2': '카드+다이어그램 — 목적/전략 카드 + 세분화. 글: 카드 + 불릿 리스트',
        'T3': '범위/개요 — Governing Message + 둥근사각형 그리드. 글: 핵심 문구 + 6개 영역',
        'T4': '다중 데이터테이블 — 복수 테이블 + 일정표. 글: 마크다운 테이블',
        'T5': '테이블+다이어그램 — 단일 테이블 + shape 설명. 글: 테이블 + 설명',
        'T6': '순수 데이터테이블 — 큰 테이블 중심. 글: 마크다운 테이블',
        'T7': '테이블+설명shape — 프로세스 화살표 + 테이블. 글: 단계별 설명 + 테이블',
        'T8': '이미지중심 — 조직도, 스크린샷, 구성도. 글: 이미지 경로 + 캡션',
        'T9': '핵심메시지/다이어그램 — 둥근사각형 라벨, 프로세스 흐름도. 글: 핵심 문구 리스트',
        'T14': '기타/특수 — 분류 외 레이아웃',
    }

    catalog = []
    for tmpl_code in sorted(tmpl_counts.keys()):
        # 대표 슬라이드 번호
        rep_slides = [s['slide_number'] for s in idx['slides'] if s['template'] == tmpl_code][:3]
        # 대표 이미지 경로
        images = []
        for sn in rep_slides:
            img = TEMPLATE_IMAGES_DIR / f"S{sn:03d}_{tmpl_code}.png"
            if img.exists():
                images.append(str(img))

        catalog.append({
            'template': tmpl_code,
            'name': TEMPLATE_NAMES.get(tmpl_code, '기타'),
            'slide_count': tmpl_counts[tmpl_code],
            'representative_slides': rep_slides,
            'preview_images': images,
        })

    return json.dumps(catalog, ensure_ascii=False, indent=2)


@mcp.tool()
def analyze_template(
    pptx_path: str,
    output_dir: str = ""
) -> str:
    """참조 PPTX를 분석하여 slide_index.json과 placeholder 템플릿을 생성합니다.

    새로운 회사 템플릿을 등록할 때 1회 실행합니다.
    참조 PPTX의 각 슬라이드를 분석하여 shape별 역할(제목/본문/카드/테이블)을 자동 분류하고,
    텍스트를 플레이스홀더(████/Lorem)로 치환한 깨끗한 템플릿을 생성합니다.

    Args:
        pptx_path: 분석할 참조 PPTX 파일 경로
        output_dir: 출력 디렉토리 (생략 시 templates/)

    Returns:
        분석 결과 요약
    """
    from template_extractor import extract_slide_index
    from template_sanitizer import sanitize_pptx, sanitize_slide_index

    pptx = Path(pptx_path)
    if not pptx.exists():
        return f"오류: PPTX 파일을 찾을 수 없습니다: {pptx_path}"

    out = Path(output_dir) if output_dir else DEFAULT_TEMPLATE_DIR
    out.mkdir(parents=True, exist_ok=True)

    idx_path = str(out / "slide_index.json")
    placeholder_path = str(out / "placeholder.pptx")

    try:
        # 1. 슬라이드 분석 → slide_index.json
        slide_index = extract_slide_index(str(pptx), idx_path)

        # 2. 플레이스홀더 PPTX 생성
        sanitize_pptx(str(pptx), idx_path, placeholder_path)

        # 3. slide_index.json 텍스트도 치환
        sanitize_slide_index(idx_path)

        total = slide_index['total_slides']
        from collections import Counter
        dist = Counter(s['template'] for s in slide_index['slides'])

        result = f"분석 완료!\n"
        result += f"  슬라이드: {total}장\n"
        result += f"  slide_index: {idx_path}\n"
        result += f"  placeholder: {placeholder_path}\n"
        result += f"\n템플릿 분포:\n"
        for t, c in dist.most_common():
            result += f"  {t}: {c}장\n"

        return result

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def create_pptx(
    extended_md: str,
    output_file: str = "",
    project_dir: str = "",
    template_dir: str = ""
) -> str:
    """확장 마크다운으로부터 PowerPoint 프레젠테이션을 생성합니다.

    확장 MD 포맷:
    ---config
    reference_pptx: path/to/placeholder.pptx  (생략 시 기본 템플릿 사용)
    ---

    ---slide
    template: T1
    ref_slide: 5
    ---
    @governing_message: 거버닝 메시지 텍스트
    @카드1_제목: 첫 번째 카드 제목
    @카드1_내용: 첫 번째 카드 내용
    @content_1: 본문 텍스트

    Args:
        extended_md: 확장 마크다운 텍스트 (---slide 구분자 포함)
        output_file: 출력 PPTX 파일명
        project_dir: 프로젝트 폴더 (지정 시 output/ 하위에 저장)
        template_dir: 템플릿 디렉토리 (생략 시 기본 templates/)

    Returns:
        생성된 파일 경로
    """
    tmpl_dir = Path(template_dir) if template_dir else DEFAULT_TEMPLATE_DIR
    idx_path = tmpl_dir / "slide_index.json"
    placeholder = tmpl_dir / "placeholder.pptx"

    if not idx_path.exists():
        return "오류: slide_index.json이 없습니다. analyze_template()을 먼저 실행하세요."

    out_path = _resolve_output(project_dir, output_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        # MD 파싱
        md_data = parse_md(extended_md)

        # reference_pptx 기본값 설정
        if 'reference_pptx' not in md_data.get('config', {}):
            md_data.setdefault('config', {})['reference_pptx'] = str(placeholder)

        config_ref = md_data['config'].get('reference_pptx', '')
        if not Path(config_ref).is_absolute():
            md_data['config']['reference_pptx'] = str(BASE_DIR / config_ref)

        slide_index = load_slide_index(str(idx_path))

        # PPTX 생성
        build_presentation(md_data, slide_index, str(out_path))

        if out_path.exists():
            size = out_path.stat().st_size
            slide_count = len(md_data.get('slides', []))
            return f"PPTX 생성 완료!\n파일: {out_path}\n크기: {size:,} bytes\n슬라이드: {slide_count}장"
        return f"오류: 파일 생성 실패: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# 진입점
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
