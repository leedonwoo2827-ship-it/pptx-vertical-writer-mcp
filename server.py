# -*- coding: utf-8 -*-
"""
pptx-vertical-writer MCP Server  (2단계)

Claude Desktop에서 제안서 확장 마크다운(Extended MD)을 작성하고,
작성된 MD로부터 PPTX를 생성할 수 있는 MCP 서버입니다.

파이프라인:
  1단계 ppt-block-maker     → 원본 PPTX 분석, 템플릿 생성
  2단계 이 서버              → 템플릿 매칭 + AI 기반 제안서 MD 작성
  3단계 md2verticalpptx     → 확장 MD → PPTX 변환 (CLI)

실행: python server.py
"""

import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR / "src"))

from mcp.server.fastmcp import FastMCP
from md_parser import parse_md, split_slide_blocks
from slide_builder import build_presentation, load_slide_index, build_single_slide, merge_pptx_files_safe, verify_pptx

mcp = FastMCP("pptx-vertical-writer")

DEFAULT_TEMPLATE_DIR = BASE_DIR / "templates"
DEFAULT_SLIDE_INDEX = DEFAULT_TEMPLATE_DIR / "slide_index.json"


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
# MCP 도구 (4개)
# ---------------------------------------------------------------------------

@mcp.tool()
def create_pptx(
    extended_md: str = "",
    md_file: str = "",
    output_file: str = "",
    project_dir: str = "",
    template_dir: str = ""
) -> str:
    """확장 마크다운으로부터 PowerPoint 프레젠테이션을 생성합니다.

    2가지 입력 방식:
    1. md_file: 확장 MD 파일 경로 (긴 문서용, 권장)
    2. extended_md: 확장 MD 텍스트 직접 전달 (짧은 문서용)

    md_file이 지정되면 extended_md보다 우선합니다.

    Args:
        extended_md: 확장 마크다운 텍스트 (짧은 문서용)
        md_file: 확장 마크다운 파일 경로 (긴 문서용, 권장)
        output_file: 출력 PPTX 파일명
        project_dir: 프로젝트 폴더 (지정 시 output/ 하위에 저장)
        template_dir: 템플릿 디렉토리 (생략 시 기본 templates/)

    Returns:
        생성된 파일 경로
    """
    tmpl_dir = Path(template_dir) if template_dir else DEFAULT_TEMPLATE_DIR
    idx_path = tmpl_dir / "slide_index.json"

    if not idx_path.exists():
        return "오류: slide_index.json이 없습니다."

    out_path = _resolve_output(project_dir, output_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        # md_file 우선, 없으면 extended_md 사용
        md_text = extended_md
        if md_file:
            md_path = Path(md_file)
            if not md_path.is_absolute() and project_dir:
                md_path = Path(project_dir) / md_file
            if not md_path.exists():
                return f"오류: MD 파일을 찾을 수 없습니다: {md_path}"
            md_text = md_path.read_text(encoding='utf-8')

        if not md_text.strip():
            return "오류: 확장 MD 내용이 비어있습니다. extended_md 또는 md_file을 지정하세요."

        # MD 파싱
        md_data = parse_md(md_text)

        # reference_pptx 기본값 (slides_dir 모드에서는 사용 안 하지만 fallback용)
        if 'reference_pptx' not in md_data.get('config', {}):
            md_data.setdefault('config', {})['reference_pptx'] = str(tmpl_dir / "placeholder.pptx")

        config_ref = md_data['config'].get('reference_pptx', '')
        if not Path(config_ref).is_absolute():
            if project_dir:
                resolved = Path(project_dir) / config_ref
                if resolved.exists():
                    md_data['config']['reference_pptx'] = str(resolved)
                else:
                    md_data['config']['reference_pptx'] = str(BASE_DIR / config_ref)
            else:
                md_data['config']['reference_pptx'] = str(BASE_DIR / config_ref)

        slide_index = load_slide_index(str(idx_path))

        # slides_dir 탐색: project_dir/templates/slides/ 가 있으면 분할 모드 (고속)
        slides_dir = None
        if project_dir:
            sd = Path(project_dir) / "templates" / "slides"
            if sd.is_dir() and any(sd.glob("S*.pptx")):
                slides_dir = str(sd)

        # PPTX 생성
        build_presentation(md_data, slide_index, str(out_path), slides_dir=slides_dir)

        if out_path.exists():
            size = out_path.stat().st_size
            slide_count = len(md_data.get('slides', []))
            return f"PPTX 생성 완료!\n파일: {out_path}\n크기: {size:,} bytes\n슬라이드: {slide_count}장"
        return f"오류: 파일 생성 실패: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def parse_md_slides(
    md_file: str = "",
    extended_md: str = "",
    project_dir: str = "",
    template_dir: str = ""
) -> str:
    """확장 마크다운을 파싱하여 슬라이드 목록을 반환합니다.

    각 슬라이드의 index, ref_slide, template, 원본 MD 텍스트를 포함합니다.
    이 결과를 build_slide에 하나씩 전달하여 개별 PPTX를 생성할 수 있습니다.

    Args:
        md_file: 확장 마크다운 파일 경로
        extended_md: 확장 마크다운 텍스트 (md_file 우선)
        project_dir: 프로젝트 폴더
        template_dir: 템플릿 디렉토리 (생략 시 기본값)

    Returns:
        JSON: {config: {...}, slides: [{index, ref_slide, template, slide_md}, ...], total: N}
    """
    import json

    try:
        md_text = extended_md
        if md_file:
            md_path = Path(md_file)
            if not md_path.is_absolute() and project_dir:
                md_path = Path(project_dir) / md_file
            if not md_path.exists():
                return f"오류: MD 파일을 찾을 수 없습니다: {md_path}"
            md_text = md_path.read_text(encoding='utf-8')

        if not md_text.strip():
            return "오류: MD 내용이 비어있습니다."

        config, blocks = split_slide_blocks(md_text)
        return json.dumps({
            'config': config,
            'slides': blocks,
            'total': len(blocks),
        }, ensure_ascii=False)

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def build_slide(
    slide_md: str,
    slide_index_num: int = 0,
    output_dir: str = "",
    project_dir: str = "",
    template_dir: str = ""
) -> str:
    """단일 슬라이드 PPTX를 생성합니다.

    parse_md_slides로 얻은 slide_md를 전달하면,
    해당 템플릿을 복사하고 텍스트를 교체하여 1장짜리 PPTX를 저장합니다.

    Args:
        slide_md: 단일 슬라이드의 확장 MD 텍스트 (---slide 블록 1개)
        slide_index_num: 슬라이드 순번 (0-based, 파일명에 사용)
        output_dir: 출력 디렉토리 (개별 PPTX 저장 위치)
        project_dir: 프로젝트 폴더
        template_dir: 템플릿 디렉토리 (생략 시 기본값)

    Returns:
        생성된 PPTX 파일 경로
    """
    tmpl_dir = Path(template_dir) if template_dir else DEFAULT_TEMPLATE_DIR
    idx_path = tmpl_dir / "slide_index.json"

    if not idx_path.exists():
        return "오류: slide_index.json이 없습니다."

    try:
        # slides_dir 탐색
        slides_dir = None
        if project_dir:
            sd = Path(project_dir) / "templates" / "slides"
            if sd.is_dir() and any(sd.glob("S*.pptx")):
                slides_dir = str(sd)
        if not slides_dir:
            sd = tmpl_dir / "slides"
            if sd.is_dir() and any(sd.glob("S*.pptx")):
                slides_dir = str(sd)
        if not slides_dir:
            return "오류: templates/slides/ 디렉토리를 찾을 수 없습니다."

        # 출력 경로
        if output_dir:
            out_dir = Path(output_dir)
        elif project_dir:
            out_dir = Path(project_dir) / "output" / "slides"
        else:
            out_dir = Path.home() / "Documents" / "slides"
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"slide_{slide_index_num:03d}.pptx"

        # MD 파싱
        md_data = parse_md(slide_md)
        slides = md_data.get('slides', [])
        if not slides:
            return f"오류: slide_md에서 슬라이드를 파싱할 수 없습니다."

        slide_data = slides[0]

        # slide_index 로드
        slide_index = load_slide_index(str(idx_path))
        slides_info = {s['slide_number']: s for s in slide_index.get('slides', [])}

        # 단일 슬라이드 생성
        result_path = build_single_slide(slide_data, slides_info, str(out_path), slides_dir)
        size = Path(result_path).stat().st_size
        return f"생성 완료: {result_path} ({size:,} bytes)"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def merge_slides(
    slide_files: str,
    output_file: str = "",
    project_dir: str = ""
) -> str:
    """여러 개별 PPTX 파일을 하나로 합칩니다.

    build_slide로 생성한 개별 파일들을 순서대로 합쳐 최종 PPTX를 만듭니다.

    Args:
        slide_files: PPTX 파일 경로 목록 (JSON 배열 또는 줄바꿈 구분)
        output_file: 최종 출력 파일명
        project_dir: 프로젝트 폴더 (지정 시 output/ 하위에 저장)

    Returns:
        합쳐진 PPTX 파일 경로
    """
    import json as _json

    try:
        # 파일 목록 파싱
        try:
            files = _json.loads(slide_files)
        except (_json.JSONDecodeError, TypeError):
            files = [f.strip() for f in slide_files.strip().split('\n') if f.strip()]

        if not files:
            return "오류: 합칠 파일 목록이 비어있습니다."

        # 존재하지 않거나 손상된 파일 필터링
        existing = [f for f in files if Path(f).exists()]
        if not existing:
            return "오류: 지정된 파일이 모두 존재하지 않습니다."
        valid = [f for f in existing if verify_pptx(f)]
        skipped = len(existing) - len(valid)
        if not valid:
            return "오류: 모든 PPTX 파일이 손상되었습니다."
        if skipped:
            existing = valid

        out_path = _resolve_output(project_dir, output_file, fallback="merged.pptx")
        out_path.parent.mkdir(parents=True, exist_ok=True)

        merge_pptx_files_safe(existing, str(out_path))

        if out_path.exists():
            size = out_path.stat().st_size
            msg = f"병합 완료!\n파일: {out_path}\n크기: {size:,} bytes\n슬라이드: {len(existing)}장"
            if skipped:
                msg += f"\n경고: 손상된 파일 {skipped}개 건너뜀"
            return msg
        return f"오류: 병합 파일 생성 실패: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# 진입점
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
