# -*- coding: utf-8 -*-
"""
pptx-vertical-writer MCP Server

확장 마크다운(Extended MD) → PowerPoint 프레젠테이션 변환.
프로젝트 폴더의 templates/slides/ 에서 1장짜리 PPTX를 가져와 조립합니다.

실행: python server.py
"""

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
# MCP 도구 (1개)
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


# ---------------------------------------------------------------------------
# 진입점
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
