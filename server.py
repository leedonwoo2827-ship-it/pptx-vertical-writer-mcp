# -*- coding: utf-8 -*-
"""
pptx-vertical-writer MCP Server  (2단계)

Claude Desktop에서 제안서 확장 마크다운(Extended MD)을 작성하는 MCP 서버입니다.
PPTX 변환은 3단계 CLI(md2verticalpptx)에서 수행합니다.

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
from md_parser import split_slide_blocks

mcp = FastMCP("pptx-vertical-writer")


# ---------------------------------------------------------------------------
# MCP 도구 (1개)
# ---------------------------------------------------------------------------

@mcp.tool()
def parse_md_slides(
    md_file: str = "",
    extended_md: str = "",
    project_dir: str = ""
) -> str:
    """확장 마크다운을 파싱하여 슬라이드 목록을 반환합니다.

    작성 중인 proposal-body.md를 검증하거나 슬라이드 구성을 확인할 때 사용합니다.

    Args:
        md_file: 확장 마크다운 파일 경로
        extended_md: 확장 마크다운 텍스트 (md_file 우선)
        project_dir: 프로젝트 폴더

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


# ---------------------------------------------------------------------------
# 진입점
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
