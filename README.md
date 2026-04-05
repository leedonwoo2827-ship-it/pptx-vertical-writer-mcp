# pptx-vertical-writer-mcp

A3 세로형 PPTX 템플릿 기반 제안서 작성 MCP 서버.

확장 마크다운(Extended MD)을 파싱하여 참조 PPTX 슬라이드를 복제·텍스트 교체·병합합니다.
PowerPoint COM API를 사용하여 레이아웃/서식/이미지를 100% 보존합니다.

## 요구사항

- Windows + Microsoft PowerPoint 설치
- Python 3.10+

## 설치

```bash
git clone https://github.com/leedonwoo2827-ship-it/pptx-vertical-writer-mcp.git
cd pptx-vertical-writer-mcp
pip install -r requirements.txt
```

## MCP 도구 (4개)

| 도구 | 설명 |
|---|---|
| `create_pptx(extended_md, md_file, output_file, project_dir)` | 확장 MD → PPTX 일괄 생성 (20장 배치 + 자동 병합) |
| `parse_md_slides(md_file, extended_md)` | 확장 MD 파싱 → 슬라이드 목록 JSON 반환 |
| `build_slide(slide_md, slide_index_num)` | 단일 슬라이드 PPTX 생성 (MCP 타임아웃 방지용) |
| `merge_slides(slide_files, output_file)` | 여러 PPTX 파일을 하나로 병합 (InsertFromFile) |

## 확장 마크다운 포맷

```markdown
---config
reference_pptx: templates/placeholder_vol3.pptx
---

---slide
template: T1
ref_slide: 3005
---
@governing_message: 핵심 메시지
@breadcrumb: III. 사업관리 > 1. 투입인력
@카드1_제목: 카드 제목
@카드1_내용: 카드 본문

---slide
template: T6
ref_slide: 3022
---
@governing_message: 성과관리 방법론

| 구분 | 방법론 | 비고 |
|---|---|---|
| 정보수집 | 설문조사 | 분기별 |
```

## 프로젝트 구조

```
pptx-vertical-writer-mcp/
├── server.py              # MCP 서버 (4개 도구)
├── src/
│   ├── md_parser.py       # 확장 MD 파서
│   ├── slide_builder.py   # PowerPoint COM 빌더 + InsertFromFile 병합
│   ├── template_extractor.py  # 참조 PPTX 분석 → slide_index.json
│   ├── template_matcher.py    # 콘텐츠 → 템플릿 매칭
│   ├── template_sanitizer.py  # 민감정보 블록처리 (████)
│   └── template_splitter.py   # PPTX → 1장짜리 분할
├── templates/
│   ├── slide_index.json   # 175슬라이드 메타데이터 (10종 T0~T9)
│   └── slides/            # S2001.pptx ~ S3047.pptx (1장짜리)
├── skills/                # Claude 스킬 정의
└── agents/                # 에이전트 정의
```

## CLI 빌드 도구

PPTX 빌드를 터미널에서 직접 실행하려면 [md2pptx](https://github.com/leedonwoo2827-ship-it/md2pptx)를 사용하세요.

```bash
python -m md2pptx body.md -t ./templates/slides -o output/result.pptx
```
