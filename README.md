# pptx-vertical-writer-mcp

A3 세로형 PPTX 제안서 작성 MCP 서버 — 3단계 파이프라인의 **2단계**.

## 파이프라인 개요

```
1단계  ppt-block-maker        원본 PPTX → 블록처리 템플릿 + 메타데이터 + 참조 MD
       ↓
2단계  pptx-vertical-writer   참조자료 + AI → 확장 MD 작성 → PPTX 생성 (이 저장소)
       ↓
3단계  md2verticalpptx        확장 MD → 최종 PPTX (CLI 도구)
```

- 1단계: [ppt-block-maker](https://github.com/leedonwoo2827-ship-it/ppt-block-maker)
- 3단계: [md2verticalpptx](https://github.com/leedonwoo2827-ship-it/md2verticalpptx)

## 이 서버가 하는 일

Claude Desktop에서 MCP 서버로 연결하여:

1. 1단계에서 생성된 `docs/`, `templates/slides/`를 참조
2. RFP, rawdata, references를 읽고 확장 마크다운(proposal-body.md) 작성
3. PowerPoint COM API로 PPTX를 직접 생성하거나, 3단계 CLI에 전달

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
| `create_pptx` | 확장 MD → PPTX 일괄 생성 (20장 배치 + 자동 병합) |
| `parse_md_slides` | 확장 MD 파싱 → 슬라이드 목록 JSON 반환 |
| `build_slide` | 단일 슬라이드 PPTX 생성 (MCP 타임아웃 방지용) |
| `merge_slides` | 여러 PPTX를 하나로 병합 (InsertFromFile, 클립보드 미사용) |

### 대량 문서 워크플로우 (20장 이상)

```
parse_md_slides(md_file)  →  build_slide(slide_md) x N회  →  merge_slides(slide_files)
```

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
@note: 출처 주석 및 발표자 노트

---slide
template: T6
ref_slide: 3022
---
@governing_message: 성과관리 방법론

| 구분 | 방법론 | 비고 |
|---|---|---|
| 정보수집 | 설문조사 | 분기별 |
```

## 프로젝트 폴더 구조 (사용자 작업 폴더)

```
프로젝트폴더/
├── docs/              ← GUIDE.md + T0~T9.md + slides/ (1단계 산출물)
├── templates/
│   ├── slide_index.json
│   └── slides/        ← S2001.pptx ~ S3047.pptx (1장짜리)
├── rfp/               ← RFP, 제안요청서
├── rawdata/           ← 통계, 보고서, 발표자료
├── references/        ← 기존 제안서
├── start_prompt.md    ← 1단계에서 자동 생성된 시작 프롬프트
└── output/            ← 생성된 PPTX
```

## 서버 소스 구조

```
pptx-vertical-writer-mcp/
├── server.py              # MCP 서버 (4개 도구)
├── src/
│   ├── md_parser.py       # 확장 MD 파서
│   └── slide_builder.py   # PowerPoint COM 빌더 + InsertFromFile 병합
├── templates/
│   └── slide_index.json   # 기본 슬라이드 메타데이터
├── skills/                # Claude 스킬 정의
└── agents/                # 에이전트 정의
```

## 3단계 CLI 빌드

PPTX 빌드를 터미널에서 직접 실행하려면 [md2verticalpptx](https://github.com/leedonwoo2827-ship-it/md2verticalpptx)를 사용:

```bash
python -m md2pptx proposal-body.md -t templates/slides -o output/result.pptx --continue-on-error -v
```