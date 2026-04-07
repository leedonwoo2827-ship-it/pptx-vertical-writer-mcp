# pptx-vertical-writer

Claude Desktop에서 제안서 본문을 작성하는 MCP 서버입니다.
[pptxpipe](https://github.com/leedonwoo2827-ship-it/pptxpipe)가 1단계(분석)와 3단계(PPTX 변환)를, 이 서버가 2단계(AI 글쓰기)를 담당합니다.

## 파이프라인

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│  1단계  pptxpipe           PPTX 분석 + 템플릿 분리           │
│         (명령 프롬프트)     → slide_index, 카탈로그, 가이드   │
│                    │                                        │
│                    ▼  start_prompt.md                       │
│                                                             │
│  2단계  pptx-vertical-writer   AI가 마크다운으로 본문 작성    │
│         (Claude Desktop)       → proposal-body-partN.md     │
│                    │                                        │
│                    ▼  작성된 마크다운                        │
│                                                             │
│  3단계  pptxpipe build     마크다운 → PPTX 변환              │
│         (명령 프롬프트)     → 최종 result.pptx               │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

| 단계 | 저장소 | 도구 |
|------|--------|------|
| **1단계** PPTX 분석 | [pptxpipe](https://github.com/leedonwoo2827-ship-it/pptxpipe) | 명령 프롬프트 |
| **2단계** AI 글쓰기 | **이 저장소** | Claude Desktop |
| **3단계** PPTX 변환 | [pptxpipe](https://github.com/leedonwoo2827-ship-it/pptxpipe) | 명령 프롬프트 |

## 설치

```bash
git clone https://github.com/leedonwoo2827-ship-it/pptx-vertical-writer.git
cd pptx-vertical-writer
pip install -r requirements.txt
```

### Claude Desktop 연결

`claude_desktop_config.json`에 아래를 추가합니다.

```json
{
  "mcpServers": {
    "pptx-vertical-writer": {
      "command": "python",
      "args": ["server.py"],
      "cwd": "설치경로/pptx-vertical-writer"
    }
  }
}
```

## 사용 방법

1. 1단계(pptxpipe)에서 생성된 `start_prompt.md`를 Claude Desktop에 붙여넣기
2. 프로젝트 폴더를 Claude Desktop에 연결 (📎 → Add folder)
3. AI가 RFP·참고자료를 읽고 backbone → 본문 순서로 작성
4. 완료된 `proposal-body-partN.md`를 3단계(pptxpipe build)로 PPTX 변환

## MCP 도구

| 도구 | 설명 |
|---|---|
| `parse_md_slides` | 확장 MD 파싱 → 슬라이드 목록 JSON 반환 (작성 중 검증용) |

## 확장 마크다운 포맷

AI가 작성하는 제안서 본문의 형식입니다.

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

## 프로젝트 폴더 구조

사용자가 작업하는 프로젝트 폴더입니다. 1단계에서 자동 생성됩니다.

```
프로젝트폴더/
├── input/             ← 원본 PPTX
├── templates/         ← slide_index.json + slides/S????.pptx
├── docs/              ← GUIDE.md + T0~T9.md + CATALOG.md
├── rfp/               ← RFP, 제안요청서
├── rawdata/           ← 통계, 보고서, 발표자료
├── references/        ← 기존 제안서
├── start_prompt.md    ← 1단계에서 자동 생성된 시작 프롬프트
└── output/            ← 생성된 PPTX (3단계 결과물)
```
