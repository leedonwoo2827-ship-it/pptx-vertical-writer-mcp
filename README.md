# pptx-vertical-writer

Claude Desktop에서 제안서 본문을 작성하는 MCP 서버입니다.
[pptxpipe](https://github.com/leedonwoo2827-ship-it/pptxpipe)가 1단계(분석)와 3단계(PPTX 변환)를, 이 서버가 2단계(AI 글쓰기)를 담당합니다.

## 파이프라인

```
1단계  pptxpipe vision-extract        PPTX 비전 분석 + 템플릿 분리
       (명령 프롬프트)                 → slide_index.json, metadata/, specs/, screenshots/
                  │
                  ▼  start_prompt.md + run.bat

2단계  pptx-vertical-writer           AI가 대화하며 마크다운 본문 작성
       (Claude Desktop)               → proposal-body-extended.md
                  │
                  ▼  작성된 확장 마크다운

3단계  pptxpipe vbuild                마크다운 → PPTX 변환
       (명령 프롬프트)                 → 최종 result.pptx
```

| 단계 | 저장소 | 도구 | 처리 방식 |
|------|--------|------|-----------|
| **1단계** PPTX 비전 분석 | [pptxpipe](https://github.com/leedonwoo2827-ship-it/pptxpipe) | 명령 프롬프트 | 배치 (자동) |
| **2단계** AI 글쓰기 | **이 저장소** | Claude Desktop | 대화형 (HITL) |
| **3단계** PPTX 변환 | [pptxpipe](https://github.com/leedonwoo2827-ship-it/pptxpipe) | 명령 프롬프트 | 배치 (자동) |

## 왜 2단계는 대화형(Interactive)인가

제안서 글쓰기는 **비결정적(non-deterministic) 출력**입니다. 같은 입력(RFP + 자료)을 줘도 매번 다른 결과가 나옵니다. 이것이 1/3단계(CLI 배치)와 근본적으로 다른 점입니다.

```
backbone 초안 → 사용자 피드백 → backbone 수정 → 승인 → 본문 작성
     ↑                                                    ↓
     └──────── 섹션별 검토 / 보강 요청 ←──────────────────┘
```

| 상황 | CLI 배치 | Claude Desktop 대화형 |
|------|----------|----------------------|
| 목차 뼈대가 RFP 의도와 다름 | 처음부터 재실행 | "2장 구조를 이렇게 바꿔줘" |
| 특정 섹션 분량 부족 | 전체 재생성 | "여기 5장 더 써줘" |
| 출처 자료 반영 누락 | 파라미터 수정 후 재실행 | "rawdata/OO.pdf p.3 참고해서 보강해" |
| 톤/문체 조정 | 불가 | "좀 더 구체적 수치로 써줘" |

- **1단계** = 배치 — 입력 고정, 출력 결정적 → CLI
- **2단계** = 대화형 — 판단 필요, 출력 비결정적 → Claude Desktop (Human-in-the-Loop)
- **3단계** = 배치 — 입력 고정, 출력 결정적 → CLI

2단계를 CLI로 만들면 "한 번에 완벽한 프롬프트"를 짜야 하는데, 제안서 수준의 복잡한 문서에서는 현실적으로 불가능합니다 (one-shot generation의 한계).

## 설치

### 마켓플레이스에서 설치 (권장)

Claude Desktop에서 바로 설치할 수 있습니다.

1. Claude Desktop 좌측 하단 **사용자 지정** 클릭
2. 개인 플러그인 옆 **+** 버튼 → **마켓플레이스 추가** 선택
3. URL 입력란에 아래 주소를 붙여넣고 **동기화** 클릭:
   ```
   https://github.com/leedonwoo2827-ship-it/pptx-vertical-writer
   ```
4. 개인 플러그인 목록에 **Pptx vertical writer**가 나타나면 설치 완료

### 수동 설치

마켓플레이스를 사용하지 않는 경우, 직접 다운로드하여 연결할 수 있습니다.

```bash
git clone https://github.com/leedonwoo2827-ship-it/pptx-vertical-writer.git
cd pptx-vertical-writer
pip install -r requirements.txt
```

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

1. 1단계(pptxpipe `vision-extract`)에서 생성된 `start_prompt.md`를 Claude Desktop에 붙여넣기
2. 프로젝트 폴더를 Claude Desktop에 연결
3. AI가 RFP/참고자료를 읽고 backbone → 본문 순서로 작성
4. **사용자가 대화하며 확인/수정** — 뼈대 조정, 분량 추가, 톤 변경 등
5. 완료된 `proposal-body-extended.md`를 3단계(`run build`)로 PPTX 변환

## MCP 도구

| 도구 | 설명 |
|---|---|
| `parse_md_slides` | 확장 MD 파싱 → 슬라이드 목록 JSON 반환 (작성 중 검증용) |

**파라미터:**
- `md_file`: 확장 마크다운 파일 경로
- `extended_md`: 마크다운 텍스트 (md_file 대신 직접 입력)
- `project_dir`: 프로젝트 폴더 경로 (상대경로 해석용)

**반환:** `{ config, slides: [{ index, ref_slide, template, slide_md }], total }`

## 확장 마크다운 포맷 (비전 기반)

AI가 작성하는 제안서 본문의 형식입니다. area_id는 `templates/specs/slide_NNNN.md` 규격서에서 확인합니다.

```markdown
---slide
# [S2001] 사업추진 배경 및 목적
template: slide_2001
ref_slide: 2001
---
@g1_shape_2: 사업추진 배경
@g2_shape_1: 1-1. 사업이해도 > 1-1-1. 사업추진 배경 및 목적
@g2_shape_3: 나이지리아 공립초중학교의 디지털 교육환경 구축을 통해 교육 불평등을 해소
@shape_7: 교육격차 해소와 품질 개선
@table_4_r0c0: Output 1.1: 디지털 학습환경 조성
@table_4_r1c0: Output 1.2: 통합 교육관리 시스템 구축
@note: [rawdata] 1차사업 슬라이드, p.3-5 / [AI] 일반 지식 기반 작성
```

**area_id 규칙:**
- `shape_N`: 일반 텍스트 영역 (N = shape 인덱스)
- `g{N}_shape_{M}`: 그룹 N 내 shape M
- `table_{N}_r{R}c{C}`: 테이블 N의 R행 C열 (0부터 시작)
- `@note`: 슬라이드 노트 (출처 주석, 발표자 메모)

## 프로젝트 폴더 구조

1단계(`pptxpipe vision-extract`)가 자동 생성합니다.

```
프로젝트폴더/
├── input/                 ← 원본 PPTX (vol2.pptx, vol3.pptx)
├── templates/
│   ├── slides/            ← 개별 슬라이드 PPTX (S2001.pptx ~ S3047.pptx)
│   ├── specs/             ← 슬라이드 규격서 (slide_2001.md ~ slide_3047.md)
│   ├── metadata/          ← shape 좌표/크기 JSON (slide_2001.json ~ )
│   ├── screenshots/       ← 슬라이드 스크린샷 PNG
│   └── slide_index.json   ← 전체 슬라이드 카탈로그
├── rfp/                   ← RFP, 제안요청서, 목차
├── rawdata/               ← 통계, 보고서, 발표자료
├── references/            ← 기존 제안서
├── start_prompt.md        ← 1단계에서 자동 생성된 시작 프롬프트
├── run.bat                ← 1단계에서 자동 생성된 CLI 실행 스크립트
└── output/                ← 생성된 PPTX (3단계 결과물)
```

## CLI 명령어 (run.bat)

`start_prompt.md` 생성 시 `run.bat`이 함께 만들어집니다.

```
run extract          1단계: 스크린샷 + 메타데이터 + 규격서 추출
run spec             규격서만 재생성
run build body.md    3단계: 마크다운 → PPTX 변환
run all [body.md]    전체 파이프라인 (추출 + 변환)
```
