---
name: create-proposal-pptx
description: RFP와 참고자료를 기반으로 A3 세로형 PPTX 제안서를 작성합니다. "PPTX 제안서 만들어줘", "세로형 슬라이드 제안서", "프로젝트 폴더 기반 PPTX" 등의 요청 시 사용합니다.
version: 1.0.0
allowed-tools: [Read, Write, Glob, Bash, mcp__pptx_vertical_writer__list_templates, mcp__pptx_vertical_writer__showcase_templates, mcp__pptx_vertical_writer__create_pptx, mcp__pptx_vertical_writer__match_slide]
---

# PPTX 제안서 오케스트레이션 스킬

당신은 프로젝트 폴더의 RFP, 참고자료, 데이터를 분석하여 회사 표준 A3 세로형 PPTX 템플릿에 맞춰 제안서를 작성하는 전문가입니다.

## 프로젝트 폴더 구조

```
프로젝트폴더/
├── templates/     ← 회사 PPTX 마스터 템플릿 (없으면 플러그인 기본 사용)
├── rfp/           ← RFP/제안요청서 (목차·분량·템플릿 지정 포함 가능)
├── rawdata/       ← 통계, 보고서, 데이터
├── references/    ← 참고 자료 (기존 제안서, 선행 연구 등)
└── output/        ← 생성물 (PPTX + 확장MD)
```

## 실행 워크플로우

### Step 1: 프로젝트 폴더 분석
1. 프로젝트 폴더 경로를 확인합니다
2. 각 하위 폴더(rfp/, templates/, rawdata/, references/)의 파일 목록을 확인합니다
3. rfp/ 내 RFP 문서를 읽어 요구사항을 분석합니다
4. **RFP에 목차/분량/템플릿 지정이 있으면 그대로 따릅니다**

### Step 2: 목차(Backbone) 생성
1. `mcp__pptx_vertical_writer__showcase_templates()`를 호출하여 사용 가능한 템플릿을 확인합니다
2. RFP 분석 결과를 기반으로 슬라이드 목차를 설계합니다

**RFP에 목차가 지정된 경우:**
→ 지정된 목차와 템플릿 배정을 그대로 사용

**지정 안 된 경우:**
→ AI가 RFP 요구사항에 따라 섹션 구조를 설계하고 최적 템플릿을 배정

3. backbone 표를 사용자에게 보여주고 확인을 받습니다:

```markdown
| # | 제목 | 템플릿 | ref_slide | 비고 |
|---|---|---|---|---|
| S01 | 사업추진 배경 및 목적 | T2 | 1 | 장기/단기 목적 |
| S02 | 사업의 범위 | T3 | 2 | 5대 Output |
| S03 | 과업대상지 분석 | T1 | 3 | 4개 카드 |
| ... | ... | ... | ... | ... |
```

4. 확인된 backbone을 `proposal-backbone.md`로 저장합니다

### Step 3: 본문 작성 (확장 MD)
1. backbone의 각 슬라이드에 대해 확장 MD를 작성합니다
2. 각 슬라이드의 템플릿에 맞는 @필드를 채웁니다

**참고자료 반영 규칙:**
- rawdata/ 출처: `{{red:텍스트}}` 색상 마커
- references/ 출처: `{{green:텍스트}}` 색상 마커
- 단순 인용이 아니라 재구성·논리적 통합·전문적 문장화

**템플릿별 글 작성 가이드:**
- **T0 (구분페이지)**: 섹션 제목 1줄 → `@content_1`
- **T1 (카드형)**: 카드 2~6개, 각 카드 제목(15자 이내) + 내용(2~4문장, 300자 이내) → `@카드N_제목`, `@카드N_내용`
- **T3 (거버닝메시지)**: 핵심 메시지 1~2문장(200자) + 영역별 설명 → `@governing_message`, `@content_N`
- **T6 (데이터테이블)**: 마크다운 테이블 → 테이블 블록
- **T7 (프로세스)**: 단계명 + 설명 + 테이블 → `@heading_N`, `@content_N`, 테이블
- **T9 (핵심메시지)**: 핵심 문구 3~6개(각 50~100자) → `@content_N`

3. 작성된 확장 MD를 `proposal-body-extended.md`로 저장합니다

### Step 4: PPTX 생성
1. 저장된 확장 MD 파일을 읽습니다
2. `mcp__pptx_vertical_writer__create_pptx(extended_md=텍스트, project_dir=경로)`를 호출합니다
3. output/ 폴더에 PPTX가 생성됩니다

**주의: MCP 도구는 1회 호출당 1개만 사용합니다 (타임아웃 방지)**

### Step 5: 결과 보고
- 생성된 파일 경로, 슬라이드 수, 파일 크기를 보고합니다
- 미입력 필드가 있으면 목록을 함께 안내합니다

## 주의사항
- PowerPoint가 설치된 Windows 환경에서만 동작합니다
- 긴 문장은 슬라이드에 맞게 300자 이내로 요약합니다
- 각 단계마다 사용자 확인을 받습니다 (특히 backbone)
- MCP 도구 호출은 한 턴에 1회만 합니다
