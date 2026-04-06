---
name: create-proposal-pptx
description: RFP와 참고자료를 기반으로 A3 세로형 PPTX 제안서를 작성합니다. "PPTX 제안서 만들어줘", "세로형 슬라이드 제안서", "프로젝트 폴더 기반 PPTX" 등의 요청 시 사용합니다.
version: 2.0.0
allowed-tools: [Read, Write, Glob, Bash, mcp__pptx_vertical_writer__list_templates, mcp__pptx_vertical_writer__showcase_templates, mcp__pptx_vertical_writer__create_pptx, mcp__pptx_vertical_writer__match_slide, mcp__pptx_vertical_writer__add_template]
---

# PPTX 제안서 오케스트레이션 스킬

당신은 프로젝트 폴더의 RFP, 참고자료, 데이터를 분석하고, 로컬 템플릿 MD 파일을 참조하여 A3 세로형 PPTX 제안서를 작성하는 전문가입니다.

## 프로젝트 폴더 구조

```
프로젝트폴더/
├── docs/          ← GUIDE.md + T0~T9.md (템플릿 구조 + 복사용 스니펫)
├── templates/     ← 블록처리된 참조 PPTX (placeholder_vol2/vol3.pptx)
├── rfp/           ← RFP/제안요청서
├── rawdata/       ← 통계, 보고서, 데이터
├── references/    ← 참고 자료 (기존 제안서, 선행 연구 등)
└── output/        ← 생성물 (PPTX + 확장MD)
```

## 실행 워크플로우

### Step 0: 스타트 프롬프트 확인 (있는 경우)
1. 프로젝트 폴더에 `start_prompt.md`가 있는지 확인합니다
2. **있으면**: start_prompt.md를 읽고 그 지시를 따릅니다 (파일 목록, 파트 구분, 볼륨 범위가 이미 정리되어 있음)
3. **없으면**: 아래 Step 1부터 직접 분석합니다

> `start_prompt.md`는 1단계 도구(ppt-block-maker)가 자동 생성합니다.
> 이 파일에는 프로젝트 내 참고자료 목록, 슬라이드 현황, 작업 규칙이 포함되어 있습니다.

### Step 1: 프로젝트 폴더 분석
1. 프로젝트 폴더 경로를 확인합니다
2. `docs/GUIDE.md`를 읽어 전체 구조를 파악합니다
3. `rfp/사업개요.md`, `rfp/목차.md`가 있으면 우선 읽습니다
4. rfp/ 내 RFP 문서를 읽어 요구사항을 분석합니다
5. **RFP에 목차/분량/템플릿 지정이 있으면 그대로 따릅니다**

### Step 2: 목차(Backbone) 생성
1. RFP 분석 결과를 기반으로 각 섹션의 템플릿 타입(T0~T9)을 결정합니다
2. `docs/T?.md`를 읽어 내용에 맞는 ref_slide를 선택합니다:
   - 카드가 3개 필요하면 → T1.md에서 카드=4인 ref_slide 선택
   - 큰 테이블이면 → T6.md에서 적합한 ref_slide 선택
3. backbone 표를 사용자에게 보여주고 확인을 받습니다:

```markdown
| # | 제목 | 템플릿 | ref_slide | 비고 |
|---|---|---|---|---|
| S01 | 사업추진 배경 | T2 | 77 | 카드2+세분화 |
| S02 | 사업의 범위 | T3 | 1 | 거버닝+영역 |
| S03 | 현황 분석 | T1 | 8 | 카드4개 |
```

4. 확인된 backbone을 `proposal-backbone.md`로 저장합니다

### Step 3: 본문 작성 (확장 MD)
1. `docs/slides/S????.md`가 있으면 해당 슬라이드의 원본 텍스트 + @필드 구조를 참고합니다
2. 없으면 `docs/T?.md`에서 해당 ref_slide의 **복사용 스니펫**을 가져옵니다
3. @필드를 실제 내용으로 채웁니다
3. 출처 주석을 추가합니다:
   ```
   @카드1_내용: 실제 내용
   <!-- [rawdata] 파일명, p.페이지 -->
   ```
4. 작성된 확장 MD를 `proposal-body-extended.md`로 저장합니다

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
- 색상 마커(`{{red:}}`, `{{green:}}`) 대신 HTML 주석 출처 표기를 사용합니다
