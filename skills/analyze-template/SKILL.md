---
name: analyze-template
description: 참조 PPTX를 분석하여 슬라이드 템플릿을 등록합니다. "템플릿 분석", "PPTX 분석", "새 템플릿 등록", "참조 PPT 등록" 등의 요청 시 사용합니다.
version: 1.0.0
allowed-tools: [Read, Write, Glob, Bash, mcp__pptx_vertical_writer__analyze_template, mcp__pptx_vertical_writer__list_templates]
---

# PPTX 템플릿 분석 스킬

당신은 참조 PPTX 파일을 분석하여 재사용 가능한 슬라이드 템플릿으로 등록하는 전문가입니다.

## 워크플로우

### 1단계: 참조 PPTX 확인
- 사용자가 제공한 PPTX 파일 경로를 확인합니다
- 파일이 존재하는지 검증합니다

### 2단계: 분석 실행
- `mcp__pptx_vertical_writer__analyze_template(pptx_path)`를 호출합니다
- 이 도구는:
  1. 모든 슬라이드의 shape 구조를 분석
  2. shape별 역할(제목/본문/카드/테이블)을 자동 분류
  3. 10개 템플릿 타입(T0~T9)으로 분류
  4. 텍스트를 플레이스홀더(████/Lorem)로 치환한 깨끗한 템플릿 생성
  5. slide_index.json과 placeholder.pptx를 templates/ 폴더에 저장

### 3단계: 결과 보고
- 분석된 슬라이드 수와 템플릿 분포를 사용자에게 보여줍니다
- 각 템플릿 타입의 용도를 설명합니다

### 4단계: PNG 미리보기 생성 (선택)
- 사용자가 원하면 PowerPoint COM으로 각 슬라이드의 PNG 미리보기를 생성합니다
- template-images/all-slides/ 폴더에 저장합니다

## 템플릿 타입 설명
| 코드 | 이름 | 용도 |
|---|---|---|
| T0 | 구분페이지 | 섹션 구분, 챕터 시작 |
| T1 | 카드형 다중 | 현황/문제점/개선, 전략/방법론 |
| T2 | 카드+다이어그램 | 목적/전략 + 세분화 |
| T3 | 범위/개요 | Governing Message + 영역 그리드 |
| T4 | 다중 데이터테이블 | 복수 테이블 + 일정표 |
| T5 | 테이블+다이어그램 | 테이블 + shape 설명 |
| T6 | 순수 데이터테이블 | 큰 테이블 중심 |
| T7 | 테이블+설명shape | 프로세스 + 테이블 |
| T8 | 이미지중심 | 조직도, 스크린샷 |
| T9 | 핵심메시지 | 라벨 그리드, 흐름도 |

## 주의사항
- PowerPoint가 설치된 Windows 환경에서만 동작합니다
- 참조 PPTX의 원본 텍스트는 플레이스홀더로 치환되어 민감 정보가 제거됩니다
- 생성된 templates/ 폴더의 파일은 깃허브에 안전하게 올릴 수 있습니다
