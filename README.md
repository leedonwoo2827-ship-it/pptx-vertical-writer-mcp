# pptx-vertical-writer-mcp

회사 표준 A3 세로형 PPTX 템플릿 기반 제안서 작성 도구.
RFP 분석 → 목차 설계 → 본문 작성 → PPTX 생성까지 오케스트레이션합니다.

## 2가지 핵심 기능

### 기능1: 제안서 오케스트레이션
```
프로젝트 폴더 (rfp/ + rawdata/ + references/)
  → RFP 분석 → 목차+템플릿 배정 → 본문 작성 → PPTX 생성
```

### 기능2: 템플릿 관리
```
기존 템플릿 쇼케이스 (10종 T0~T9, PNG 미리보기)
  + 새 슬라이드 분석 → 기존 템플릿 매칭 또는 신규 등록
```

## 설치

```bash
git clone https://github.com/leedonwoo2827-ship-it/pptx-vertical-writer-mcp.git
cd pptx-vertical-writer-mcp
pip install -r requirements.txt
```

## 요구사항
- Windows + Microsoft PowerPoint 설치
- Python 3.10+

## MCP 도구 (6개)

| 도구 | 설명 |
|---|---|
| `list_templates()` | 템플릿 목록 (간략) |
| `showcase_templates()` | 템플릿 쇼케이스 (PNG + 필드 가이드) |
| `match_slide(pptx, num)` | 슬라이드↔템플릿 매칭 |
| `add_template(pptx, num, name)` | 새 템플릿 등록 |
| `analyze_template(pptx)` | 참조 PPTX 전체 분석 |
| `create_pptx(md, output)` | 확장 MD → PPTX 생성 |

## 스킬 (4개)

| 스킬 | 설명 |
|---|---|
| `create-proposal-pptx` | RFP→분석→MD→PPTX 전체 워크플로우 |
| `manage-templates` | 템플릿 쇼케이스/매칭/등록 |
| `write-pptx` | 확장 MD → PPTX 단독 변환 |
| `analyze-template` | 참조 PPTX 분석 |

## 프로젝트 폴더 구조
```
프로젝트폴더/
├── templates/     ← PPTX 마스터 템플릿 (없으면 기본 사용)
├── rfp/           ← RFP/제안요청서
├── rawdata/       ← 통계, 보고서, 데이터
├── references/    ← 기존 제안서, 참고 자료
└── output/        ← 생성물
```
