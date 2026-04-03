# pptx-vertical-writer-mcp

회사 표준 A3 세로형 PPTX 템플릿 기반 제안서 작성 도구.
RFP 분석 → 목차 설계 → 본문 작성 → PPTX 생성까지 오케스트레이션합니다.

## 핵심 기능

### 기능1: 제안서 오케스트레이션
```
프로젝트 폴더 (rfp/ + rawdata/ + references/)
  → RFP 분석 → docs/T?.md로 템플릿 매칭 → 스니펫 조립 → PPTX 생성
```

### 기능2: 템플릿 관리
```
기존 템플릿 쇼케이스 (10종 T0~T9)
  + 새 슬라이드 분석 → 기존 템플릿 매칭 또는 신규 등록
  + add_template() 시 T?.md 자동 업데이트
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
| `showcase_templates()` | 템플릿 쇼케이스 (필드 가이드) |
| `match_slide(pptx, num)` | 슬라이드↔템플릿 매칭 |
| `add_template(pptx, num, name)` | 새 템플릿 등록 + MD 자동 업데이트 |
| `analyze_template(pptx)` | 참조 PPTX 전체 분석 |
| `create_pptx(md, output)` | 확장 MD → PPTX 생성 |

## 프로젝트 폴더 구조

```
프로젝트폴더/
├── docs/          ← GUIDE.md + T0~T9.md (템플릿 구조 + 복사용 스니펫)
├── templates/     ← 블록처리된 참조 PPTX (placeholder_vol2/vol3.pptx)
├── rfp/           ← RFP/제안요청서
├── rawdata/       ← 통계, 보고서, 데이터
├── references/    ← 기존 제안서, 참고 자료
└── output/        ← 생성물
```

## 새 템플릿 등록 시 주의사항

`analyze_template()` 또는 `add_template()`로 새 PPTX를 분석/등록할 때, 원본 파일명과 텍스트가 `slide_index.json`에 기록됩니다.

**민감 정보 제거 절차:**
1. 분석 후 `slide_index.json`의 `source_pptx` 필드를 익명화 (예: `vol2`, `vol3`)
2. `template_sanitizer.py`로 PPTX 블록처리 (텍스트 → ████)
3. 블록처리 후에도 `slide_index.json`의 `text`, `table_preview` 필드에 원본이 남을 수 있으니 재확인
4. 커밋 전 `grep -r "민감단어" --include="*.json" --include="*.py"` 로 최종 검수

```bash
# 블록처리
python src/template_sanitizer.py <원본.pptx> <출력.pptx>

# slide_index.json 텍스트도 블록처리
python -c "
from src.template_sanitizer import sanitize_slide_index
sanitize_slide_index('templates/slide_index.json')
"

# 민감 정보 검수
grep -rn "회사명\|프로젝트명\|고객사" --include="*.json" --include="*.py"
```

## 스킬 (4개)

| 스킬 | 설명 |
|---|---|
| `create-proposal-pptx` | RFP→분석→MD→PPTX 전체 워크플로우 |
| `manage-templates` | 템플릿 쇼케이스/매칭/등록 |
| `write-pptx` | 확장 MD → PPTX 단독 변환 |
| `analyze-template` | 참조 PPTX 분석 |
