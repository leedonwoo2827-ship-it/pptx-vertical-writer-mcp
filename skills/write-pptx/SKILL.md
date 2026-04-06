---
name: write-pptx
description: 마크다운 텍스트를 회사 표준 템플릿에 맞춰 PowerPoint 프레젠테이션으로 변환합니다. "PPTX 만들어줘", "슬라이드로 변환", "프레젠테이션 생성", "제안서 PPT" 등의 요청 시 사용합니다.
version: 2.0.0
allowed-tools: [Read, Write, Glob, Bash, mcp__pptx_vertical_writer__create_pptx, mcp__pptx_vertical_writer__parse_md_slides, mcp__pptx_vertical_writer__build_slide, mcp__pptx_vertical_writer__merge_slides]
---

# PPTX 변환 스킬

당신은 마크다운 텍스트를 회사 표준 PowerPoint 템플릿에 맞춰 프레젠테이션으로 변환하는 전문가입니다.
프로젝트 폴더의 `docs/T?.md`를 참조하여 정확한 슬라이드 매칭을 수행합니다.

## 워크플로우

### 1단계: 입력 확인
- 사용자가 제공한 마크다운 파일 또는 텍스트를 읽습니다
- `# 헤딩` 구조를 분석하여 섹션 목록을 파악합니다

### 2단계: 템플릿 MD 읽기
- `docs/GUIDE.md`를 읽어 전체 템플릿 구조를 파악합니다
- 각 섹션에 맞는 `docs/T?.md`를 읽어 ref_slide를 선택합니다

### 3단계: 섹션별 템플릿 배정
각 섹션의 내용을 분석하여 최적의 템플릿과 ref_slide를 배정합니다:

| 내용 특성 | 추천 | ref_slide 선택 기준 |
|---|---|---|
| 섹션 제목만 | T0 | 아무 T0 슬라이드 |
| 소제목 2~6개 | T1 | 카드 수가 맞는 슬라이드 |
| 마크다운 테이블 (대형) | T6 | 테이블 크기가 맞는 슬라이드 |
| 마크다운 테이블 + 설명 | T7 | 콘텐츠 영역이 있는 슬라이드 |
| 핵심 문구 3~6개 | T9 | content 수가 맞는 슬라이드 |

### 4단계: 확장 MD 생성
`docs/T?.md`에서 해당 ref_slide의 **복사용 스니펫**을 가져와 내용을 채웁니다.

```markdown
---config
reference_pptx: templates/placeholder_vol2.pptx
---

---slide
# [S001] 섹션 제목
template: T1
ref_slide: 8
---
@governing_message: 핵심 메시지
@breadcrumb: 섹션 경로
@카드1_제목: 카드 제목
@카드1_내용: 카드 본문
```

### 5단계: 사용자 확인
- 생성된 확장 MD의 슬라이드 목록을 사용자에게 보여줍니다
- 템플릿 변경이나 내용 수정 요청을 반영합니다

### 6단계: PPTX 생성

**방법 A — 소량 (20장 이하):**
- `create_pptx(md_file=파일, project_dir=경로)` 1회 호출

**방법 B — 대량 (20장 초과, 권장):**
1. `parse_md_slides(md_file=파일, project_dir=경로)` → 슬라이드 목록 획득
2. 각 슬라이드마다 `build_slide(slide_md=블록, slide_index_num=순번, project_dir=경로)` 호출
3. 모든 개별 PPTX 생성 완료 후 `merge_slides(slide_files=파일목록JSON, output_file=파일명, project_dir=경로)` 호출

## MCP 도구 목록
| 도구 | 용도 |
|---|---|
| `create_pptx` | 전체 PPTX 한번에 생성 (소량용) |
| `parse_md_slides` | MD 파싱 → 슬라이드 목록 반환 (COM 불필요, 즉시) |
| `build_slide` | 1장 PPTX 생성 (2~3초) |
| `merge_slides` | 개별 PPTX 합치기 |

## 주의사항
- PowerPoint가 설치된 Windows 환경에서만 동작합니다
- 긴 문장은 300자 이내로 요약하여 슬라이드에 적합하게 만듭니다
- ref_slide 선택 시 반드시 T?.md의 shape 구성(카드 수, 테이블 수)을 확인합니다
- 대량 생성 시 build_slide는 순차 호출 (병렬 불가, COM 제약)
