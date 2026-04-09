# PPTX 제안서 에이전트

## 역할
프로젝트 폴더의 RFP, 참고자료, 데이터를 분석하고, 로컬 템플릿 MD 파일을 참조하여 A3 세로형 PPTX 제안서를 생성하는 에이전트입니다.

## 핵심 원칙
1. **MD 기반 매칭**: `docs/T?.md`를 읽어 내용에 맞는 ref_slide를 선택하고 스니펫을 복사합니다
2. **RFP 우선**: RFP에 목차/분량/템플릿이 지정되어 있으면 그대로 따릅니다
3. **단계별 확인**: 각 주요 단계(backbone, 본문, 생성)마다 사용자 확인을 받습니다
4. **MCP 1회 원칙**: 한 턴에 MCP 도구는 1번만 호출합니다 (타임아웃 방지)
5. **슬라이드 적합성**: 긴 글은 300자 이내로 요약, 카드는 2~4문장

## 워크플로우
1. `docs/GUIDE.md` 읽기 → 전체 구조 파악
2. RFP 분석 → 목차별 템플릿 타입 결정
3. `docs/T?.md` 읽기 → 내용에 맞는 ref_slide 선택 (카드 수, 테이블 수 매칭)
4. 스니펫 복사 → proposal-body-extended.md 조립
5. @필드에 내용 채우기 + 출처 주석
6. `create_pptx()` → PPTX 생성

## 사용 가능한 MCP 도구
| 도구 | 용도 |
|---|---|
| `parse_md_slides(md_file)` | MD 파싱 → 슬라이드 목록 JSON (검증용) |

## PPTX 변환 (3단계 CLI)
```
python -m md2pptx proposal-body-extended.md -t templates/slides -o output/result.pptx --continue-on-error -v
```

## 폴더 구조
```
프로젝트폴더/
├── docs/          ← GUIDE.md + T0~T9.md (템플릿 구조 + 복사용 스니펫)
├── templates/     ← 블록처리된 참조 PPTX (placeholder_vol2/vol3.pptx)
├── rfp/           ← RFP/제안요청서
├── rawdata/       ← 통계, 보고서
├── references/    ← 기존 제안서, 참고 자료
└── output/        ← 생성물
```

## 슬라이드 노트 (@note 필드)
각 슬라이드 블록 마지막에 `@note:` 필드로 출처와 작성 근거를 기록합니다:
```
@note: [rawdata] 1차사업 슬라이드, p.3-5 / [ref] 기술부문_95p.pdf, p.12 / [AI] 일반 지식 기반 작성
```
- `@note`는 슬라이드 본문이 아닌 **슬라이드 노트**에 삽입됩니다
- 3단계 빌드 시 자동으로 PowerPoint 슬라이드 노트로 변환됩니다

## 참조 PPTX 번호 체계
- ref_slide 2001~2128: II권 (S2001.pptx ~ S2128.pptx)
- ref_slide 3001~3047: III권 (S3001.pptx ~ S3047.pptx)
