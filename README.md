# pptx-vertical-writer-mcp

회사 표준 A3 세로형 PPTX 템플릿 기반으로 마크다운을 PowerPoint 프레젠테이션으로 변환하는 MCP 서버 + Claude Code 스킬.

## 특징
- 참조 PPTX의 레이아웃/디자인을 100% 보존하면서 텍스트만 교체
- 10개 템플릿 타입(T0~T9) 자동 분류
- 플레이스홀더(████/Lorem) 기반 깨끗한 템플릿
- PowerPoint COM API 사용 (Windows + PowerPoint 필수)

## 설치

```bash
pip install -r requirements.txt
```

## MCP 도구
| 도구 | 설명 |
|---|---|
| `list_templates()` | 사용 가능한 템플릿 목록 반환 |
| `analyze_template(pptx_path)` | 참조 PPTX 분석 → 템플릿 등록 |
| `create_pptx(extended_md)` | 확장 MD → PPTX 생성 |

## 스킬
| 스킬 | 설명 |
|---|---|
| `write-pptx` | 마크다운 → 템플릿 추천 → PPTX 생성 |
| `analyze-template` | 참조 PPTX 분석 및 템플릿 등록 |

## 요구사항
- Windows + Microsoft PowerPoint 설치
- Python 3.10+
