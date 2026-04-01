"""
pptx-template-filler: 참조 PPTX 슬라이드를 템플릿으로 사용하여
확장 마크다운에서 PPTX를 생성하는 CLI 도구.

사용법:
    python src/main.py input.md -o output.pptx
    python src/main.py input.md -o output.pptx --index templates/slide_index.json
"""
import argparse
import json
import os
import sys

from md_parser import parse_md
from slide_builder import build_presentation, load_slide_index


def main():
    parser = argparse.ArgumentParser(
        description='참조 PPTX 템플릿 기반 MD→PPTX 변환 도구'
    )
    parser.add_argument('input_md', help='입력 마크다운 파일 경로')
    parser.add_argument('-o', '--output', default='output/result.pptx',
                        help='출력 PPTX 파일 경로 (기본: output/result.pptx)')
    parser.add_argument('--index', default='templates/slide_index.json',
                        help='slide_index.json 경로 (기본: templates/slide_index.json)')

    args = parser.parse_args()

    # MD 파일 읽기
    if not os.path.exists(args.input_md):
        print(f'Error: 입력 파일을 찾을 수 없습니다: {args.input_md}')
        sys.exit(1)

    with open(args.input_md, 'r', encoding='utf-8') as f:
        md_text = f.read()

    print(f'입력 MD: {args.input_md}')

    # MD 파싱
    md_data = parse_md(md_text)
    print(f'설정: {md_data["config"]}')
    print(f'슬라이드 수: {len(md_data["slides"])}')

    # slide_index 로드
    if not os.path.exists(args.index):
        print(f'Error: slide_index.json을 찾을 수 없습니다: {args.index}')
        sys.exit(1)

    slide_index = load_slide_index(args.index)
    print(f'참조 인덱스: {slide_index["total_slides"]}슬라이드')

    # 출력 디렉토리 확인
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    # PPTX 생성
    print(f'\nPPTX 생성 시작...')
    build_presentation(md_data, slide_index, args.output)

    print(f'\n완료!')


if __name__ == '__main__':
    main()
