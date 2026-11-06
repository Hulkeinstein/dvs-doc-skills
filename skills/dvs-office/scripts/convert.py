"""Convert an office file with LibreOffice (headless).

  python convert.py report.docx --to pdf
  python convert.py old.xls --to xlsx --out converted
  python convert.py slides.ppt --to pptx

Prints the output path. The source file is never changed.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from office_common import soffice_convert  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('file')
    ap.add_argument('--to', required=True, help='pdf, docx, xlsx, pptx, odt, csv ...')
    ap.add_argument('--out', default=None, help='output folder (default: next to the file)')
    a = ap.parse_args()
    src = Path(a.file).resolve()
    if src.suffix.lower().lstrip('.') == a.to and not a.out:
        sys.exit('같은 형식으로 같은 폴더에 바꾸면 원본을 덮어씁니다 — --out 폴더를 주세요.')
    print(soffice_convert(src, a.to, a.out))


if __name__ == '__main__':
    main()
