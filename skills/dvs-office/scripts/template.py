"""Open or save Office templates (.dotx .potx .xltx) with libraries that only accept documents.

  python template.py open 양식.dotx --out 작업.docx     # template -> normal document (styles, layouts kept)
  python template.py save 작업.docx --out 새양식.dotx   # normal document -> template

Only the declared content type of the main part changes; nothing else in the file is touched.
"""
import argparse
import sys
import zipfile
from pathlib import Path

PAIRS = {  # template main part  <->  document main part
    'wordprocessingml.template.main+xml': 'wordprocessingml.document.main+xml',
    'presentationml.template.main+xml': 'presentationml.presentation.main+xml',
    'spreadsheetml.template.main+xml': 'spreadsheetml.sheet.main+xml',
}
EXT = {'.dotx': '.docx', '.potx': '.pptx', '.xltx': '.xlsx'}


def swap(src, dst, to_template):
    if Path(dst).resolve() == Path(src).resolve():
        sys.exit('원본과 같은 이름으로 저장하지 않습니다.')
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        changed = False
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == '[Content_Types].xml':
                text = data.decode('utf-8')
                for tmpl, doc in PAIRS.items():
                    a, b = (doc, tmpl) if to_template else (tmpl, doc)
                    if a in text:
                        text, changed = text.replace(a, b), True
                data = text.encode('utf-8')
            zout.writestr(item, data)
    if not changed:
        Path(dst).unlink()
        sys.exit('바꿀 형식이 없습니다 — 이미 ' + ('양식' if to_template else '일반 문서') + '입니다.')
    print(dst)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('action', choices=['open', 'save'])
    ap.add_argument('file')
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    if a.action == 'open' and Path(a.out).suffix.lower() not in EXT.values():
        sys.exit('--out은 .docx / .pptx / .xlsx 이름으로 주세요.')
    if a.action == 'save' and Path(a.out).suffix.lower() not in EXT:
        sys.exit('--out은 .dotx / .potx / .xltx 이름으로 주세요.')
    swap(a.file, a.out, a.action == 'save')


if __name__ == '__main__':
    main()
