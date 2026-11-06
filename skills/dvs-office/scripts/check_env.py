"""Report which packages and programs the DVS office skills need, and how to install the missing ones.

  python check_env.py                 # everything
  python check_env.py --need xlsx pdf # only what these tasks need
  python check_env.py --fonts         # Korean fonts found on this computer

Tasks: xlsx docx pptx pdf render recalc ocr. Exit code 0 = all needed items present, 1 = something missing.
Nothing is installed here: the agent shows the printed commands to the user and asks first.
"""
import argparse
import importlib.util
import json
import sys

sys.path.insert(0, __import__('os').path.dirname(__file__))
from office_common import find_fonts, find_soffice, find_tesseract  # noqa: E402

# import name -> pip name
PKGS = {
    'openpyxl': 'openpyxl', 'pandas': 'pandas', 'docx': 'python-docx', 'pptx': 'python-pptx',
    'pypdf': 'pypdf', 'pdfplumber': 'pdfplumber', 'reportlab': 'reportlab', 'pypdfium2': 'pypdfium2',
    'pytesseract': 'pytesseract', 'PIL': 'pillow', 'lxml': 'lxml',
}
NEED = {
    'xlsx': ['openpyxl', 'pandas'],
    'docx': ['docx', 'lxml', 'PIL'],
    'pptx': ['pptx', 'lxml', 'PIL'],
    'pdf': ['pypdf', 'pdfplumber', 'reportlab'],
    'render': ['pypdfium2', 'PIL', 'soffice'],
    'recalc': ['openpyxl', 'soffice'],
    'ocr': ['pypdfium2', 'pytesseract', 'PIL', 'tesseract'],
}
PROGRAM_INSTALL = {
    'soffice': {'win': 'winget install -e --id TheDocumentFoundation.LibreOffice', 'mac': 'brew install --cask libreoffice', 'linux': 'sudo apt-get install -y libreoffice'},
    'tesseract': {'win': 'winget install -e --id UB-Mannheim.TesseractOCR  (한국어: 설치 화면에서 Korean 언어 데이터 선택)', 'mac': 'brew install tesseract tesseract-lang', 'linux': 'sudo apt-get install -y tesseract-ocr tesseract-ocr-kor'},
}


def platform_key():
    return 'win' if sys.platform.startswith('win') else 'mac' if sys.platform == 'darwin' else 'linux'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--need', nargs='*', choices=sorted(NEED), help='only check what these tasks need')
    ap.add_argument('--fonts', action='store_true', help='list Korean fonts')
    a = ap.parse_args()
    if a.fonts:
        print(json.dumps(find_fonts(), ensure_ascii=False, indent=2))
        return 0
    wanted = sorted({x for t in (a.need or NEED) for x in NEED[t]})
    have, missing_pip, missing_prog = {}, [], []
    for name in wanted:
        if name == 'soffice':
            path = find_soffice()
        elif name == 'tesseract':
            path = find_tesseract()
        else:
            path = 'ok' if importlib.util.find_spec(name) else None
        have[name] = path or 'MISSING'
        if not path:
            (missing_prog if name in PROGRAM_INSTALL else missing_pip).append(name)
    report = {'python': sys.version.split()[0], 'found': have, 'install': []}
    if missing_pip:
        report['install'].append(f'{sys.executable} -m pip install ' + ' '.join(PKGS[m] for m in missing_pip))
    for p in missing_prog:
        report['install'].append(PROGRAM_INSTALL[p][platform_key()])
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 1 if report['install'] else 0


if __name__ == '__main__':
    sys.exit(main())
