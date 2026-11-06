"""Shared helpers: locate LibreOffice / Tesseract / Korean fonts, and convert files with LibreOffice."""
import glob
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SOFFICE_GUESSES = [
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    '/usr/bin/soffice', '/usr/local/bin/soffice', '/snap/bin/libreoffice',
]
TESSERACT_GUESSES = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    '/opt/homebrew/bin/tesseract', '/usr/local/bin/tesseract', '/usr/bin/tesseract',
]
# (display name, file patterns) — first found wins
KOREAN_FONTS = [
    ('Malgun Gothic', ['C:/Windows/Fonts/malgun.ttf']),
    ('Apple SD Gothic Neo', ['/System/Library/Fonts/AppleSDGothicNeo.ttc']),
    ('NanumGothic', ['C:/Windows/Fonts/NanumGothic.ttf', '~/Library/Fonts/NanumGothic.ttf', '/usr/share/fonts/truetype/nanum/NanumGothic.ttf']),
    ('Noto Sans KR', ['/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc', '/usr/share/fonts/**/NotoSansKR*.ttf', '~/Library/Fonts/NotoSansKR*.ttf']),
]


def _first(paths):
    for p in paths:
        hits = glob.glob(os.path.expanduser(p), recursive=True)
        if hits:
            return hits[0]
    return None


def find_soffice():
    return shutil.which('soffice') or shutil.which('libreoffice') or _first(SOFFICE_GUESSES)


def find_tesseract():
    return shutil.which('tesseract') or _first(TESSERACT_GUESSES)


def find_fonts():
    """[{name, path}] of Korean fonts present, in preference order."""
    out = []
    for name, pats in KOREAN_FONTS:
        p = _first(pats)
        if p:
            out.append({'name': name, 'path': p})
    return out


def korean_font():
    """(name, path) of the preferred Korean font, or (None, None)."""
    f = find_fonts()
    return (f[0]['name'], f[0]['path']) if f else (None, None)


# Make Calc recalculate every formula when it opens .xlsx/.ods (default is "never" for Excel files,
# which would leave formulas written by openpyxl without values).
RECALC_XCU = '''<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
'''


def soffice_convert(src, fmt, outdir=None, timeout=180, recalc=True):
    """Convert with LibreOffice headless. fmt e.g. 'pdf', 'xlsx', 'docx'. Returns the output path.
    A private profile directory is used so a running LibreOffice window does not block the conversion,
    and (recalc=True) spreadsheet formulas are recalculated on load."""
    exe = find_soffice()
    if not exe:
        sys.exit('LibreOffice(soffice)가 없습니다 — python check_env.py --need render 로 설치 방법을 확인하세요.')
    src = Path(src).resolve()
    outdir = Path(outdir or src.parent).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as prof:
        if recalc:
            (Path(prof) / 'user').mkdir(parents=True, exist_ok=True)
            (Path(prof) / 'user' / 'registrymodifications.xcu').write_text(RECALC_XCU, encoding='utf-8')
        cmd = [exe, f'-env:UserInstallation={Path(prof).as_uri()}', '--headless', '--norestore',
               '--convert-to', fmt, '--outdir', str(outdir), str(src)]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    out = outdir / (src.stem + '.' + fmt.split(':')[0])
    if r.returncode != 0 or not out.exists():
        sys.exit(f'변환 실패({fmt}): {r.stderr.strip() or r.stdout.strip()}')
    return out
