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


class SofficeError(SystemExit):
    """LibreOffice failed or did not answer. Subclass of SystemExit: callers that do not catch it
    exit with the message (no traceback); recalc.py catches it to fall back to recalculate-on-open."""


# Reused across runs so only the first run pays for creating a LibreOffice profile.
PROFILE_DIR = Path(tempfile.gettempdir()) / 'dvs-office-lo-profile'


def _prepare_profile(prof, recalc):
    if recalc:
        (Path(prof) / 'user').mkdir(parents=True, exist_ok=True)
        xcu = Path(prof) / 'user' / 'registrymodifications.xcu'
        if not xcu.exists():
            xcu.write_text(RECALC_XCU, encoding='utf-8')


def _kill_tree(proc):
    # soffice.exe starts soffice.bin as a child; killing only the parent leaves the child hanging.
    if os.name == 'nt':
        subprocess.run(['taskkill', '/F', '/T', '/PID', str(proc.pid)], capture_output=True)
    else:
        import signal
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    proc.wait()


def _run_soffice(exe, prof, fmt, outdir, src, timeout):
    cmd = [exe, f'-env:UserInstallation={Path(prof).as_uri()}', '--headless', '--norestore',
           '--convert-to', fmt, '--outdir', str(outdir), str(src)]
    # Without these LibreOffice asks the default printer for page metrics on start and on PDF export;
    # an unreachable (e.g. Wi-Fi) printer stalls each run ~50 s. Conversions do not need a printer.
    env = dict(os.environ, SAL_DISABLE_PRINTERLIST='1', SAL_DISABLE_DEFAULTPRINTER='1')
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, env=env,
                            start_new_session=(os.name != 'nt'))
    try:
        out, err = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill_tree(proc)
        return None
    return proc.returncode, (err or out or '').strip()


def soffice_convert(src, fmt, outdir=None, timeout=90, recalc=True):
    """Convert with LibreOffice headless. fmt e.g. 'pdf', 'xlsx', 'docx'. Returns the output path.
    A private profile directory is used so a running LibreOffice window does not block the conversion,
    and (recalc=True) spreadsheet formulas are recalculated on load.
    If LibreOffice does not answer within `timeout` seconds it is killed and tried once more with a
    fresh profile; a second failure raises SofficeError."""
    exe = find_soffice()
    if not exe:
        raise SofficeError('LibreOffice(soffice)가 없습니다 — python check_env.py --need render 로 설치 방법을 확인하세요.')
    src = Path(src).resolve()
    outdir = Path(outdir or src.parent).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / (src.stem + '.' + fmt.split(':')[0])
    print(f'LibreOffice로 처리 중입니다(보통 몇 초 · 처음 한 번은 더 걸림, 최대 {timeout}초) — 끝날 때까지 기다려 주세요.', file=sys.stderr, flush=True)
    _prepare_profile(PROFILE_DIR, recalc)
    r = _run_soffice(exe, PROFILE_DIR, fmt, outdir, src, timeout)
    if r is None:  # no answer: retry once with a brand-new profile
        print(f'LibreOffice가 {timeout}초 안에 끝나지 않아 끄고 한 번 더 시도합니다.', file=sys.stderr, flush=True)
        shutil.rmtree(PROFILE_DIR, ignore_errors=True)
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as prof:
            _prepare_profile(prof, recalc)
            r = _run_soffice(exe, prof, fmt, outdir, src, timeout)
    if r is None:
        raise SofficeError(f'LibreOffice가 두 번 모두 {timeout}초 안에 응답하지 않았습니다(변환 {fmt} 실패).')
    code, msg = r
    if code != 0 or not out.exists():
        raise SofficeError(f'변환 실패({fmt}): {msg}')
    return out
