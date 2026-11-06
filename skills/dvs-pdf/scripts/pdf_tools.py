"""PDF helpers built on pypdf, pdfplumber, reportlab, pypdfium2 and (for OCR) pytesseract.

  python pdf_tools.py info f.pdf
  python pdf_tools.py text f.pdf [--pages 1-3]
  python pdf_tools.py tables f.pdf [--pages 2] [--xlsx out.xlsx]
  python pdf_tools.py images f.pdf --outdir imgs
  python pdf_tools.py ocr f.pdf [--lang kor+eng] [--searchable out.pdf] [--pages 1-2]
  python pdf_tools.py merge a.pdf b.pdf --out all.pdf
  python pdf_tools.py pages f.pdf --pages 3,1,5-7 --out new.pdf
  python pdf_tools.py split f.pdf --outdir parts
  python pdf_tools.py rotate f.pdf [--pages 2] --degrees 90 --out new.pdf
  python pdf_tools.py watermark f.pdf (--text 대외비 | --stamp stamp.pdf) --out new.pdf
  python pdf_tools.py number f.pdf [--format "{n} / {total}"] --out new.pdf
  python pdf_tools.py encrypt f.pdf --password P --out new.pdf
  python pdf_tools.py decrypt f.pdf --password P --out new.pdf
  python pdf_tools.py fields form.pdf
  python pdf_tools.py fill form.pdf --data values.json --out filled.pdf
  python pdf_tools.py overlay form.pdf --data positions.json --out filled.pdf
  python pdf_tools.py create notes.md --out new.pdf [--title 제목]

Page numbers start at 1. Output never overwrites the input.
"""
import argparse
import io
import os
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'dvs-office' / 'scripts'))


def parse_pages(spec, total):
    if not spec:
        return list(range(1, total + 1))
    nums = []
    for part in spec.split(','):
        part = part.strip()
        if '-' in part:
            lo, hi = part.split('-')
            nums.extend(range(int(lo), int(hi) + 1))
        elif part:
            nums.append(int(part))
    bad = [n for n in nums if not 1 <= n <= total]
    if bad:
        sys.exit(f'없는 쪽 번호: {bad} (전체 {total}쪽)')
    return nums


def guard_out(src, out):
    if Path(out).resolve() == Path(src).resolve():
        sys.exit('원본과 같은 이름으로 저장하지 않습니다 — --out에 다른 이름을 주세요.')
    return out


def reader(path, password=None):
    from pypdf import PdfReader
    r = PdfReader(path)
    if r.is_encrypted:
        # Many PDFs only carry a permissions password: they open with an empty password.
        if not r.decrypt(password or ''):
            sys.exit('여는 암호가 걸린 PDF입니다 — 사용자에게 암호를 받아 decrypt 명령에 주세요.')
    return r


# ---------- korean font for reportlab ----------
_FONT = None


def register_korean_font():
    """Register a Hangul-capable TTF with reportlab and return its name (cached)."""
    global _FONT
    if _FONT:
        return _FONT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    try:
        from office_common import korean_font
        path = korean_font()[1]
    except Exception:
        path = None
    candidates = [path] if path else []
    # reportlab reads TrueType outlines only — CFF fonts such as Noto Sans CJK .ttc/.otf fail and are skipped.
    candidates += ['C:/Windows/Fonts/malgun.ttf', '/System/Library/Fonts/Supplemental/AppleGothic.ttf',
                   '/Library/Fonts/NanumGothic.ttf', str(Path.home() / 'Library/Fonts/NanumGothic.ttf'),
                   '/usr/share/fonts/truetype/nanum/NanumGothic.ttf', '/System/Library/Fonts/AppleSDGothicNeo.ttc']
    for c in candidates:
        if c and Path(c).exists():
            try:
                kw = {'subfontIndex': 0} if str(c).lower().endswith('.ttc') else {}
                pdfmetrics.registerFont(TTFont('KR', str(c), **kw))
                bold = Path(str(c).replace('malgun.ttf', 'malgunbd.ttf'))
                if bold.exists() and bold != Path(c):
                    pdfmetrics.registerFont(TTFont('KR-Bold', str(bold)))
                else:
                    pdfmetrics.registerFont(TTFont('KR-Bold', str(c), **kw))
                from reportlab.lib.fonts import addMapping
                addMapping('KR', 0, 0, 'KR'); addMapping('KR', 1, 0, 'KR-Bold')
                addMapping('KR', 0, 1, 'KR'); addMapping('KR', 1, 1, 'KR-Bold')
                _FONT = 'KR'
                return _FONT
            except Exception:
                continue
    sys.exit('PDF에 쓸 한글 글꼴(TrueType)을 찾지 못했습니다 — 나눔고딕을 설치하세요 '
             '(Mac: brew install --cask font-nanum-gothic · Ubuntu: sudo apt install fonts-nanum).')


# ---------- reading ----------
def cmd_info(a):
    from pypdf import PdfReader
    r = PdfReader(a.file)
    enc = r.is_encrypted
    if enc and not r.decrypt(''):
        print(json.dumps({'file': a.file, 'encrypted': True, 'needs_password': True}, ensure_ascii=False))
        return
    p0 = r.pages[0]
    sample = ''.join((r.pages[i].extract_text() or '') for i in range(min(3, len(r.pages))))
    fields = r.get_fields() or {}
    print(json.dumps({'file': a.file, 'pages': len(r.pages), 'encrypted': bool(enc), 'needs_password': False,
                      'page_size_pt': [round(float(p0.mediabox.width)), round(float(p0.mediabox.height))],
                      'rotation_first_page': p0.rotation,
                      'has_text': len(sample.strip()) > 20, 'form_fields': len(fields),
                      'metadata': {k: str(v) for k, v in (r.metadata or {}).items()}}, ensure_ascii=False, indent=1))


def cmd_text(a):
    import pdfplumber
    with pdfplumber.open(a.file) as pdf:
        for n in parse_pages(a.pages, len(pdf.pages)):
            t = pdf.pages[n - 1].extract_text(layout=True) or ''
            print(f'===== {n}쪽 =====')
            print('\n'.join(line.rstrip() for line in t.splitlines()).strip())


def cmd_tables(a):
    import pdfplumber
    found = []
    with pdfplumber.open(a.file) as pdf:
        for n in parse_pages(a.pages, len(pdf.pages)):
            for k, tb in enumerate(pdf.pages[n - 1].extract_tables(), 1):
                rows = [[(c or '').strip() for c in row] for row in tb if any(c for c in row)]
                if rows:
                    found.append({'page': n, 'table': k, 'rows': rows})
    if a.xlsx:
        import openpyxl
        from openpyxl.styles import Font
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for t in found:
            ws = wb.create_sheet(f"{t['page']}쪽-표{t['table']}"[:31])
            for row in t['rows']:
                ws.append([_num(v) for v in row])
            for c in ws[1]:
                c.font = Font(bold=True)
        if not found:
            wb.create_sheet('표 없음')
        wb.save(guard_out(a.file, a.xlsx))
        print(json.dumps({'tables': len(found), 'xlsx': a.xlsx}, ensure_ascii=False))
    else:
        print(json.dumps(found, ensure_ascii=False, indent=1))


NUMBER = re.compile(r'-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?')


def _num(v):
    """Turn '1,234' / '12.5' into numbers; keep codes ('00123'), long ids (>15 digits) and odd text as text."""
    s = v.strip()
    if not NUMBER.fullmatch(s):
        return v
    digits = s.replace(',', '').lstrip('-')
    if (len(digits) > 1 and digits[0] == '0' and not digits.startswith('0.')) or len(digits.replace('.', '')) > 15:
        return v
    s = s.replace(',', '')
    return float(s) if '.' in s else int(s)


def cmd_images(a):
    r = reader(a.file)
    out = Path(a.outdir); out.mkdir(parents=True, exist_ok=True)
    saved = []
    for n, page in enumerate(r.pages, 1):
        for k, img in enumerate(page.images, 1):
            dst = out / f'p{n}-{k}{Path(img.name).suffix or ".png"}'
            dst.write_bytes(img.data)
            saved.append(str(dst))
    print(json.dumps({'images': len(saved), 'files': saved[:50]}, ensure_ascii=False))


def cmd_ocr(a):
    import pypdfium2 as pdfium
    import pytesseract
    from office_common import find_tesseract
    exe = find_tesseract()
    if exe:
        pytesseract.pytesseract.tesseract_cmd = exe
    # Language data can live in a user folder (no admin rights needed): ~/.dvs-office/tessdata
    user_dir = Path.home() / '.dvs-office' / 'tessdata'
    if user_dir.is_dir():
        os.environ['TESSDATA_PREFIX'] = str(user_dir)  # env var avoids Windows quoting problems with --tessdata-dir
    # Images rendered from PDF carry no DPI, so Tesseract misjudges the scale and reports "Empty page".
    # psm 6 (one block of text) reads ordinary documents and tables far better than the default here.
    cfg = f'--dpi 300 --psm {a.psm}'
    langs = pytesseract.get_languages(config=cfg)
    missing = [l for l in a.lang.split('+') if l not in langs]
    if missing:
        sys.exit(f'Tesseract 언어 데이터가 없습니다: {missing} — 있는 것: {langs}. '
                 f'SKILL.md 「스캔본 한국어 준비」대로 {user_dir}를 만드세요.')
    doc = pdfium.PdfDocument(a.file)
    pages = parse_pages(a.pages, len(doc))
    pdf_parts = []
    for n in pages:
        img = doc[n - 1].render(scale=300 / 72).to_pil()
        print(f'===== {n}쪽 =====')
        print(pytesseract.image_to_string(img, lang=a.lang, config=cfg).strip())
        if a.searchable:
            pdf_parts.append(pytesseract.image_to_pdf_or_hocr(img, lang=a.lang, extension='pdf', config=cfg))
    if a.searchable:
        from pypdf import PdfReader, PdfWriter
        w = PdfWriter()
        for blob in pdf_parts:
            for p in PdfReader(io.BytesIO(blob)).pages:
                w.add_page(p)
        with open(guard_out(a.file, a.searchable), 'wb') as f:
            w.write(f)
        print(f'검색되는 PDF: {a.searchable}', file=sys.stderr)


# ---------- page operations ----------
def write(w, src, out):
    with open(guard_out(src, out), 'wb') as f:
        w.write(f)
    print(out)


def cmd_merge(a):
    from pypdf import PdfWriter
    w = PdfWriter()
    for f in a.files:
        w.append(reader(f))
    write(w, a.files[0], a.out)


def cmd_pages(a):
    from pypdf import PdfWriter
    r = reader(a.file)
    w = PdfWriter()
    for n in parse_pages(a.pages, len(r.pages)):
        w.add_page(r.pages[n - 1])
    write(w, a.file, a.out)


def cmd_split(a):
    from pypdf import PdfWriter
    r = reader(a.file)
    out = Path(a.outdir); out.mkdir(parents=True, exist_ok=True)
    stem = Path(a.file).stem
    for n, p in enumerate(r.pages, 1):
        w = PdfWriter(); w.add_page(p)
        with open(out / f'{stem}-{n:03d}.pdf', 'wb') as f:
            w.write(f)
    print(json.dumps({'files': len(r.pages), 'outdir': str(out)}, ensure_ascii=False))


def cmd_rotate(a):
    from pypdf import PdfWriter
    if a.degrees % 90:
        sys.exit('각도는 90의 배수만 됩니다.')
    r = reader(a.file)
    targets = set(parse_pages(a.pages, len(r.pages)))
    w = PdfWriter()
    for n, p in enumerate(r.pages, 1):
        w.add_page(p)
        if n in targets:
            w.pages[-1].rotate(a.degrees)
    write(w, a.file, a.out)


def stamp_on(page, draw=None, stamp=None):
    """Put a reportlab drawing (draw(c, W, H) in visible-page coordinates) or a stamp page on top of `page`.
    Rotation is baked into the content first and the visible box (cropbox) origin is honoured,
    so marks land in the same place the reader sees on rotated or cropped pages."""
    from pypdf import PdfReader, Transformation
    from reportlab.pdfgen import canvas
    if page.rotation:
        page.transfer_rotation_to_content()
    box = page.cropbox
    x0, y0, W, H = float(box.left), float(box.bottom), float(box.width), float(box.height)
    if stamp is not None:
        page.merge_transformed_page(stamp, Transformation().translate(x0, y0))
        return
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(max(x0, 0) + W, max(y0, 0) + H))
    c.translate(x0, y0)
    draw(c, W, H)
    c.save()
    buf.seek(0)
    page.merge_page(PdfReader(buf).pages[0])


def cmd_watermark(a):
    from pypdf import PdfReader, PdfWriter
    r = reader(a.file)
    w = PdfWriter()
    stamp = PdfReader(a.stamp).pages[0] if a.stamp else None
    font = register_korean_font() if a.text else None
    def draw(c, W, H):
        c.setFont(font, min(W, H) / 8)
        c.setFillColorRGB(0.5, 0.5, 0.5, alpha=0.25)
        c.translate(W / 2, H / 2); c.rotate(45)
        c.drawCentredString(0, 0, a.text)
    for p in r.pages:
        stamp_on(p, draw=None if stamp is not None else draw, stamp=stamp)
        w.add_page(p)
    write(w, a.file, a.out)


def cmd_number(a):
    from pypdf import PdfWriter
    r = reader(a.file)
    font = register_korean_font()
    total = len(r.pages)
    w = PdfWriter()
    for n, p in enumerate(r.pages, 1):
        label = a.format.format(n=n, total=total)
        def draw(c, W, H, label=label):
            c.setFont(font, 9); c.setFillColorRGB(0.3, 0.3, 0.3)
            c.drawCentredString(W / 2, 22, label)
        stamp_on(p, draw=draw)
        w.add_page(p)
    write(w, a.file, a.out)


def cmd_encrypt(a):
    from pypdf import PdfWriter
    r = reader(a.file)
    w = PdfWriter(clone_from=r)
    w.encrypt(user_password=a.password, owner_password=a.password, algorithm='AES-256')
    write(w, a.file, a.out)


def cmd_decrypt(a):
    from pypdf import PdfWriter
    r = reader(a.file, a.password)
    w = PdfWriter(clone_from=r)
    write(w, a.file, a.out)


# ---------- forms ----------
def cmd_fields(a):
    r = reader(a.file)
    fields = r.get_fields() or {}
    out = []
    for name, f in fields.items():
        kind = {'/Tx': '글', '/Btn': '체크/선택', '/Ch': '목록', '/Sig': '서명'}.get(f.get('/FT'), str(f.get('/FT')))
        d = {'name': name, 'type': kind, 'value': str(f.get('/V', ''))}
        if f.get('/_States_'):
            d['choices'] = [str(s) for s in f['/_States_']]
        if f.get('/Opt'):
            d['choices'] = [str(o[1] if isinstance(o, list) else o) for o in f['/Opt']]
        out.append(d)
    pages = {}
    for n, p in enumerate(r.pages, 1):
        for an in p.get('/Annots') or []:
            an = an.get_object()
            if an.get('/T'):
                pages[str(an['/T'])] = {'page': n, 'rect': [round(float(x)) for x in an['/Rect']]}
    for d in out:
        d.update(pages.get(d['name'].split('.')[-1], {}))
    print(json.dumps(out if out else {'fields': 0, 'hint': '칸이 없는 양식입니다 — overlay 명령을 쓰세요.'}, ensure_ascii=False, indent=1))


def cmd_fill(a):
    from pypdf import PdfWriter
    r = reader(a.file)
    values = json.loads(Path(a.data).read_text(encoding='utf-8'))
    known = set((r.get_fields() or {}).keys())
    unknown = [k for k in values if k not in known]
    if unknown:
        sys.exit(f'양식에 없는 칸 이름: {unknown} — fields 명령으로 이름을 확인하세요.')
    w = PdfWriter(clone_from=r)
    from pypdf.generic import NameObject
    acro = w._root_object.get('/AcroForm')
    if acro is not None and '/XFA' in acro.get_object():
        # XFA forms (common in government PDFs) make Acrobat ignore the normal field values.
        del acro.get_object()[NameObject('/XFA')]
        print('⚠ XFA 양식이라 XFA 부분을 지우고 일반 칸에 채웠습니다 — 결과를 그림으로 꼭 확인하세요.', file=sys.stderr)
    for page in w.pages:
        w.update_page_form_field_values(page, values, auto_regenerate=False)
    # Viewers redraw the fields from the values (needed for Hangul — the form's own font usually has no Korean).
    w.set_need_appearances_writer(True)
    write(w, a.file, a.out)


def cmd_overlay(a):
    from pypdf import PdfWriter
    r = reader(a.file)
    items = json.loads(Path(a.data).read_text(encoding='utf-8'))
    font = register_korean_font()
    w = PdfWriter()
    for n, p in enumerate(r.pages, 1):
        mine = [it for it in items if int(it['page']) == n]
        if mine:
            def draw(c, W, H, mine=mine):
                for it in mine:
                    c.setFont(font, float(it.get('size', 11)))
                    c.drawString(float(it['x']), float(it['y']), str(it['text']))
            stamp_on(p, draw=draw)
        w.add_page(p)
    write(w, a.file, a.out)


# ---------- create ----------
def md_escape(s):
    from xml.sax.saxutils import escape
    return escape(s).replace(' ', '&nbsp;')


def md_inline(s):
    from xml.sax.saxutils import escape
    s = escape(s)
    s = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', s)
    s = re.sub(r'`(.+?)`', r'<font backColor="#EEEEEE">\1</font>', s)
    return s


def cmd_create(a):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    font = register_korean_font()
    base = ParagraphStyle('base', fontName=font, fontSize=10.5, leading=16, wordWrap='CJK')
    h = {1: ParagraphStyle('h1', parent=base, fontName='KR-Bold', fontSize=20, leading=28, spaceBefore=6, spaceAfter=10),
         2: ParagraphStyle('h2', parent=base, fontName='KR-Bold', fontSize=15, leading=22, spaceBefore=10, spaceAfter=6),
         3: ParagraphStyle('h3', parent=base, fontName='KR-Bold', fontSize=12.5, leading=18, spaceBefore=8, spaceAfter=4)}
    lines = Path(a.file).read_text(encoding='utf-8').splitlines()
    story, para, bullets, table, code = [], [], [], [], []
    ordered = False
    mono = ParagraphStyle('code', parent=base, fontSize=9, leading=13, backColor=colors.HexColor('#F2F2F2'),
                          borderPadding=4, leftIndent=4)

    def flush():
        nonlocal para, bullets, table, ordered
        if para:
            story.append(Paragraph(md_inline(' '.join(para)), base)); story.append(Spacer(1, 4)); para = []
        if bullets:
            items = [ListItem(Paragraph(md_inline(b), base), leftIndent=14) for b in bullets]
            if ordered:
                story.append(ListFlowable(items, bulletType='1', bulletFormat='%s.', bulletFontName=font, leftIndent=14))
            else:
                story.append(ListFlowable(items, bulletType='bullet', start='•', bulletFontName=font, leftIndent=14))
            bullets, ordered = [], False
        if table:
            rows = [r for r in table if not re.fullmatch(r'[\s|:-]+', r)]
            data = [[Paragraph(md_inline(c.strip()), base) for c in r.strip().strip('|').split('|')] for r in rows]
            t = Table(data, repeatRows=1, hAlign='LEFT')
            t.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#999999')),
                                   ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E8EEF7')),
                                   ('VALIGN', (0, 0), (-1, -1), 'TOP')]))
            story.append(t); story.append(Spacer(1, 6)); table = []

    if a.title:
        story.append(Paragraph(md_inline(a.title), h[1]))
    in_code = False
    for line in lines:
        if line.strip().startswith('```'):
            if in_code:
                story.append(Paragraph('<br/>'.join(md_escape(c) for c in code) or ' ', mono)); story.append(Spacer(1, 6))
                code, in_code = [], False
            else:
                flush(); in_code = True
            continue
        if in_code:
            code.append(line); continue
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            flush(); story.append(Paragraph(md_inline(m.group(2)), h[min(len(m.group(1)), 3)])); continue
        mo = re.match(r'^\s*\d+[.)]\s+(.*)', line)
        if mo:
            if para or table or (bullets and not ordered):
                flush()
            ordered = True
            bullets.append(mo.group(1)); continue
        if line.strip().startswith('|'):
            if para or bullets:
                flush()
            table.append(line); continue
        if re.match(r'^\s*[-*]\s+', line):
            if para or table or ordered:
                flush()
            bullets.append(re.sub(r'^\s*[-*]\s+', '', line)); continue
        if not line.strip():
            flush(); continue
        if bullets or table:
            flush()
        para.append(line.strip())
    flush()
    out = guard_out(a.file, a.out)
    SimpleDocTemplate(out, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=20 * mm, bottomMargin=20 * mm,
                      title=a.title or Path(a.file).stem).build(story)
    print(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    def add(name, *extra, file=True):
        p = sub.add_parser(name)
        if file:
            p.add_argument('file')
        for args, kw in extra:
            p.add_argument(*args, **kw)
    OUT = (('--out',), {'required': True})
    PAGES = (('--pages',), {})
    add('info'); add('text', PAGES); add('tables', PAGES, (('--xlsx',), {}))
    add('images', (('--outdir',), {'required': True}))
    add('ocr', PAGES, (('--lang',), {'default': 'kor+eng'}), (('--searchable',), {}), (('--psm',), {'default': '6'}))
    add('merge', (('files',), {'nargs': '+'}), OUT, file=False)
    add('pages', (('--pages',), {'required': True}), OUT)
    add('split', (('--outdir',), {'required': True}))
    add('rotate', PAGES, (('--degrees',), {'type': int, 'default': 90}), OUT)
    add('watermark', (('--text',), {}), (('--stamp',), {}), OUT)
    add('number', (('--format',), {'default': '{n} / {total}'}), OUT)
    add('encrypt', (('--password',), {'required': True}), OUT)
    add('decrypt', (('--password',), {'required': True}), OUT)
    add('fields')
    add('fill', (('--data',), {'required': True}), OUT)
    add('overlay', (('--data',), {'required': True}), OUT)
    add('create', OUT, (('--title',), {}))
    a = ap.parse_args()
    if a.cmd == 'watermark' and not (a.text or a.stamp):
        sys.exit('--text 또는 --stamp 중 하나를 주세요.')
    if a.cmd == 'merge' and len(a.files) < 2:
        sys.exit('합칠 파일을 두 개 이상 주세요.')
    globals()['cmd_' + a.cmd](a)


if __name__ == '__main__':
    main()
