"""Word (.docx) helpers built on python-docx + lxml.

  python docx_tools.py text in.docx                          # headings/paragraphs/tables in order, tracked changes as [+new+] [-old-], comments
  python docx_tools.py toc in.docx [--after "보고서 제목"] [--out o.docx]
  python docx_tools.py page-numbers in.docx [--format "- {PAGE} / {NUMPAGES} -"] [--replace] [--out o.docx]
  python docx_tools.py header in.docx --text "팀 주간 보고" [--footer] [--replace] [--out o.docx]
  python docx_tools.py replace in.docx --old A --new B [--out o.docx]          # keeps formatting, line breaks and tabs
  python docx_tools.py track-replace in.docx --old A --new B [--author 이름] [--out o.docx]   # shows as a tracked change
  python docx_tools.py track-insert in.docx --after 글 --text 새글 [--author 이름] [--out o.docx]  # tracked insertion only
  python docx_tools.py comment in.docx --find 글 --text 메모 [--author 이름] [--out o.docx]
  python docx_tools.py accept in.docx [--out o.docx]          # accept all tracked changes
  python docx_tools.py reject in.docx [--out o.docx]          # reject all tracked changes

Without --out, editing commands write <name>-수정.docx next to the input (the input is never overwritten).
Templates (.dotx): open/save them through dvs-office/scripts/template.py first.
"""
import argparse
import copy
import datetime
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


# ---------- shared ----------
def out_path(a):
    out = a.out or str(Path(a.file).with_name(Path(a.file).stem + '-수정' + Path(a.file).suffix))
    if Path(out).resolve() == Path(a.file).resolve():
        sys.exit('원본과 같은 이름으로 저장하지 않습니다 — --out에 다른 이름을 주세요.')
    return out


def all_parts(doc):
    """Root XML elements to edit: body + every header/footer (each part once)."""
    roots, seen = [doc.element.body], set()
    for s in doc.sections:
        for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer, s.even_page_header, s.even_page_footer):
            if hf is not None and not hf.is_linked_to_previous and id(hf._element) not in seen:
                seen.add(id(hf._element))
                roots.append(hf._element)
    return roots


def paragraphs_xml(root):
    return list(root.iter(qn('w:p')))


def owning_p(el):
    el = el.getparent()
    while el is not None and el.tag != qn('w:p'):
        el = el.getparent()
    return el


def visible_runs(p, include_ins=True):
    """Runs whose text shows in paragraph p: direct runs and runs inside hyperlinks, smart tags and content controls.
    Runs in deleted text (and, if include_ins is False, in tracked insertions) are skipped, as are runs of nested
    paragraphs (text boxes), which are visited as their own paragraphs."""
    skip = {qn('w:del'), qn('w:moveFrom')} | (set() if include_ins else {qn('w:ins'), qn('w:moveTo')})
    out = []
    for r in p.iter(qn('w:r')):
        if owning_p(r) is not p:
            continue
        anc, ok = r.getparent(), True
        while anc is not p:
            if anc.tag in skip:
                ok = False
                break
            anc = anc.getparent()
        if ok:
            out.append(r)
    return out


def text_nodes(p, include_ins=True):
    """The w:t elements of the visible runs, in order. Editing these leaves tabs, breaks and drawings in place."""
    return [t for r in visible_runs(p, include_ins) for t in r if t.tag == qn('w:t')]


def new_t(text, tag='w:t'):
    t = OxmlElement(tag)
    t.text = text
    t.set(XML_SPACE, 'preserve')
    return t


def split_at(t, i, j):
    """Split the run holding text node t so that t.text[i:j] becomes its own run (same formatting). Returns that run.
    Other run content (tabs, breaks, drawings) is moved, not copied, so nothing is duplicated."""
    r = t.getparent()
    rPr = r.find(qn('w:rPr'))
    content = [c for c in r if c is not rPr]
    idx = content.index(t)
    text = t.text or ''

    def make():
        nr = r.makeelement(r.tag, dict(r.attrib))
        if rPr is not None:
            nr.append(copy.deepcopy(rPr))
        return nr
    before, mid, after = make(), make(), make()
    for c in content[:idx]:
        before.append(c)
    if i > 0:
        before.append(new_t(text[:i]))
    mid.append(new_t(text[i:j]))
    if j < len(text):
        after.append(new_t(text[j:]))
    for c in content[idx + 1:]:
        after.append(c)
    base = 1 if rPr is not None else 0
    for nr in (before, mid, after):
        if len(nr) > base:
            r.addprevious(nr)
    r.getparent().remove(r)
    return mid


def find_text(p, needle, include_ins=False):
    for t in text_nodes(p, include_ins):
        k = (t.text or '').find(needle)
        if k >= 0:
            return t, k
    return None, -1


def next_id(doc):
    ids = [int(x) for x in re.findall(r'w:id="(\d+)"', doc.element.xml)]
    return (max(ids) + 1) if ids else 1


def stamp():
    return datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


def change_mark(tag, wid, author, when):
    el = OxmlElement(tag)
    el.set(qn('w:id'), str(wid)); el.set(qn('w:author'), author); el.set(qn('w:date'), when)
    return el


def add_field(paragraph, instr, placeholder=''):
    """Insert a complex field (PAGE, NUMPAGES, TOC ...) at the end of a paragraph."""
    def r_with(child):
        r = OxmlElement('w:r')
        r.append(child)
        paragraph._p.append(r)
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin'); r_with(b)
    it = OxmlElement('w:instrText'); it.set(XML_SPACE, 'preserve'); it.text = f' {instr} '; r_with(it)
    s = OxmlElement('w:fldChar'); s.set(qn('w:fldCharType'), 'separate'); r_with(s)
    r_with(new_t(placeholder))
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end'); r_with(e)


def mark_update_fields(doc):
    """Ask Word to refresh fields (TOC, page counts) when the file is opened."""
    settings = doc.settings.element
    el = settings.find(qn('w:updateFields'))
    if el is None:
        el = OxmlElement('w:updateFields')
        settings.append(el)
    el.set(qn('w:val'), 'true')


# ---------- text ----------
def blocks(container):
    """Paragraphs and tables in order, looking inside content controls (w:sdt)."""
    for el in container:
        if el.tag in (qn('w:p'), qn('w:tbl')):
            yield el
        elif el.tag == qn('w:sdt'):
            content = el.find(qn('w:sdtContent'))
            if content is not None:
                yield from blocks(content)


def para_text_marked(p):
    parts = []
    for el in p.iter():
        if owning_p(el) is not p and el is not p:
            continue
        if el.tag == qn('w:t'):
            anc = el.getparent().getparent()
            parts.append(f'[+{el.text or ""}+]' if anc is not None and anc.tag in (qn('w:ins'), qn('w:moveTo')) else (el.text or ''))
        elif el.tag == qn('w:delText'):
            parts.append(f'[-{el.text or ""}-]')
        elif el.tag == qn('w:tab'):
            parts.append('\t')
        elif el.tag in (qn('w:br'), qn('w:cr')):
            parts.append('\n')
    return ''.join(parts)


def table_rows(tbl):
    rows = []
    for tr in tbl.findall(qn('w:tr')):
        cells = []
        for tc in tr.findall(qn('w:tc')):
            texts = []
            for b in blocks(tc):
                texts.append(para_text_marked(b) if b.tag == qn('w:p') else '[표: ' + ' / '.join(' | '.join(r) for r in table_rows(b)) + ']')
            cells.append('\n'.join(t for t in texts if t))
        rows.append(cells)
    return rows


def cmd_text(a):
    doc = Document(a.file)
    out = []
    for el in blocks(doc.element.body):
        if el.tag == qn('w:p'):
            style = el.find(f'{{{W}}}pPr/{{{W}}}pStyle')
            out.append({'style': style.get(qn('w:val')) if style is not None else 'Normal', 'text': para_text_marked(el)})
        else:
            out.append({'table': table_rows(el)})
    comments = []
    for rel in doc.part.rels.values():
        if rel.reltype.endswith('/comments'):
            from lxml import etree
            root = etree.fromstring(rel.target_part.blob)
            for c in root.iter(f'{{{W}}}comment'):
                comments.append({'id': c.get(f'{{{W}}}id'), 'author': c.get(f'{{{W}}}author'),
                                 'text': ''.join(t.text or '' for t in c.iter(f'{{{W}}}t'))})
    sec = doc.sections[0]
    print(json.dumps({'page': f'{sec.page_width.cm:.1f}x{sec.page_height.cm:.1f}cm' if sec.page_width else None,
                      'blocks': out, 'comments': comments}, ensure_ascii=False, indent=1))


# ---------- toc / page numbers / header ----------
def cmd_toc(a):
    doc = Document(a.file)
    anchor = None
    if a.after:
        anchor = next((p for p in doc.paragraphs if a.after in p.text), None)
        if anchor is None:
            sys.exit(f'「{a.after}」 문단을 찾지 못했습니다.')
    heading = doc.add_paragraph()  # added at the end, moved below
    heading.add_run(a.title).bold = True
    toc = doc.add_paragraph()
    add_field(toc, 'TOC \\o "1-3" \\h \\z \\u', '목차 — 워드에서 필드 업데이트(F9)를 누르면 채워집니다')
    if anchor is not None:
        anchor._p.addnext(heading._p)
        heading._p.addnext(toc._p)
    else:
        body = doc.element.body
        body.insert(0, toc._p)
        body.insert(0, heading._p)
    mark_update_fields(doc)
    doc.save(out_path(a))
    print(out_path(a))


def target_paragraph(hf, replace):
    """Paragraph to write into: wipe the part with --replace, reuse a single empty paragraph, otherwise add a new one
    so existing logos/addresses stay."""
    if replace:
        for el in list(hf._element):
            if el.tag in (qn('w:p'), qn('w:tbl'), qn('w:sdt')):
                hf._element.remove(el)
        return hf.add_paragraph()
    paras = hf.paragraphs
    if len(paras) == 1 and not paras[0]._p.findall('.//' + qn('w:r')) and len(hf.tables) == 0:
        return paras[0]
    return hf.add_paragraph()


def cmd_page_numbers(a):
    doc = Document(a.file)
    done = set()
    for s in doc.sections:
        footer = s.footer
        footer.is_linked_to_previous = False
        if id(footer._element) in done:
            continue
        done.add(id(footer._element))
        has_page = any('PAGE' in (it.text or '') for it in footer._element.iter(qn('w:instrText'))) or \
            any('PAGE' in (f.get(qn('w:instr')) or '') for f in footer._element.iter(qn('w:fldSimple')))
        if has_page and not a.replace:
            sys.exit('이미 쪽 번호가 있습니다 — 바꾸려면 --replace를 주세요(바닥글의 다른 내용도 지워집니다).')
        p = target_paragraph(footer, a.replace)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for piece in re.split(r'(\{PAGE\}|\{NUMPAGES\})', a.format):
            if piece == '{PAGE}':
                add_field(p, 'PAGE', '1')
            elif piece == '{NUMPAGES}':
                add_field(p, 'NUMPAGES', '1')
            elif piece:
                p.add_run(piece)
    mark_update_fields(doc)
    doc.save(out_path(a))
    print(out_path(a))


def cmd_header(a):
    doc = Document(a.file)
    done = set()
    for s in doc.sections:
        hf = s.footer if a.footer else s.header
        hf.is_linked_to_previous = False
        if id(hf._element) in done:
            continue
        done.add(id(hf._element))
        p = target_paragraph(hf, a.replace)
        p.add_run(a.text)
        p.alignment = {'left': WD_ALIGN_PARAGRAPH.LEFT, 'center': WD_ALIGN_PARAGRAPH.CENTER, 'right': WD_ALIGN_PARAGRAPH.RIGHT}[a.align]
    doc.save(out_path(a))
    print(out_path(a))


# ---------- replace (keeps formatting, handles text split across runs) ----------
def replace_in_paragraph(p, old, new):
    nodes = text_nodes(p)
    texts = [t.text or '' for t in nodes]
    count, start = 0, 0
    while True:
        full = ''.join(texts)
        i = full.find(old, start)
        if i < 0:
            break
        end = i + len(old)
        pos = 0
        first = None
        for k, tx in enumerate(texts):
            s, e = pos, pos + len(tx)
            pos = e
            if e <= i or s >= end:
                continue
            lo, hi = max(i, s) - s, min(end, e) - s
            if first is None:
                first = k
                texts[k] = tx[:lo] + new + tx[hi:]
            else:
                texts[k] = tx[:lo] + tx[hi:]
        count += 1
        start = i + len(new)  # never re-match inside the replacement
    for t, tx in zip(nodes, texts):
        if (t.text or '') != tx:
            t.text = tx
            t.set(XML_SPACE, 'preserve')
    return count


def cmd_replace(a):
    doc = Document(a.file)
    n = sum(replace_in_paragraph(p, a.old, a.new) for root in all_parts(doc) for p in paragraphs_xml(root))
    doc.save(out_path(a))
    print(json.dumps({'replaced': n, 'file': out_path(a)}, ensure_ascii=False))


# ---------- tracked changes ----------
def cmd_track_replace(a):
    doc = Document(a.file)
    wid, when, done, skipped = next_id(doc), stamp(), 0, 0
    for root in all_parts(doc):
        for p in paragraphs_xml(root):
            for _ in range(500):
                t, k = find_text(p, a.old)
                if t is None:
                    if a.old in ''.join(x.text or '' for x in text_nodes(p, False)):
                        skipped += 1  # spans pieces with different formatting
                    break
                mid = split_at(t, k, k + len(a.old))
                d = change_mark('w:del', wid, a.author, when); wid += 1
                mid.addprevious(d)
                dr = copy.deepcopy(mid)
                for x in list(dr):
                    if x.tag == qn('w:t'):
                        dr.remove(x)
                dr.append(new_t(a.old, 'w:delText'))
                d.append(dr)
                if a.new:
                    ins = change_mark('w:ins', wid, a.author, when); wid += 1
                    d.addnext(ins)
                    ir = copy.deepcopy(mid)
                    for x in list(ir):
                        if x.tag == qn('w:t'):
                            x.text = a.new
                    ins.append(ir)
                mid.getparent().remove(mid)
                done += 1
    doc.save(out_path(a))
    print(json.dumps({'tracked': done, 'skipped_across_formatting': skipped, 'file': out_path(a)}, ensure_ascii=False))


def cmd_track_insert(a):
    doc = Document(a.file)
    for root in all_parts(doc):
        for p in paragraphs_xml(root):
            t, k = find_text(p, a.after)
            if t is None:
                continue
            mid = split_at(t, k, k + len(a.after))
            ins = change_mark('w:ins', next_id(doc), a.author, stamp())
            ir = copy.deepcopy(mid)
            for x in ir:
                if x.tag == qn('w:t'):
                    x.text = a.text
            ins.append(ir)
            mid.addnext(ins)
            doc.save(out_path(a))
            print(json.dumps({'inserted_after': a.after, 'file': out_path(a)}, ensure_ascii=False))
            return
    sys.exit(f'「{a.after}」을(를) 한 덩어리 글에서 찾지 못했습니다 — text 명령으로 실제 글을 보고 더 짧게 주세요.')


def unwrap(el):
    parent = el.getparent()
    for child in list(el):
        el.addprevious(child)
    parent.remove(el)


CHANGE_TAGS = ('w:rPrChange', 'w:pPrChange', 'w:sectPrChange', 'w:tblPrChange', 'w:trPrChange', 'w:tcPrChange',
               'w:tblGridChange', 'w:numberingChange')
# children of a property block that a *Change element does not describe and that must survive a reject
KEEP_ON_REJECT = {qn('w:rPr'), qn('w:sectPr'), qn('w:headerReference'), qn('w:footerReference')}


def apply_changes(root, accept):
    gone, keep = ('w:del', 'w:ins') if accept else ('w:ins', 'w:del')
    gone_mv, keep_mv = ('w:moveFrom', 'w:moveTo') if accept else ('w:moveTo', 'w:moveFrom')
    n = 0
    # 1) table rows inserted/deleted as a whole
    for tr in list(root.iter(qn('w:tr'))):
        trPr = tr.find(qn('w:trPr'))
        if trPr is None:
            continue
        if trPr.find(qn(gone)) is not None:
            tr.getparent().remove(tr)
            n += 1
            continue
        m = trPr.find(qn(keep))
        if m is not None:
            trPr.remove(m)
            n += 1
    # 2) paragraph marks — a removed mark joins the paragraph with the next one (last first, so chains work)
    for p in reversed(list(root.iter(qn('w:p')))):
        rPr = p.find(f'{{{W}}}pPr/{{{W}}}rPr')
        if rPr is None:
            continue
        k = rPr.find(qn(keep))
        if k is not None:
            rPr.remove(k)
            n += 1
        g = rPr.find(qn(gone))
        if g is not None:
            rPr.remove(g)
            n += 1
            nxt = p.getnext()
            if nxt is not None and nxt.tag == qn('w:p'):
                for c in list(nxt):
                    if c.tag != qn('w:pPr'):
                        p.append(c)
                nxt.getparent().remove(nxt)
    # 3) runs
    for tag in (gone, gone_mv):
        for el in list(root.iter(qn(tag))):
            if el.getparent() is not None:
                el.getparent().remove(el)
                n += 1
    for tag in (keep, keep_mv):
        for el in list(root.iter(qn(tag))):
            if not accept:
                for dt in el.iter(qn('w:delText')):
                    dt.tag = qn('w:t')
                for di in el.iter(qn('w:delInstrText')):
                    di.tag = qn('w:instrText')
            unwrap(el)
            n += 1
    for tag in ('w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveToRangeStart', 'w:moveToRangeEnd'):
        for el in list(root.iter(qn(tag))):
            el.getparent().remove(el)
    # 4) formatting changes
    for tag in CHANGE_TAGS:
        for el in list(root.iter(qn(tag))):
            holder = el.getparent()
            holder.remove(el)
            if not accept and len(el):
                old = el[0]
                for child in list(holder):
                    if child.tag not in KEEP_ON_REJECT:
                        holder.remove(child)
                pos = 0
                while pos < len(holder) and holder[pos].tag in (qn('w:headerReference'), qn('w:footerReference')):
                    pos += 1
                for child in reversed(list(old)):
                    holder.insert(pos, child)
            n += 1
    return n


def cmd_accept_reject(a, accept):
    doc = Document(a.file)
    n = sum(apply_changes(root, accept) for root in all_parts(doc))
    doc.save(out_path(a))
    print(json.dumps({'changes': n, 'file': out_path(a)}, ensure_ascii=False))


# ---------- comments ----------
def cmd_comment(a):
    doc = Document(a.file)
    if not hasattr(doc, 'add_comment'):
        sys.exit('python-docx 1.2 이상이 필요합니다: pip install -U python-docx')
    from docx.text.run import Run
    for p in doc.paragraphs + [p for t in doc.tables for row in t.rows for c in row.cells for p in c.paragraphs]:
        t, k = find_text(p._p, a.find)
        if t is None:
            continue
        mid = split_at(t, k, k + len(a.find))
        doc.add_comment(Run(mid, p), text=a.text, author=a.author, initials=a.author[:2])
        doc.save(out_path(a))
        print(json.dumps({'commented': a.find, 'file': out_path(a)}, ensure_ascii=False))
        return
    sys.exit(f'「{a.find}」을(를) 한 덩어리 글에서 찾지 못했습니다 — text 명령으로 실제 글을 보고 더 짧게 주세요.')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    def add(name, *extra):
        p = sub.add_parser(name)
        p.add_argument('file')
        p.add_argument('--out')
        for args, kw in extra:
            p.add_argument(*args, **kw)
    REPLACE = (('--replace',), {'action': 'store_true'})
    add('text')
    add('toc', (('--after',), {}), (('--title',), {'default': '목차'}))
    add('page-numbers', (('--format',), {'default': '- {PAGE} -'}), REPLACE)
    add('header', (('--text',), {'required': True}), (('--footer',), {'action': 'store_true'}), REPLACE,
        (('--align',), {'default': 'center', 'choices': ['left', 'center', 'right']}))
    add('replace', (('--old',), {'required': True}), (('--new',), {'required': True}))
    add('track-replace', (('--old',), {'required': True}), (('--new',), {'default': ''}), (('--author',), {'default': 'AI'}))
    add('track-insert', (('--after',), {'required': True}), (('--text',), {'required': True}), (('--author',), {'default': 'AI'}))
    add('comment', (('--find',), {'required': True}), (('--text',), {'required': True}), (('--author',), {'default': 'AI'}))
    add('accept')
    add('reject')
    a = ap.parse_args()
    for key in ('old', 'find', 'after'):
        if getattr(a, key, None) == '':
            sys.exit(f'--{key}에 빈 글은 줄 수 없습니다.')
    {'text': cmd_text, 'toc': cmd_toc, 'page-numbers': cmd_page_numbers, 'header': cmd_header,
     'replace': cmd_replace, 'track-replace': cmd_track_replace, 'track-insert': cmd_track_insert,
     'comment': cmd_comment,
     'accept': lambda x: cmd_accept_reject(x, True), 'reject': lambda x: cmd_accept_reject(x, False)}[a.cmd](a)


if __name__ == '__main__':
    main()
