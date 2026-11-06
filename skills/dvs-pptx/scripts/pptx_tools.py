"""PowerPoint (.pptx) helpers built on python-pptx.

  python pptx_tools.py text deck.pptx                         # per slide: layout, title, text boxes, tables, notes (JSON)
  python pptx_tools.py layouts template.pptx                  # layout names and placeholder idx/type
  python pptx_tools.py replace deck.pptx --old A --new B [--out o.pptx]
  python pptx_tools.py keep deck.pptx --slides 1,3,2 [--out o.pptx]      # keep only these, in this order
  python pptx_tools.py split deck.pptx --slides 1-5 [--out o.pptx]
  python pptx_tools.py clear template.pptx --out empty.pptx          # remove all slides, keep layouts/theme
  python pptx_tools.py duplicate deck.pptx --slide 3 [--out o.pptx]
  python pptx_tools.py merge a.pptx b.pptx [...] --out all.pptx
  python pptx_tools.py notes deck.pptx --slide 2 --text "..." [--out o.pptx]

Without --out, editing commands write <name>-수정.pptx (the input is never overwritten).
Slide numbers start at 1.
"""
import argparse
import copy
import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def out_path(a, default_suffix='-수정'):
    p = Path(a.file)
    out = getattr(a, 'out', None) or str(p.with_name(p.stem + default_suffix + p.suffix))
    if Path(out).resolve() == p.resolve():
        sys.exit('원본과 같은 이름으로 저장하지 않습니다 — --out에 다른 이름을 주세요.')
    return out


def parse_slides(spec, total):
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
        sys.exit(f'없는 슬라이드 번호: {bad} (전체 {total}장)')
    return nums


def set_korean_font(text_frame, name='맑은 고딕'):
    """Set both the latin and east-asian font of every run (PowerPoint uses a:ea for Hangul)."""
    for p in text_frame.paragraphs:
        for r in p.runs:
            r.font.name = name
            rPr = r._r.get_or_add_rPr()
            ea = rPr.find(qn('a:ea'))
            if ea is None:
                ea = rPr.makeelement(qn('a:ea'), {})
                rPr.append(ea)
            ea.set('typeface', name)


# ---------- reading ----------
def shape_info(sh):
    d = {'name': sh.name, 'kind': str(sh.shape_type).split('.')[-1].split(' ')[0] if sh.shape_type else 'placeholder'}
    if sh.is_placeholder:
        d['placeholder'] = sh.placeholder_format.idx
    if sh.has_text_frame and sh.text_frame.text.strip():
        d['text'] = sh.text_frame.text
    if getattr(sh, 'has_table', False) and sh.has_table:
        d['table'] = [[c.text for c in row.cells] for row in sh.table.rows]
    if getattr(sh, 'has_chart', False) and sh.has_chart:
        d['chart'] = str(sh.chart.chart_type)
    if sh.shape_type is not None and 'GROUP' in str(sh.shape_type):
        d['children'] = [shape_info(c) for c in sh.shapes]
    return d


def cmd_text(a):
    prs = Presentation(a.file)
    slides = []
    for i, s in enumerate(prs.slides, 1):
        notes = notes_text(s)
        title = s.shapes.title.text if s.shapes.title is not None else ''
        slides.append({'slide': i, 'layout': s.slide_layout.name, 'title': title,
                       'shapes': [shape_info(sh) for sh in s.shapes], 'notes': notes})
    print(json.dumps({'size_in': [round(prs.slide_width / 914400, 2), round(prs.slide_height / 914400, 2)],
                      'slides': slides}, ensure_ascii=False, indent=1))


def cmd_layouts(a):
    prs = Presentation(a.file)
    out = []
    for i, lay in enumerate(prs.slide_layouts):
        out.append({'index': i, 'name': lay.name,
                    'placeholders': [{'idx': ph.placeholder_format.idx, 'type': str(ph.placeholder_format.type).split('.')[-1].split(' ')[0], 'name': ph.name}
                                     for ph in lay.placeholders]})
    print(json.dumps(out, ensure_ascii=False, indent=1))


# ---------- editing ----------
def iter_text_frames(shapes):
    for sh in shapes:
        if sh.has_text_frame:
            yield sh.text_frame
        if getattr(sh, 'has_table', False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    yield c.text_frame
        if sh.shape_type is not None and 'GROUP' in str(sh.shape_type):
            yield from iter_text_frames(sh.shapes)


def replace_in_paragraph(p, old, new):
    runs = list(p.runs)
    texts = [r.text for r in runs]
    count, start = 0, 0
    while True:
        full = ''.join(texts)
        i = full.find(old, start)
        if i < 0:
            break
        end = i + len(old)
        pos, starts = 0, []
        for t in texts:
            starts.append(pos)
            pos += len(t)
        first = max(k for k in range(len(runs)) if starts[k] <= i and (len(texts[k]) or starts[k] == i))
        for k in range(len(runs)):
            s, e = starts[k], starts[k] + len(texts[k])
            if e <= i or s >= end:
                continue
            lo, hi = max(i, s) - s, min(end, e) - s
            texts[k] = texts[k][:lo] + (new if k == first else '') + texts[k][hi:]
        count += 1
        start = i + len(new)  # never re-match inside the replacement
    for r, t in zip(runs, texts):
        if r.text != t:
            r.text = t
    return count


def cmd_replace(a):
    prs = Presentation(a.file)
    n = 0
    for s in prs.slides:
        frames = list(iter_text_frames(s.shapes))
        if s.has_notes_slide and s.notes_slide.notes_text_frame is not None:
            frames.append(s.notes_slide.notes_text_frame)
        for tf in frames:
            for p in tf.paragraphs:
                n += replace_in_paragraph(p, a.old, a.new)
    prs.save(out_path(a))
    print(json.dumps({'replaced': n, 'file': out_path(a)}, ensure_ascii=False))


def drop_slide(prs, index):
    sldIdLst = prs.slides._sldIdLst
    sldId = sldIdLst[index]
    prs.part.drop_rel(sldId.get(qn('r:id')))
    sldIdLst.remove(sldId)


def keep_slides(prs, order):
    """Keep only the 1-based slide numbers in `order`, in that order (no repeats)."""
    if len(set(order)) != len(order):
        sys.exit('같은 번호를 두 번 쓸 수 없습니다 — 복제는 duplicate 명령을 쓰세요.')
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    keep = [ids[n - 1] for n in order]
    for el in ids:
        if el not in keep:
            prs.part.drop_rel(el.get(qn('r:id')))
            sldIdLst.remove(el)
    for el in keep:  # re-append in requested order
        sldIdLst.remove(el)
        sldIdLst.append(el)


def cmd_clear(a):
    """Remove every slide but keep masters/layouts/theme — the starting point for building on a template."""
    prs = Presentation(a.file)
    lst = prs.slides._sldIdLst
    for el in list(lst):
        prs.part.drop_rel(el.get(qn('r:id')))
        lst.remove(el)
    prs.save(out_path(a))
    print(json.dumps({'slides': 0, 'layouts': len(prs.slide_layouts), 'file': out_path(a)}, ensure_ascii=False))


def cmd_keep(a):
    prs = Presentation(a.file)
    order = parse_slides(a.slides, len(prs.slides))
    keep_slides(prs, order)
    prs.save(out_path(a))
    print(json.dumps({'slides': len(prs.slides), 'file': out_path(a)}, ensure_ascii=False))


SKIP_RELS = ('/slideLayout', '/notesSlide', '/slide', '/comments', '/commentAuthors')
CLONE_ALWAYS = ('/chart', '/package', '/oleObject', '/diagramData', '/diagramLayout', '/diagramQuickStyle',
                '/diagramColors', '/diagramDrawing', '/chartUserShapes')


def _template(partname):
    """'/ppt/media/image3.png' -> '/ppt/media/image%d.png' for next_partname()."""
    stem, dot, ext = str(partname).rpartition('.')
    return re.sub(r'\d+$', '', stem) + '%d' + dot + ext


def clone_part(part, package, cache):
    """Copy a part (and the parts it points to) into `package` under a fresh partname.
    Copying the part object itself would put two files with the same name in the zip."""
    if id(part) in cache:
        return cache[id(part)]
    from pptx.opc.package import PartFactory
    new = PartFactory(package.next_partname(_template(part.partname)), part.content_type, package, part.blob)
    cache[id(part)] = new
    rid_map = {}
    for rid, rel in part.rels.items():
        if rel.reltype.endswith(SKIP_RELS):
            continue
        if rel.is_external:
            rid_map[rid] = new.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rid_map[rid] = new.relate_to(clone_part(rel.target_part, package, cache), rel.reltype)
    el = getattr(new, '_element', None)
    if el is not None:
        remap_rids(el, rid_map)
    return new


def remap_rids(root, rid_map):
    """Point r:* attributes at the new relationship ids; drop links whose target was not copied (e.g. jump to slide 5)."""
    for el in list(root.iter()):
        for attr in list(el.attrib):
            if not attr.startswith('{%s}' % R_NS):
                continue
            v = el.get(attr)
            if v in rid_map:
                el.set(attr, rid_map[v])
            elif el.tag.endswith(('}hlinkClick', '}hlinkHover')) and el.getparent() is not None:
                el.getparent().remove(el)
                break


def copy_slide_into(dst_prs, src_slide, layout=None, position=None):
    """Copy a slide (shapes, background, transition/animation, hidden flag, notes, pictures, charts, links)
    into a new slide of dst_prs."""
    same_deck = src_slide.part.package is dst_prs.part.package
    layout = layout or pick_layout(dst_prs, src_slide.slide_layout)
    new = dst_prs.slides.add_slide(layout)
    for ph in list(new.placeholders):
        ph._element.getparent().remove(ph._element)
    cache, rid_map = {}, {}
    for rid, rel in src_slide.part.rels.items():
        if rel.reltype.endswith(SKIP_RELS):
            continue
        if rel.is_external:
            rid_map[rid] = new.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif same_deck and not rel.reltype.endswith(CLONE_ALWAYS):
            rid_map[rid] = new.part.relate_to(rel.target_part, rel.reltype)  # pictures/media can be shared
        else:
            rid_map[rid] = new.part.relate_to(clone_part(rel.target_part, new.part.package, cache), rel.reltype)
    src_el, dst_el = src_slide._element, new._element
    src_cSld, dst_cSld = src_el.find(qn('p:cSld')), dst_el.find(qn('p:cSld'))
    bg = src_cSld.find(qn('p:bg'))
    if bg is not None:
        dst_cSld.insert(0, copy.deepcopy(bg))
    tree = dst_cSld.find(qn('p:spTree'))
    for el in src_cSld.find(qn('p:spTree')):
        if el.tag in (qn('p:nvGrpSpPr'), qn('p:grpSpPr')):
            continue
        tree.append(copy.deepcopy(el))
    for tag in ('p:transition', 'p:timing'):
        el = src_el.find(qn(tag))
        if el is not None:
            dst_el.append(copy.deepcopy(el))
    if src_el.get('show') is not None:
        dst_el.set('show', src_el.get('show'))
    remap_rids(dst_el, rid_map)
    notes = notes_text(src_slide)
    if notes:
        tf = new.notes_slide.notes_text_frame
        if tf is not None:
            tf.text = notes
    if position is not None:
        lst = dst_prs.slides._sldIdLst
        el = lst[-1]
        lst.remove(el)
        lst.insert(position, el)
    return new


def notes_text(slide):
    if not slide.has_notes_slide:
        return ''
    tf = slide.notes_slide.notes_text_frame
    return tf.text if tf is not None else ''


def pick_layout(prs, src_layout):
    for lay in prs.slide_layouts:
        if lay.name == src_layout.name:
            return lay
    blank = [l for l in prs.slide_layouts if len(l.placeholders) == 0]
    return blank[0] if blank else prs.slide_layouts[-1]


def cmd_duplicate(a):
    prs = Presentation(a.file)
    n = parse_slides(str(a.slide), len(prs.slides))[0]
    src = prs.slides[n - 1]
    copy_slide_into(prs, src, layout=src.slide_layout, position=n)
    prs.save(out_path(a))
    print(json.dumps({'duplicated': n, 'slides': len(prs.slides), 'file': out_path(a)}, ensure_ascii=False))


def cmd_merge(a):
    if any(Path(a.out).resolve() == Path(f).resolve() for f in a.files):
        sys.exit('합칠 파일과 같은 이름으로 저장하지 않습니다.')
    base = Presentation(a.files[0])
    for f in a.files[1:]:
        other = Presentation(f)
        if (other.slide_width, other.slide_height) != (base.slide_width, base.slide_height):
            print(f'⚠ {f}의 슬라이드 크기가 달라 배치가 어긋날 수 있습니다.', file=sys.stderr)
        for s in other.slides:
            copy_slide_into(base, s)
    base.save(a.out)
    print(json.dumps({'slides': len(base.slides), 'file': a.out}, ensure_ascii=False))


def cmd_notes(a):
    prs = Presentation(a.file)
    n = parse_slides(str(a.slide), len(prs.slides))[0]
    tf = prs.slides[n - 1].notes_slide.notes_text_frame
    if tf is None:
        sys.exit('이 파일의 노트 양식에 본문 자리가 없어 노트를 넣을 수 없습니다.')
    tf.text = a.text
    prs.save(out_path(a))
    print(out_path(a))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    for name in ('text', 'layouts'):
        sub.add_parser(name).add_argument('file')
    p = sub.add_parser('replace'); p.add_argument('file'); p.add_argument('--old', required=True); p.add_argument('--new', required=True); p.add_argument('--out')
    for name in ('keep', 'split'):
        p = sub.add_parser(name); p.add_argument('file'); p.add_argument('--slides', required=True); p.add_argument('--out')
    p = sub.add_parser('clear'); p.add_argument('file'); p.add_argument('--out')
    p = sub.add_parser('duplicate'); p.add_argument('file'); p.add_argument('--slide', required=True); p.add_argument('--out')
    p = sub.add_parser('merge'); p.add_argument('files', nargs='+'); p.add_argument('--out', required=True)
    p = sub.add_parser('notes'); p.add_argument('file'); p.add_argument('--slide', required=True); p.add_argument('--text', required=True); p.add_argument('--out')
    a = ap.parse_args()
    if a.cmd == 'merge' and len(a.files) < 2:
        sys.exit('합칠 파일을 두 개 이상 주세요.')
    {'text': cmd_text, 'layouts': cmd_layouts, 'replace': cmd_replace, 'keep': cmd_keep, 'split': cmd_keep, 'clear': cmd_clear,
     'duplicate': cmd_duplicate, 'merge': cmd_merge, 'notes': cmd_notes}[a.cmd](a)


if __name__ == '__main__':
    main()
