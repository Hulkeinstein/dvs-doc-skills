"""Render a .docx/.pptx/.xlsx/.odt/.pdf file to one PNG per page so the result can be looked at.

  python render.py report.docx --out check --dpi 80
  python render.py deck.pptx --pages 1-3

Office files are first converted to PDF with LibreOffice; PDFs are rendered directly (pypdfium2).
Prints a JSON list of the PNG paths.
"""
import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from office_common import soffice_convert  # noqa: E402


def parse_pages(spec, n):
    if not spec:
        return list(range(n))
    out = []
    for part in spec.split(','):
        a, _, b = part.partition('-')
        lo, hi = int(a), int(b or a)
        out += [i - 1 for i in range(lo, min(hi, n) + 1)]
    return out


def render_pdf(pdf, outdir, dpi, pages, stem):
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(pdf))
    doc.init_forms()  # draw filled-in form fields too, so a filled form can be checked by eye
    paths = []
    for i in parse_pages(pages, len(doc)):
        img = doc[i].render(scale=dpi / 72, may_draw_forms=True).to_pil()
        p = outdir / f'{stem}-{i + 1:03d}.png'
        img.save(p)
        paths.append(str(p))
    return paths


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('file')
    ap.add_argument('--out', default=None, help='folder for PNGs (default: <file>-pages next to the file)')
    ap.add_argument('--dpi', type=int, default=80)
    ap.add_argument('--pages', default=None, help='e.g. 1-3,5 (1-based)')
    a = ap.parse_args()
    src = Path(a.file).resolve()
    outdir = Path(a.out or src.with_name(src.stem + '-pages')).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    if src.suffix.lower() == '.pdf':
        paths = render_pdf(src, outdir, a.dpi, a.pages, src.stem)
    else:
        with tempfile.TemporaryDirectory() as tmp:
            pdf = soffice_convert(src, 'pdf', tmp)
            paths = render_pdf(pdf, outdir, a.dpi, a.pages, src.stem)
    print(json.dumps({'pages': len(paths), 'images': paths}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
