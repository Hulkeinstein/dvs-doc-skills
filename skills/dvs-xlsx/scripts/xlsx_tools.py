"""Small spreadsheet helpers.

  python xlsx_tools.py inspect sales.xlsx            # sheets, size, header row, formulas, merged cells, charts/images
  python xlsx_tools.py to-csv sales.xlsx --sheet 1월 # one sheet -> sales-1월.csv (utf-8-sig, opens fine in Excel)
  python xlsx_tools.py from-csv data.csv             # csv/tsv -> data.xlsx with a formatted header row
  python xlsx_tools.py fit report.xlsx [--out x.xlsx] # widen columns so no text is cut off in Excel
"""
import argparse
import csv
import io
import re
import json
import math
import sys
from pathlib import Path


def inspect(path):
    import openpyxl
    wb = openpyxl.load_workbook(path)
    out = []
    for ws in wb:
        formulas = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('='))
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1), [])][:30]
        out.append({
            'sheet': ws.title, 'range': ws.dimensions, 'rows': ws.max_row, 'cols': ws.max_column,
            'header_row_1': header, 'formulas': formulas,
            'merged': [str(r) for r in ws.merged_cells.ranges][:20],
            'charts': len(getattr(ws, '_charts', [])), 'images': len(getattr(ws, '_images', [])),
            'freeze': ws.freeze_panes,
        })
    return {'file': str(path), 'sheets': out}


def to_csv(path, sheet):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb.active
    dst = Path(path).with_name(f'{Path(path).stem}-{ws.title}.csv')
    with open(dst, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        for row in ws.iter_rows(values_only=True):
            w.writerow(['' if v is None else v for v in row])
    return str(dst)


def from_csv(path):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    src = Path(path)
    raw = src.read_bytes()
    for enc in ('utf-8-sig', 'cp949', 'utf-8'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        sys.exit('글자 인코딩을 알 수 없습니다.')
    delim = '\t' if src.suffix.lower() == '.tsv' or text.count('\t') > text.count(',') else ','
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = src.stem[:31]
    for row in csv.reader(io.StringIO(text, newline=''), delimiter=delim):  # keeps line breaks inside quoted cells
        ws.append([_num(v) for v in row])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', start_color='E8EEF7')
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        width = max(sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(c.value or '')) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(60, width + 2)
    dst = src.with_suffix('.xlsx')
    if dst.exists():
        dst = src.with_name(src.stem + '-변환.xlsx')
    wb.save(dst)
    return str(dst)


def _text_width(text, size, bold):
    """Width in Excel column units. Korean glyphs are drawn wider in Excel / Excel for the web than in
    LibreOffice, so a picture made with LibreOffice can look fine while the real file is cut off."""
    w = max((sum(2.5 if ord(ch) > 0x2E80 else 1.2 for ch in line) for line in text.split('\n')), default=0)
    return w * (size / 11) * (1.08 if bold else 1)


def fit(path, out=None):
    """Widen every column so no text is cut off (wrap-text cells are left alone)."""
    import openpyxl
    from openpyxl.styles import Alignment
    wb = openpyxl.load_workbook(path)
    changed = []
    for ws in wb:
        filter_row = None
        if ws.auto_filter.ref:
            filter_row = int(re.search(r'\d+', ws.auto_filter.ref).group())
        merged = {c for r in ws.merged_cells.ranges for row in ws.iter_rows(
            min_row=r.min_row, max_row=r.max_row, min_col=r.min_col, max_col=r.max_col) for c in row}
        need = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is None or c in merged:
                    continue
                if c.alignment and c.alignment.wrap_text:
                    # wrapping only helps if the row is tall enough to show a second line;
                    # a one-line row (e.g. a header with filter buttons) still cuts the text
                    h = ws.row_dimensions[c.row].height
                    if h and h >= ((c.font.sz if c.font and c.font.sz else 11) * 2.6):
                        continue
                if isinstance(c.value, str) and c.value.startswith('='):
                    continue  # result unknown until Excel calculates it
                if isinstance(c.value, str) and ws.cell(c.row, c.column + 1).value is None:
                    continue  # Excel lets text run on into an empty cell to the right (titles, notes)
                text = f'{c.value:,}' if isinstance(c.value, (int, float)) else str(c.value)
                size = (c.font.sz or 11) if c.font else 11
                w = _text_width(text, size, bool(c.font and c.font.b)) + 3  # cell padding on both sides
                if c.row == filter_row:
                    w += 3  # room for the filter button
                need[c.column_letter] = max(need.get(c.column_letter, 0), w)
        for col, w in need.items():
            cur = ws.column_dimensions[col].width or 8.43
            w = math.ceil(w)  # whole numbers, so running fit twice changes nothing
            if w <= cur:
                continue
            if w > 80:  # too long for one line: wrap instead of a giant column
                ws.column_dimensions[col].width = max(cur, 60)
                for c in ws[col]:
                    c.alignment = Alignment(wrap_text=True, vertical=c.alignment.vertical if c.alignment else None)
                changed.append(f'{ws.title}!{col} 줄바꿈')
            else:
                ws.column_dimensions[col].width = w
                changed.append(f'{ws.title}!{col} {cur:.0f}→{w:.0f}')
    wb.save(out or path)
    return {'file': str(out or path), 'changed': changed}


NUMBER = re.compile(r'-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?')


def _num(v):
    """Turn '1,234' / '12.5' into numbers; keep codes ('00123'), long ids (>15 digits — Excel would round them)
    and anything else as text."""
    s = v.strip()
    if not NUMBER.fullmatch(s):
        return v
    digits = s.replace(',', '').lstrip('-')
    if (len(digits) > 1 and digits[0] == '0' and not digits.startswith('0.')) or len(digits.replace('.', '')) > 15:
        return v
    s = s.replace(',', '')
    return float(s) if '.' in s else int(s)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('inspect'); p.add_argument('file')
    p = sub.add_parser('to-csv'); p.add_argument('file'); p.add_argument('--sheet')
    p = sub.add_parser('from-csv'); p.add_argument('file')
    p = sub.add_parser('fit'); p.add_argument('file'); p.add_argument('--out')
    a = ap.parse_args()
    if a.cmd == 'inspect':
        print(json.dumps(inspect(a.file), ensure_ascii=False, indent=2, default=str))
    elif a.cmd == 'fit':
        print(json.dumps(fit(a.file, a.out), ensure_ascii=False, indent=2))
    elif a.cmd == 'to-csv':
        print(to_csv(a.file, a.sheet))
    else:
        print(from_csv(a.file))


if __name__ == '__main__':
    main()
