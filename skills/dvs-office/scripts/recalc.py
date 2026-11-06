"""Recalculate every formula in an .xlsx with LibreOffice and report error cells.

  python recalc.py sales.xlsx            # report only — the file is not changed
  python recalc.py sales.xlsx --show B12 Summary!C3   # also print computed values of these cells
  python recalc.py new.xlsx --mark       # file you just made: also set "recalculate when opened"
  python recalc.py new.xlsx --write-values new-값.xlsx   # also save a copy whose formula cells carry computed values

How it works: a temporary copy is recalculated by LibreOffice and read back with openpyxl(data_only).
Exit code 0 = no formula errors, 1 = errors found (listed in the JSON),
2 = LibreOffice did not answer: nothing was checked (with --mark the file is still marked
"recalculate when opened"). Tell the user the check could not be done — never report 0 errors.
"""
import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from office_common import SofficeError, soffice_convert  # noqa: E402

ERRORS = ('#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NUM!', '#NULL!', 'Err:', '#SPILL!', '#CALC!')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('file')
    ap.add_argument('--show', nargs='*', default=[], help='cells to print, e.g. B12 or Sheet2!C3')
    ap.add_argument('--mark', action='store_true', help='mark the file "recalculate on open" (re-saves it in place with openpyxl — only for files this task created)')
    ap.add_argument('--write-values', metavar='OUT.xlsx', help='save the LibreOffice-recalculated workbook (formulas + cached values) as OUT.xlsx')
    a = ap.parse_args()
    import openpyxl
    src = Path(a.file).resolve()
    if a.write_values and (Path(a.write_values).resolve() == src or not a.write_values.lower().endswith('.xlsx')):
        sys.exit('--write-values에는 원본과 다른 .xlsx 이름을 주세요.')
    macro = src.suffix.lower() in ('.xlsm', '.xltm')
    formulas = openpyxl.load_workbook(src, keep_vba=macro)  # formulas as text; keep macros if re-saved
    n_formulas = sum(1 for ws in formulas for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('='))
    with tempfile.TemporaryDirectory() as tmp:
        try:
            calc_path = soffice_convert(src, 'xlsx', tmp)
        except SofficeError as e:
            if a.mark:
                formulas.calculation.fullCalcOnLoad = True
                formulas.save(src)
            print(json.dumps({'formulas': n_formulas, 'recalculated': False, 'marked_recalc_on_open': a.mark,
                              'message': f'{e.code} 수식 오류 검사는 하지 못했습니다'
                                         + (' — 엑셀에서 열 때 다시 계산되도록 표시했습니다.' if a.mark else '.')},
                             ensure_ascii=False, indent=2))
            return 2
        calc = openpyxl.load_workbook(calc_path, data_only=True)
        errors = []
        for ws in calc:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith(ERRORS):
                        src_cell = formulas[ws.title][c.coordinate].value if ws.title in formulas.sheetnames else None
                        errors.append({'cell': f'{ws.title}!{c.coordinate}', 'value': c.value, 'formula': src_cell})
        shown = {}
        for ref in a.show:
            sheet, _, cell = ref.rpartition('!')
            ws = calc[sheet] if sheet else calc.active
            shown[ref] = ws[cell].value
        if a.write_values:
            # Previews, pandas and openpyxl(data_only) read cached values; openpyxl-made files have none.
            shutil.copyfile(calc_path, a.write_values)
    if a.mark:
        # Excel then recalculates on open. Only for files this task created with openpyxl:
        # re-saving with openpyxl drops charts/images it did not create.
        formulas.calculation.fullCalcOnLoad = True
        formulas.save(src)
    print(json.dumps({'formulas': n_formulas, 'errors': errors, 'values': shown}, ensure_ascii=False, indent=2, default=str))
    return 1 if errors else 0


if __name__ == '__main__':
    sys.exit(main())
