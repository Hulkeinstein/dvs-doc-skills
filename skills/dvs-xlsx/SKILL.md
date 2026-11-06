---
name: dvs-xlsx
description: Creates, reads, edits, cleans and checks Excel spreadsheets (.xlsx, .xlsm, .csv, .tsv) — 엑셀 · 스프레드시트 · 표 · 매출표 · 가계부. Use whenever the result should be a spreadsheet file or the user points at an existing spreadsheet (even casually, like "다운로드 폴더에 있는 엑셀"): making a table from notes, adding columns or totals, summarising by month, fixing broken formulas, tidying messy rows, adding a chart, or converting between xlsx and csv. Not for Word, slides, PDF or Google Sheets.
---

# 엑셀 스킬


> 명령에 나오는 `scripts/…`는 **이 스킬 폴더 안** 경로다. 사용자 작업 폴더에서 실행할 때는 스킬 폴더를 붙인 전체 경로로 부른다(다른 스킬의 도구도 같다 — 예: `<스킬 폴더들이 있는 곳>/dvs-office/scripts/render.py`). Mac · Linux에서 `python`이 없으면 `python3`로 부른다.
## 먼저

1. `dvs-office` 스킬의 `python scripts/check_env.py --need xlsx recalc`로 도구를 확인한다(없으면 묻고 설치).
2. 있는 파일을 다룰 때는 손대기 전에 구조부터 본다:
   `python scripts/xlsx_tools.py inspect 파일.xlsx` — 시트 · 크기 · 머리글 · 수식 수 · 병합 칸 · 차트 유무.

## 만들기 · 고치기 — openpyxl로 직접 코드를 쓴다

- **계산은 엑셀이 하게 한다.** 합계 · 평균 · 비율은 파이썬에서 계산한 숫자를 넣지 말고 `=SUM(B2:B13)`처럼 수식으로 넣는다. 그래야 사용자가 숫자를 바꾸면 결과도 바뀐다.
- 입력값은 한곳(예: 「가정」 칸이나 시트)에 두고 수식이 그 칸을 가리키게 한다. 같은 숫자를 여러 칸에 적어 넣지 않는다.
- 첫 줄 머리글: 굵게, 연한 배경, 틀 고정(`ws.freeze_panes = 'A2'`), 자동 필터.
- 열 너비는 내용에 맞춘다(한글은 글자당 폭 약 2로 계산). 날짜 · 금액 · 비율은 표시 형식을 준다(`yyyy-mm-dd`, `#,##0`, `0.0%`).
- 큰 표는 pandas로 다루고, 서식은 openpyxl로 입힌다. pandas가 없거나 PC 보안 정책이 막으면(Windows 「앱 제어」) openpyxl만으로 읽고 고친다.
- 있는 파일을 고칠 때는 `openpyxl.load_workbook(파일)`로 열어 필요한 칸만 바꾸고 **새 이름으로 저장**한다. 매크로 파일(.xlsm)은 `keep_vba=True`. 엑셀 양식(.xltx)은 `dvs-office`의 `template.py open`으로 먼저 바꾼다.
- ⚠ openpyxl은 원래 파일의 차트 · 그림 · 피벗 일부를 저장할 때 잃는다. `inspect`에서 차트나 그림이 보이면 사용자에게 먼저 알리고, 값만 바꾸는 일이면 그대로 진행하되 결과를 그림으로 확인한다.

## 정리하기 (지저분한 표)

- 머리글이 몇 줄 아래에 있거나 빈 줄 · 합계 줄이 섞여 있으면: pandas로 읽어(`header=`, `skiprows=`) 진짜 머리글을 찾고, 빈 줄 · 중복 · 앞뒤 공백을 지우고, 날짜 · 숫자 형식을 맞춘다.
- 무엇을 지웠는지(몇 줄 · 어떤 기준) 사용자에게 말한다. 원본은 남긴다.

## 차트

- `openpyxl.chart`(BarChart · LineChart · PieChart)로 넣는다. 제목 · 축 이름을 붙이고, 데이터 범위는 수식이 있는 칸을 가리킨다.

## csv · tsv

- 숫자로 보이는 글 중 앞자리 0(우편번호 · 상품 코드)과 16자리 이상(주문 번호)은 글자로 남긴다(엑셀은 15자리까지만 정확하다).
- `python scripts/xlsx_tools.py to-csv 파일.xlsx --sheet 시트명` / `from-csv 파일.csv` (한글 csv는 `utf-8-sig`로 저장해야 엑셀에서 안 깨진다 — 스크립트가 그렇게 한다).

## 넘기기 전 확인 (빠뜨리지 않는다)

1. 수식이 있으면 `dvs-office`의 `python scripts/recalc.py 파일.xlsx --mark --show <핵심 칸>` → 오류 0, 핵심 값이 맞는지 본다. 이번에 만든 파일이 아니면 `--mark` 없이. 사용자가 다른 프로그램(구글 시트 미리보기 · 데이터 분석)으로 숫자를 읽을 거면 `--write-values 파일-값.xlsx`도.
2. 필요하면 `python scripts/render.py 파일.xlsx`로 그림을 보고 잘린 열 · ##### 표시가 없는지 본다.
3. 저장 위치와 시트 이름, 무엇을 넣었는지 한두 줄로 알려 준다.
