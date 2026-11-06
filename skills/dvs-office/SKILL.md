---
name: dvs-office
description: Shared helpers for office files (엑셀 · 워드 · 파워포인트 · PDF). Checks which Python packages and programs (LibreOffice, Tesseract) are installed and prints install commands, renders .docx/.pptx/.xlsx/.pdf pages to PNG images so the agent can look at the result, converts office files to PDF, and recalculates Excel formulas then reports formula errors. Used by dvs-xlsx, dvs-docx, dvs-pptx and dvs-pdf before and after making a file.
---

# 오피스 공용 도구

엑셀 · 워드 · 파워포인트 · PDF 스킬이 함께 쓰는 도구다. 스크립트는 `scripts/` 안에 있고, 처음에는 `--help`로 쓰는 법만 본다(소스 전체를 읽지 않는다).

> 명령에 나오는 `scripts/…`는 **이 스킬 폴더 안** 경로다. 사용자 작업 폴더에서 실행할 때는 스킬 폴더를 붙인 전체 경로로 부른다(다른 스킬의 도구도 같다 — 예: `<스킬 폴더들이 있는 곳>/dvs-office/scripts/render.py`). Mac · Linux에서 `python`이 없으면 `python3`로 부른다.
## 1. 시작 전에 — 설치 확인

```
python scripts/check_env.py            # 무엇이 있고 무엇이 없는지
python scripts/check_env.py --need docx render   # 이번 일에 필요한 것만
```

- 없는 것이 있으면 출력된 설치 명령을 **사용자에게 보여 주고 「설치해도 될까요?」라고 물은 뒤** 실행한다. 묻지 않고 설치하지 않는다. 앱이 명령을 자동으로 실행하는 설정이어도 먼저 묻고 대답을 기다린다(설치는 사용자 PC를 바꾸는 일이다).
- 파이썬 자체가 없으면: Windows는 `winget install -e --id Python.Python.3.12`, Mac은 `brew install python`.
- LibreOffice(그림으로 확인 · 수식 다시 계산 · PDF로 바꾸기)와 Tesseract(스캔 글자 읽기)는 꼭 필요할 때만 설치를 권한다. 없어도 만들기 · 읽기 · 고치기는 된다.

## 2. 만든 뒤에 — 눈으로 확인

```
python scripts/render.py 보고서.docx --out 확인용 --dpi 80
```

- 쪽마다 PNG가 생긴다. **그림을 직접 열어 본다**: 글자가 넘치거나 겹치지 않는지, 한글이 네모(□)로 깨지지 않았는지, 빈 쪽이 없는지, 표가 쪽 밖으로 나가지 않았는지.
- 문제가 보이면 고치고 다시 그린다. 문제를 찾지 못했으면 「확인한 쪽 수」를 사용자에게 말한다.
- 확인용 그림 폴더는 일이 끝나면 지운다(사용자가 원하면 남긴다).

## 3. PDF로 바꾸기

```
python scripts/convert.py 보고서.docx --to pdf
```

## 4. 엑셀 수식 다시 계산 · 오류 찾기

```
python scripts/recalc.py 매출.xlsx --show 요약!B12      # 검사만 — 파일은 그대로
python scripts/recalc.py 새로만든.xlsx --mark            # 이번에 만든 파일: 「열 때 다시 계산」 표시까지
```

- 임시 사본을 LibreOffice로 다시 계산해 `#REF!` · `#DIV/0!` · `#VALUE!` · `#NAME?` · `#N/A` 같은 오류 칸을 시트 · 칸 주소와 수식으로 알려 준다(JSON). `--show`로 계산된 값을 몇 칸 찍어 맞는지 본다.
- 오류가 0이 될 때까지 고친다. 수식이 있는 파일은 **반드시** 이 단계를 거친 뒤 넘긴다.
- `--mark`는 openpyxl로 이번에 만든 파일에만 쓴다 — 그 파일을 그 자리에서 다시 저장한다(openpyxl이 만들지 않은 차트 · 그림이 빠지기 때문). 매크로 파일(.xlsm)은 매크로를 유지한다.
- 미리보기 · 다른 프로그램이 계산된 숫자를 읽어야 하면(openpyxl로 만든 파일은 수식 칸에 값이 비어 있다) `--write-values 결과-값.xlsx`로 계산 값이 든 사본을 함께 만든다. 이 사본은 LibreOffice가 저장한 것이라 차트 · 서식 일부가 달라질 수 있으니 그림으로 확인한다.

## 5. 양식 파일(.dotx · .potx · .xltx)

```
python scripts/template.py open 양식.dotx --out 작업.docx     # 양식 → 일반 문서(라이브러리로 열 수 있게)
python scripts/template.py save 작업.docx --out 새양식.dotx   # 일반 문서 → 양식
```

## 공통 규칙

- 원본 파일을 덮어쓰지 않는다. 고친 결과는 `이름-수정.확장자`처럼 새 이름으로 저장하고, 사용자가 원하면 그때 바꾼다.
- 한글 글꼴: 맑은 고딕(Windows) → Apple SD Gothic Neo(Mac) → 나눔고딕 / Noto Sans KR 순서로 있는 것을 쓴다. `python scripts/check_env.py --fonts`로 찾는다.
- 파일을 넘길 때 어디에 저장했는지 전체 경로를 알려 준다.
