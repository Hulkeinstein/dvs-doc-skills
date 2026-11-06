---
name: dvs-pdf
description: "Reads, creates and changes PDF files — PDF · 스캔 문서 · 계약서 · 영수증 · 신청서 양식. Use whenever a .pdf is involved or the user wants one: pulling out text or tables, merging several PDFs, splitting or extracting pages, rotating, adding a watermark or page numbers, filling in a PDF form, locking with or removing a password, pulling out images, reading scanned pages with OCR, or making a new PDF (from notes, markdown, or by converting a Word/PowerPoint/Excel file)."
---

# PDF 스킬


> 명령에 나오는 `scripts/…`는 **이 스킬 폴더 안** 경로다. 사용자 작업 폴더에서 실행할 때는 스킬 폴더를 붙인 전체 경로로 부른다(다른 스킬의 도구도 같다 — 예: `<스킬 폴더들이 있는 곳>/dvs-office/scripts/render.py`). Mac · Linux에서 `python`이 없으면 `python3`로 부른다.
## 먼저

1. `dvs-office` 스킬의 `python scripts/check_env.py --need pdf` (스캔 글자 읽기는 `--need pdf ocr`).
2. 무슨 PDF인지 본다: `python scripts/pdf_tools.py info 파일.pdf` — 쪽 수, 쪽 크기, 암호, 양식 칸 수, 글자층이 있는지(없으면 스캔본).

## 읽기

- 글: `python scripts/pdf_tools.py text 파일.pdf [--pages 1-3]` — 쪽마다 글(배치 유지).
- 표: `python scripts/pdf_tools.py tables 파일.pdf [--pages 2] [--xlsx 표.xlsx]` — 표를 JSON으로, `--xlsx`면 표마다 시트로 저장.
- 그림 꺼내기: `python scripts/pdf_tools.py images 파일.pdf --outdir 그림`
- **스캔본**(info에서 `has_text: false`): `python scripts/pdf_tools.py ocr 파일.pdf [--lang kor+eng] [--searchable 결과.pdf]` — 글을 읽어 출력하고, `--searchable`이면 검색 · 복사가 되는 PDF를 만든다. Tesseract 한국어 데이터(`kor`)가 필요하다.
- **스캔본 한국어 준비**(Tesseract 기본 설치에는 한국어가 없다 · 관리자 권한 없이 한 번만): 사용자 폴더에 `.dvs-office/tessdata`를 만들고, Tesseract 설치 폴더의 `tessdata` 안 `eng.traineddata` · `osd.traineddata` · `pdf.ttf` · `configs` 폴더를 복사한 뒤, 공식 저장소 https://github.com/tesseract-ocr/tessdata_fast 의 `kor.traineddata`를 받아 같은 곳에 넣는다(받기 전에 사용자에게 묻는다). 더 정확히 읽어야 하면 `tessdata_best`의 `kor.traineddata`를 쓴다(느리고 크다).
- `--searchable`은 쪽을 300dpi 그림으로 다시 만들고 그 위에 글자층을 얹는다 — 파일이 커지고, `--pages`를 주면 그 쪽만 담긴다.
- 읽은 글은 틀릴 수 있다(굵은 글씨 · 작은 글씨 · 표). 숫자 · 이름은 그림(`render.py`)과 대조해 확인하고, 여러 단으로 된 쪽이 섞여 나오면 `--psm 3`으로 다시 읽는다.
- 쪽을 눈으로 봐야 하면 `dvs-office`의 `render.py 파일.pdf --pages 1`.

## 바꾸기 (원본은 그대로, 결과는 새 파일)

| 하고 싶은 일 | 명령 |
|---|---|
| 합치기 | `pdf_tools.py merge 가.pdf 나.pdf --out 합본.pdf` |
| 쪽 뽑기 · 순서 바꾸기 | `pdf_tools.py pages 파일.pdf --pages 3,1,5-7 --out 새.pdf` |
| 쪽마다 나누기 | `pdf_tools.py split 파일.pdf --outdir 나눔` |
| 돌리기 | `pdf_tools.py rotate 파일.pdf --pages 2 --degrees 90 --out 새.pdf` |
| 워터마크 | `pdf_tools.py watermark 파일.pdf --text "대외비" --out 새.pdf` (또는 `--stamp 도장.pdf`) |
| 쪽 번호 | `pdf_tools.py number 파일.pdf --out 새.pdf [--format "{n} / {total}"]` |
| 암호 걸기 | `pdf_tools.py encrypt 파일.pdf --password 비밀번호 --out 새.pdf` |
| 암호 풀기 | `pdf_tools.py decrypt 파일.pdf --password 비밀번호 --out 새.pdf` |

- 암호는 사용자가 채팅에 준 것만 쓴다. 답할 때 암호를 다시 적지 않는다.

## 양식(신청서) 채우기

1. `python scripts/pdf_tools.py fields 양식.pdf` — 칸 이름 · 종류 · 지금 값 · 고를 수 있는 값(체크 상자 · 선택 목록)을 본다.
2. 채울 값을 JSON으로 쓴다: `{"성명": "홍길동", "동의": "/Yes"}` — 체크 상자는 `fields`가 보여 준 「켜짐 값」을 쓴다.
3. `python scripts/pdf_tools.py fill 양식.pdf --data 값.json --out 작성본.pdf` → `render.py`로 그림을 보고 글이 칸에 맞게 들어갔는지 확인한다.
- 한글 값은 「열 때 칸을 다시 그리라」는 표시로 보이게 한다. 대부분의 PDF 보기 프로그램(Acrobat · Edge · Chrome)은 이를 따르지만, 일부 휴대폰 · Mac 미리보기에서는 칸이 비어 보일 수 있다 — 그럴 때는 칸이 없는 양식 방법(`overlay`)으로 글을 얹는다.
- 관공서 양식에 흔한 XFA 양식이면 도구가 XFA 부분을 지우고 일반 칸에 채운 뒤 경고를 낸다. 그림 확인을 꼭 한다.
- **칸이 없는 양식**(그냥 그림 · 글만 있는 PDF)이면: `render.py`로 쪽을 그림으로 보고 빈칸 위치(가로 · 세로 pt, 왼쪽 아래가 0)를 정해 `pdf_tools.py overlay 양식.pdf --data 위치.json --out 작성본.pdf`로 글을 얹는다. 위치 JSON = `[{"page": 1, "x": 120, "y": 640, "text": "홍길동", "size": 11}]` — 좌표는 **화면에 보이는 쪽** 기준(돌아간 쪽도 보이는 방향 그대로). 얹은 뒤 반드시 그림으로 확인하고 어긋나면 좌표를 고쳐 다시.

## 새 PDF 만들기

- 다른 오피스 파일이 있으면 바꾸는 게 가장 깔끔하다: `dvs-office`의 `convert.py 보고서.docx --to pdf`.
- 글 · 마크다운에서 바로: `python scripts/pdf_tools.py create 원고.md --out 결과.pdf [--title 제목]` — `#` 제목, 문단, `-` 목록, `|` 표, `**굵게**`를 한글 글꼴로 A4에 찍는다.
- 더 꾸민 PDF가 필요하면 reportlab으로 코드를 직접 쓴다. 한글은 **반드시 한글 글꼴을 등록**해야 한다(`pdf_tools.py`의 `register_korean_font()`를 가져다 쓴다) — 기본 글꼴은 한글을 네모로 찍는다. 위첨자 · 아래첨자는 유니코드 문자(²) 대신 `<super>2</super>`를 쓴다.

## 넘기기 전 확인

- 만든 · 바꾼 PDF는 `render.py`로 첫 쪽과 바뀐 쪽을 그림으로 본다(한글 깨짐 · 잘림 · 돌아간 방향).
- `info`로 쪽 수가 기대와 맞는지 본다.
- 저장 위치를 알려 준다.
