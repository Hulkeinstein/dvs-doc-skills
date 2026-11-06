---
name: dvs-pptx
description: "Creates, reads and edits PowerPoint presentations (.pptx, .potx) — 파워포인트 · 발표 자료 · 슬라이드 · 피치덱 · 강의 자료. Use whenever a .pptx/.potx file is involved as input or output: making a deck from notes or a document, pulling text or speaker notes out of slides (even to use elsewhere), changing words or images on existing slides, filling a company template, merging or splitting decks, reordering or deleting slides, adding speaker notes. Not for Word, spreadsheets, PDF or Google Slides."
---

# 파워포인트 스킬

> 명령에 나오는 `scripts/…`는 **이 스킬 폴더 안** 경로다. 사용자 작업 폴더에서 실행할 때는 스킬 폴더를 붙인 전체 경로로 부른다(다른 스킬의 도구도 같다 — 예: `<스킬 폴더들이 있는 곳>/dvs-office/scripts/render.py`). Mac · Linux에서 `python`이 없으면 `python3`로 부른다.
## 먼저

1. `dvs-office` 스킬의 `python scripts/check_env.py --need pptx render`.
2. 있는 파일은 먼저 본다:
   - `python scripts/pptx_tools.py text 발표.pptx` — 슬라이드마다 제목 · 글 상자 · 표 · 발표자 노트(JSON).
   - `python scripts/pptx_tools.py layouts 양식.pptx` — 양식이 가진 레이아웃 이름과 자리(placeholder) 번호.
   - `dvs-office`의 `render.py 발표.pptx`로 그림을 보고 디자인(색 · 글꼴 · 배치)을 파악한다.

## 새로 만들기 — python-pptx로 직접 코드를 쓴다

- 크기: 16:9 (`prs.slide_width = Inches(13.333)`, `slide_height = Inches(7.5)`).
- 회사 · 사용자 양식이 있으면 **그 양식의 레이아웃과 자리**를 쓴다(`python scripts/pptx_tools.py clear 양식.pptx --out 빈양식.pptx`로 슬라이드만 비운 뒤 `Presentation('빈양식.pptx')`로 연다. `.potx`는 먼저 `dvs-office`의 `template.py open 양식.potx --out 양식.pptx`). 양식의 색 · 글꼴을 따르고 새로 꾸미지 않는다.
- 양식이 없으면 아래 디자인 원칙으로 직접 꾸민다.

### 디자인 원칙

- **주제에 맞는 색 하나를 고른다.** 기본 파랑을 습관처럼 쓰지 않는다. 주색 1 + 짙은 글자색 1 + 연한 바탕 1 + 강조색 1, 이 4가지 안에서 쓴다. 표지 · 마무리는 짙은 바탕, 본문은 밝은 바탕처럼 「샌드위치」로 짜면 정돈돼 보인다.
- **슬라이드마다 그림 요소 하나**: 사진 · 아이콘 · 도형 · 차트 · 큰 숫자 중 하나. 글만 있는 슬라이드를 연달아 두지 않는다.
- 배치를 돌려 가며 쓴다: 왼쪽 글 + 오른쪽 그림 / 2~3단 카드 / 큰 숫자 + 설명 / 단계 흐름(1→2→3) / 비교(전 vs 후) / 표.
- 글자 크기: 제목 36~44pt, 본문 16~20pt, 설명 · 출처 11~12pt. 한 슬라이드 본문은 6줄 안팎까지 — 넘치면 나눈다.
- 한글 글꼴은 `dvs-office`의 순서(맑은 고딕 → …)로. 글 상자마다 `font.name`을 정하고 동아시아 글꼴(`a:ea`)도 같이 정한다(`pptx_tools.py`의 `set_korean_font` 참고).
- 여백: 가장자리 0.5인치 이상, 요소 사이 0.3~0.5인치를 고르게. 왼쪽 정렬이 기본이고, 가운데 정렬은 표지 · 큰 숫자에만.
- 하지 말 것: 제목 밑 장식 줄, 글 상자 아래 꽉 찬 색 띠 반복, 모든 슬라이드 같은 배치, 그림 비율 찌그러뜨리기(`width`나 `height` 하나만 준다), 읽기 어려운 연한 회색 글자.
- 차트는 python-pptx의 `chart_data`(`CategoryChartData`)로 넣는다 — 그림으로 붙이지 않아야 사용자가 숫자를 고칠 수 있다.
- 발표자 노트: `slide.notes_slide.notes_text_frame.text = "…"` — 사용자가 「대본」 「노트」를 원하면 넣는다.

## 있는 파일 고치기

- 글자 바꾸기(서식 유지): `python scripts/pptx_tools.py replace 발표.pptx --old "2025" --new "2026" --out 발표-수정.pptx`
- 슬라이드 빼기 · 순서 바꾸기: `python scripts/pptx_tools.py keep 발표.pptx --slides 1,3,2,5 --out 새.pptx` (적은 번호만, 적은 순서대로 남는다)
- 복제: `python scripts/pptx_tools.py duplicate 발표.pptx --slide 3 --out 발표-수정.pptx` (3번 바로 뒤에 같은 슬라이드)
- 나누기: `python scripts/pptx_tools.py split 발표.pptx --slides 1-5 --out 앞부분.pptx` (= `keep`의 범위형)
- 합치기: `python scripts/pptx_tools.py merge 가.pptx 나.pptx --out 합본.pptx` — 그림 · 차트 · 표 · 노트 · 전환 효과까지 옮긴다. 같은 양식끼리가 가장 깔끔하다. 양식이 다르면 뒤 파일 슬라이드는 앞 파일의 같은 이름(없으면 빈) 레이아웃으로 옮겨져 **바탕 · 색 · 자리 위치가 바뀔 수 있다** — 결과를 반드시 그림으로 확인한다. 다른 슬라이드로 건너뛰는 링크는 빠진다.
- 양식으로 저장: 만든 결과를 `.potx` 양식으로 넘겨야 하면 `dvs-office`의 `template.py save 결과.pptx --out 양식.potx`.
- 노트 넣기: `python scripts/pptx_tools.py notes 발표.pptx --slide 2 --text "여기서 숫자 강조" --out 발표-수정.pptx`
- 그 밖의 수정(그림 바꾸기, 표 칸 고치기, 도형 추가)은 python-pptx로 열어 코드로 하고 새 이름으로 저장한다. 기존 자리(placeholder)에 글을 넣을 때는 `shape.text_frame`을 쓰고 새 글 상자를 겹쳐 올리지 않는다.

## 넘기기 전 확인 (빠뜨리지 않는다)

> 🔴 그림으로 보는 확인(`render.py`)을 하지 않았으면 「다 됐다」고 말하지 않는다. 할 수 없었으면(LibreOffice가 없거나 응답하지 않음) 그 사실을 사용자에게 그대로 알린다.

1. `dvs-office`의 `render.py 결과.pptx --dpi 60`으로 슬라이드 그림을 모두 연다. 처음 볼 때는 **문제가 있다고 가정하고** 찾는다:
   - 글이 상자 밖으로 넘침 · 다른 요소와 겹침 · 잘린 글
   - 한글이 네모(□)로 깨짐 · 글꼴이 섞임
   - 대비가 낮아 안 읽히는 글 · 가장자리에 붙은 요소 · 고르지 않은 간격
   - 비어 있는 자리(「제목을 입력하십시오」) 남음
2. 문제를 고치고 **다시 그려서** 고친 슬라이드를 다시 본다. 한 번에 다 고쳤다고 믿지 않는다.
3. `text` 명령으로 글 순서 · 빠진 슬라이드가 없는지 본다.
4. 저장 위치와 슬라이드 수를 알려 준다.
