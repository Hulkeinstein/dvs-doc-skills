---
name: dvs-docx
description: Creates, reads and edits Word documents (.docx, .dotx) — 워드 · 보고서 · 회의록 · 공문 · 편지 · 양식. Use whenever the user wants a report, memo, letter, minutes or template as a Word file, or wants an existing Word file read, summarised, corrected, filled in, given a table of contents / page numbers / header, have images inserted, find-and-replace, comments, or tracked changes (변경 내용 추적 · 메모). Not for PDF, spreadsheets, slides or Google Docs.
---

# 워드 스킬


> 명령에 나오는 `scripts/…`는 **이 스킬 폴더 안** 경로다. 사용자 작업 폴더에서 실행할 때는 스킬 폴더를 붙인 전체 경로로 부른다(다른 스킬의 도구도 같다 — 예: `<스킬 폴더들이 있는 곳>/dvs-office/scripts/render.py`). Mac · Linux에서 `python`이 없으면 `python3`로 부른다.
## 먼저

1. `dvs-office` 스킬의 `scripts/check_env.py --need docx`(그림 확인까지 하려면 `docx render`).
2. 있는 문서는 먼저 읽는다: `python scripts/docx_tools.py text 파일.docx` — 제목 수준 · 문단 · 표를 순서대로, 변경 내용 추적 · 메모가 있으면 함께 보여 준다.

## 새로 만들기 — python-docx로 직접 코드를 쓴다

- 구조는 **제목 스타일**로 만든다: `doc.add_heading(text, level=1~3)`. 글씨를 키워 흉내 내지 않는다(목차가 제목 스타일을 읽는다).
- 한글 글꼴을 문서 기본으로 정한다: `Normal` 스타일의 `font.name`과 동아시아 글꼴(`w:eastAsia`)을 둘 다 맑은 고딕(없으면 `dvs-office`의 `--fonts` 결과)으로. 본문 10.5~11pt, 줄 간격 1.15~1.3.
- 글머리표 · 번호는 `List Bullet` · `List Number` 스타일을 쓴다(글자로 「•」를 찍지 않는다).
- 표: `doc.add_table(rows, cols, style='Table Grid')`, 첫 줄 머리글 굵게 · 연한 배경, 숫자는 오른쪽 정렬, 표 너비는 쪽 안에.
- 그림: `doc.add_picture(path, width=Cm(15))` — 쪽 너비(A4 여백 빼고 약 16cm)를 넘기지 않는다.
- 쪽: 한국 문서는 A4(21×29.7cm), 여백 2~2.5cm.
- 만든 뒤 꾸밈은 스크립트로:
  - 목차: `python scripts/docx_tools.py toc 파일.docx --after "제목"` (워드에서 열 때 갱신하라는 표시도 넣는다)
  - 쪽 번호: `python scripts/docx_tools.py page-numbers 파일.docx --format "- {PAGE} / {NUMPAGES} -"`
  - 머리글 · 바닥글 글: `python scripts/docx_tools.py header 파일.docx --text "우리 팀 주간 보고"`
  - 둘 다 기존 머리글 · 바닥글(로고 · 주소)은 두고 새 줄로 더한다. 모두 지우고 새로 쓰려면 `--replace`(지우기 전에 사용자에게 확인).

## 있는 문서 고치기

- 글자만 바꾸기: `python scripts/docx_tools.py replace 파일.docx --old "2025" --new "2026" --out 파일-수정.docx` — 서식은 유지된다. 바꾼 개수를 확인한다(0이면 글이 여러 조각으로 나뉘어 있을 수 있다 — `text`로 실제 글을 보고 더 짧은 글로 다시).
- 문단 추가 · 표 고치기 · 그림 넣기는 python-docx로 열어 코드로 하고 새 이름으로 저장한다.
- **변경 내용 추적**(사용자가 「고친 곳이 보이게」, 「수정 표시로」, 「검토용으로」 원할 때):
  `python scripts/docx_tools.py track-replace 파일.docx --old "다음 주" --new "이번 주" --author "AI 비서" --out 파일-검토.docx`
  워드에서 빨간 줄(지운 글)과 밑줄(새 글)로 보이고, 사용자가 하나씩 수락 · 거부할 수 있다.
  글을 지우지 않고 더하기만: `python scripts/docx_tools.py track-insert 파일.docx --after "예산" --text "(확정)" --author "AI 비서" --out 파일-검토.docx`
  `skipped_across_formatting`이 0보다 크면, 찾는 글이 서로 다른 서식(굵게 등)에 걸쳐 있어 표시하지 못한 곳이 있다는 뜻이다 — 더 짧은 글로 다시 한다.
- **메모 달기**: `python scripts/docx_tools.py comment 파일.docx --find "예산" --text "금액 확인 필요" --author "AI 비서" --out 파일-메모.docx`
- 추적된 변경을 모두 반영: `python scripts/docx_tools.py accept 파일.docx --out 파일-최종.docx` (모두 되돌리기는 `reject`).

## 양식(.dotx)

- 양식을 받으면 그 양식의 스타일을 그대로 써서 내용만 채운다. python-docx는 `.dotx`를 바로 열지 못하므로 `dvs-office`의 `template.py open 양식.dotx --out 작업.docx`로 바꿔 연다(내용은 그대로, 형식 표시만 바뀐다). 결과를 다시 양식으로 넘길 때는 `template.py save 작업.docx --out 새양식.dotx`.

## 넘기기 전 확인

1. `python scripts/docx_tools.py text 결과.docx`로 글 순서 · 빠진 곳을 본다.
2. `dvs-office`의 `render.py 결과.docx`로 쪽 그림을 보고: 한글 깨짐 · 표가 쪽 밖으로 · 빈 쪽 · 그림 잘림이 없는지.
3. 목차를 넣었으면 「워드에서 열 때 목차를 갱신할지 물으면 예를 누르세요」라고 알려 준다.
