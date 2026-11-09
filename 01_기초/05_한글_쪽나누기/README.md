# 기초 5강 — 한글 쪽 나누기 · 빈 페이지 삭제 (Ctrl+Enter, 문단 부호로 빈 줄 찾기)

- 강의: https://www.coding-now.com/blog/hangul/hangul-basic-05-page-break

## 파일

| 파일 | 내용 |
|---|---|
| `실습.hwpx` | 따라 할 파일 — 가상 모임 공지 뒤 빈 줄(Enter) 37개로 '참가 신청서' 를 2쪽에 밀어 놓고 끝에 빈 줄 40개 → 3쪽(3쪽은 빈 쪽) |
| `정답.hwpx` · `정답.pdf` | 빈 줄을 지우고 '참가 신청서' 앞에 쪽 나누기 → 2쪽 |
| `measure.py` → `measure.json` | 한글 2024 자동화로 잰 값 — 빈 줄 몇 개에서 2쪽, Enter 로 민 글 vs 쪽 나누기 2쪽 첫 줄 위치(PDF), 편집 동작(HAction.Run) 결과 |
| `make_files.py` | 실습 · 정답 파일 만들기 |
| `webhwp.json` | 네이버 MYBOX 웹 한글에서 잰 것 — 쪽 메뉴 항목, Ctrl+Enter 1/1→2/2쪽, Backspace 2/2→1/1쪽, 실습 파일 3/3쪽 → 문단 부호 ↵ → Backspace 40번 2/2쪽 |
| `crop_web.py` | 강의 그림(`shots/` 원본은 저장소에 올리지 않음) — 웹 한글은 창 전체 + 빨간 테두리, PDF 2쪽 비교 |

스크립트는 Windows + 한글 2024 + `pip install pywin32 pymupdf pillow` 에서 돈다. %TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서 임시 폴더에 저장한 뒤 복사한다. HWPX 작성자 · PDF 메타데이터는 CODINGNOW.

## 핵심

- 이 PC 한글 2024 는 구독이 없어 HAction.Run('BreakPage' · 'BreakPara' · 'MoveDocEnd' · 'DeleteBack')이 모두 False — 글 넣기(InsertText, 줄바꿈 \r\n)만 된다. 쪽 나누기는 HWPX section0.xml 단락의 `pageBreak="1"` 로 넣었고, 한글로 다시 열어 저장해도 남았다.
- 빈 줄 37개면 둘째 쪽(빈 쪽)이 생김(36개는 1쪽).
- Enter 37번으로 민 '2쪽 시작' 줄: 1쪽에 3줄 더 쓰면 2쪽 y 99.4 → 147.3pt. 쪽 나누기면 99.4pt 그대로.
- 웹 한글: 쪽 메뉴 — 편집 용지 F7 · 머리말 · 꼬리말 · 새 번호로 시작 · 현재 쪽만 감추기 Alt+N+S · 쪽 나누기 Ctrl+Enter · 단 나누기 Ctrl+Shift+Enter · 단 · 다단 설정 나누기 Ctrl+Alt+Enter.
