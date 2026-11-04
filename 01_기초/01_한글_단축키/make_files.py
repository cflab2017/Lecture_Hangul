"""한글 기초 1강(단축키) 파일을 만든다 — 한글 2024 자동화(창 숨김). 내용은 모두 연습용 가상 글이다.
① 단축키표.hwpx · 단축키표.pdf — 자주 쓰는 단축키 표(한글 2024 도움말 '단축키 일람' + 웹 한글 도움말에서 옮김, 인쇄용)
② 실습.hwpx — 단축키로 고쳐 볼 회의 안내문 ③ 정답.hwpx — 제목 가운데 정렬·글씨 16pt, 끝에 3줄 2칸 표
(글자 크기는 HWPX 값으로 넣는다 — first_para_size)
이 PC 한글은 편집 화면이 열리지 않아(구독 없음) HAction.Execute 와 SetPos·SelectText 만 쓴다(hwp.Run 은 동작하지 않음).
저장 뒤 HWPX 안의 작성자(content.hpf)는 CODINGNOW 로 바꾼다. 실행: python make_files.py  (pip install pywin32 pymupdf)"""
import os, re, shutil, tempfile, zipfile
import pymupdf
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
# 한글 자동화는 %TEMP% 밖 경로에 저장하려 하면 '파일 접근 허용' 보안 창을 띄우고 멈춘다 — 임시 폴더에 저장하고 끝에 복사한다
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul01"); os.makedirs(WORK, exist_ok=True)

# (분류, 기능, 한글 2024, 웹 한글) — 웹 한글 칸: 웹 도움말 '단축키 일람'에 있는 키, 없으면 —
KEYS = [
    ("파일", "저장하기", "Alt+S · Ctrl+S", "Ctrl+S"),
    ("파일", "다른 이름으로 저장하기", "Alt+V", "—"),
    ("파일", "불러오기", "Alt+O", "—"),
    ("파일", "인쇄", "Alt+P · Ctrl+P", "Ctrl+P"),
    ("쪽", "편집 용지", "F7", "F7"),
    ("편집", "되돌리기 · 다시 실행", "Ctrl+Z · Ctrl+Shift+Z", "Ctrl+Z · Ctrl+Shift+Z"),
    ("편집", "모두 선택", "Ctrl+A", "Ctrl+A"),
    ("편집", "블록 설정(2번 낱말, 3번 문단)", "F3", "F3"),
    ("편집", "모양 복사", "Alt+C", "Ctrl+Alt+C"),
    ("편집", "찾아 바꾸기", "Ctrl+F2 · Ctrl+H", "Ctrl+H · Ctrl+F2"),
    ("편집", "강제 줄 나누기", "Shift+Enter", "Shift+Enter"),
    ("입력", "문자표", "Ctrl+F10", "Ctrl+F10"),
    ("입력", "한자로 바꾸기", "F9", "—"),
    ("입력", "각주", "Ctrl+N,N", "Ctrl+N,N · Alt+N,N"),
    ("서식", "글자 모양", "Alt+L", "Alt+L"),
    ("서식", "문단 모양", "Alt+T", "Alt+T"),
    ("서식", "스타일", "F6", "F6"),
    ("서식", "글씨 크게 · 작게", "Ctrl+] · Ctrl+[", "Ctrl+] · Ctrl+["),
    ("서식", "진하게 · 기울임 · 밑줄", "Ctrl+B · Ctrl+I · Ctrl+U", "Ctrl+B · Ctrl+I · Ctrl+U"),
    ("서식", "자간 넓게 · 좁게", "Alt+Shift+W · Alt+Shift+N", "Alt+Shift+W · Alt+Shift+N"),
    ("서식", "가운데 · 왼쪽 · 오른쪽 정렬", "Ctrl+Shift+C · L · R", "Ctrl+Shift+C · L · R"),
    ("서식", "양쪽 정렬", "Ctrl+Shift+M", "Ctrl+Shift+M"),
    ("서식", "줄 간격 넓게 · 좁게", "Alt+Shift+Z · Alt+Shift+A", "Alt+Shift+Z · Alt+Shift+A"),
    ("서식", "첫 줄 들여 쓰기", "Ctrl+F6", "Ctrl+F6"),
    ("쪽", "쪽 나누기", "Ctrl+Enter · Ctrl+J", "Ctrl+Enter"),
    ("쪽", "쪽 번호 매기기", "Ctrl+N,P", "—"),
    ("쪽", "머리말/꼬리말", "Ctrl+N,H", "—"),
    ("표", "표 만들기", "Ctrl+N,T", "Ctrl+N,T · Alt+N,T"),
    ("표", "셀 블록(3번이면 표 전체)", "F5", "F5"),
    ("표", "표 안에서 줄 추가하기", "Ctrl+Enter", "Ctrl+Enter"),
    ("표", "줄/칸 추가하기 · 지우기", "Alt+Insert · Alt+Delete", "Alt+Insert · Alt+Delete"),
    ("표", "셀 합치기 · 셀 나누기(셀 블록에서)", "M · S", "M · S"),
    ("보기", "조판 부호 보이기", "Ctrl+G,C", "Ctrl+G,C"),
    ("보기", "문단 부호 보이기", "Ctrl+G,T", "Ctrl+G,T"),
]
TITLE = "한글 단축키 표 — 자주 쓰는 34가지 (한글 2024 · 웹 한글)"
NOTE = ("'Ctrl+N,T' 처럼 쉼표가 있으면 Ctrl+N 을 함께 눌렀다가 떼고 T 를 누릅니다. 출처: 한글 2024 도움말 '단축키 일람', "
        "한컴오피스 Web 도움말 '단축키 일람'. 웹 한글은 브라우저 단축키(크롬의 Ctrl+N 새 창 등)가 먼저 동작할 수 있습니다. — 코딩나우 coding-now.com")
MEMO_TITLE = "11월 팀 회의 안내"
MEMO = [
    "이번 주 팀 회의는 화요일 오전 10시에 3층 대회의실에서 합니다.",
    "발표 자료는 월요일 오후 6시까지 팀 공유 폴더에 올려 주세요.",
    "회의가 끝나면 12시부터 같은 층 휴게실에서 점심을 함께합니다.",
]
TASKS = [
    "[따라 하기] ① 제목 줄에 커서를 두고 Ctrl+Shift+C — 가운데 정렬",
    "② 제목 줄에서 F3 을 세 번 눌러 블록을 잡고 Ctrl+] 를 여러 번 — 글씨 크게(정답 파일은 16pt)",
    "③ 문서 끝에서 Ctrl+N,T — 줄 3 · 칸 2 표를 만들어 '날짜 / 11월 10일', '장소 / 3층 대회의실' 넣기(표 안 이동은 Tab)",
    "④ Alt+T 로 문단 모양 창을 열어 줄 간격을 확인하고 Esc",
]

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
VERSION = hwp.Version


def ins(t):
    a = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", a.HSet)
    a.Text = t
    return hwp.HAction.Execute("InsertText", a.HSet)


def para_align(kind):  # 3 = 가운데
    ps = hwp.HParameterSet.HParaShape
    hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
    ps.AlignType = kind
    return hwp.HAction.Execute("ParagraphShape", ps.HSet)


def char(height=None, bold=None):
    cs = hwp.HParameterSet.HCharShape
    hwp.HAction.GetDefault("CharShape", cs.HSet)
    if height:
        cs.Height = height
    if bold is not None:
        cs.Bold = 1 if bold else 0
    return hwp.HAction.Execute("CharShape", cs.HSet)


def table(rows, cols, cells):
    tc = hwp.HParameterSet.HTableCreation
    hwp.HAction.GetDefault("TableCreate", tc.HSet)
    tc.Rows, tc.Cols = rows, cols
    assert hwp.HAction.Execute("TableCreate", tc.HSet)
    lst0 = hwp.GetPos()[0]
    for r, row in enumerate(cells):
        for c, t in enumerate(row):
            hwp.SetPos(lst0 + r * cols + c, 0, 0)
            ins(t)
    return lst0


def save(name, fmt):
    path = os.path.join(WORK, name)
    if os.path.exists(path):
        os.remove(path)
    assert hwp.SaveAs(path, fmt, ""), path


def widen(path, mm):
    """표 칸 너비(mm)를 바꾼다 — <hp:cellSz width> 는 칸마다, <hp:sz width> 는 표 전체."""
    hu = 7200 / 25.4
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "Contents/section0.xml":
                x = data.decode("utf-8")
                a = x.index("<hp:tbl "); b = x.index("</hp:tbl>", a)
                t = x[a:b]

                def tc(m):
                    c = m.group(0)
                    col = int(re.search(r'<hp:cellAddr colAddr="(\d+)"', c).group(1))
                    return re.sub(r'(<hp:cellSz width=")\d+', lambda k: k.group(1) + str(round(mm[col] * hu)), c)
                t = re.sub(r"<hp:tc .*?</hp:tc>", tc, t, flags=re.S)
                t = re.sub(r'(<hp:sz width=")\d+', lambda k: k.group(1) + str(round(sum(mm) * hu)), t, count=1)
                data = (x[:a] + t + x[b:]).encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, path)


def first_para_size(path, height):
    """첫 문단(제목) 글자 크기를 바꾼다 — 자동화의 CharShape 실행이 이 PC 에서 먹지 않아(10pt 그대로) HWPX 값으로.
    header.xml 의 charPr 0 을 복사해 새 id·height 로 넣고, 첫 문단의 run 들이 그 id 를 쓰게 한다."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx").name
    with zipfile.ZipFile(path) as zin:
        h = zin.read("Contents/header.xml").decode("utf-8")
        x = zin.read("Contents/section0.xml").decode("utf-8")
        m = re.search(r'<hh:charProperties itemCnt="(\d+)">', h)
        n = int(m.group(1))
        c0 = re.search(r'<hh:charPr id="0" .*?</hh:charPr>', h, re.S).group(0)
        c_new = re.sub(r'id="0"', f'id="{n}"', c0, count=1)
        c_new = re.sub(r'height="\d+"', f'height="{height}"', c_new, count=1)
        h = h.replace(m.group(0), f'<hh:charProperties itemCnt="{n + 1}">', 1)
        h = h.replace("</hh:charProperties>", c_new + "</hh:charProperties>", 1)
        a = x.index("<hp:p "); b = x.index("</hp:p>", a)
        para = re.sub(r'<hp:run charPrIDRef="\d+"', f'<hp:run charPrIDRef="{n}"', x[a:b])
        x = x[:a] + para + x[b:]
        with zipfile.ZipFile(tmp, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "Contents/header.xml":
                    data = h.encode("utf-8")
                elif info.filename == "Contents/section0.xml":
                    data = x.encode("utf-8")
                zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, path)


def clean_hwpx(path):
    """HWPX(zip) 안 content.hpf 의 작성자·마지막 수정한 사람을 CODINGNOW 로."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith("content.hpf"):
                s = data.decode("utf-8")
                s = re.sub(r'(<opf:meta name="(?:creator|lastsaveby)"[^>]*>)[^<]*(</opf:meta>)', r"\1CODINGNOW\2", s)
                data = s.encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, path)


def clean_pdf(path):
    doc = pymupdf.open(path)
    meta = doc.metadata
    meta.update({"author": "CODINGNOW", "creator": "Hancom Office Hangul 2024", "producer": "CODINGNOW"})
    doc.set_metadata(meta)
    tmp = path + ".tmp.pdf"
    doc.save(tmp, garbage=3, deflate=True)
    doc.close()
    shutil.move(tmp, path)


try:
    # ① 단축키 표
    hwp.Clear(1)
    ins(TITLE + "\r\n")
    hwp.SetPos(0, 0, 0)
    para_align(3)
    hwp.SelectText(0, 0, 0, len(TITLE))
    char(height=1500, bold=True)
    hwp.MovePos(3, 0, 0)
    table(len(KEYS) + 1, 4, [("분류", "기능", "한글 2024", "웹 한글(MYBOX)")] + KEYS)
    hwp.MovePos(3, 0, 0)
    ins("\r\n" + NOTE)
    save("단축키표.hwpx", "HWPX")
    first_para_size(os.path.join(WORK, "단축키표.hwpx"), 1500)
    # 칸 너비를 내용에 맞춰(분류 좁게) HWPX 값만 바꾸고 다시 열어 PDF 로 — 같은 너비면 글이 접혀 2쪽이 된다
    widen(os.path.join(WORK, "단축키표.hwpx"), [14, 48, 44, 44])
    hwp.Clear(1)
    assert hwp.Open(os.path.join(WORK, "단축키표.hwpx"), "HWPX", "")
    save("단축키표.pdf", "PDF")
    # ② 실습
    hwp.Clear(1)
    ins(MEMO_TITLE + "\r\n" + "\r\n".join(MEMO) + "\r\n\r\n" + "\r\n".join(TASKS))
    save("실습.hwpx", "HWPX")
    # ③ 정답 — 제목 가운데 + 16pt, 본문 끝에 표
    hwp.Clear(1)
    ins(MEMO_TITLE + "\r\n" + "\r\n".join(MEMO) + "\r\n")
    hwp.SetPos(0, 0, 0)
    para_align(3)
    hwp.SelectText(0, 0, 0, len(MEMO_TITLE))
    char(height=1600)
    hwp.MovePos(3, 0, 0)
    table(3, 2, [("항목", "내용"), ("날짜", "11월 10일"), ("장소", "3층 대회의실")])
    save("정답.hwpx", "HWPX")
    hwp.Clear(1)
    first_para_size(os.path.join(WORK, "정답.hwpx"), 1600)
    # 바꾼 HWPX 를 한글이 다시 여는지 확인
    assert hwp.Open(os.path.join(WORK, "정답.hwpx"), "HWPX", ""), "정답.hwpx 를 다시 못 연다"
    hwp.Clear(1)
finally:
    hwp.Quit()

for f in ("단축키표.hwpx", "실습.hwpx", "정답.hwpx"):
    clean_hwpx(os.path.join(WORK, f))
clean_pdf(os.path.join(WORK, "단축키표.pdf"))
for f in ("단축키표.hwpx", "단축키표.pdf", "실습.hwpx", "정답.hwpx"):
    shutil.copy2(os.path.join(WORK, f), os.path.join(HERE, f))
pdf = pymupdf.open(os.path.join(HERE, "단축키표.pdf"))
print("version", VERSION, "| pdf pages", pdf.page_count, "| meta", {k: pdf.metadata.get(k) for k in ("author", "creator", "producer")})
pdf.close()
