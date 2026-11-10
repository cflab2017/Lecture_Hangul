"""한글 기초 6강(쪽 번호) 실측 + 실습 · 정답 파일 — 한글 2024 자동화(창 숨김) → PDF 쪽 아래 글자로 번호 확인.
⚠️ 이 PC 한글 2024 는 구독이 없어 HAction.Run 편집 동작은 False(11/7) — 하지만 HAction.Execute 로
   쪽 번호 매기기(PageNumPos) · 현재 쪽만 감추기(PageHiding, Fields 0x20 = 쪽 번호) · 새 번호로 시작(NewNumber, HAutoNum)은 된다(11/10).
   표지 뒤 쪽 나누기는 5강처럼 HWPX section0.xml 단락의 pageBreak="1".
연습 문서: 가상 '봉사 동아리 하반기 활동 보고서' — 1쪽 표지 + 본문 3쪽.
A 쪽 번호 매기기(가운데 아래, 줄표) · B A+1쪽 감추기 · C B+2쪽에서 새 번호 1 · D A+커서를 2쪽에 둔 채 감추기 · E A+1쪽에서 새 번호 5 · F 줄표 없이
결과 measure.json, 실습.hwpx(번호 없음) · 정답.hwpx · 정답.pdf(= C). HWPX 작성자 · PDF 메타데이터는 CODINGNOW.
%TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서 임시 폴더에 저장 후 복사. 실행: python measure.py  (pywin32 · pymupdf)"""
import json, os, re, shutil, sys, tempfile, zipfile
import pymupdf
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul06")
os.makedirs(WORK, exist_ok=True)
COVER = ["", "", "", "봉사 동아리 하반기 활동 보고서", "", "2026년 11월", "", "연습용 가상 동아리 · 총무 작성"]
HEADS = ["1. 활동 요약", "2. 월별 활동", "3. 다음 계획"]
BODY = "연습용으로 만든 가상 동아리 보고서의 본문입니다. 쪽 번호가 어떻게 붙는지 보려고 문단을 여러 개 넣었습니다."
FIRST = HEADS[0]

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
R = {"version": hwp.Version}


def ins(t):
    a = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", a.HSet)
    a.Text = t
    return hwp.HAction.Execute("InsertText", a.HSet)


def save(name, fmt):
    p = os.path.join(WORK, name)
    if os.path.exists(p):
        os.remove(p)
    assert hwp.SaveAs(p, fmt, ""), p
    return p


def patch_hwpx(path, break_text=None):
    """작성자 CODINGNOW, break_text 로 시작하는 단락에 pageBreak="1"."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith("content.hpf"):
                s = data.decode("utf-8")
                s = re.sub(r'(<opf:meta name="(?:creator|lastsaveby)"[^>]*>)[^<]*(</opf:meta>)', r"\1CODINGNOW\2", s)
                data = s.encode("utf-8")
            if break_text and info.filename == "Contents/section0.xml":
                s = data.decode("utf-8")
                k = s.index(f">{break_text}<")
                i = s.rindex("<hp:p ", 0, k)
                j = s.index(">", i)
                s = s[:i] + s[i:j].replace('pageBreak="0"', 'pageBreak="1"') + s[j:]
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


def nums(pdf):
    """쪽마다 아래쪽(밑에서 90pt 안)에 찍힌 글자 — 쪽 번호."""
    out = []
    with pymupdf.open(pdf) as d:
        for pg in d:
            ws = pg.get_text("words")
            out.append(" ".join(w[4] for w in ws if w[1] > pg.rect.height - 90) or "(없음)")
    return out


def pagenum(side="-"):
    ps = hwp.HParameterSet.HPageNumPos
    hwp.HAction.GetDefault("PageNumPos", ps.HSet)
    ps.DrawPos = hwp.PageNumPosition("BottomCenter")
    ps.NumberFormat = hwp.NumberFormat("Digit")
    ps.SideChar = ord(side) if side else 0
    return hwp.HAction.Execute("PageNumPos", ps.HSet)


def hide_here():
    ph = hwp.HParameterSet.HPageHiding
    hwp.HAction.GetDefault("PageHiding", ph.HSet)
    ph.Fields = 0x20  # 쪽 번호 감추기
    return hwp.HAction.Execute("PageHiding", ph.HSet)


def new_number(n):
    an = hwp.HParameterSet.HAutoNum
    hwp.HAction.GetDefault("NewNumber", an.HSet)
    an.NumType = 0  # 쪽 번호
    an.NewNumber = n
    return hwp.HAction.Execute("NewNumber", an.HSet)


def para_of_page(page):
    for i in range(0, 400):
        if not hwp.SetPos(0, i, 0):
            return None
        if hwp.KeyIndicator()[3] == page:
            return i
    return None


def base_doc():
    hwp.Clear(1)
    lines = COVER + [HEADS[0]] + [BODY] * 14 + [HEADS[1]] + [BODY] * 14 + [HEADS[2]] + [BODY] * 10
    ins("\r\n".join(lines))
    p = save("base.hwpx", "HWPX")
    patch_hwpx(p, break_text=FIRST)
    return p


def case(name, steps):
    assert hwp.Open(BASE, "HWPX", ""), BASE
    res = {"steps": []}
    for s in steps:
        res["steps"].append([s[0], s[1](*s[2:])])
    res["pages"] = hwp.PageCount
    res["numbers"] = nums(save(name + ".pdf", "PDF"))
    R[name] = res
    return res


try:
    BASE = base_doc()
    assert hwp.Open(BASE, "HWPX", "")
    R["base_pages"] = hwp.PageCount
    P2 = para_of_page(2)
    R["page2_first_para"] = P2
    at = lambda i: hwp.SetPos(0, i, 0)  # noqa: E731
    case("A_pagenum", [("커서 1쪽", at, 0), ("쪽 번호 매기기", pagenum)])
    case("B_hide_cover", [("커서 1쪽", at, 0), ("쪽 번호 매기기", pagenum), ("커서 1쪽", at, 0), ("현재 쪽만 감추기", hide_here)])
    case("C_hide_and_restart", [("커서 1쪽", at, 0), ("쪽 번호 매기기", pagenum), ("커서 1쪽", at, 0), ("현재 쪽만 감추기", hide_here),
                                ("커서 2쪽 첫 문단", at, P2), ("새 번호로 시작 1", new_number, 1)])
    case("D_hide_wrong_page", [("커서 1쪽", at, 0), ("쪽 번호 매기기", pagenum), ("커서 2쪽 첫 문단", at, P2), ("현재 쪽만 감추기", hide_here)])
    case("E_start_5", [("커서 1쪽", at, 0), ("쪽 번호 매기기", pagenum), ("커서 1쪽", at, 0), ("새 번호로 시작 5", new_number, 5)])
    case("F_no_dash", [("커서 1쪽", at, 0), ("쪽 번호 매기기(줄표 없음)", pagenum, "")])
    # 실습 · 정답 파일
    assert hwp.Open(BASE, "HWPX", "")
    prac = save("실습.hwpx", "HWPX")
    patch_hwpx(prac)
    R["실습_numbers"] = nums(save("실습_확인.pdf", "PDF"))
    assert hwp.Open(BASE, "HWPX", "")
    at(0); pagenum(); at(0); hide_here(); at(P2); new_number(1)
    ans = save("정답.hwpx", "HWPX")
    patch_hwpx(ans)
    assert hwp.Open(ans, "HWPX", "")
    pdf = save("정답.pdf", "PDF")
    clean_pdf(pdf)
    R["정답_numbers"] = nums(pdf)
    hwp.Clear(1)
finally:
    hwp.Quit()
for f in ("실습.hwpx", "정답.hwpx", "정답.pdf"):
    shutil.copyfile(os.path.join(WORK, f), os.path.join(HERE, f))
with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1)
for k, v in R.items():
    if isinstance(v, dict) and "numbers" in v:
        print(k, v["numbers"], [s for s in v["steps"]])
print("base", R["base_pages"], "2쪽 첫 문단", R["page2_first_para"], "실습", R["실습_numbers"], "정답", R["정답_numbers"])
