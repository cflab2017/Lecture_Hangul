"""한글 기초 5강(쪽 나누기 · 빈 페이지) 실측 — 한글 2024 자동화(창 숨김) + HWPX 직접 고치기 → PDF 쪽 수 · 줄 위치.
⚠️ 이 PC 한글 2024 는 구독이 없어 HAction.Run('BreakPage'·'BreakPara'·'MoveDocEnd'·'DeleteBack') 이 모두 False(11/7 실측) —
   글 넣기(InsertText, 줄바꿈 \\r\\n 포함)만 된다. 그래서 쪽 나누기는 HWPX section0.xml 단락의 pageBreak="1" 로 넣는다
   (한글이 Ctrl+Enter 로 저장하는 것과 같은 속성 — 같은 파일을 한글로 다시 열어 PDF 로 저장해 쪽 수 확인).
- 빈 줄(Enter) 몇 개에서 둘째 쪽이 생기나(PageCount 이분 탐색)
- 맨 끝 빈 단락에 쪽 나누기 → 빈 둘째 쪽
- Enter 로 민 글 vs 쪽 나누기로 넘긴 글: 1쪽에 3줄 더 쓰면 2쪽 첫 줄이 무엇·어디(y)
웹 한글(MYBOX)에서 Ctrl+Enter 1/1→2/2쪽, 둘째 쪽 맨 앞 Backspace 2/2→1/1쪽은 shots/w3·w5(11/7).
%TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서 임시 폴더에 저장한다. 실행: python measure.py (pywin32 · pymupdf)"""
import json, os, re, shutil, sys, tempfile, zipfile
import pymupdf
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul05")
os.makedirs(WORK, exist_ok=True)
NOTICE = ["11월 정기 모임 안내", "", "일시: 11월 21일(토) 오후 2시", "장소: 시립 도서관 3층 세미나실", "문의: 총무 — 단체 대화방"]
PAGE2 = "2쪽 시작: 참가 신청서"
MORE = ["추가 1", "추가 2", "추가 3"]

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
R = {"version": hwp.Version, "run_actions": {}}


def ins(t):
    a = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", a.HSet)
    a.Text = t
    return hwp.HAction.Execute("InsertText", a.HSet)


def write(text):
    hwp.Clear(1)
    ins(text)
    return hwp.PageCount


def save(name, fmt):
    p = os.path.join(WORK, name)
    if os.path.exists(p):
        os.remove(p)
    assert hwp.SaveAs(p, fmt, ""), p
    return p


def pdf_info(p):
    with pymupdf.open(p) as d:
        out = []
        for pg in d:
            ws = pg.get_text("words")
            if not ws:
                out.append({"empty": True})
                continue
            top = min(w[1] for w in ws)
            first = " ".join(w[4] for w in sorted([w for w in ws if abs(w[1] - top) < 2], key=lambda w: w[0]))
            out.append({"first_line": first, "y": round(top, 1)})
        return {"pages": d.page_count, "per_page": out}


def with_break(text, para_index_from_end, name):
    """text 를 쓰고 HWPX 로 저장 → 뒤에서 n 번째 단락에 pageBreak="1" → 한글로 다시 열어 PDF."""
    write(text)
    src = save(name + ".hwpx", "HWPX")
    tmp = src + ".tmp"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "Contents/section0.xml":
                s = data.decode("utf-8")
                starts = [m.start() for m in re.finditer(r"<hp:p ", s)]
                i = starts[-para_index_from_end]
                j = s.index(">", i)
                s = s[:i] + s[i:j].replace('pageBreak="0"', 'pageBreak="1"') + s[j:]
                data = s.encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, src)
    assert hwp.Open(src, "HWPX", ""), src
    pc = hwp.PageCount
    return {"page_count_prop": pc, **pdf_info(save(name + ".pdf", "PDF"))}


try:
    for act in ("BreakPage", "BreakPara", "MoveDocEnd", "DeleteBack"):
        R["run_actions"][act] = hwp.HAction.Run(act)
    base = "\r\n".join(NOTICE)
    R["base_pages"] = write(base)
    # 빈 줄 몇 개에서 2쪽?
    lo, hi = 1, 120
    while lo < hi:
        mid = (lo + hi) // 2
        if write(base + "\r\n" * mid) >= 2:
            hi = mid
        else:
            lo = mid + 1
    R["enters_until_page2"] = lo
    write(base + "\r\n" * lo)
    R["enters_page2_pdf"] = pdf_info(save("enters.pdf", "PDF"))
    R["enters_minus1_pages"] = write(base + "\r\n" * (lo - 1))
    # 맨 끝 빈 단락에 쪽 나누기 → 빈 2쪽
    R["break_on_last_empty_para"] = with_break(base + "\r\n", 1, "break_end")
    # Enter 로 민 글 vs 쪽 나누기 — 앞 쪽에 3줄 더하면?
    pad = lo  # 빈 줄 lo 개 뒤 줄이 2쪽 첫 줄(lo-1 개면 그 줄도 1쪽 끝에 남음 — 11/7 실측)
    write(base + "\r\n" * pad + PAGE2)
    R["push_enter_before"] = pdf_info(save("enter_before.pdf", "PDF"))
    write("\r\n".join(NOTICE + MORE) + "\r\n" * pad + PAGE2)
    R["push_enter_after"] = pdf_info(save("enter_after.pdf", "PDF"))
    R["push_break_before"] = with_break(base + "\r\n" + PAGE2, 1, "break_before")
    R["push_break_after"] = with_break("\r\n".join(NOTICE + MORE) + "\r\n" + PAGE2, 1, "break_after")
    hwp.Clear(1)
finally:
    hwp.Quit()
with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
