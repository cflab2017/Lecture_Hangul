"""한글 기초 5강(쪽 나누기 · 빈 페이지) 실습 · 정답 파일 — 한글 2024 자동화(창 숨김) + HWPX 단락 pageBreak 속성. 내용은 연습용 가상 공지.
실습.hwpx: 1쪽 공지 뒤 빈 줄(Enter) 37개로 '참가 신청서'를 2쪽에 밀어 놓고, 맨 끝에도 빈 줄 40개 → 빈 3쪽.
           (1쪽에 한 줄만 더 써도 신청서 제목이 2쪽 아래로 밀리고, 3쪽은 빈 쪽)
정답.hwpx: 빈 줄을 지우고 '참가 신청서' 단락 앞에 쪽 나누기(Ctrl+Enter = pageBreak="1") — 2쪽, 제목은 늘 2쪽 맨 위.
⚠️ 이 PC 한글은 구독이 없어 HAction.Run 편집 동작이 False 라(measure.py) 쪽 나누기는 HWPX XML 로 넣고 한글로 다시 열어 저장한다.
HWPX 작성자(content.hpf)·PDF 메타데이터는 CODINGNOW. %TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서 임시 폴더에 저장 후 복사.
실행: python make_files.py  (pywin32 · pymupdf)"""
import json, os, re, shutil, sys, tempfile, zipfile
import pymupdf
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul05")
os.makedirs(WORK, exist_ok=True)
NOTICE = ["11월 정기 모임 안내", "", "일시: 11월 21일(토) 오후 2시", "장소: 시립 도서관 3층 세미나실", "문의: 총무 — 단체 대화방"]
FORM = ["참가 신청서", "", "이름:", "연락처:", "참석 여부:  참석 / 불참"]
ENTERS = 37  # measure.json enters_until_page2 — 이만큼 빈 줄 뒤 줄이 2쪽 첫 줄
TAIL = 40

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False


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


def pages(pdf):
    with pymupdf.open(pdf) as d:
        return d.page_count, [(" ".join(w[4] for w in pg.get_text("words")[:3]) or "(빈 쪽)") for pg in d]


R = {}
try:
    # 실습
    hwp.Clear(1)
    ins("\r\n".join(NOTICE) + "\r\n" * ENTERS + "\r\n".join(FORM) + "\r\n" * TAIL)
    p = save("실습.hwpx", "HWPX")
    patch_hwpx(p)
    R["실습"] = pages(save("실습_확인.pdf", "PDF"))
    # 정답
    hwp.Clear(1)
    ins("\r\n".join(NOTICE) + "\r\n" + "\r\n".join(FORM))
    a = save("정답.hwpx", "HWPX")
    patch_hwpx(a, break_text=FORM[0])
    assert hwp.Open(a, "HWPX", ""), a
    a2 = save("정답_다시.hwpx", "HWPX")  # 한글이 다시 저장해도 쪽 나누기가 남는지
    with zipfile.ZipFile(a2) as z:
        R["정답_다시저장_pageBreak"] = len(re.findall(r'pageBreak="1"', z.read("Contents/section0.xml").decode("utf-8")))
    pdf = save("정답.pdf", "PDF")
    R["정답"] = pages(pdf)
    hwp.Clear(1)
finally:
    hwp.Quit()
clean_pdf(pdf)
for src, dst in ((p, "실습.hwpx"), (a, "정답.hwpx"), (pdf, "정답.pdf")):
    shutil.copy(src, os.path.join(HERE, dst))
with zipfile.ZipFile(os.path.join(HERE, "정답.hwpx")) as z:
    R["hpf_creator"] = re.findall(r'name="(?:creator|lastsaveby)"[^>]*>([^<]*)<', z.read("Contents/content.hpf").decode("utf-8"))
print(json.dumps(R, ensure_ascii=False, indent=1))
