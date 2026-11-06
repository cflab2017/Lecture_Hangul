"""한글 기초 4강(특수문자 · 문자표) 파일 — 한글 2024 자동화(창 숨김)로 실습.hwpx · 정답.hwpx · 기호_비교.pdf 를 만들고 잰다.
내용은 모두 연습용 가상 공지(동아리 11월 모임).
실습: 기호 없이 쓴 공지. 정답: 웹 한글 문자표(Ctrl+F10)에서 찾은 기호(★ ▶ ㎡ ① ② ③ ※ ☎)를 넣은 공지.
기호_비교.pdf: 생김새가 비슷한 기호 짝(가운뎃점 · 와 아래아 ㆍ, ㈜ 와 (주), × 와 x, ① 과 (1))을 같은 글꼴로 놓고 PDF 에서 글자 폭(pt)을 잰다.
이 PC 한글은 편집 화면이 열리지 않아(구독 없음) 문자표 대화상자는 웹 한글에서 재고(charmap_web.json · webhwp.json), 여기서는 글자만 넣는다.
%TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서 임시 폴더에 저장하고 끝에 복사한다. HWPX 작성자(content.hpf)·PDF 메타데이터는 CODINGNOW 로.
실행: python make_files.py  (pip install pywin32 pymupdf)"""
import json, os, re, shutil, sys, tempfile, zipfile
import pymupdf
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul04")
os.makedirs(WORK, exist_ok=True)
PRACTICE = [
    "11월 정기 모임 안내",
    "",
    "일시: 11월 21일(토) 오후 2시",
    "장소: 시립 도서관 3층 세미나실(약 40제곱미터)",
    "준비물: 1. 읽은 책  2. 메모  3. 회비 1만 원",
    "참고: 도서관 주차는 2시간까지 무료입니다.",
    "문의: 총무 — 단체 대화방",
]
ANSWER = [
    "★ 11월 정기 모임 안내 ★",
    "",
    "▶ 일시: 11월 21일(토) 오후 2시",
    "▶ 장소: 시립 도서관 3층 세미나실(약 40㎡)",
    "▶ 준비물: ① 읽은 책  ② 메모  ③ 회비 1만 원",
    "※ 도서관 주차는 2시간까지 무료입니다.",
    "☎ 문의: 총무 — 단체 대화방",
]
# 생김새가 비슷한 짝 — (이름표, 글)
PAIRS = [
    ("가운뎃점 U+00B7", "책·차·예약금"),
    ("아래아 U+318D", "책ㆍ차ㆍ예약금"),
    ("원문자 U+2460", "① 읽은 책"),
    ("괄호 숫자 (1)", "(1) 읽은 책"),
    ("괄호문자 U+321C", "㈜코딩나우"),
    ("괄호 + 주", "(주)코딩나우"),
    ("곱하기 U+00D7", "3×4"),
    ("영문 x", "3x4"),
]

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
VERSION = hwp.Version


def ins(t):
    a = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", a.HSet)
    a.Text = t
    return hwp.HAction.Execute("InsertText", a.HSet)


def save(name, fmt):
    path = os.path.join(WORK, name)
    if os.path.exists(path):
        os.remove(path)
    assert hwp.SaveAs(path, fmt, ""), path
    return path


def clean_hwpx(path):
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


def write(lines):
    hwp.Clear(1)
    ins("\r\n".join(lines))


R = {"version": VERSION}
made = {}
try:
    write(PRACTICE)
    made["실습.hwpx"] = save("실습.hwpx", "HWPX")
    write(ANSWER)
    made["정답.hwpx"] = save("정답.hwpx", "HWPX")
    made["정답.pdf"] = save("정답.pdf", "PDF")
    hwp.Clear(1)
    ins("\r\n".join(f"[{a}] {b}" for a, b in PAIRS))
    made["기호_비교.pdf"] = save("기호_비교.pdf", "PDF")
    hwp.Clear(1)
finally:
    hwp.Quit()

# PDF 에서 글자마다 폭(pt) · 글꼴 — 짝마다 비교
doc = pymupdf.open(made["기호_비교.pdf"])
chars = {}
for b in doc[0].get_text("rawdict")["blocks"]:
    for ln in b.get("lines", []):
        for sp in ln["spans"]:
            for c in sp["chars"]:
                ch = c["c"]
                if ch in "·ㆍ①(1)㈜주×x":
                    w = round(c["bbox"][2] - c["bbox"][0], 2)
                    chars.setdefault(ch, {"U": f"U+{ord(ch):04X}", "width_pt": w, "font": sp["font"]})
doc.close()
R["pair_widths"] = chars
# HWPX 안에 들어간 글자(코드) — 정답
with zipfile.ZipFile(made["정답.hwpx"]) as z:
    sec = z.read("Contents/section0.xml").decode("utf-8")
txt = "".join(re.findall(r"<hp:t>([^<]*)</hp:t>", sec))
R["answer_symbols"] = {ch: f"U+{ord(ch):04X}" for ch in "★▶㎡①②③※☎" if ch in txt}
pdfdoc = pymupdf.open(made["정답.pdf"])
R["answer_pdf_text"] = pdfdoc[0].get_text().strip().splitlines()
pdfdoc.close()

for nm, p in made.items():
    if nm.endswith(".hwpx"):
        clean_hwpx(p)
    else:
        clean_pdf(p)
    if nm != "정답.pdf":
        shutil.copy(p, os.path.join(HERE, nm))
json.dump(R, open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1)[:2500])
