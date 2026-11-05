"""한글 기초 2강(줄간격) 파일을 만들고 줄 사이 거리를 잰다 — 한글 2024 자동화(창 숨김). 내용은 모두 연습용 가상 안내문.
① 줄간격_비교.pdf — 같은 문단을 글자에 따라 160%(기본) · 100% · 200% · 0% · 고정 값 20pt · 고정 값 7pt · 최소 7pt 로 놓은 비교 쪽
   → PDF 글자 위치(pymupdf)로 줄 윗부분 사이 거리(pt)를 재서 measure.json 에 적는다.
② 실습.hwpx — 본문이 100%(답답함) · 한 문단은 고정 값 7pt(글자 겹침) ③ 정답.hwpx — 본문 160%, 마지막 알림 문단 200%
이 PC 한글은 편집 화면이 열리지 않아(구독 없음) HAction.Execute 만 쓴다. %TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서
임시 폴더에 저장하고 끝에 복사한다. 저장 뒤 HWPX 작성자(content.hpf)·PDF 메타데이터는 CODINGNOW 로.
실행: python make_files.py  (pip install pywin32 pymupdf)"""
import json, os, re, shutil, sys, tempfile, zipfile
import pymupdf
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul02"); os.makedirs(WORK, exist_ok=True)
SAMPLE = ("11월 15일(토) 오전 9시부터 12시까지 지하 주차장 바닥 도색 작업이 있습니다. 작업 시간에는 지하 주차장을 "
          "쓸 수 없으니 차량을 지상 주차장이나 단지 앞 공영 주차장으로 옮겨 주시기 바랍니다.")
# (이름표, 종류, % 또는 pt) — 종류 0 = 글자에 따라(%), 1 = 고정 값, 2 = 여백만 지정, 3 = 최소
# 고정 계열을 자동화(HParaShape.LineSpacing)로 넣으면 저장 값이 절반이 됐다(2000 → HWPX 1000 → PDF 줄 간격 10pt, 11/5 실측) — pt × 200 으로 넣는다
CASES = [
    ("글자에 따라 160% (기본)", 0, 160),
    ("글자에 따라 100%", 0, 100),
    ("글자에 따라 200%", 0, 200),
    ("글자에 따라 0%", 0, 0),
    ("고정 값 20pt", 1, 20),
    ("고정 값 7pt (글자 10pt 보다 작게)", 1, 7),
    ("최소 7pt", 3, 7),
    ("여백만 지정 0pt", 2, 0),
]
TITLE = "지하 주차장 바닥 도색 작업 안내"
BODY = [
    SAMPLE,
    "작업 뒤 바닥이 마르는 동안에는 냄새가 날 수 있습니다. 지하 1층과 연결된 승강기 문은 작업 시간 동안 열리지 않도록 잠급니다.",
    "작업에 대해 궁금한 점은 관리사무소(내선 100)로 문의해 주시기 바랍니다. 입주민 여러분의 양해를 부탁드립니다.",
    "※ 지하 주차장에 남아 있는 차량은 작업에 따라 다른 곳으로 옮겨질 수 있습니다.",
]

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
VERSION = hwp.Version


def ins(t):
    a = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", a.HSet)
    a.Text = t
    return hwp.HAction.Execute("InsertText", a.HSet)


def spacing(kind, value):
    """kind 0 이면 value 는 %, 아니면 pt."""
    ps = hwp.HParameterSet.HParaShape
    hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
    ps.LineSpacingType = kind
    ps.LineSpacing = value if kind == 0 else int(value * 200)
    return hwp.HAction.Execute("ParagraphShape", ps.HSet)


def save(name, fmt):
    path = os.path.join(WORK, name)
    if os.path.exists(path):
        os.remove(path)
    assert hwp.SaveAs(path, fmt, ""), path


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
    doc.save(tmp, garbage=3, deflate=True); doc.close()
    shutil.move(tmp, path)


def hwpx_spacing(path):
    """HWPX 의 문단 모양(paraPr)마다 줄 간격 종류·값 — 저장된 값이 넣은 값과 같은지 확인용."""
    with zipfile.ZipFile(path) as z:
        h = z.read("Contents/header.xml").decode("utf-8")
    return re.findall(r'<hh:lineSpacing type="(\w+)" value="(-?\d+)" unit="(\w+)"', h)


R = {"version": VERSION}
try:
    # ① 비교 쪽
    hwp.Clear(1)
    for label, kind, value in CASES:
        spacing(0, 160)
        ins(f"[{label}]\r\n")
        assert spacing(kind, value), label
        ins(SAMPLE + "\r\n")
        spacing(0, 160); ins("\r\n")  # 빈 문단 — 겹친 줄이 다음 이름표를 덮지 않게
    spacing(0, 160)
    save("줄간격_비교.hwpx", "HWPX")
    R["hwpx_lineSpacing"] = hwpx_spacing(os.path.join(WORK, "줄간격_비교.hwpx"))
    save("줄간격_비교.pdf", "PDF")
    # ② 실습 — 본문 100%, 마지막 알림 문단 고정 값 7pt
    hwp.Clear(1)
    spacing(0, 160); ins(TITLE + "\r\n")
    for i, t in enumerate(BODY):
        spacing(*((1, 7) if i == len(BODY) - 1 else (0, 100)))
        ins(t + "\r\n")
    save("실습.hwpx", "HWPX"); save("실습.pdf", "PDF")
    # ③ 정답 — 본문 160%, 알림 문단 200%
    hwp.Clear(1)
    spacing(0, 160); ins(TITLE + "\r\n")
    for i, t in enumerate(BODY):
        spacing(*((0, 200) if i == len(BODY) - 1 else (0, 160)))
        ins(t + "\r\n")
    save("정답.hwpx", "HWPX"); save("정답.pdf", "PDF")
    hwp.Clear(1)
finally:
    hwp.Quit()


def line_tops(pdf_path):
    """쪽의 글 줄마다 (윗부분 y, 아랫부분 y, 글) — 같은 y 의 조각은 한 줄로 합친다."""
    doc = pymupdf.open(pdf_path); page = doc[0]
    rows = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            text = "".join(s["text"] for s in ln["spans"]).strip()
            if text:
                rows.append((round(ln["bbox"][1], 2), round(ln["bbox"][3], 2), text))
    doc.close()
    rows.sort()
    merged = []
    for y0, y1, t in rows:
        if merged and abs(merged[-1][0] - y0) < 0.5:
            merged[-1] = (merged[-1][0], max(merged[-1][1], y1), merged[-1][2] + t)
        else:
            merged.append((y0, y1, t))
    return merged


lines = line_tops(os.path.join(WORK, "줄간격_비교.pdf"))
cases, cur = {}, None
for y0, y1, t in lines:
    m = re.match(r"\[(.+?)\]", t)
    if m:
        cur = m.group(1); cases[cur] = {"lines": []}
    elif cur:
        cases[cur]["lines"].append((y0, y1, t[:12]))
for k, v in cases.items():
    tops = [a for a, _, _ in v["lines"]]
    v["line_count"] = len(tops)
    v["pitch_pt"] = [round(b - a, 2) for a, b in zip(tops, tops[1:])]
    v["glyph_height_pt"] = [round(b - a, 2) for a, b, _ in v["lines"]][:1]
R["cases"] = cases

for f in ("줄간격_비교.hwpx", "실습.hwpx", "정답.hwpx"):
    clean_hwpx(os.path.join(WORK, f))
for f in ("줄간격_비교.pdf", "실습.pdf", "정답.pdf"):
    clean_pdf(os.path.join(WORK, f))
for f in ("줄간격_비교.pdf", "실습.hwpx", "정답.hwpx"):
    shutil.copy2(os.path.join(WORK, f), os.path.join(HERE, f))
for f in ("실습.pdf", "정답.pdf"):  # 화면용(저장소에는 올리지 않음)
    os.makedirs(os.path.join(HERE, "shots"), exist_ok=True)
    shutil.copy2(os.path.join(WORK, f), os.path.join(HERE, "shots", f))
json.dump(R, open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({"version": VERSION, "hwpx": R["hwpx_lineSpacing"], **{k: (v["line_count"], v["pitch_pt"], v["glyph_height_pt"]) for k, v in cases.items()}}, ensure_ascii=False, indent=1))
