"""한글 기초 3강(들여쓰기) 파일을 만들고 줄 시작 위치를 잰다 — 한글 2024 자동화(창 숨김). 내용은 모두 연습용 가상 안내문.
① 들여쓰기_비교.pdf — 같은 문단을 첫 줄 보통 · 들여쓰기 10pt · 20pt · 내어쓰기 20pt · 띄어쓰기 두 칸 · 왼쪽 여백 20pt · 번호 항목(내어쓰기 없음/있음)으로 놓은 비교 쪽
   → PDF 글자 위치(pymupdf)로 문단마다 첫 줄과 둘째 줄이 시작하는 x(mm)를 재서 measure.json 에 적는다.
② 실습.hwpx — 들여쓰기 없음(첫 문단은 띄어쓰기 두 칸으로 민 상태) ③ 정답.hwpx — 본문 첫 줄 들여쓰기 10pt, 번호 항목 내어쓰기('1. ' 폭)
이 PC 한글은 편집 화면이 열리지 않아(구독 없음) HAction.Execute 만 쓴다. %TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서
임시 폴더에 저장하고 끝에 복사한다. 저장 뒤 HWPX 작성자(content.hpf)·PDF 메타데이터는 CODINGNOW 로.
실행: python make_files.py  (pip install pywin32 pymupdf)"""
import json, os, re, shutil, sys, tempfile, zipfile
import pymupdf
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul03"); os.makedirs(WORK, exist_ok=True)
MM = 25.4 / 72
SAMPLE = ("11월 사내 교육은 18일(화) 오후 2시부터 4시까지 본관 3층 대회의실에서 열립니다. 이번 교육은 새로 바뀐 출장비 정산 방법을 "
          "다루며, 부서마다 한 사람 이상 참석해 주시기 바랍니다. 노트북을 가져오시면 실습을 함께 할 수 있습니다.")
ITEM = "1. 신청: 참석할 사람은 11월 14일(금)까지 사내 게시판 교육 신청란에 이름과 부서를 적어 주시기 바랍니다. 자리가 차면 신청을 마감합니다."
# (이름표, 첫 줄 종류, 값 pt, 왼쪽 여백 pt, 글) — 첫 줄: 0 보통 · + 들여쓰기 · - 내어쓰기
# 줄 간격 고정 값처럼 자동화로 넣는 pt 값은 × 200 (2강에서 확인) — 여기서도 맞는지 PDF 로 다시 잰다
CASES = [
    ("보통", 0, 0, SAMPLE),
    ("들여쓰기 10pt", 1, 10, SAMPLE),
    ("들여쓰기 20pt", 1, 20, SAMPLE),
    ("내어쓰기 20pt", -1, 20, SAMPLE),
    ("띄어쓰기 두 칸으로 민 첫 줄", 0, 0, "  " + SAMPLE),
    ("왼쪽 여백 20pt", 0, 0, SAMPLE, 20),
    ("번호 항목 — 내어쓰기 없음", 0, 0, ITEM),
]
TITLE = "11월 사내 교육 안내"
BODY = [
    SAMPLE,
    "교육 자료는 당일 회의실 입구에서 나누어 드리고, 끝난 뒤에는 사내 게시판 자료실에도 올립니다. 궁금한 점은 인사팀(내선 210)으로 문의해 주세요.",
]
ITEMS = [
    ITEM,
    "2. 준비물: 출장비 정산에 쓰는 사원증과 법인 카드를 가져오시면 시스템에 직접 접속해 볼 수 있습니다. 노트북이 없으면 회의실 PC 를 씁니다.",
    "3. 문의: 일정이 겹쳐 참석이 어려운 부서는 인사팀에 알려 주시면 12월 교육 일정에 먼저 넣어 드립니다.",
]

hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
VERSION = hwp.Version


def ins(t):
    a = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", a.HSet)
    a.Text = t
    return hwp.HAction.Execute("InsertText", a.HSet)


def shape(first=0, value=0.0, left=0.0):
    """first: 0 보통, 1 들여쓰기, -1 내어쓰기 (value pt) · left: 왼쪽 여백 pt."""
    ps = hwp.HParameterSet.HParaShape
    hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
    ps.Indentation = int(first * value * 200)
    ps.LeftMargin = int(left * 200)
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


def hwpx_margins(path):
    """HWPX 문단 모양(paraPr)의 첫 줄(intent)·왼쪽(left) 값 — 넣은 값이 어떻게 저장됐는지 확인용."""
    with zipfile.ZipFile(path) as z:
        h = z.read("Contents/header.xml").decode("utf-8")
    return [(a, b) for a, b in re.findall(r'<hc:intent value="(-?\d+)"[^/]*/>\s*<hc:left value="(-?\d+)"', h)]


def lines(pdf_path):
    """쪽의 글 줄마다 (y, x0 mm, 글)."""
    doc = pymupdf.open(pdf_path); out = []
    for pno, page in enumerate(doc):
        for b in page.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                spans = [s for s in ln["spans"] if s["text"].strip()]
                if spans:
                    x0 = min(s["bbox"][0] for s in spans)
                    out.append((pno, round(ln["bbox"][1], 2), round(x0 * MM, 2), "".join(s["text"] for s in ln["spans"])))
    doc.close()
    out.sort()
    merged = []
    for p, y, x, t in out:
        if merged and merged[-1][0] == p and abs(merged[-1][1] - y) < 0.5:
            merged[-1] = (p, y, min(merged[-1][2], x), merged[-1][3] + t)
        else:
            merged.append((p, y, x, t))
    return merged


def word_x(pdf_path, word):
    """글자 조각(word) 가 처음 나오는 x(mm) — '신청:' 이 시작하는 자리 = '1. ' 뒤."""
    doc = pymupdf.open(pdf_path)
    for page in doc:
        for w in page.get_text("words"):
            if w[4].startswith(word):
                x = round(w[0] * MM, 2); doc.close(); return x
    doc.close()


R = {"version": VERSION}
try:
    # ① 비교 쪽
    hwp.Clear(1)
    for c in CASES:
        label, first, value, text = c[:4]
        left = c[4] if len(c) > 4 else 0
        shape(); ins(f"[{label}]\r\n")
        assert shape(first, value, left), label
        ins(text + "\r\n")
        shape(); ins("\r\n")
    save("들여쓰기_비교.hwpx", "HWPX"); save("들여쓰기_비교.pdf", "PDF")
    R["hwpx_intent_left"] = hwpx_margins(os.path.join(WORK, "들여쓰기_비교.hwpx"))
finally:
    pass

L = lines(os.path.join(WORK, "들여쓰기_비교.pdf"))
cases, cur = {}, None
for p, y, x, t in L:
    m = re.match(r"\[(.+?)\]", t)
    if m:
        cur = m.group(1); cases[cur] = {"label_x_mm": x, "lines": []}
    elif cur:
        cases[cur]["lines"].append((x, t[:14]))
for k, v in cases.items():
    xs = [x for x, _ in v["lines"]]
    v["first_x_mm"], v["second_x_mm"] = xs[0], (xs[1] if len(xs) > 1 else None)
    v["line_count"] = len(xs)
body_x = cases["보통"]["first_x_mm"]
for k, v in cases.items():
    v["first_from_body_mm"] = round(v["first_x_mm"] - body_x, 2)
    v["second_from_body_mm"] = round(v["second_x_mm"] - body_x, 2) if v["second_x_mm"] is not None else None
item_text_x = word_x(os.path.join(WORK, "들여쓰기_비교.pdf"), "신청:")
R["item_text_after_number_mm"] = round(item_text_x - body_x, 2)
HANG = round(R["item_text_after_number_mm"] / MM, 1)   # '1. ' 폭(pt) — 정답 내어쓰기 값
R["hang_pt"] = HANG

try:
    # ①-2 번호 항목 + 내어쓰기('1. ' 폭) — 같은 비교 쪽 끝에 더해 다시 저장
    hwp.Clear(1)
    for c in CASES + [("번호 항목 — 내어쓰기 '1. ' 폭", -1, HANG, ITEM)]:
        label, first, value, text = c[:4]
        left = c[4] if len(c) > 4 else 0
        shape(); ins(f"[{label}]\r\n")
        assert shape(first, value, left), label
        ins(text + "\r\n")
        shape(); ins("\r\n")
    save("들여쓰기_비교.hwpx", "HWPX"); save("들여쓰기_비교.pdf", "PDF")
    # ② 실습 — 첫 문단은 띄어쓰기 두 칸, 번호 항목은 내어쓰기 없음
    hwp.Clear(1)
    shape(); ins(TITLE + "\r\n")
    ins("  " + BODY[0] + "\r\n"); ins(BODY[1] + "\r\n")
    for t in ITEMS:
        ins(t + "\r\n")
    save("실습.hwpx", "HWPX"); save("실습.pdf", "PDF")
    # ③ 정답 — 본문 첫 줄 들여쓰기 10pt, 번호 항목 내어쓰기 HANG pt
    hwp.Clear(1)
    shape(); ins(TITLE + "\r\n")
    for t in BODY:
        shape(1, 10); ins(t + "\r\n")
    for t in ITEMS:
        shape(-1, HANG); ins(t + "\r\n")
    shape()
    save("정답.hwpx", "HWPX"); save("정답.pdf", "PDF")
    hwp.Clear(1)
finally:
    hwp.Quit()

L = lines(os.path.join(WORK, "들여쓰기_비교.pdf"))
cur = None
for p, y, x, t in L:
    m = re.match(r"\[(.+?)\]", t)
    if m:
        cur = m.group(1)
        cases.setdefault(cur, {"label_x_mm": x, "lines": []})
        if cur.startswith("번호 항목 — 내어쓰기 '1."):
            cases[cur]["lines"] = []
    elif cur and cur.startswith("번호 항목 — 내어쓰기 '1."):
        cases[cur]["lines"].append((x, t[:14]))
v = cases["번호 항목 — 내어쓰기 '1. ' 폭"]
xs = [x for x, _ in v["lines"]]
v.update({"first_x_mm": xs[0], "second_x_mm": xs[1], "line_count": len(xs),
          "first_from_body_mm": round(xs[0] - body_x, 2), "second_from_body_mm": round(xs[1] - body_x, 2)})
# 띄어쓰기로 민 첫 줄은 줄 상자(bbox)가 빈칸부터 잡혀 30.0 이 나온다 — 첫 글자(낱말) 위치로 다시 잰다
doc = pymupdf.open(os.path.join(WORK, "들여쓰기_비교.pdf"))
ws = sorted(doc[0].get_text("words"), key=lambda w: (round(w[1], 1), w[0]))
lab = next(w for w in ws if w[4].startswith("[띄어쓰기"))
first = next(w for w in ws if w[1] > lab[3] - 0.5 and w[4].startswith("11월"))
normal_lab = next(w for w in ws if w[4].startswith("[보통]"))
normal_first = next(w for w in ws if w[1] > normal_lab[3] - 0.5 and w[4].startswith("11월"))
doc.close()
cases["띄어쓰기 두 칸으로 민 첫 줄"]["first_word_from_body_mm"] = round((first[0] - normal_first[0]) * MM, 2)
R["cases"] = cases
R["body_x_mm"] = body_x
R["answer_lines"] = [(round(x - body_x, 2), t[:12]) for p, y, x, t in lines(os.path.join(WORK, "정답.pdf"))]
R["practice_lines"] = [(round(x - body_x, 2), t[:12]) for p, y, x, t in lines(os.path.join(WORK, "실습.pdf"))]

for f in ("들여쓰기_비교.hwpx", "실습.hwpx", "정답.hwpx"):
    clean_hwpx(os.path.join(WORK, f))
for f in ("들여쓰기_비교.pdf", "실습.pdf", "정답.pdf"):
    clean_pdf(os.path.join(WORK, f))
for f in ("들여쓰기_비교.pdf", "실습.hwpx", "정답.hwpx"):
    shutil.copy2(os.path.join(WORK, f), os.path.join(HERE, f))
os.makedirs(os.path.join(HERE, "shots"), exist_ok=True)
for f in ("실습.pdf", "정답.pdf"):  # 화면용(저장소에는 올리지 않음)
    shutil.copy2(os.path.join(WORK, f), os.path.join(HERE, "shots", f))
json.dump(R, open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({"version": VERSION, "hwpx": R["hwpx_intent_left"], "hang_pt": HANG, "item_after_num_mm": R["item_text_after_number_mm"],
                  **{k: (v["line_count"], v["first_from_body_mm"], v["second_from_body_mm"]) for k, v in cases.items()}}, ensure_ascii=False, indent=1))
