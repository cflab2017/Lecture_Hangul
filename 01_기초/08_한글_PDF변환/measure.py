"""한글 기초 8강(PDF 변환) — 실습 파일을 만들고 한글 2024(창 숨김)에서 PDF 로 저장한 결과를 잰다.
실습: 가상 '마을 도서관 11월 소식' 2쪽(2쪽은 쪽 나누기로 시작) · 쪽 번호(아래 가운데, PageNumPos — 6강과 같은 방법) · 하이퍼링크 1개(되면).
잰 것: 쪽 수 · 용지 크기 · 글꼴(포함 여부) · 글자 복사(추출) · 쪽 번호 · 링크 · PDF 문서 정보(지은이 등) · 파일 크기.
%TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서 임시 폴더에서 만들고 끝에 이 폴더로 복사한다(6 · 7강과 같음).
결과 measure.json, 파일 실습.hwpx · 정답.pdf. 실행: python measure.py  (pywin32 · pymupdf)"""
import json, os, re, shutil, sys, tempfile, zipfile
import pymupdf
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul08")
os.makedirs(WORK, exist_ok=True)
R = {}

P1 = ["마을 도서관 11월 소식",
      "11월부터 평일 운영 시간이 오후 8시까지로 한 시간 늘어납니다. 주말은 지금처럼 오후 5시까지 엽니다.",
      "11월 22일(토) 오후 2시에는 어린이 열람실에서 그림책 읽어 주기 모임이 열립니다. 신청은 안내 데스크에서 받습니다.",
      "도서관 누리집: https://www.coding-now.com (연습용 주소)",
      "빌린 책은 2주 안에 돌려주세요. 한 번 연장하면 1주를 더 빌릴 수 있습니다."]
P2_HEAD = "이달의 새 책"
BOOKS = [f"{i}. 새 책 {i} — 생활 · 취미 분야 (연습용 목록)" for i in range(1, 13)]


def ins(hwp, t):
    a = hwp.HParameterSet.HInsertText
    hwp.HAction.GetDefault("InsertText", a.HSet)
    a.Text = t
    return hwp.HAction.Execute("InsertText", a.HSet)


def page_num(hwp):
    ps = hwp.HParameterSet.HPageNumPos
    hwp.HAction.GetDefault("PageNumPos", ps.HSet)
    ps.DrawPos = hwp.PageNumPosition("BottomCenter")
    return hwp.HAction.Execute("PageNumPos", ps.HSet)


def hyperlink(hwp):
    """'누리집' 줄 끝에 하이퍼링크(되면). 안 되면 False 를 남긴다."""
    try:
        hl = hwp.HParameterSet.HHyperLink
        hwp.HAction.GetDefault("InsertHyperlink", hl.HSet)
        hl.Text = "코딩나우"
        hl.Command = "https://www.coding-now.com;1;0;0;"
        return bool(hwp.HAction.Execute("InsertHyperlink", hl.HSet))
    except Exception as e:
        return f"err {e}"


def patch_hwpx(path, break_text):
    """작성자 CODINGNOW, break_text 로 시작하는 단락에 pageBreak="1"(6강 measure.patch_hwpx 와 같음)."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith("content.hpf"):
                s = data.decode("utf-8")
                s = re.sub(r'(<opf:meta name="(?:creator|lastsaveby)"[^>]*>)[^<]*(</opf:meta>)', r"\1CODINGNOW\2", s)
                data = s.encode("utf-8")
            if info.filename == "Contents/section0.xml":
                s = data.decode("utf-8")
                k = s.index(f">{break_text}<")
                i = s.rindex("<hp:p ", 0, k)
                j = s.index(">", i)
                s = s[:i] + s[i:j].replace('pageBreak="0"', 'pageBreak="1"') + s[j:]
                data = s.encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, path)


def hpf_meta(path):
    with zipfile.ZipFile(path) as z:
        n = [x for x in z.namelist() if x.endswith("content.hpf")][0]
        s = z.read(n).decode("utf-8")
    return dict(re.findall(r'<opf:meta name="([^"]+)"[^>]*>([^<]*)</opf:meta>', s))


def pdf_info(path, src_text):
    d = pymupdf.open(path)
    fonts = {}
    for pg in d:
        for f in pg.get_fonts(full=True):
            # (xref, ext, type, basefont, name, encoding, referencer)
            fonts[f[3]] = {"type": f[2], "embedded": f[1] not in ("", "n/a"), "ext": f[1]}
    text = "".join(pg.get_text() for pg in d)
    out = {
        "pages": d.page_count,
        "page_size_pt": [round(d[0].rect.width, 1), round(d[0].rect.height, 1)],
        "fonts": fonts,
        "links": [l.get("uri") for pg in d for l in pg.get_links()],
        "metadata": {k: v for k, v in d.metadata.items() if v},
        "bottom_text_per_page": [" ".join(w[4] for w in pg.get_text("words") if w[1] > pg.rect.height - 90) or "(없음)" for pg in d],
        "first_sentence_found": src_text in text.replace("\n", ""),
        "size_bytes": os.path.getsize(path),
    }
    d.close()
    return out


hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
try:
    hwp.Clear(1)
    ins(hwp, "\r\n".join(P1) + "\r\n" + P2_HEAD + "\r\n" + "\r\n".join(BOOKS))
    R["page_num_execute"] = bool(page_num(hwp))
    hwp.MovePos(3, 0, 0)  # 문서 끝(moveBottomOfFile)
    ins(hwp, "\r\n자세한 안내: ")
    R["hyperlink_execute"] = hyperlink(hwp)
    src = os.path.join(WORK, "실습.hwpx")
    if os.path.exists(src):
        os.remove(src)
    assert hwp.SaveAs(src, "HWPX", "")
    hwp.Clear(1)
    patch_hwpx(src, P2_HEAD)
    R["hwpx_meta_after_patch"] = hpf_meta(src)
    # ① HWPX 를 열어 PDF 로 저장(파일 → 다른 이름으로 저장 → PDF 와 같은 저장 형식)
    assert hwp.Open(src, "HWPX", "")
    R["hwp_page_count"] = hwp.PageCount
    pdf1 = os.path.join(WORK, "변환.pdf")
    if os.path.exists(pdf1):
        os.remove(pdf1)
    R["saveas_pdf_ok"] = bool(hwp.SaveAs(pdf1, "PDF", ""))
    R["pdf"] = pdf_info(pdf1, P1[1][:30])
    R["hwpx_size_bytes"] = os.path.getsize(src)
    # ② 같은 문서를 HWP 로 저장한 뒤 PDF — 결과가 같은지
    hwpf = os.path.join(WORK, "실습.hwp")
    if os.path.exists(hwpf):
        os.remove(hwpf)
    hwp.SaveAs(hwpf, "HWP", "")
    R["hwp_size_bytes"] = os.path.getsize(hwpf)
    hwp.Clear(1)
    assert hwp.Open(hwpf, "HWP", "")
    pdf2 = os.path.join(WORK, "변환_hwp.pdf")
    if os.path.exists(pdf2):
        os.remove(pdf2)
    hwp.SaveAs(pdf2, "PDF", "")
    R["pdf_from_hwp"] = {k: v for k, v in pdf_info(pdf2, P1[1][:30]).items() if k in ("pages", "page_size_pt", "metadata", "size_bytes", "links")}
    hwp.Clear(1)
    # ③ 새 문서(저장 안 한 것)에서 바로 PDF — 문서 정보가 어디서 오는지
    ins(hwp, "새 문서에서 바로 PDF 로 저장해 본 줄")
    pdf3 = os.path.join(WORK, "새문서.pdf")
    if os.path.exists(pdf3):
        os.remove(pdf3)
    hwp.SaveAs(pdf3, "PDF", "")
    R["pdf_from_new_doc_metadata"] = pdf_info(pdf3, "새 문서에서")["metadata"]
    hwp.Clear(1)
finally:
    hwp.Quit()

# 저장소용: 실습.hwpx · 정답.pdf(문서 정보 지은이 · 만든 프로그램을 CODINGNOW 로 고친 사본)
shutil.copy(src, os.path.join(HERE, "실습.hwpx"))
ans = os.path.join(HERE, "정답.pdf")
d = pymupdf.open(pdf1)
meta = d.metadata
meta.update({"author": "CODINGNOW", "producer": "CODINGNOW"})
d.set_metadata(meta)
d.save(ans, garbage=3, deflate=True)
d.close()
R["answer_pdf_metadata"] = {k: v for k, v in pymupdf.open(ans).metadata.items() if v}
with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=2, default=str)
print(json.dumps(R, ensure_ascii=False, indent=1, default=str))
