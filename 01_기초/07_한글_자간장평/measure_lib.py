"""한글 기초 7강 공용 — 한글 2024 자동화(창 숨김)로 글 넣기 · 저장, HWPX charPr(ratio · spacing) 고치기, PDF 줄 읽기.
편집 동작 HAction.Run 은 이 PC 에서 False(구독 없음, 11/7)라 글자 모양은 HWPX XML 을 직접 고친다.
%TEMP% 밖에 저장하면 '파일 접근 허용' 창이 떠서 임시 폴더(WORK)에 저장한다."""
import os, re, shutil, tempfile, zipfile
import pymupdf
import win32com.client as win32

WORK = os.path.join(os.environ["TEMP"], "claude", "hangul07")
os.makedirs(WORK, exist_ok=True)
LANGS = ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")
BASE_TEXT = ("이번 달 독서 모임은 셋째 주 토요일 오후 두 시에 마을 도서관 세미나실에서 열립니다. 각자 읽은 책에서 기억에 남는 문장을 하나씩 골라 와 주시고, "
             "회비는 모임 날 현장에서 걷습니다. 처음 오시는 분도 편하게 오셔서 이야기를 들으셔도 됩니다.")

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


def make(name, paras):
    hwp.Clear(1)
    ins("\r\n".join(paras))
    return save(name, "HWPX")


def _char_pr(base, nid, sh):
    c = base.replace('<hh:charPr id="0" ', f'<hh:charPr id="{nid}" ', 1)
    for tag in ("ratio", "spacing"):
        if tag in sh:
            vals = sh[tag]
            attrs = " ".join(f'{l}="{vals.get(l, 100 if tag == "ratio" else 0)}"' for l in LANGS)
            c = re.sub(rf"<hh:{tag} [^>]*/>", f"<hh:{tag} {attrs}/>", c, count=1)
    return c


def _rewrite(path, fn, author):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".hwpx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        hdr = zin.read("Contents/header.xml").decode("utf-8")
        sec = zin.read("Contents/section0.xml").decode("utf-8")
        hdr, sec, out = fn(hdr, sec)
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "Contents/header.xml":
                data = hdr.encode("utf-8")
            elif info.filename == "Contents/section0.xml":
                data = sec.encode("utf-8")
            elif author and info.filename.endswith("content.hpf"):
                s = data.decode("utf-8")
                s = re.sub(r'(<opf:meta name="(?:creator|lastsaveby)"[^>]*>)[^<]*(</opf:meta>)', r"\1CODINGNOW\2", s)
                data = s.encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, path)
    return out


def _add_shapes(hdr, shapes):
    cnt = int(re.search(r'<hh:charProperties itemCnt="(\d+)"', hdr).group(1))
    j = hdr.index('<hh:charPr id="0" ')
    base = hdr[j:hdr.index("</hh:charPr>", j) + len("</hh:charPr>")]
    ids, add = [], ""
    for k, sh in enumerate(shapes):
        ids.append(cnt + k)
        add += _char_pr(base, cnt + k, sh)
    e = hdr.index("</hh:charProperties>")
    hdr = hdr[:e] + add + hdr[e:]
    hdr = hdr.replace(f'<hh:charProperties itemCnt="{cnt}"', f'<hh:charProperties itemCnt="{cnt + len(shapes)}"', 1)
    return hdr, ids


def _set_para(sec, k, nid):
    i = sec.rindex("<hp:p ", 0, k)
    pe = sec.index("</hp:p>", k)
    return sec[:i] + re.sub(r'charPrIDRef="0"', f'charPrIDRef="{nid}"', sec[i:pe]) + sec[pe:]


def patch(path, shapes, author=True):
    """shapes: {문단 맨 앞 글자: {"ratio": {...}, "spacing": {...}}} — 그 글자로 시작하는 첫 문단의 글자 모양을 바꾼다."""
    def fn(hdr, sec):
        hdr, ids = _add_shapes(hdr, list(shapes.values()))
        for mark, nid in zip(shapes, ids):
            sec = _set_para(sec, sec.index(f"<hp:t>{mark}"), nid)
        return hdr, sec, dict(zip(shapes, ids))
    return _rewrite(path, fn, author)


def patch_nth(path, mark, shapes, author=False):
    """shapes: {n(1부터): {...}} — mark 로 시작하는 n 번째 문단의 글자 모양을 바꾼다."""
    def fn(hdr, sec):
        hdr, ids = _add_shapes(hdr, list(shapes.values()))
        pos, k = [], -1
        while True:
            k = sec.find(f"<hp:t>{mark}", k + 1)
            if k < 0:
                break
            pos.append(k)
        for n, nid in sorted(zip(shapes, ids), reverse=True):  # 뒤에서부터(앞 위치가 안 바뀌게)
            sec = _set_para(sec, pos[n - 1], nid)
        return hdr, sec, dict(zip(shapes, ids))
    return _rewrite(path, fn, author)


def rows_of(pdf):
    """PDF 의 줄(위에서 아래 순서) — 글자 · x 범위 · 낱말."""
    rows = []
    with pymupdf.open(pdf) as d:
        for pg in d:
            by = {}
            for w in pg.get_text("words"):
                by.setdefault((w[5], w[6]), []).append(w)
            for key in sorted(by):
                ws = sorted(by[key], key=lambda w: w[0])
                rows.append({"text": " ".join(w[4] for w in ws), "x0": round(ws[0][0], 1), "x1": round(ws[-1][2], 1), "words": ws})
    return rows


def paras_starting(pdf, start):
    """start 로 시작하는 줄마다 새 문단 — 문단별 줄 글자 목록. 빈 줄은 PDF 에 안 나온다."""
    groups = []
    for r in rows_of(pdf):
        if r["text"].startswith(start):
            groups.append([])
        if groups:
            groups[-1].append(r["text"])
    return groups
