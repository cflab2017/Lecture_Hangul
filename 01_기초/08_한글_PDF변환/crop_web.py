"""강의 페이지용 그림 — 한글 2024 가 만든 정답.pdf 의 쪽을 그대로 그림으로(웹 한글 화면 대신, 11/12: 9225 크롬이 꺼져 있었고
바탕화면을 다른 세션이 쓰는 중이라 새 창을 띄우지 않음). 1쪽 위쪽 본문 + 아래 쪽 번호, 2쪽 본문 끝(링크) + 쪽 번호에 빨간 테두리.
결과 coding-now public/lectures/hangul/basic-08-pdf-page1.png · page2.png. 실행: python crop_web.py (pymupdf · pillow)"""
import io, os
import pymupdf
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\content_hub\homepage\public\lectures\hangul"
Z = 1.6  # 1pt → 1.6px


def page_png(pg):
    pix = pg.get_pixmap(matrix=pymupdf.Matrix(Z, Z))
    return Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")


def box(im, rect, pad=5):
    x0, y0, x1, y1 = (v * Z for v in rect)
    ImageDraw.Draw(im).rectangle((x0 - pad, y0 - pad, x1 + pad, y1 + pad), outline=(220, 30, 30), width=3)


def union(ws):
    return (min(w[0] for w in ws), min(w[1] for w in ws), max(w[2] for w in ws), max(w[3] for w in ws))


def make(pg, name, body_frac=None):
    im = page_png(pg)
    foot = [w for w in pg.get_text("words") if w[1] > pg.rect.height - 90]
    if foot:
        box(im, union(foot))
    for l in pg.get_links():
        box(im, tuple(l["from"]))
    w, h = im.size
    body = [b for b in pg.get_text("blocks") if b[1] < pg.rect.height - 90]
    cut = int((max(b[3] for b in body) + 30) * Z) if body_frac is None else int(h * body_frac)
    top, bot = im.crop((0, 0, w, cut)), im.crop((0, h - int(110 * Z), w, h))
    out = Image.new("RGB", (w, top.height + bot.height + 24), "white")
    out.paste(top, (0, 0))
    ImageDraw.Draw(out).line((0, top.height + 12, w, top.height + 12), fill=(180, 180, 180), width=2)
    out.paste(bot, (0, top.height + 24))
    out = out.resize((out.width * 3 // 4, out.height * 3 // 4))
    out.save(os.path.join(DST, name))
    print(name, out.size)


doc = pymupdf.open(os.path.join(HERE, "정답.pdf"))
make(doc[0], "basic-08-pdf-page1.png")
make(doc[1], "basic-08-pdf-page2.png")
