"""들여쓰기_비교.pdf 와 shots/(웹 한글 화면)에서 강의 페이지용 그림을 만들어 coding-now public/lectures/hangul/ 에 넣는다."""
import os
import pymupdf
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\hangul"
S = lambda n: Image.open(os.path.join(HERE, "shots", n)).convert("RGB")
k = 72 / 25.4
font = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 26)
out = {}
# ① 비교 쪽 — 한글 2024 가 저장한 PDF 그대로(본문 폭 + 양옆 조금)
doc = pymupdf.open(os.path.join(HERE, "들여쓰기_비교.pdf"))
pg = doc[0]
ws = pg.get_text("words")
y0 = min(w[1] for w in ws) / k - 3; y1 = max(w[3] for w in ws) / k + 3
pix = pg.get_pixmap(dpi=170, clip=pymupdf.Rect(26 * k, y0 * k, 184 * k, y1 * k))
out["basic-03-compare.png"] = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
# ② 문단 모양 — 들여쓰기 10pt(정답 첫 문단) · ③ Shift+Tab 뒤 내어쓰기 13.7pt
out["basic-03-dialog.png"] = S("f_dlg_indent.png")
out["basic-03-dialog-shifttab.png"] = S("f_dlg_shifttab.png")
# ④ Shift+Tab 전·뒤 — 번호 항목 1·2 만
a = S("f_practice_before.png"); b = S("f_practice_shifttab.png")
box = (0, 286, a.width, 440)
ta, tb = a.crop(box), b.crop(box)
W, H = ta.size
im = Image.new("RGB", (W, (H + 46) * 2 + 10), "white"); d = ImageDraw.Draw(im)
for i, (lab, t) in enumerate((("Shift+Tab 전 — 둘째 줄이 '1.' 아래", ta), ("'신청' 앞에서 Shift+Tab 뒤 — 둘째 줄이 '신청' 아래", tb))):
    y = i * (H + 56)
    d.text((4, y + 4), lab, fill=(71, 85, 105), font=font)
    im.paste(t, (0, y + 40)); d.rectangle([0, y + 40, W - 1, y + 40 + H - 1], outline=(148, 163, 184), width=2)
out["basic-03-shifttab.png"] = im
for n, m in out.items():
    m.save(os.path.join(DST, n), optimize=True); print(n, m.size)
