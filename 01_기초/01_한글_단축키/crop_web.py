"""shots/*.png(웹 한글 화면 — 네이버 MYBOX 한컴오피스 Web, shortcut_test_webhwp.py 로 찍음)와 단축키표.pdf 를 강의 페이지용 그림으로 자른다.
결과는 인자로 받은 폴더(사이트의 public/lectures/hangul)에 basic-01-*.png. 실행: python crop_web.py <저장 폴더>"""
import os, sys
import pymupdf
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.join(HERE, "shots")
OUT = sys.argv[1]
FONT = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 22)


def labeled(img, text, h=40):
    c = Image.new("RGB", (img.width, img.height + h), "white")
    d = ImageDraw.Draw(c)
    d.rectangle((0, 0, img.width, h), fill=(17, 24, 39))
    d.text((14, h // 2), text, fill="white", font=FONT, anchor="lm")
    c.paste(img, (0, h))
    return c


# ① Alt+L → 글자 모양(하나만 크게)
a = Image.open(os.path.join(S, "k_alt_l.png")).crop((410, 398, 1028, 772)).convert("RGB")
labeled(a, "Alt+L  →  글자 모양").save(os.path.join(OUT, "basic-01-web-altl.png"), optimize=True)

# ② 나머지 넷 — 어떤 창이 뜨는지 한눈에(2×2)
parts = [("k_alt_t.png", (410, 340, 1030, 830), "Alt+T  →  문단 모양"),
         ("k_ctrl_f10.png", (335, 316, 1100, 845), "Ctrl+F10  →  문자표"),
         ("k_f7.png", (435, 360, 1000, 805), "F7  →  편집 용지"),
         ("k_ctrl_h.png", (500, 470, 940, 700), "Ctrl+H · Ctrl+F2  →  찾아 바꾸기")]
W, Hh = 560, 440
tiles = []
for f, box, lab in parts:
    im = Image.open(os.path.join(S, f)).crop(box).convert("RGB")
    im.thumbnail((W, Hh - 40))
    t = Image.new("RGB", (W, Hh - 40), (243, 244, 246)); t.paste(im, ((W - im.width) // 2, (Hh - 40 - im.height) // 2))
    tiles.append(labeled(t, lab))
g = Image.new("RGB", (W * 2 + 12, Hh * 2 + 12), "white")
for i, t in enumerate(tiles):
    g.paste(t, ((i % 2) * (W + 12), (i // 2) * (Hh + 12)))
g.save(os.path.join(OUT, "basic-01-web-grid.png"), optimize=True)

# ③ 웹 한글 도움말 '단축키 일람'의 주의(브라우저 단축키가 먼저)
Image.open(os.path.join(S, "guide_tab.png")).crop((395, 140, 1395, 505)).convert("RGB").save(os.path.join(OUT, "basic-01-web-guide.png"), optimize=True)

# ④ 단축키표.pdf 첫 쪽
doc = pymupdf.open(os.path.join(HERE, "단축키표.pdf"))
pix = doc[0].get_pixmap(dpi=110)
im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
bbox = Image.eval(im.convert("L"), lambda v: 255 if v < 250 else 0).getbbox()
im.crop((max(bbox[0] - 30, 0), max(bbox[1] - 30, 0), min(bbox[2] + 30, im.width), min(bbox[3] + 30, im.height))).save(os.path.join(OUT, "basic-01-cheatsheet.png"), optimize=True)
for f in sorted(os.listdir(OUT)):
    if f.startswith("basic-01"):
        print(f, Image.open(os.path.join(OUT, f)).size)
