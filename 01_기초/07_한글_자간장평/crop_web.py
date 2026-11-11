"""shots/ 원본(웹 한글 MYBOX)으로 강의 페이지용 그림을 만든다.
11/6 규칙: 웹 한글 화면은 창 전체(제목 줄의 문서 이름 · 계정 글자만 잘라 냄) + 빨간 테두리, 그 뒤 확대.
결과는 coding-now public/lectures/hangul/basic-07-*.png."""
import os
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\hangul"
RED = (220, 30, 30)
TOP = 36  # 웹 한글 제목 줄(문서 이름 · 계정 글자)
FONT = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 22)


def S(n):
    return Image.open(os.path.join(HERE, "shots", n)).convert("RGB")


def boxed(im, boxes):
    d = ImageDraw.Draw(im)
    for b in boxes:
        d.rectangle(b, outline=RED, width=3)
    return im


def save(im, name):
    im.save(os.path.join(DST, name))
    print(name, im.size)


# 1) 글자 모양 창(Alt+L) — 언어 '대표' · 장평 · 자간 칸, 고른 문단
im = boxed(S("w3_charshape.png"), [(517, 541, 631, 567), (687, 571, 801, 597), (687, 601, 801, 627)])
save(im.crop((0, TOP, im.width, im.height)), "basic-07-charshape.png")

# 2) 자간 좁게 세 번 뒤 창 전체 — 3줄이 된 문단
im = boxed(S("w8_n3.png"), [(405, 412, 990, 500)])
save(im.crop((0, TOP, im.width, im.height)), "basic-07-spacing-window.png")

# 3) 확대 비교 — 기본 4줄 / 자간 -3% 3줄 / 장평 97% 3줄, 오른쪽에 글자 모양 창의 값
rows = [("기본 — 4줄, 마지막 줄에 '요.' 두 글자", "w1_open.png", None),
        ("Alt+Shift+N 세 번 → 자간 -3% — 3줄", "w8_n3.png", "w9_n3_dialog.png"),
        ("Alt+Shift+J 세 번 → 장평 97% — 3줄", "w14_j3.png", "w15_j3_dialog.png")]
PARA = (405, 412, 990, 500)
VAL = (640, 570, 800, 627)
scale = 1.6
pw, ph = int((PARA[2] - PARA[0]) * scale), int((PARA[3] - PARA[1]) * scale)
vw, vh = int((VAL[2] - VAL[0]) * scale), int((VAL[3] - VAL[1]) * scale)
W = pw + vw + 60
H = len(rows) * (ph + 50) + 20
out = Image.new("RGB", (W, H), "white")
d = ImageDraw.Draw(out)
y = 10
for label, doc, dlg in rows:
    d.text((10, y), label, fill=(30, 30, 30), font=FONT)
    y += 36
    p = S(doc).crop(PARA).resize((pw, ph))
    out.paste(p, (10, y))
    d.rectangle((10, y, 10 + pw, y + ph), outline=(200, 200, 200), width=1)
    if dlg:
        v = S(dlg).crop(VAL).resize((vw, vh))
        out.paste(v, (pw + 40, y + (ph - vh) // 2))
        d.rectangle((pw + 40, y + (ph - vh) // 2, pw + 40 + vw, y + (ph - vh) // 2 + vh), outline=RED, width=3)
    y += ph + 14
save(out, "basic-07-compare.png")
