"""shots/ 원본(웹 한글 MYBOX)과 measure.py 의 PDF 로 강의 페이지용 그림을 만든다.
11/6 규칙: 웹 한글 화면은 창 전체(제목 줄의 계정 글자만 잘라 냄) + 빨간 테두리, 그 뒤 확대 — 대화상자는 '창 전체(작게) + 확대' 한 장으로.
결과는 coding-now public/lectures/hangul/basic-06-*.png."""
import os
import pymupdf
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\hangul"
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul06")
RED = (220, 30, 30)
TOP = 36  # 웹 한글 제목 줄(문서 이름 · 계정 글자)
FONT = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 22)
SMALL = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 18)


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


def with_zoom(src, boxes, zoom, name):
    """왼쪽: 창 전체(작게, 확대한 곳은 빨간 틀) · 오른쪽: 그 부분 확대."""
    full = boxed(S(src), boxes).crop((0, TOP, 1438, 1168))
    d = ImageDraw.Draw(full)
    z = (zoom[0], zoom[1] - TOP, zoom[2], zoom[3] - TOP)
    d.rectangle(z, outline=RED, width=6)
    small = full.resize((int(full.width * 0.42), int(full.height * 0.42)))
    big = boxed(S(src), boxes).crop(zoom)
    H = max(small.height, big.height)
    out = Image.new("RGB", (small.width + big.width + 30, H + 20), (244, 244, 244))
    out.paste(small, (10, 10))
    out.paste(big, (small.width + 20, 10 + (H - big.height) // 2))
    save(out, name)


def strips(parts, name, title=None):
    """쪽 아래쪽 띠 여러 개를 세로로 — (그림 파일, 자를 곳, 이름표)"""
    ims = []
    for src, box, label, mark in parts:
        im = S(src).crop(box)
        ImageDraw.Draw(im).rectangle((mark[0] - 24, mark[1] - 16, mark[0] + 24, mark[1] + 16), outline=RED, width=3)
        im = im.resize((int(im.width * 1.4), int(im.height * 1.4)))
        lab = Image.new("RGB", (im.width, 34), (244, 244, 244))
        ImageDraw.Draw(lab).text((10, 4), label, font=FONT, fill=(30, 30, 30))
        both = Image.new("RGB", (im.width, im.height + 34), "white")
        both.paste(lab, (0, 0)); both.paste(im, (0, 34))
        ims.append(both)
    W = max(i.width for i in ims)
    H = sum(i.height for i in ims) + 12 * (len(ims) + 1)
    out = Image.new("RGB", (W + 24, H), (244, 244, 244))
    y = 12
    for i in ims:
        out.paste(i, (12, y)); y += i.height + 12
    save(out, name)


# ① 쪽 → 꼬리말 ▸ 가운데 쪽 번호 (창 위쪽 — 메뉴까지)
im = boxed(S("w14_footer_menu.png"), [(340, 38, 390, 67), (342, 122, 578, 146), (585, 178, 722, 202)])
save(im.crop((0, TOP, 1438, 560)), "basic-06-footer-menu.png")
# ② 1쪽에서 넣은 뒤 — 1쪽 아래 '1' · 상태 표시줄 1/3쪽
im = boxed(S("w15_page1_number.png"), [(674, 1060, 718, 1092), (8, 1144, 66, 1167)])
save(im.crop((0, TOP, 1438, 1168)), "basic-06-page1-number.png")
# ③ 현재 쪽만 감추기 — 꼬리말 체크 · 안내문
with_zoom("w16_hide_footer_dialog.png", [(634, 520, 706, 552), (530, 625, 836, 690), (836, 496, 916, 526)], (480, 430, 960, 720), "basic-06-hide-footer.png")
# ④ 함정: '쪽 번호'만 체크해 감추기 → 꼬리말로 넣은 3 그대로(3쪽)
dlg = boxed(S("w7_hide_pagenum_checked.png"), [(548, 540, 620, 562)]).crop((480, 430, 960, 720))
bot = boxed(S("w8_after_hide_pagenum.png"), [(674, 1060, 718, 1092)]).crop((380, 980, 1010, 1130))
W = max(dlg.width, bot.width) + 24
out = Image.new("RGB", (W, dlg.height + bot.height + 34 * 2 + 36), (244, 244, 244))
d = ImageDraw.Draw(out)
d.text((12, 6), "① 3쪽에서 '쪽 번호'만 체크 → 설정", font=FONT, fill=(30, 30, 30))
out.paste(dlg, (12, 40))
d.text((12, 40 + dlg.height + 6), "② 꼬리말로 넣은 번호 3 은 그대로", font=FONT, fill=(30, 30, 30))
out.paste(bot, (12, 40 + dlg.height + 40))
save(out, "basic-06-trap-pagenum-only.png")
# ⑤ 함정: 커서가 3쪽일 때 꼬리말 쪽 번호 → 1 · 2쪽엔 번호 없음
strips([("w10_cover_cursor.png", (380, 980, 1010, 1128), "1쪽 아래 — 번호 없음", (316, 95)),
        ("w12_page2_bottom.png", (380, 250, 1010, 420), "2쪽 아래 — 번호 없음", (316, 116)),
        ("w4_fit_page.png", (380, 1000, 1010, 1128), "3쪽 아래 — 3", (316, 77))], "basic-06-trap-cursor.png")
# ⑥ 새 번호로 시작 — 쪽 번호 · 시작 번호 1 · 넣기
with_zoom("w19_newnum_dialog.png", [(548, 553, 652, 578), (534, 616, 720, 650), (830, 525, 912, 556)], (480, 440, 960, 700), "basic-06-new-number.png")
# ⑦ 결과: 1쪽 없음 · 2쪽 1 · 3쪽 2
strips([("w17_cover_hidden.png", (380, 980, 1010, 1128), "1쪽(표지) — 감춤", (316, 95)),
        ("w20_after_newnum.png", (380, 250, 1010, 420), "2쪽 — 1", (316, 116)),
        ("w21_page3_bottom.png", (380, 1000, 1010, 1128), "3쪽 — 2", (316, 78))], "basic-06-result.png")
# ⑧ 정답 PDF(한글 2024) 세 쪽
pages = []
with pymupdf.open(os.path.join(HERE, "정답.pdf")) as doc:
    for k, pg in enumerate(doc):
        p = os.path.join(WORK, f"answer_p{k + 1}.png")
        pg.get_pixmap(dpi=45).save(p)
        pages.append(Image.open(p).convert("RGB"))
gap = 18
out = Image.new("RGB", (sum(i.width for i in pages) + gap * (len(pages) + 1), max(i.height for i in pages) + 60), (244, 244, 244))
d = ImageDraw.Draw(out)
x = gap
for k, im in enumerate(pages):
    out.paste(im, (x, 44))
    d.rectangle((x - 1, 43, x + im.width, 44 + im.height), outline=(170, 170, 170))
    d.text((x, 10), f"{k + 1}쪽", font=FONT, fill=(30, 30, 30))
    x += im.width + gap
save(out, "basic-06-answer-pdf.png")
