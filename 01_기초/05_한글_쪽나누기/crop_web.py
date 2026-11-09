"""shots/ 원본(웹 한글 MYBOX)과 measure.py 의 PDF 로 강의 페이지용 그림을 만든다.
11/6 규칙: 웹 한글 화면은 창 전체(제목 줄의 계정 글자만 잘라 냄) + 빨간 테두리, 그 뒤 확대.
결과는 coding-now public/lectures/hangul/basic-05-*.png."""
import os
import pymupdf
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\hangul"
WORK = os.path.join(os.environ["TEMP"], "claude", "hangul05")
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


# ① 쪽 탭 → 쪽 나누기 Ctrl+Enter (창 위쪽 — 메뉴가 펼쳐진 부분까지)
im = boxed(S("w1_page_tab.png"), [(342, 40, 390, 66), (345, 197, 582, 227)])
save(im.crop((0, TOP, 1438, 560)), "basic-05-page-menu.png")
# ② Ctrl+Enter 뒤 2/2쪽 — 창 전체 + 상태 표시줄
im = boxed(S("w3_after_ctrl_enter.png"), [(8, 1144, 66, 1167)])
save(im.crop((0, TOP, 1438, 1168)), "basic-05-ctrl-enter.png")
# ③ 실습 파일 — 문단 부호 켠 빈 3쪽(↵ 줄) · 문단 부호 단추 · 3/3쪽
im = boxed(S("p6_para_marks.png"), [(1060, 74, 1102, 152), (404, 1032, 436, 1118), (8, 1144, 66, 1167)])
save(im.crop((0, TOP, 1438, 1168)), "basic-05-blank-page.png")
# ④ 상태 표시줄 확대 — 빈 줄 지우기 전 3/3쪽 → 지운 뒤 2/2쪽
a = S("p6_para_marks.png").crop((0, 1142, 360, 1168)).resize((720, 52))
b = S("p7_after_bs40.png").crop((0, 1142, 360, 1168)).resize((720, 52))
out = Image.new("RGB", (720, 52 * 2 + 70), "white")
d = ImageDraw.Draw(out)
d.text((8, 4), "빈 줄 지우기 전", fill=(60, 60, 60), font=FONT)
out.paste(a, (0, 34))
d.text((8, 34 + 52 + 4), "Backspace 로 빈 줄을 지운 뒤", fill=(60, 60, 60), font=FONT)
out.paste(b, (0, 34 + 52 + 34))
out = out.crop((0, 0, 720, 34 + 52 + 34 + 52))
boxed(out, [(4, 34, 110, 86), (4, 120, 110, 172)])
save(out, "basic-05-status-zoom.png")
# ⑤ PDF 2쪽 비교 — Enter 로 민 글 vs 쪽 나누기(1쪽에 3줄 더 쓴 뒤)
def page2(pdf):
    with pymupdf.open(os.path.join(WORK, pdf)) as doc:
        pix = doc[1].get_pixmap(dpi=150, clip=pymupdf.Rect(40, 60, 330, 175))  # 2쪽 위쪽 — 본문 시작 줄(y 99.4)과 밀린 줄(147.3)이 보이게
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


e, k = page2("enter_after.pdf"), page2("break_after.pdf")
W, H = e.width, e.height
cmp_ = Image.new("RGB", (W * 2 + 30, H + 44), (235, 238, 242))
d = ImageDraw.Draw(cmp_)
d.text((10, 8), "Enter 로 민 경우 — 3줄 아래로", fill=RED, font=FONT)
d.text((W + 40, 8), "쪽 나누기 — 맨 위 그대로", fill=(20, 110, 60), font=FONT)
cmp_.paste(e, (0, 44))
cmp_.paste(k, (W + 30, 44))
# 본문이 시작하는 줄 높이(PDF y 99.4pt)에 점선 — 두 그림 모두
yy = 44 + int((99.4 - 60) * 150 / 72)
for x in range(0, cmp_.width, 12):
    d.line([(x, yy), (x + 6, yy)], fill=(120, 120, 120), width=1)
save(cmp_, "basic-05-push-compare.png")
