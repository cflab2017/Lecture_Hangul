"""강의 페이지 그림 — ① 줄간격_비교.pdf 를 그림으로(200dpi) ② 웹 한글(MYBOX) 문단 모양 창(Alt+T, 1강 시험 때 찍은 화면) 자르기.
결과는 coding-now public/lectures/hangul/ 에. 웹 한글 화면은 제목 표시줄(문서 이름·계정)을 잘라 낸다."""
import os
import pymupdf
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\hangul"
doc = pymupdf.open(os.path.join(HERE, "줄간격_비교.pdf"))
pix = doc[0].get_pixmap(dpi=200, clip=pymupdf.Rect(78, 96, 520, 625))
pix.save(os.path.join(DST, "basic-02-compare.png")); print("basic-02-compare.png", pix.width, pix.height)
web = Image.open(os.path.join(HERE, "shots", "web_alt_t.png"))
for name, box in (("basic-02-web-dialog.png", (418, 350, 1020, 818)), ("basic-02-web-toolbar.png", (720, 158, 990, 187))):
    im = web.crop(box); im.save(os.path.join(DST, name), optimize=True); print(name, im.size)
