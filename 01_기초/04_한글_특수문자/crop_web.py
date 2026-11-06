"""웹 한글 화면(shots/w_*.png, 1438×1168 CSS px)과 기호_비교.pdf 에서 강의 페이지용 그림을 만들어 coding-now public/lectures/hangul/ 에 넣는다.
화면 오른쪽 위 계정 글자는 잘라 낸다. 실행: python crop_web.py  (pymupdf · Pillow 필요)"""
import os
import pymupdf
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\hangul"
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
DLG = (342, 324, 1098, 844)          # 문자표 대화상자
out = {
    "basic-04-charmap.png": S("w_k1_star.png").crop(DLG),        # 기타 기호 → ★ 한 번 = 입력 문자 ★
    "basic-04-circled.png": S("w_k6_circled.png").crop(DLG),     # 원문자 → ①②③ 차례로 = 입력 문자 ①②③ · 최근 ☎☆★
    "basic-04-unit.png": S("w_k7_unit.png").crop(DLG),           # 단위기호 → ㎡
    "basic-04-araea.png": S("w_k8_araea.png").crop(DLG),         # 호환용 한글 자모 → ㆍ(아래아)
    "basic-04-answer.png": S("w_k10_answer.png").crop((0, 36, 960, 560)),       # 정답 공지 — 메뉴줄·도구 상자가 보이게(11/6 사용자 '너무 일부분만 하면 위치 파악 어려워')
}
# 창 전체 + 빨간 테두리 — 문자표가 어디서 열리는지(도구 상자 '문자표' 단추 · 가운데 창)
win = S("w_k1_star.png").convert("RGB")
dr = ImageDraw.Draw(win)
for b in ((724, 78, 772, 138), (342, 324, 1098, 844)):
    dr.rounded_rectangle(b, radius=8, outline=(220, 38, 38), width=6)
out["basic-04-window.png"] = win.crop((0, 36, 1438, 880))   # 제목 줄(계정 글자) 은 잘라 냄
doc = pymupdf.open(os.path.join(HERE, "기호_비교.pdf"))
pg = doc[0]
k = 72 / 25.4
pix = pg.get_pixmap(dpi=200, clip=pymupdf.Rect(25 * k, 32 * k, 110 * k, 78 * k))
out["basic-04-pairs.png"] = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
doc.close()
os.makedirs(DST, exist_ok=True)
for n, m in out.items():
    m.save(os.path.join(DST, n), optimize=True)
    print(n, m.size)
