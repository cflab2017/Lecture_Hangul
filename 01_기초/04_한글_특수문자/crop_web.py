"""웹 한글 화면(shots/w_*.png, 1438×1168 CSS px)과 기호_비교.pdf 에서 강의 페이지용 그림을 만들어 coding-now public/lectures/hangul/ 에 넣는다.
화면 오른쪽 위 계정 글자는 잘라 낸다. 실행: python crop_web.py  (pymupdf · Pillow 필요)"""
import os
import pymupdf
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\hangul"
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
DLG = (342, 324, 1098, 844)          # 문자표 대화상자
out = {
    "basic-04-charmap.png": S("w_k1_star.png").crop(DLG),        # 기타 기호 → ★ 한 번 = 입력 문자 ★
    "basic-04-circled.png": S("w_k6_circled.png").crop(DLG),     # 원문자 → ①②③ 차례로 = 입력 문자 ①②③ · 최근 ☎☆★
    "basic-04-unit.png": S("w_k7_unit.png").crop(DLG),           # 단위기호 → ㎡
    "basic-04-araea.png": S("w_k8_araea.png").crop(DLG),         # 호환용 한글 자모 → ㆍ(아래아)
    "basic-04-answer.png": S("w_k10_answer.png").crop((300, 330, 1000, 540)),  # 정답 공지 — 웹 한글에서 연 화면
}
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
