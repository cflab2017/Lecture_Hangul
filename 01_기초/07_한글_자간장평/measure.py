"""한글 기초 7강(자간 · 장평) 실측 + 실습 · 정답 파일 — 한글 2024 자동화(창 숨김)로 글을 넣고, 글자 모양은 HWPX 의 charPr(ratio · spacing)를
직접 고쳐(편집 동작 HAction.Run 은 이 PC 에서 False — 구독 없음, 11/7) 한글로 다시 열어 PDF 로 저장한 뒤 줄 수 · 글자 폭을 잰다(measure_lib.py).
문단은 calib.py 로 고른 '가을 독서 모임 안내' 본문 — 기본(글자 10pt)이면 4줄이고 마지막 줄에 '요.' 두 글자만 남는다.
- 같은 문단: 기본 · 자간 -5 · -10 · +10 · 장평 95 · 90 — 줄 수 · 마지막 줄
- 몇 %부터 한 줄이 줄어드나: 자간 -1 ~ -5 · 장평 99 ~ 95
- 한글 · 영문이 섞인 문장: 기본 · 한글만 자간 -15 · 모든 언어 자간 -15 — 낱말 폭(pt)
실습.hwpx = 기본, 정답.hwpx = 안내 문단에 자간 -3%(4줄 → 3줄, Alt+Shift+N 세 번과 같음). 결과 measure.json. 실행: python measure.py  (pywin32 · pymupdf)"""
import json, os, shutil, sys
import measure_lib as M

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
TEXT = M.BASE_TEXT + " 늦으시는 분은 미리 연락해 주세요."
MIXED = "모임 공지 Book Club Notice 회비 안내 Monthly Fee"
ALL = lambda v: {l: v for l in M.LANGS}  # noqa: E731
R = {"version": M.hwp.Version, "text": TEXT}


def sweep(name, cases):
    """cases: [(이름, 글자 모양 또는 None)] — 같은 문단을 여러 벌 넣고 문단마다 줄 수 · 마지막 줄."""
    paras = []
    for _ in cases:
        paras += [TEXT, ""]
    p = M.make(name + ".hwpx", paras[:-1])
    M.patch(p, {}, author=False)
    shapes = {k + 1: sh for k, (_, sh) in enumerate(cases) if sh}
    if shapes:
        M.patch_nth(p, "이번 달", shapes)
    assert M.hwp.Open(p, "HWPX", "")
    g = M.paras_starting(M.save(name + ".pdf", "PDF"), "이번 달")
    return [{"case": c, "lines": len(x), "last_line": x[-1], "last_chars": len(x[-1].replace(" ", "")), "first_line_end": x[0][-12:]}
            for (c, _), x in zip(cases, g)]


def width(rows, word, after):
    seen = False
    for r in rows:
        if r["text"].startswith(after):
            seen = True
        if seen:
            for w in r["words"]:
                if w[4] == word:
                    return round(w[2] - w[0], 2)
    return None


try:
    R["compare"] = sweep("compare", [("기본", None), ("자간 -5%", {"spacing": ALL(-5)}), ("자간 -10%", {"spacing": ALL(-10)}),
                                     ("자간 +10%", {"spacing": ALL(10)}), ("장평 95%", {"ratio": ALL(95)}), ("장평 90%", {"ratio": ALL(90)})])
    R["threshold"] = sweep("threshold", [(f"자간 {v}%", {"spacing": ALL(v)}) for v in (-1, -2, -3, -4, -5)] +
                           [(f"장평 {v}%", {"ratio": ALL(v)}) for v in (99, 98, 97, 96, 95)])
    # 한글 · 영문 섞인 문장
    p = M.make("mixed.hwpx", ["(아) " + MIXED, "", "(사) " + MIXED, "", "(자) " + MIXED])
    M.patch(p, {"(사)": {"spacing": {"hangul": -15}}, "(자)": {"spacing": ALL(-15)}}, author=False)
    assert M.hwp.Open(p, "HWPX", "")
    rows = M.rows_of(M.save("mixed.pdf", "PDF"))
    R["mixed"] = {}
    for m, name in (("(아)", "기본"), ("(사)", "한글만 자간 -15%"), ("(자)", "모든 언어 자간 -15%")):
        r = next(x for x in rows if x["text"].startswith(m))
        R["mixed"][name] = {"공지": width(rows, "공지", m), "Notice": width(rows, "Notice", m), "line_pt": round(r["x1"] - r["x0"], 1)}
    # 실습 · 정답 — 가을 독서 모임 안내
    DOC = ["가을 독서 모임 안내", "", TEXT, "", "준비물: 읽은 책 한 권, 필기구", "문의: 모임 총무(연습용 가상 연락처)"]
    for name, shapes, out in (("practice", {}, "실습.hwpx"), ("answer", {"이번 달": {"spacing": ALL(-3)}}, "정답.hwpx")):
        p = M.make(name + ".hwpx", DOC)
        M.patch(p, shapes)
        shutil.copy(p, os.path.join(HERE, out))
        assert M.hwp.Open(p, "HWPX", "")
        g = M.paras_starting(M.save(name + ".pdf", "PDF"), "이번 달")
        R[name] = {"lines": len(g[0]) - 2, "rows": g[0]}  # 뒤 두 줄(준비물 · 문의)은 따로 문단
finally:
    M.hwp.Clear(1)
    M.hwp.Quit()
json.dump(R, open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
