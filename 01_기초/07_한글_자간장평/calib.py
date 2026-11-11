"""measure.py 의 안내 문단 길이 고르기 — 끝에 글자를 하나씩 더 붙인 문단 여러 벌을 한 문서에 넣고 PDF 로 줄 수 · 마지막 줄 글자 수를 잰다.
마지막 줄에 두 글자만 남는 길이를 찾고, 그 문단에 자간 -5% · 장평 95% 를 주면 한 줄이 줄어드는지 본다. 결과 calib.json. 실행: python calib.py"""
import json, os, sys
import measure_lib as M

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
TAILS = ["", " 늦으시는 분은 미리 연락해 주세요.", " 늦게 오시는 분은 미리 연락해 주세요.", " 늦으시면 총무에게 미리 연락해 주세요.",
         " 늦게 오시는 분은 총무에게 미리 알려 주세요.", " 늦으시는 분은 총무에게 미리 문자를 보내 주세요.", " 늦게 오시면 총무에게 미리 문자를 보내 주세요.",
         " 늦게 오시는 분은 총무에게 미리 문자를 보내 주세요.", " 늦으시는 분은 전날까지 총무에게 미리 알려 주세요.",
         " 늦게 오시는 분은 전날까지 총무에게 미리 알려 주세요.", " 늦으시는 분은 모임 전날까지 총무에게 미리 알려 주세요.",
         " 늦게 오시는 분은 모임 전날까지 총무에게 꼭 알려 주세요.", " 늦게 오실 분은 모임 전날까지 총무에게 미리 꼭 알려 주세요."]
R = {}
try:
    ks = list(range(len(TAILS)))
    texts = [M.BASE_TEXT + x for x in TAILS]
    paras = []
    for t in texts:
        paras += [t, ""]
    p = M.make("calib.hwpx", paras)
    M.patch(p, {}, author=False)
    assert M.hwp.Open(p, "HWPX", "")
    groups = M.paras_starting(M.save("calib.pdf", "PDF"), "이번 달")
    R["candidates"] = [{"k": k, "tail": TAILS[k], "lines": len(g), "last_line": g[-1], "last_chars": len(g[-1].replace(" ", ""))}
                       for k, g in zip(ks, groups)]
    pick = next((c for c in R["candidates"] if c["last_chars"] == 2), None)
    R["pick"] = pick
    if pick:
        t = M.BASE_TEXT + pick["tail"]
        p = M.make("calib2.hwpx", [t, "", t, "", t])
        # 같은 문단 세 벌 = 기본 · 자간 -5 · 장평 95 (앞 표시 글자 없이 — 표시가 줄 나눔을 바꾸지 않게)
        M.patch(p, {}, author=False)
        M.patch_nth(p, "이번 달", {2: {"spacing": {l: -5 for l in M.LANGS}}, 3: {"ratio": {l: 95 for l in M.LANGS}}})
        assert M.hwp.Open(p, "HWPX", "")
        g = M.paras_starting(M.save("calib2.pdf", "PDF"), "이번 달")
        R["verify"] = [{"case": c, "lines": len(x), "last_line": x[-1]} for c, x in zip(["기본", "자간 -5%", "장평 95%"], g[0:3])]
finally:
    M.hwp.Clear(1)
    M.hwp.Quit()
json.dump(R, open(os.path.join(HERE, "calib.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({"pick": R.get("pick"), "verify": R.get("verify"), "near": [(c["k"], c["lines"], c["last_chars"]) for c in R.get("candidates", [])]}, ensure_ascii=False))
