"""웹한글(MYBOX 한컴오피스 Web)에서 단축키를 실제로 눌러 어떤 창이 뜨는지 본다 — 한글 강의 기초 1강(단축키).
열린 작업 탭(webhwp)에서: 본문 빈 곳을 눌러 커서를 두고 → 키 → 2초 → 화면 저장 → Esc → 창이 닫혔는지 확인.
대화상자만 열고 닫는다(문서 내용을 바꾸는 키는 넣지 않는다). 결과: <저장 폴더>/k_<이름>.png 와 표준 출력의 창 제목.
인자: 저장 폴더 [시험 이름 …](없으면 전부)."""
import os, sys, time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from tab import attach, work_tab, shot
from act import click

OUT = sys.argv[1]
ONLY = set(sys.argv[2:])
d = attach(); work_tab(d, prefer="HWP_EDITOR")  # 도움말 탭(cloud-hwp/help)도 주소에 webhwp 가 들어가므로 편집기 주소로 고른다
if "HWP_EDITOR" not in d.current_url:
    raise SystemExit("웹한글 편집기 탭이 아님")
A, C, S = Keys.ALT, Keys.CONTROL, Keys.SHIFT


def combo(mods, key):
    ch = ActionChains(d)
    for m in mods:
        ch.key_down(m)
    ch.send_keys(key)
    for m in reversed(mods):
        ch.key_up(m)
    ch.perform()


TESTS = {
    "alt_l": [((A,), "l")],                 # 글자 모양
    "alt_t": [((A,), "t")],                 # 문단 모양
    "ctrl_f10": [((C,), Keys.F10)],         # 문자표
    "f7": [((), Keys.F7)],                  # 편집 용지
    "ctrl_n_t": [((C,), "n"), ((), "t")],   # 표 만들기(한글 2024 도움말)
    "alt_n_t": [((A,), "n"), ((), "t")],    # 웹한글 메뉴 표시 확인용
    "ctrl_f2": [((C,), Keys.F2)],           # 찾아 바꾸기
    "ctrl_h": [((C,), "h")],                # 찾아 바꾸기(다른 키)
    "ctrl_k_n": [((C,), "k"), ((), "n")],   # 문단 번호 모양
}


def dialogs():
    return d.execute_script("""return [...document.querySelectorAll('div,section')].filter(e=>{
        if(!e.offsetParent) return false; const cs=getComputedStyle(e); const r=e.getBoundingClientRect();
        return (cs.position==='absolute'||cs.position==='fixed') && r.width>200 && r.height>120 && r.y>60 && +cs.zIndex>0;
      }).map(e=>{const t=(e.querySelector('[class*=title],h1,h2,h3')||e).innerText.trim().split('\\n')[0];
        const r=e.getBoundingClientRect(); return [t.slice(0,30), Math.round(r.x), Math.round(r.y), Math.round(r.width), Math.round(r.height)];}).slice(0,4)""")


for name, steps in TESTS.items():
    if ONLY and name not in ONLY:
        continue
    click(d, 700, 900); time.sleep(0.8)       # 본문 아래쪽 빈 곳 — 커서만 둔다
    before = len(d.window_handles)
    for mods, key in steps:
        combo(mods, key); time.sleep(0.6)
    time.sleep(2.0)
    dl = dialogs()
    print(name, "→", dl, "| tabs", before, "→", len(d.window_handles))
    shot(d, os.path.join(OUT, f"k_{name}.png"))
    ActionChains(d).send_keys(Keys.ESCAPE).perform(); time.sleep(1.2)
    if dialogs():
        ActionChains(d).send_keys(Keys.ESCAPE).perform(); time.sleep(1.0)
    left = dialogs()
    if left:
        print("  !! 창이 남아 있음:", left)
        shot(d, os.path.join(OUT, f"k_{name}_left.png"))
        break
print("done")
