# -*- coding: utf-8 -*-
"""
딱 한 번만 실행하는 파일이에요.

브라우저 창이 뜨면 사장님이 직접 에이블리 파트너센터에 로그인하세요.
로그인이 끝난 걸 이 파일이 알아서 알아채고 '로그인된 상태'를 세션.json 으로
저장해요. 터미널로 돌아와 엔터를 누를 필요 없어요.
그 뒤로는 자동 다운로드가 이 세션을 씁니다.

    python3 -m pip install playwright
    python3 -m playwright install chromium
    python3 에이블리_세션저장.py
"""

import json
import os
import sys
import time

여기 = os.path.dirname(os.path.abspath(__file__))
설정파일 = os.path.join(여기, "설정.json")
세션파일 = os.path.join(여기, "세션.json")
# 로그인한 상태를 이 폴더에 남겨둬요. 그래서 한 번 로그인하면 다음부터는 안 해도 돼요.
프로필폴더 = os.path.join(여기, ".크롬프로필")
파트너센터 = "https://my.a-bly.com/"
발주관리 = "https://my.a-bly.com/sales/order/prepare"

기다리는시간 = 1800   # 최대 30분까지 로그인을 기다려요

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("❌ playwright 가 없어요. 먼저 아래 두 줄을 실행해 주세요:")
    print("   python3 -m pip install playwright")
    print("   python3 -m playwright install chromium")
    sys.exit(1)


def 로그인정보읽기():
    """설정.json 에 에이블리 아이디·비밀번호가 있으면 가져와요. 없으면 (None, None).

    환경변수(ABLY_ID / ABLY_PW)가 있으면 그게 우선이에요 — 클라우드용이에요.
    """
    아이디 = os.environ.get("ABLY_ID", "")
    비번 = os.environ.get("ABLY_PW", "")
    if not (아이디 and 비번):
        설정 = {}
        try:
            # 클라우드에서는 설정.json 통째로를 SETTINGS_JSON 시크릿에 넣어 써요
            통째로 = os.environ.get("SETTINGS_JSON")
            if 통째로:
                설정 = json.loads(통째로)
            elif os.path.exists(설정파일):
                with open(설정파일, encoding="utf-8") as f:
                    설정 = json.load(f)
        except Exception as e:
            print(f"   ⚠️ 설정을 못 읽었어요: {e}")
        아이디 = 아이디 or 설정.get("에이블리아이디", "")
        비번 = 비번 or 설정.get("에이블리비밀번호", "")
    아이디, 비번 = (아이디 or "").strip(), (비번 or "").strip()
    return (아이디, 비번) if (아이디 and 비번) else (None, None)


def 자동로그인(쪽, 아이디, 비번):
    """로그인 칸을 찾아 채우고 로그인 버튼을 눌러요. 성공 여부를 돌려줘요."""
    아이디칸 = [
        "input[placeholder='이메일']",
        "input[type='email']",
        "input[name*='id' i]",
        "form input[type='text']",
    ]
    비번칸 = ["input[placeholder='비밀번호']", "input[type='password']"]
    try:
        # ⚠️ 로그인 폼이 그려지기를 먼저 기다려요.
        #    my.a-bly.com/ 로 들어가면 화면을 그린 다음에 /login 으로 넘어가서,
        #    바로 채우려 들면 '아이디 칸을 못 찾았어요' 로 실패해요.
        쪽.wait_for_selector("input[type='password']", timeout=30000, state="visible")
    except Exception:
        print("   ⚠️ 로그인 칸이 안 나타났어요")
        return False
    try:
        for 후보 in 아이디칸:
            칸 = 쪽.locator(후보)
            if 칸.count():
                칸.first.fill(아이디, timeout=10000)
                break
        else:
            print("   ⚠️ 아이디 칸을 못 찾았어요")
            return False

        for 후보 in 비번칸:
            칸 = 쪽.locator(후보)
            if 칸.count():
                칸.first.fill(비번, timeout=10000)
                break
        else:
            print("   ⚠️ 비밀번호 칸을 못 찾았어요")
            return False

        단추 = 쪽.get_by_role("button", name="로그인")
        if 단추.count() == 0:
            단추 = 쪽.locator("button[type='submit'], input[type='submit']")
        if 단추.count() == 0:
            쪽.keyboard.press("Enter")
        else:
            단추.first.click(timeout=10000)
        # networkidle 은 에이블리에서 영영 안 오니 기다리지 않아요
        쪽.wait_for_load_state("domcontentloaded", timeout=30000)
        return True
    except Exception as e:
        print(f"   ⚠️ 자동 로그인 중 문제: {e}")
        return False


def 로그인됐나(쪽, 기다릴까=True, 최대=20):
    """아직 로그인 전인지 봐요.

    ⚠️ goto 직후 바로 주소만 보면 안 돼요. 에이블리는 화면을 그린 다음에
    /login?redirect=... 로 넘기기 때문에, 넘어가기 '전'에 읽으면 로그인된 걸로
    착각해요.

    ⚠️ 그렇다고 networkidle 을 기다리면 안 돼요. 에이블리 화면은 연결을 계속
    열어둬서 networkidle 이 영영 안 와요. 그걸 기다리다 타임아웃 나면 멀쩡한
    세션도 '로그인 안 됨'으로 잘못 판정해요 (실제로 이것 때문에 한 번 날렸어요).
    그래서 주소가 로그인으로 튕기는지를 잠깐 지켜보는 방식으로 확인해요.
    """
    끝 = time.time() + (최대 if 기다릴까 else 0)
    while True:
        try:
            if "login" in 쪽.url:
                return False        # 로그인으로 튕겼으면 확실히 아니에요
            본문 = 쪽.locator("body").inner_text(timeout=5000)
            if "PASSWORD" in 본문.upper():
                return False        # 로그인 폼이 보이면 아직이에요
            if len(본문.strip()) > 300:
                return True         # 내용이 실하게 그려졌으면 들어간 거예요
        except Exception:
            pass
        if time.time() >= 끝:
            break
        time.sleep(1)
    # 시간이 다 됐는데 튕기지도, 폼이 보이지도 않으면 들어간 걸로 봐요
    try:
        return "login" not in 쪽.url
    except Exception:
        return False


def 세션이진짜되나(p):
    """저장한 세션.json 을 '새 창'에 끼워서 발주관리가 열리는지 진짜로 해봐요.

    쿠키 이름을 넘겨짚는 것보다 이게 확실해요. 자동 발송이 실제로 쓰는 방식과
    똑같이(창 없이) 열어보는 거라, 여기서 되면 진짜 되는 거예요.
    """
    브라우저 = p.chromium.launch(headless=True)
    try:
        화면 = 브라우저.new_context(storage_state=세션파일, locale="ko-KR")
        쪽 = 화면.new_page()
        쪽.goto(발주관리, wait_until="domcontentloaded", timeout=60000)
        return 로그인됐나(쪽)
    except Exception:
        return False
    finally:
        브라우저.close()


def 메인():
    with sync_playwright() as p:
        # ⚠️ launch() + new_context() 를 쓰면 크로미움이 우리가 모르는 창을 따로
        #    열어서, 사장님이 그쪽에 로그인하면 영영 못 알아채요. 실제로 네 번 헛돌았어요.
        #    프로필 폴더를 정해서 여는 방식은 창이 하나로 딱 정해져서 그럴 일이 없고,
        #    로그인한 상태가 폴더에 남아 다음 실행 때 다시 로그인 안 해도 돼요.
        화면 = p.chromium.launch_persistent_context(
            프로필폴더,
            headless=False,
            args=["--start-maximized"],
            no_viewport=True,
            locale="ko-KR",
            accept_downloads=True,
        )
        브라우저 = 화면.browser
        쪽 = 화면.pages[0] if 화면.pages else 화면.new_page()
        쪽.goto(파트너센터, wait_until="domcontentloaded")
        쪽.bring_to_front()
        # 다른 창(파이참 등)에 가려서 못 보고 지나치는 일이 많아서 앞으로 꺼내요
        os.system("""osascript -e 'tell application "System Events" to set frontmost of """
                  """(first process whose name contains "Chromium") to true' 2>/dev/null""")

        print()
        아이디, 비번 = 로그인정보읽기()
        if 아이디:
            print(f"🔑 설정.json 에 적힌 아이디({아이디})로 자동 로그인해볼게요…")
            if 자동로그인(쪽, 아이디, 비번):
                print("   눌렀어요. 들어갔는지 확인 중…")
            else:
                print("   자동 로그인이 안 됐어요 → 창에서 직접 로그인해 주세요.")
        else:
            print("🌐 브라우저가 열렸어요.")
            print("   에이블리 셀러(my.a-bly.com)에 직접 로그인해 주세요.")
            print("   (설정.json 에 '에이블리아이디'·'에이블리비밀번호'를 채우면 다음부턴 자동으로 해요)")
        print("   로그인이 끝나면 알아서 저장하고 창을 닫아요 (엔터 안 눌러도 돼요).")
        print()

        # ⚠️ 여기서 '지금 로그인됐나?' 를 페이지 주소로 알아내려던 시도가 여러 번
        #    실패했어요. 창이 여러 개 뜨거나 추적이 어긋나면 화면은 대시보드인데
        #    코드는 로그인 화면을 보고 있었어요.
        #    그래서 이제는 묻지 않고, 그냥 저장해본 다음 '그 세션이 진짜 되는지'
        #    새 창에서 확인해요. 되면 그게 로그인된 거예요. 확인이 곧 감지예요.
        마감 = time.time() + 기다리는시간
        알림찍은시각 = 0
        다음시도 = 0
        while time.time() < 마감:
            if time.time() >= 다음시도:
                다음시도 = time.time() + 5
                try:
                    화면.storage_state(path=세션파일)
                except Exception:
                    pass
                else:
                    if 세션이진짜되나(p):
                        화면.close()
                        print("✅ 로그인 확인! (저장한 세션으로 발주관리가 열렸어요)")
                        print(f"   저장 완료 → {세션파일}")
                        print("   이 파일에는 로그인 정보가 들어 있어요. 남한테 주면 안 돼요!")
                        print("   (깃에는 안 올라가게 이미 막아뒀어요)")
                        print()
                        print("클라우드에서 돌리려면 이 파일 내용을 통째로 복사해서")
                        print("GitHub → Settings → Secrets → ABLY_SESSION 에 붙여넣으세요.")
                        return 0
                    if os.path.exists(세션파일):
                        os.remove(세션파일)   # 아직 안 된 세션은 남겨두지 않아요

            남은시간 = int(마감 - time.time())
            if time.time() - 알림찍은시각 > 20:
                print(f"   ⏳ 로그인 기다리는 중… ({남은시간}초 남음)")
                알림찍은시각 = time.time()
            time.sleep(2)

        화면.close()
        print("❌ 시간 안에 로그인이 안 끝나서 그냥 나왔어요. 다시 실행해 주세요.")
        return 1


if __name__ == "__main__":
    sys.exit(메인())
