# -*- coding: utf-8 -*-
"""
결제 완료(발주 관리)에 새 주문이 뜨면, 사람 손 없이 끝까지:

  ① 발주 관리에서 해당 주문 체크 → [상품준비중 처리]
  ② 고객 메일로 테마 다운로드 링크 발송 💌
  ③ 발송 관리에서 해당 주문 체크 → [발송 처리 → 개별 상태 변경]

이메일을 안 적었거나 링크가 없는 주문은 ①②③ 전부 건드리지 않고
보류_수동처리.csv 에 남겨요 (에이블리에는 '결제 완료'로 그대로 있어요).

  python3 에이블리_전자동.py              # 연습: 아무것도 안 누르고 계획만 보여줘요
  python3 에이블리_전자동.py --진짜        # 진짜로 다 해요
  python3 에이블리_전자동.py --진짜 --보이게   # 브라우저 창을 보면서

매 단계 화면을 과정기록/ 폴더에 남겨요. 창 모양이 예상과 다르면
거기서 멈추고 스크린샷을 남기니, 그걸 보고 스크립트를 다듬으면 돼요.
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime

여기 = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, 여기)

import 에이블리_자동발송 as 발송기
import 에이블리_세션저장 as 세션기

세션파일 = os.path.join(여기, "세션.json")
받는곳 = os.path.join(여기, "주문엑셀")
기록폴더 = os.path.join(여기, "과정기록")

파트너센터 = "https://my.a-bly.com/"
발주관리 = "https://my.a-bly.com/sales/order/prepare"
발송관리 = "https://my.a-bly.com/sales/order/send"

# 실제 화면에서 확인한 버튼·메뉴 이름 (바뀔까 봐 후보 여러 개)
발주엑셀후보 = ["전체 주문 엑셀 다운로드", "전체 주문 엑셀 받기", "엑셀 다운로드"]
발송엑셀후보 = ["전체 주문 엑셀 받기", "전체 주문 엑셀 다운로드", "엑셀 받기"]
준비중버튼후보 = ["상품준비중 처리", "상품 준비중 처리"]
발송처리버튼후보 = ["발송 처리", "발송처리"]
개별변경후보 = ["개별 상태 변경", "개별상태변경"]
확인버튼후보 = ["확인", "네", "예", "변경", "저장", "적용", "처리"]

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("❌ playwright 가 없어요:  python3 -m pip install playwright && python3 -m playwright install chromium")
    sys.exit(1)


# ─────────────────────────────────────── 작은 도구들

def 세션읽기():
    """저장돼 있는 세션을 가져와요. 없으면 None (그때는 자동 로그인으로 만들어요)."""
    시크릿 = os.environ.get("ABLY_SESSION")
    if 시크릿:
        return json.loads(시크릿)
    if os.path.exists(세션파일):
        with open(세션파일, encoding="utf-8") as f:
            return json.load(f)
    return None


def 준비물점검():
    """시작하자마자 뭐가 있고 뭐가 없는지 알려줘요.

    클라우드(GitHub Actions)에서 시크릿 이름을 하나 틀리면 그냥 죽어버려서
    무엇 때문인지 알 수가 없었어요. 그래서 맨 앞에서 있는 것/없는 것을 찍어요.
    값은 절대 안 찍고, 있는지 없는지만 봐요.
    """
    설정있음 = os.path.exists(os.path.join(여기, "설정.json"))
    통째로있음 = bool(os.environ.get("SETTINGS_JSON"))
    항목 = [
        ("SETTINGS_JSON", "설정 통째로 (이거 하나면 아래 4개는 없어도 돼요)"),
        ("ABLY_ID", "에이블리 아이디"),
        ("ABLY_PW", "에이블리 비밀번호"),
        ("GMAIL_ADDR", "보내는 메일 주소"),
        ("GMAIL_APP_PASSWORD", "메일 앱비밀번호"),
        ("THEME_LINKS_JSON", "테마 링크표"),
    ]
    print("🧾 준비물 점검")
    print(f"   설정.json 파일: {'있음' if 설정있음 else '없음 (클라우드에서는 정상)'}")
    빠진것 = []
    for 이름, 설명 in 항목:
        값 = os.environ.get(이름, "")
        print(f"   {이름:20} {'✅ 있음' if 값 else '❌ 없음'}  ({설명})")
        if not 값:
            빠진것.append(이름)

    if not 설정있음 and not 통째로있음 and 빠진것:
        # 테마링크표는 저장소 파일로도 되니 없어도 괜찮아요
        치명적 = [x for x in 빠진것 if x not in ("THEME_LINKS_JSON", "SETTINGS_JSON")]
        if 치명적:
            print()
            print(f"🛑 시크릿이 빠졌어요: {', '.join(치명적)}")
            print("   깃허브 → Settings → Secrets and variables → Actions 에서")
            print("   위 이름 그대로(대문자·밑줄까지) 등록했는지 확인해 주세요.")
            sys.exit(1)
    print()


def 자동로그인해서세션만들기(p):
    """아이디·비밀번호로 창 없이 로그인해서 세션을 새로 만들어요.

    세션은 언젠가 만료돼요. 클라우드(GitHub Actions)에서는 사람이 다시 로그인해
    줄 수가 없으니, 만료되면 여기서 스스로 다시 로그인해요. 이게 있어야 컴퓨터를
    꺼놔도 계속 돌아가요.
    """
    아이디, 비번 = 세션기.로그인정보읽기()
    if not 아이디:
        print("ℹ️  로그인에 쓸 아이디·비밀번호가 없어요.")
        print("    (클라우드라면 ABLY_ID / ABLY_PW 시크릿, 맥이라면 설정.json)")
        return None

    print(f"🔑 아이디({아이디})·비밀번호로 자동 로그인 중…")
    브라우저 = p.chromium.launch(headless=True)
    try:
        화면 = 브라우저.new_context(locale="ko-KR", accept_downloads=True)
        쪽 = 화면.new_page()
        쪽.goto(파트너센터, wait_until="domcontentloaded", timeout=60000)
        세션기.자동로그인(쪽, 아이디, 비번)

        # 눌렀다고 끝이 아니에요. 실제로 발주관리가 열리는지 확인될 때까지 지켜봐요.
        마감 = time.time() + 60
        while time.time() < 마감:
            상태 = 화면.storage_state()
            if 세션살아있나(p, 상태):
                print("   ✅ 자동 로그인 성공")
                try:                       # 다음 실행에서 바로 쓰라고 남겨둬요
                    with open(세션파일, "w", encoding="utf-8") as f:
                        json.dump(상태, f, ensure_ascii=False)
                except Exception:
                    pass
                return 상태
            time.sleep(3)

        # 왜 안 됐는지 알 수 있게 화면을 남겨요 (Actions 아티팩트로 받아볼 수 있어요)
        print("   ❌ 자동 로그인이 안 됐어요")
        try:
            print(f"      지금 주소: {쪽.url}")
            본문 = 쪽.locator("body").inner_text(timeout=5000)
            print(f"      화면 글자: {본문[:200].strip()!r}")
        except Exception as e:
            print(f"      화면을 못 읽었어요: {e}")
        os.makedirs(기록폴더, exist_ok=True)
        try:
            쪽.screenshot(path=os.path.join(기록폴더, "자동로그인실패.png"), full_page=True)
            with open(os.path.join(기록폴더, "자동로그인실패.html"), "w", encoding="utf-8") as f:
                f.write(쪽.content())
            print(f"      화면을 {os.path.basename(기록폴더)}/ 에 남겼어요")
        except Exception:
            pass
        return None
    finally:
        브라우저.close()


def 세션살아있나(p, 세션):
    """이 세션으로 발주관리가 진짜 열리는지 창 없이 확인해요."""
    if not 세션:
        return False
    브라우저 = p.chromium.launch(headless=True)
    try:
        쪽 = 브라우저.new_context(storage_state=세션, locale="ko-KR").new_page()
        쪽.goto(발주관리, wait_until="domcontentloaded", timeout=60000)
        return 세션기.로그인됐나(쪽)
    except Exception:
        return False
    finally:
        브라우저.close()


def 쓸수있는세션(p):
    """저장된 세션 → 안 되면 자동 로그인. 둘 다 실패하면 멈춰요."""
    세션 = 세션읽기()
    if 세션살아있나(p, 세션):
        return 세션
    if 세션:
        print("ℹ️  저장된 세션이 만료됐어요 → 다시 로그인할게요")
    세션 = 자동로그인해서세션만들기(p)
    if 세션:
        return 세션
    print("❌ 에이블리에 들어갈 수가 없어요.")
    print("   설정.json 의 '에이블리아이디'·'에이블리비밀번호'를 채우거나,")
    print("   (클라우드라면) ABLY_ID / ABLY_PW 시크릿을 넣어주세요.")
    print("   아니면 맥에서 `python3 에이블리_세션저장.py` 를 한 번 돌려주세요.")
    sys.exit(1)


def 찰칵(쪽, 이름):
    """지금 화면을 과정기록/ 에 남겨요."""
    os.makedirs(기록폴더, exist_ok=True)
    경로 = os.path.join(기록폴더, f"{datetime.now():%Y%m%d_%H%M%S}_{이름}.png")
    try:
        쪽.screenshot(path=경로, full_page=False)
    except Exception:
        pass
    return 경로


def 총건수(쪽):
    try:
        글 = 쪽.locator("text=/총\\s*[\\d,]+\\s*건/").first.inner_text(timeout=8000)
        m = re.search(r"총\s*([\d,]+)\s*건", 글)
        if m:
            return int(m.group(1).replace(",", ""))
    except Exception:
        pass
    return None


def 엑셀받기(쪽, 후보들, 별명):
    os.makedirs(받는곳, exist_ok=True)
    for 이름 in 후보들:
        버튼 = 쪽.get_by_role("button", name=이름).or_(쪽.get_by_text(이름, exact=False))
        try:
            if 버튼.count() == 0:
                continue
            with 쪽.expect_download(timeout=45000) as 받기:
                버튼.first.click(timeout=5000)
            받은것 = 받기.value
            경로 = os.path.join(받는곳, 받은것.suggested_filename or f"{datetime.now():%Y%m%d_%H%M}_{별명}.xlsx")
            받은것.save_as(경로)
            return 경로
        except Exception:
            continue
    return None


def 목록에보이는것만(쪽, 주문들, 화면이름):
    """엑셀에 들어 있어도 '지금 이 화면 목록'에 없는 주문은 빼요.  ★안전장치★

    에이블리 엑셀 버튼 이름이 '전체 주문 엑셀 다운로드' 라서, 지난 주문(배송 완료 등)까지
    통째로 딸려올 수 있어요. 엑셀에는 주문 상태 열이 없어서 엑셀만 봐서는 구분이 안 돼요.
    이 대조가 없으면 예전 고객에게 메일이 다시 나갈 수 있어요.

    못 읽으면 '아무것도 안 한다' 쪽으로 실패해요 (보내는 것보다 안 보내는 게 안전).
    """
    if not 주문들:
        return []
    try:
        본문 = 쪽.locator("body").inner_text(timeout=10000)
    except Exception:
        print(f"   ⚠️ {화면이름} 화면 글자를 못 읽었어요 → 안전을 위해 이 화면 주문은 전부 건너뜁니다")
        찰칵(쪽, f"{화면이름}_화면읽기실패")
        return []

    보임 = [j for j in 주문들 if j["주문번호"] in 본문]
    안보임 = [j for j in 주문들 if j["주문번호"] not in 본문]
    if 안보임:
        print(f"   ℹ️ {화면이름}: 엑셀 {len(주문들)}건 중 화면 목록에 있는 {len(보임)}건만 처리해요")
        print(f"      (화면에 없는 {len(안보임)}건은 지난 주문이거나 다음 페이지 — 건드리지 않아요)")
    return 보임


def 줄체크(쪽, 상품주문번호):
    """표에서 해당 주문 줄을 찾아 맨 앞 체크박스를 켜요."""
    줄 = 쪽.locator(f"tr:has-text('{상품주문번호}')")
    if 줄.count() == 0:
        # <tr> 이 아니라 div 로 만든 목록일 수도 있어요
        줄 = 쪽.locator(f"[class*='row']:has-text('{상품주문번호}')")
    if 줄.count() == 0:
        return False
    첫줄 = 줄.first
    상자 = 첫줄.locator("input[type='checkbox']")
    if 상자.count() > 0:
        for 강제로 in (False, True):
            try:
                상자.first.check(timeout=4000, force=강제로)
                return True
            except Exception:
                continue
    # 체크박스가 숨겨진 꾸밈형이면 첫 칸 '안'의 눌러지는 것만 클릭
    # (칸 전체를 누르면 주문 상세 페이지로 새어나갈 수 있어서 범위를 좁혀요)
    try:
        꾸밈 = 첫줄.locator("td, [class*='cell']").first.locator("label, span, i, svg, div").first
        꾸밈.click(timeout=3000)
        return True
    except Exception:
        return False


def 동의체크(쪽, 대화):
    """'네, 확인 했습니다.' 같은 동의 체크박스를 켜요.
    이걸 켜야 [확인] 버튼이 살아나요 (실제 화면에서 확인함)."""
    켰다 = False
    try:
        상자들 = 대화.locator("input[type='checkbox']")
        for i in range(상자들.count()):
            상자 = 상자들.nth(i)
            for 강제로 in (False, True):
                try:
                    if 상자.is_checked():
                        break
                    상자.check(timeout=2500, force=강제로)
                    켰다 = True
                    break
                except Exception:
                    continue
    except Exception:
        pass
    if not 켰다:
        # 숨겨진 input 이면 글자를 눌러요
        for 글 in ["네, 확인 했습니다", "확인 했습니다", "확인했습니다"]:
            try:
                항목 = 대화.get_by_text(글, exact=False)
                if 항목.count() > 0:
                    항목.first.click(timeout=2500)
                    켰다 = True
                    break
            except Exception:
                continue
    if 켰다:
        쪽.wait_for_timeout(800)  # [확인] 버튼이 살아날 시간을 줘요
    return 켰다


def 대화상자(쪽):
    """떠 있는 팝업(모달)을 찾아요."""
    for 고르개 in [".ant-modal-content", "[role='dialog']", ".ant-modal", ".modal-content"]:
        칸 = 쪽.locator(고르개)
        try:
            if 칸.count() > 0 and 칸.first.is_visible():
                return 칸.first
        except Exception:
            continue
    return None


def 확인누르기(쪽, 어디, 단계이름):
    """팝업(또는 화면)에서 확인류 버튼을 눌러요."""
    for 이름 in 확인버튼후보:
        try:
            버튼 = 어디.get_by_role("button", name=이름)
            if 버튼.count() > 0 and 버튼.first.is_enabled():
                버튼.first.click(timeout=4000)
                쪽.wait_for_timeout(1500)
                return True
        except Exception:
            continue
    찰칵(쪽, f"{단계이름}_확인버튼못찾음")
    return False


# 예전에 팝업 안 '상태 선택칸'을 자동으로 고르는 코드가 있었는데 없앴어요.
# 발주/발송 팝업은 둘 다 "~하시겠습니까?" 단순 확인창이라 고를 게 없고,
# 발송 팝업의 택배사 드롭다운을 잘못 열었다 Escape 로 닫으면
# 팝업 전체가 닫혀버릴 수 있어서 오히려 위험했어요.
# 디지털 상품이라 택배사·송장번호는 비워둬도 발송 처리됩니다 (실제로 확인함).


# ─────────────────────────────────────── 단계들

def 화면열기(쪽, 주소, 이름):
    # ⚠️ networkidle 로 기다리면 안 돼요. 에이블리 화면은 연결을 계속 열어둬서
    #    networkidle 이 영영 안 오고, 60초 뒤 타임아웃으로 통째로 죽어요.
    #    대신 표(또는 '총 N건' 글자)가 그려질 때까지 기다려요.
    쪽.goto(주소, wait_until="domcontentloaded", timeout=60000)
    try:
        쪽.wait_for_selector("table, [class*='row'], text=/총\\s*[\\d,]+\\s*건/",
                             timeout=20000, state="attached")
    except Exception:
        pass          # 목록이 0건이면 표가 아예 없을 수도 있어요
    if "login" in 쪽.url:
        print("❌ 세션이 만료됐어요. `python3 에이블리_세션저장.py` 를 다시 돌려주세요.")
        찰칵(쪽, f"{이름}_로그인튕김")
        sys.exit(2)


def 표에서처리(쪽, 대상들, 버튼후보, 단계이름, 개별메뉴=None):
    """주문들 체크 → 버튼 누르기 → (팝업 뜨면) 상태 고르고 확인."""
    들어온주소 = 쪽.url
    체크됨 = []
    for 번호 in 대상들:
        if 줄체크(쪽, 번호):
            체크됨.append(번호)
        else:
            print(f"   ⚠️ 표에서 {번호} 줄을 못 찾았어요 (건너뜀)")
    if not 체크됨:
        print(f"   {단계이름}: 체크된 주문이 없어서 넘어가요")
        return []

    # 체크하다가 주문 상세 페이지로 새어나갔으면 여기서 멈춰요 (엉뚱한 버튼 누르기 방지)
    if 쪽.url != 들어온주소:
        print(f"   ❌ {단계이름}: 체크하다 다른 화면으로 넘어갔어요 → 중단합니다")
        찰칵(쪽, f"{단계이름}_화면이탈")
        return []

    찰칵(쪽, f"{단계이름}_체크후")

    눌렀다 = False
    for 이름 in 버튼후보:
        try:
            버튼 = 쪽.get_by_role("button", name=이름)
            if 버튼.count() > 0 and 버튼.first.is_enabled():
                버튼.first.click(timeout=5000)
                눌렀다 = True
                break
        except Exception:
            continue
    if not 눌렀다:
        print(f"   ❌ {단계이름}: 버튼을 못 찾았어요")
        찰칵(쪽, f"{단계이름}_버튼못찾음")
        return []

    쪽.wait_for_timeout(1000)

    # '발송 처리'는 드롭다운이라 '개별 상태 변경'을 한 번 더 눌러야 해요
    if 개별메뉴:
        for 이름 in 개별메뉴:
            try:
                항목 = 쪽.get_by_text(이름, exact=False)
                if 항목.count() > 0:
                    항목.first.click(timeout=4000)
                    쪽.wait_for_timeout(1000)
                    break
            except Exception:
                continue

    대화 = 대화상자(쪽)
    if 대화:
        찰칵(쪽, f"{단계이름}_팝업")
        # 택배사·송장번호는 건드리지 않아요 (디지털 상품이라 비워도 통과)
        if not 확인누르기(쪽, 대화, 단계이름):
            print(f"   ⚠️ {단계이름}: 팝업 확인 버튼을 못 눌렀어요 — 과정기록/ 스크린샷을 봐주세요")
            return []

    # 발송 처리 뒤엔 '결과 안내' 창이 하나 더 떠요:
    # "네, 확인 했습니다." 체크를 해야 확인 버튼이 살아나는 방식 (실제 화면에서 확인함)
    for _ in range(2):
        쪽.wait_for_timeout(1200)
        뒤대화 = 대화상자(쪽)
        if not 뒤대화:
            break
        찰칵(쪽, f"{단계이름}_결과창")
        동의체크(쪽, 뒤대화)
        if not 확인누르기(쪽, 뒤대화, f"{단계이름}_결과"):
            break
    쪽.wait_for_timeout(1000)
    찰칵(쪽, f"{단계이름}_처리후")
    print(f"   ✓ {단계이름}: {len(체크됨)}건")
    return 체크됨


# ─────────────────────────────────────── 메인

def 메인():
    ap = argparse.ArgumentParser()
    ap.add_argument("--진짜", dest="진짜", action="store_true", help="진짜로 버튼 누르고 메일 보내요")
    ap.add_argument("--보이게", dest="보이게", action="store_true")
    ap.add_argument("--중복확인없이강행", dest="강행", action="store_true",
                    help="위험! 중복 확인이 안 되는 상태에서도 발송해요")
    ap.add_argument("--최대", dest="최대", type=int, default=10,
                    help="한 번에 보낼 수 있는 최대 통수 (안전벨트, 기본 10)")
    인자 = ap.parse_args()

    준비물점검()
    설정 = 발송기.설정읽기()
    링크표 = 발송기.링크표읽기()
    기록 = 발송기.발송기록읽기()

    # 중복 발송 안전장치 — 자동발송기와 똑같은 규칙을 써요 (자세한 이유는 그쪽 주석에)
    보낸것, 막힘 = 발송기.이미보낸것모으기(설정, 인자.진짜, 인자.강행)
    if 막힘:
        sys.exit(3)

    with sync_playwright() as p:
        세션 = 쓸수있는세션(p)
        브라우저 = p.chromium.launch(headless=not 인자.보이게)
        화면 = 브라우저.new_context(storage_state=세션, locale="ko-KR", accept_downloads=True)
        쪽 = 화면.new_page()

        # ── 새 주문 확인 (발주 관리 = 결제 완료)
        print("🌐 발주 관리(결제 완료) 확인 중…")
        화면열기(쪽, 발주관리, "발주관리")
        새건수 = 총건수(쪽)
        print(f"   결제 완료 주문: {새건수 if 새건수 is not None else '?'}건")

        발주주문들 = []
        if 새건수 != 0:
            엑셀 = 엑셀받기(쪽, 발주엑셀후보, "발주")
            if 엑셀:
                # 엑셀에 지난 주문까지 들어올 수 있어서, 지금 이 화면에 보이는 것만 남겨요
                발주주문들 = 목록에보이는것만(쪽, 발송기.엑셀읽기(엑셀), "발주 관리")
            else:
                print("   ⚠️ 발주 엑셀을 못 받았어요")
                찰칵(쪽, "발주_엑셀실패")

        # ── 발송 관리에 남아있는 것도 확인 (지난번에 반쯤 처리된 주문 구출용)
        화면열기(쪽, 발송관리, "발송관리")
        발송대기건수 = 총건수(쪽)
        발송관리주문들 = []
        if 발송대기건수 and 발송대기건수 > 0:
            엑셀2 = 엑셀받기(쪽, 발송엑셀후보, "발송관리")
            if 엑셀2:
                발송관리주문들 = 목록에보이는것만(쪽, 발송기.엑셀읽기(엑셀2), "발송 관리")

        # ── 분류
        전체주문 = {j["주문번호"]: j for j in (발주주문들 + 발송관리주문들)}.values()
        보낼것, 보류, _ = 발송기.분류(list(전체주문), 링크표, 보낸것)

        발주번호들 = {j["주문번호"] for j in 발주주문들}
        # ①에서 체크할 것: 발주 관리에 있으면서, 메일 보낼 수 있거나 이미 보낸 것
        준비중대상 = [it["주문"]["주문번호"] for it in 보낼것 if it["주문"]["주문번호"] in 발주번호들]
        준비중대상 += [번 for 번 in (보낸것 & 발주번호들)]

        print()
        print(f"📋 계획")
        print(f"   ① 상품준비중 처리: {len(준비중대상)}건")
        print(f"   ② 메일 발송: {len(보낼것)}건")
        for it in 보낼것:
            print(f"      · {it['주문']['주문자']} | {it['테마']} | {it['기종']} → {it['이메일']}")
        if 보류:
            print(f"   ⏸ 건드리지 않고 보류: {len(보류)}건 (→ 보류_수동처리.csv)")
            for it in 보류:
                print(f"      · {it['주문자']} | {it['옵션']} ↳ {it['사유']}")
            발송기.보류저장(보류)

        if not 보낼것 and not 준비중대상:
            print("\n💤 처리할 게 없어요. 끝!")
            브라우저.close()
            return

        if not 인자.진짜:
            print("\n🧪 연습 모드 — 버튼도 안 눌렀고 메일도 안 나갔어요.")
            print("   진짜로 하려면:  python3 에이블리_전자동.py --진짜")
            브라우저.close()
            return

        # 안전벨트: 평소보다 너무 많으면 일단 멈춰요 (엑셀에 지난 주문이 딸려온 경우 대비)
        if len(보낼것) > 인자.최대:
            print(f"\n🛑 보낼 메일이 {len(보낼것)}통이라 안전벨트({인자.최대}통)에 걸렸어요.")
            print("   위 목록이 정말 다 보내야 할 새 주문이 맞는지 먼저 확인해 주세요.")
            print(f"   맞으면:  python3 에이블리_전자동.py --진짜 --최대 {len(보낼것)}")
            브라우저.close()
            sys.exit(3)

        # ── ① 상품준비중 처리
        if 준비중대상:
            print("\n① 상품준비중 처리")
            화면열기(쪽, 발주관리, "발주관리")
            표에서처리(쪽, 준비중대상, 준비중버튼후보, "상품준비중처리")

        # ── ② 메일 발송
        발송성공번호 = []
        if 보낼것:
            if not 발송기.계정채워짐(설정):
                print("\n❌ 메일 계정이 설정 안 돼서 ②③은 못 해요. 설정.json 을 채워주세요.")
                브라우저.close()
                sys.exit(1)
            print("\n② 메일 발송")
            성공목록, 실패 = 발송기.메일들보내기(설정, 보낼것, 기록)
            발송성공번호 = [it["주문"]["주문번호"] for it in 성공목록]

        # ── ③ 발송 처리 (메일이 나간 주문만!)
        발송처리대상 = set(발송성공번호) | (보낸것 & {j["주문번호"] for j in 전체주문})
        if 발송처리대상:
            print("\n③ 발송 처리")
            화면열기(쪽, 발송관리, "발송관리")
            표에서처리(쪽, sorted(발송처리대상), 발송처리버튼후보, "발송처리", 개별메뉴=개별변경후보)

        브라우저.close()

    print()
    print(f"🎉 끝! 메일 {len(발송성공번호)}건 발송 + 상태 처리까지 완료")
    if 보류:
        print(f"⚠️  보류 {len(보류)}건은 보류_수동처리.csv 를 봐주세요")


if __name__ == "__main__":
    메인()
