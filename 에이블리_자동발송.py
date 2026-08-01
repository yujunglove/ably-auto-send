# -*- coding: utf-8 -*-
"""
에이블리 주문 엑셀 → 고객 이메일로 테마 다운로드 링크 자동 발송

기본은 '연습 모드'예요. 진짜로 보내려면 --진짜발송 을 붙여야 해요.

    python3 에이블리_자동발송.py                      # 연습 (아무것도 안 나감)
    python3 에이블리_자동발송.py --진짜발송            # 실제 발송
    python3 에이블리_자동발송.py --파일 어쩌구.xlsx    # 특정 엑셀 지정
"""

import argparse
import csv
import email as 이메일모듈
import imaplib
import json
import os
import re
import smtplib
import ssl
import sys
import unicodedata
from email.header import decode_header, make_header
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr

import openpyxl

여기 = os.path.dirname(os.path.abspath(__file__))
설정파일 = os.path.join(여기, "설정.json")
링크표파일 = os.path.join(여기, "테마링크.json")
발송기록파일 = os.path.join(여기, "발송기록.json")
보류파일 = os.path.join(여기, "보류_수동처리.csv")

# 엑셀에서 찾을 열 이름 (에이블리가 순서를 바꿔도 이름으로 찾아요)
# 발주 관리/배송 완료 등 화면마다 열 구성이 조금 달라서, 꼭 필요한 것만 필수예요
필수열 = {
    "주문번호": "상품주문번호",
    "상품명": "상품명",
    "옵션": "옵션 정보",
}
있으면좋은열 = {
    "주문자": "주문자명",
    "연락처": "연락처",
    "결제일": "결제일",
}


# ─────────────────────────────────────── 설정 읽기

def 설정읽기():
    설정 = {}
    # 클라우드에서는 설정.json 을 통째로 시크릿(SETTINGS_JSON)에 넣어 써요.
    # 이게 있어야 서비스(덤) 테마·블로그링크·가게이름까지 그대로 따라와요.
    통째로 = os.environ.get("SETTINGS_JSON")
    if 통째로:
        try:
            설정 = json.loads(통째로)
        except Exception as e:
            print(f"❌ SETTINGS_JSON 을 못 읽었어요 (JSON 형식 확인): {e}")
            sys.exit(1)
    elif os.path.exists(설정파일):
        with open(설정파일, encoding="utf-8") as f:
            설정 = json.load(f)
    elif not (os.environ.get("GMAIL_ADDR") and os.environ.get("GMAIL_APP_PASSWORD")):
        # 클라우드(GitHub Actions)에서는 파일 없이 환경변수로만 돌 수 있어요
        print(f"❌ 설정 파일이 없어요: {설정파일}")
        print("   설정_예시.json 을 복사해서 설정.json 으로 만들고 채워주세요.")
        sys.exit(1)
    # 환경변수가 있으면 그게 우선 (GitHub Actions용 — 변수 이름은 영문만 돼요)
    설정["보내는메일"] = os.environ.get("GMAIL_ADDR") or 설정.get("보내는메일", "")
    설정["앱비밀번호"] = os.environ.get("GMAIL_APP_PASSWORD") or 설정.get("앱비밀번호", "")

    # 예시 파일 그대로 두면(한글 안내문이 들어 있으면) 안 채운 걸로 봐요
    if not 계정채워짐(설정):
        설정["앱비밀번호"] = ""
    return 설정


def 계정채워짐(설정):
    주소 = 설정.get("보내는메일", "")
    비번 = (설정.get("앱비밀번호") or "").replace(" ", "")
    if "@" not in 주소 or 주소.startswith("여기에"):
        return False
    if not 비번 or not 비번.isascii() or len(비번) < 8:
        return False
    return True


def 메일로그인(서버, 설정):
    """네이버는 아이디만, 지메일은 전체 주소로 로그인해서 둘 다 시도해요."""
    주소, 비번 = 설정["보내는메일"], 설정["앱비밀번호"]
    마지막오류 = None
    for 이름 in [주소.split("@")[0], 주소]:
        try:
            서버.login(이름, 비번)
            return
        except Exception as e:
            마지막오류 = e
    raise 마지막오류


def 메일서버들(설정):
    """보내는 메일 주소를 보고 SMTP/IMAP 서버를 골라요 (네이버·지메일 자동)."""
    주소 = 설정.get("보내는메일", "").lower()
    if 주소.endswith("@naver.com"):
        return "smtp.naver.com", "imap.naver.com"
    if 주소.endswith("@gmail.com"):
        return "smtp.gmail.com", "imap.gmail.com"
    return 설정.get("SMTP서버", "smtp.gmail.com"), 설정.get("IMAP서버", "imap.gmail.com")


def 링크표읽기():
    # GitHub Actions에서는 시크릿(JSON 통째로)으로 넘길 수 있어요
    시크릿 = os.environ.get("THEME_LINKS_JSON")
    if 시크릿:
        return json.loads(시크릿)
    if not os.path.exists(링크표파일):
        print(f"❌ 테마 링크표가 없어요: {링크표파일}")
        sys.exit(1)
    with open(링크표파일, encoding="utf-8") as f:
        return json.load(f)


def 발송기록읽기():
    if not os.path.exists(발송기록파일):
        return {}
    with open(발송기록파일, encoding="utf-8") as f:
        return json.load(f)


def 발송기록쓰기(기록):
    with open(발송기록파일, "w", encoding="utf-8") as f:
        json.dump(기록, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────── 보낸편지함으로 중복 확인
# 메일 제목에 "(주문 640650610)" 이 들어 있어서, 지메일 보낸편지함만 뒤지면
# 기록 파일이 없어도 이미 보낸 주문을 알 수 있어요. (클라우드에서 돌릴 때 중요!)

주문번호패턴 = re.compile(r"주문\s*(\d{6,})")


def _제목들에서_주문번호(서버, 번호들):
    찾은것 = set()
    결과, 응답 = 서버.fetch(b",".join(번호들), "(BODY.PEEK[HEADER.FIELDS (SUBJECT)])")
    if 결과 == "OK":
        for 조각 in 응답:
            if not isinstance(조각, tuple):
                continue
            제목 = str(make_header(decode_header(
                이메일모듈.message_from_bytes(조각[1]).get("Subject", "")
            )))
            for m in 주문번호패턴.finditer(제목):
                찾은것.add(m.group(1))
    return 찾은것


def 보낸편지함_주문번호들(설정, 며칠전까지=90):
    주소, 비번 = 설정.get("보내는메일"), 설정.get("앱비밀번호")
    if not 주소 or not 비번:
        return set(), "메일 계정 정보가 없어요"
    _, IMAP주소 = 메일서버들(설정)
    try:
        서버 = imaplib.IMAP4_SSL(IMAP주소)
        메일로그인(서버, 설정)
    except Exception as e:
        return set(), f"메일함 접속 실패({IMAP주소}): {e}"

    try:
        if "gmail" in IMAP주소:
            # 지메일은 전체보관함에서 지메일 전용 검색으로 한 방에
            전체함 = None
            결과, 목록 = 서버.list()
            if 결과 == "OK":
                for 줄 in 목록:
                    글 = 줄.decode("utf-8", "ignore") if isinstance(줄, bytes) else str(줄)
                    if "\\All" in 글:
                        전체함 = 글.split(' "/" ')[-1].strip().strip('"')
                        break
            서버.select(전체함 or "INBOX", readonly=True)
            결과, 데이터 = 서버.search(None, "X-GM-RAW", f'in:sent newer_than:{며칠전까지}d subject:주문')
            if 결과 != "OK" or not 데이터 or not 데이터[0]:
                return set(), ""
            return _제목들에서_주문번호(서버, 데이터[0].split()), ""

        # 네이버 등 일반 메일: 보낸메일함을 찾아서 최근 것들 제목만 훑어요
        보낸함 = None
        결과, 목록 = 서버.list()
        if 결과 == "OK":
            for 줄 in 목록:
                글 = 줄.decode("utf-8", "ignore") if isinstance(줄, bytes) else str(줄)
                if "\\Sent" in 글 or "Sent" in 글:
                    보낸함 = 글.split(' "/" ')[-1].strip().strip('"')
                    break
        if not 보낸함:
            return set(), "보낸메일함 폴더를 못 찾았어요"
        # 폴더 이름에 띄어쓰기가 있으면(예: Sent Messages) 따옴표로 감싸야 해요
        서버.select(f'"{보낸함.strip(chr(34))}"', readonly=True)

        기준일 = (datetime.now() - timedelta(days=며칠전까지)).strftime("%d-%b-%Y")
        결과, 데이터 = 서버.search(None, "SINCE", 기준일)
        if 결과 != "OK" or not 데이터 or not 데이터[0]:
            return set(), ""
        번호들 = 데이터[0].split()[-500:]   # 최근 500통이면 충분해요
        return _제목들에서_주문번호(서버, 번호들), ""
    except Exception as e:
        return set(), f"보낸편지함 확인 중 문제: {e}"
    finally:
        try:
            서버.logout()
        except Exception:
            pass


# ─────────────────────────────────────── 엑셀 읽기

def 최신엑셀찾기(폴더):
    """폴더에서 '에이블리'가 들어간 제일 최근 xlsx 를 골라요."""
    폴더 = os.path.expanduser(폴더)
    후보 = []
    if os.path.isdir(폴더):
        for 이름 in os.listdir(폴더):
            if 이름.startswith("~$") or not 이름.endswith(".xlsx"):
                continue
            if "에이블리" not in unicodedata.normalize("NFC", 이름):
                continue
            전체 = os.path.join(폴더, 이름)
            후보.append((os.path.getmtime(전체), 전체))
    if not 후보:
        return None
    후보.sort(reverse=True)
    return 후보[0][1]


def 엑셀읽기(경로):
    wb = openpyxl.load_workbook(경로, data_only=True, read_only=True)
    ws = wb.worksheets[0]

    행들 = list(ws.iter_rows(values_only=True))
    if not 행들:
        return []

    # 머리글 줄 찾기 (보통 1행이지만 안내문이 위에 붙는 경우 대비)
    머리글줄 = None
    for i, 행 in enumerate(행들[:10]):
        값 = [str(c).strip() if c is not None else "" for c in 행]
        if "상품주문번호" in 값 and "옵션 정보" in 값:
            머리글줄 = i
            열이름 = 값
            break
    if 머리글줄 is None:
        print("❌ 엑셀에서 '상품주문번호'/'옵션 정보' 열을 못 찾았어요. 에이블리 양식이 바뀌었을 수 있어요.")
        sys.exit(1)

    위치 = {}
    for 키, 이름 in 필수열.items():
        if 이름 not in 열이름:
            print(f"❌ '{이름}' 열이 없어요.")
            sys.exit(1)
        위치[키] = 열이름.index(이름)
    for 키, 이름 in 있으면좋은열.items():
        위치[키] = 열이름.index(이름) if 이름 in 열이름 else None

    주문들 = []
    for 행 in 행들[머리글줄 + 1:]:
        if 행 is None:
            continue
        def 값(키):
            i = 위치[키]
            if i is None:
                return ""
            return str(행[i]).strip() if i < len(행) and 행[i] is not None else ""
        주문번호 = 값("주문번호")
        if not 주문번호:
            continue
        주문들.append({
            "주문번호": 주문번호,
            "상품명": 값("상품명"),
            "옵션": 값("옵션"),
            "주문자": 값("주문자") or "고객",
            "연락처": 값("연락처"),
            "결제일": 값("결제일"),
        })
    return 주문들


# ─────────────────────────────────────── 옵션 해석

메일패턴 = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

기종사전 = {
    "아이폰": "아이폰", "아이폰(ios)": "아이폰", "ios": "아이폰", "iphone": "아이폰",
    "갤럭시": "갤럭시", "안드로이드": "갤럭시", "android": "갤럭시", "갤럭시(안드로이드)": "갤럭시",
    "둘다": "둘 다", "둘 다": "둘 다", "둘다(아이폰+갤럭시)": "둘 다", "both": "둘 다",
}


def 이름정리(글자):
    """띄어쓰기·대소문자·특수문자 무시하고 비교하려고 납작하게 만들어요."""
    글자 = unicodedata.normalize("NFC", 글자)
    글자 = re.sub(r"[\s\-_·,\.]+", "", 글자)
    return 글자.lower()


def 옵션해석(옵션):
    """'아이폰/CHROME NOIR/abc@naver.com' → (기종, 테마명, 이메일)"""
    옵션 = unicodedata.normalize("NFC", 옵션)

    # 이메일은 어디에 있든 정규식으로 뽑아요 (고객이 순서를 바꿔 적는 경우 대비)
    메일찾기 = 메일패턴.search(옵션)
    이메일 = 메일찾기.group(0).strip().rstrip(".").lower() if 메일찾기 else ""

    나머지 = 옵션
    if 이메일:
        나머지 = 옵션[:메일찾기.start()] + 옵션[메일찾기.end():]

    조각 = [c.strip() for c in 나머지.split("/") if c.strip()]

    기종 = ""
    테마 = ""
    for c in 조각:
        후보 = 기종사전.get(이름정리(c))
        if 후보 and not 기종:
            기종 = 후보
        elif not 테마:
            테마 = c
    if not 테마 and 조각:
        테마 = 조각[-1]

    return 기종, 테마, 이메일


def 링크찾기(링크표, 테마, 기종):
    """테마 이름으로 링크표에서 다운로드 주소를 찾아요. 이름이 조금 달라도 잡아요.
    돌려주는 것: (링크들, 문제, 정식이름) — 정식이름은 링크표에 적힌 원래 이름이에요."""
    # '_설명' 같은 메모 항목은 테마가 아니니 빼요
    납작한표 = {이름정리(k): (k, v) for k, v in 링크표.items() if not k.startswith("_")}
    키 = 이름정리(테마)

    짝 = 납작한표.get(키)
    if 짝 is None:
        # 부분 일치도 한 번 시도 (예: '베이비' → '레오파드 베이비')
        맞는것 = [(원래, v) for k, (원래, v) in 납작한표.items() if k and (k in 키 or 키 in k)]
        if len(맞는것) == 1:
            짝 = 맞는것[0]
    if 짝 is None:
        return None, "링크표에 없는 테마", 테마
    정식이름, 항목 = 짝

    if 기종 == "둘 다":
        아 = 항목.get("아이폰")
        갤 = 항목.get("갤럭시")
        if not 아 or not 갤:
            return None, "둘 다인데 한쪽 링크가 비어 있음", 정식이름
        return {"아이폰": 아, "갤럭시": 갤}, "", 정식이름
    if not 기종:
        return None, "기종을 못 읽음", 정식이름

    주소 = 항목.get(기종)
    if not 주소:
        return None, f"{기종} 링크가 비어 있음", 정식이름
    return {기종: 주소}, "", 정식이름


# ─────────────────────────────────────── 메일 만들기

def 메일본문(설정, 주문, 테마, 링크들):
    가게 = 설정.get("가게이름", "죠미메이드")
    블로그 = 설정.get("블로그링크", "")
    서비스 = 설정.get("서비스테마", {})   # {"이름": "실버체인", "아이폰": 링크, "갤럭시": 링크}

    줄 = []
    줄.append(f"안녕하세요, {가게}입니다 🎀")
    줄.append("카카오톡 테마를 구매해 주셔서 정말 감사합니다!")
    줄.append("")
    줄.append("Gmail 정책상 파일이 안 열릴 수도 있어 구글 드라이브로 보내드립니다!")
    for 기종, 주소 in 링크들.items():
        줄.append(f"👉 '{테마}' ({기종}용)")
        줄.append(주소)
    줄.append("")
    줄.append("테마 파일을 전달드렸으니 사용하시는 기종에 맞는 파일인지 확인한 후 적용해 주세요.")
    줄.append("파일에 문제가 있거나 정상적으로 적용되지 않는 경우에는 구매확정 전 문의를 남겨주시면 확인해 드리겠습니다.")

    # 서비스(덤) 테마 — 주문한 기종에 맞는 링크가 있을 때만 끼워요
    서비스이름 = 서비스.get("이름", "")
    덤링크들 = [(기, 서비스.get(기)) for 기 in 링크들 if 서비스.get(기)]
    if 서비스이름 and 덤링크들:
        줄.append("")
        줄.append(f"🎁 서비스 테마 '{서비스이름}'도 함께 드립니다. 이 테마는 곧 판매 예정입니다 :)")
        for 기, 링 in 덤링크들:
            줄.append(링)

    if 블로그:
        줄.append("")
        줄.append(f"{가게} 블로그도 놀러오세요 💕")
        줄.append("˚⊹♡ 놀러와요 죠미의 숲 ♡⊹˚₊")
        줄.append(블로그)

    줄.append("")
    줄.append("📌 이용 전 주의사항")
    줄.append("· 본 상품은 디지털 파일로, 파일 전달 후에는 단순 변심이나 기종 착오로 인한 교환·환불이 어렵습니다.")
    줄.append("· 구매자 본인만 사용할 수 있으며, 파일 공유·재판매·무단 배포·수정 후 배포는 금지되어 있습니다.")
    줄.append("· 카카오톡 업데이트 또는 기기 설정에 따라 일부 화면이 상세 이미지와 다르게 보일 수 있습니다.")
    줄.append("· 파일 삭제에 대비하여 전달받은 파일을 별도로 보관해 주세요.")
    줄.append("")
    줄.append("예쁘게 제작한 테마인 만큼 오래오래 잘 사용해 주세요 💗")
    return "\n".join(줄)


def 메일만들기(설정, 주문, 테마, 링크들, 받는사람):
    가게 = 설정.get("가게이름", "죠미메이드")
    쪽지 = EmailMessage()
    # 끝의 (주문 번호)는 '보낸편지함 검색'으로 중복 발송을 막는 데 쓰여요 — 빼면 안 돼요
    쪽지["Subject"] = f"안녕하세요. {가게}입니다. '{테마}' 보내드립니다! (주문 {주문['주문번호']})"
    쪽지["From"] = formataddr((가게, 설정["보내는메일"]))
    쪽지["To"] = 받는사람
    if 설정.get("숨은참조"):
        쪽지["Bcc"] = 설정["숨은참조"]
    쪽지.set_content(메일본문(설정, 주문, 테마, 링크들))
    return 쪽지


# ─────────────────────────────────────── 분류·발송 (전자동 스크립트에서도 그대로 가져다 써요)

def 분류(주문들, 링크표, 보낸것):
    """주문 목록을 '보낼 수 있는 것 / 사람이 봐야 하는 것 / 이미 보낸 것'으로 나눠요."""
    보낼것, 보류, 이미보냄 = [], [], []
    for 주문 in 주문들:
        if 주문["주문번호"] in 보낸것:
            이미보냄.append(주문)
            continue
        기종, 테마, 이메일 = 옵션해석(주문["옵션"])
        if not 이메일:
            보류.append({**주문, "사유": "이메일을 안 적으셨어요", "기종": 기종, "테마": 테마})
            continue
        링크들, 문제, 정식이름 = 링크찾기(링크표, 테마, 기종)
        if 링크들 is None:
            보류.append({**주문, "사유": 문제, "기종": 기종, "테마": 정식이름})
            continue
        보낼것.append({"주문": 주문, "기종": 기종, "테마": 정식이름, "이메일": 이메일, "링크들": 링크들})
    return 보낼것, 보류, 이미보냄


def 보류저장(보류):
    if not 보류:
        return
    with open(보류파일, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["주문번호", "주문자", "연락처", "상품명", "옵션", "기종", "테마", "사유"])
        w.writeheader()
        for it in 보류:
            w.writerow({k: it.get(k, "") for k in w.fieldnames})


def 메일들보내기(설정, 보낼것, 기록):
    """한 통 보낼 때마다 기록을 저장해서, 중간에 멈춰도 중복이 안 나요."""
    맥락 = ssl.create_default_context()
    성공목록, 실패 = [], []
    SMTP주소, _ = 메일서버들(설정)
    with smtplib.SMTP_SSL(SMTP주소, 465, context=맥락) as 서버:
        메일로그인(서버, 설정)
        for it in 보낼것:
            쪽지 = 메일만들기(설정, it["주문"], it["테마"], it["링크들"], it["이메일"])
            try:
                서버.send_message(쪽지)
            except Exception as e:
                실패.append((it, str(e)))
                print(f"   ❌ {it['이메일']} — {e}")
                continue
            기록[it["주문"]["주문번호"]] = {
                "보낸시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "받는사람": it["이메일"],
                "테마": it["테마"],
                "기종": it["기종"],
                "주문자": it["주문"]["주문자"],
            }
            발송기록쓰기(기록)
            성공목록.append(it)
            print(f"   ✅ {it['주문']['주문자']} → {it['이메일']}")
    return 성공목록, 실패


# ─────────────────────────────────────── 중복 발송 안전장치

def 이미보낸것모으기(설정, 진짜발송, 강행=False):
    """'이미 보낸 주문번호' 를 모아요. 못 모으면 발송을 막아요.  ★안전장치★

    중복 발송을 막는 방법이 두 가지 있어요:
      ① 발송기록.json  — 내 컴퓨터에 남는 기록
      ② 보낸편지함 조회 — 메일 제목의 '(주문 12345)' 를 IMAP 으로 찾기

    맥에서 돌릴 때는 ①이 항상 있어서 괜찮아요. 그런데 GitHub Actions 같은
    클라우드에서는 ①이 매번 빈 상태로 시작해요(깃에 안 올리는 파일이라서).
    그래서 그때 ②까지 실패하면 '보낸 적 없음' 으로 보여서, 지난 고객 전원에게
    메일이 다시 나가요. 실제로 그럴 뻔했어요.

    그래서 둘 다 없으면 그냥 멈춰요. 안 보내는 쪽이 두 번 보내는 것보다 나아요.
    돌려주는 것: (보낸것집합, 막혔는지)
    """
    기록 = 발송기록읽기()
    보낸것 = set(기록.keys())
    기록파일있음 = os.path.exists(발송기록파일)
    메일함확인함 = False

    if 설정.get("중복확인_보낸편지함", True) and 계정채워짐(설정):
        메일에서찾은것, 문제 = 보낸편지함_주문번호들(설정)
        if 문제:
            print(f"ℹ️  보낸편지함 확인은 건너뛰었어요 ({문제})")
        else:
            메일함확인함 = True
            새로안것 = 메일에서찾은것 - 보낸것
            if 새로안것:
                print(f"📬 보낸편지함에서 이미 보낸 주문 {len(새로안것)}건을 더 찾았어요")
            보낸것 |= 메일에서찾은것
        print()

    if 진짜발송 and not 메일함확인함 and not 기록파일있음:
        print("🛑 중복 발송을 막을 방법이 하나도 없어요. 그래서 발송을 멈췄어요.")
        print(f"   · 발송기록 파일 없음 ({os.path.basename(발송기록파일)})")
        print("   · 보낸편지함 확인도 실패")
        print("   이대로 보내면 예전에 보낸 고객에게 메일이 또 갈 수 있어요.")
        print("   → 메일 설정(IMAP)을 고치고 다시 돌려주세요.")
        if not 강행:
            return 보낸것, True
        print("   ⚠️ --중복확인없이강행 을 주셔서 그대로 진행해요. 정말 조심하세요!")

    return 보낸것, False


# ─────────────────────────────────────── 실행

def 실행(엑셀경로, 진짜발송, 강행=False):
    설정 = 설정읽기()
    링크표 = 링크표읽기()

    주문들 = 엑셀읽기(엑셀경로)
    print(f"📄 {os.path.basename(엑셀경로)} — 주문 {len(주문들)}건")
    print()

    보낸것, 막힘 = 이미보낸것모으기(설정, 진짜발송, 강행)
    if 막힘:
        sys.exit(3)
    기록 = 발송기록읽기()

    보낼것, 보류, 이미보냄 = 분류(주문들, 링크표, 보낸것)

    # ── 결과 보여주기
    if 이미보냄:
        print(f"⏭  이미 보낸 주문 {len(이미보냄)}건은 건너뛰었어요")
    print(f"✅ 보낼 수 있는 주문: {len(보낼것)}건")
    for it in 보낼것:
        print(f"   · {it['주문']['주문자']} | {it['테마']} | {it['기종']} → {it['이메일']}")
    print()

    if 보류:
        print(f"⚠️  사람이 봐야 하는 주문: {len(보류)}건  (→ {os.path.basename(보류파일)})")
        for it in 보류:
            print(f"   · {it['주문자']} ({it['연락처']}) | {it['상품명'][:28]} | 옵션: {it['옵션']}")
            print(f"     ↳ {it['사유']}")
        보류저장(보류)
        print()

    if not 보낼것:
        print("보낼 게 없어요. 끝!")
        return

    if not 진짜발송:
        print("🧪 연습 모드였어요 — 메일은 하나도 안 나갔어요.")
        print("   진짜로 보내려면:  python3 에이블리_자동발송.py --진짜발송")
        return

    # ── 실제 발송
    if not 계정채워짐(설정):
        print("❌ 설정.json 의 '보내는메일'/'앱비밀번호'를 아직 안 채우셨어요.")
        print("   지메일 → 계정 관리 → 보안 → 2단계 인증 → 앱 비밀번호 에서 16자리를 만들어 넣어주세요.")
        sys.exit(1)

    print("📮 보내는 중…")
    성공목록, 실패 = 메일들보내기(설정, 보낼것, 기록)

    print()
    print(f"🎉 {len(성공목록)}건 발송 완료" + (f" / {len(실패)}건 실패" if 실패 else ""))


def 메인():
    ap = argparse.ArgumentParser()
    ap.add_argument("--파일", dest="파일", help="에이블리 주문 엑셀 경로")
    ap.add_argument("--진짜발송", dest="진짜발송", action="store_true", help="진짜로 메일을 보내요")
    ap.add_argument("--중복확인없이강행", dest="강행", action="store_true",
                    help="위험! 중복 확인이 안 되는 상태에서도 발송해요")
    인자 = ap.parse_args()

    경로 = 인자.파일
    if not 경로:
        설정 = 설정읽기()
        경로 = 최신엑셀찾기(설정.get("엑셀폴더", "~/Downloads"))
    if not 경로 or not os.path.exists(경로):
        print("❌ 주문 엑셀을 못 찾았어요. --파일 로 경로를 알려주세요.")
        sys.exit(1)

    실행(경로, 인자.진짜발송, 인자.강행)


if __name__ == "__main__":
    메인()
