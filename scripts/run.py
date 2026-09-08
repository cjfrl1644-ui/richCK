# -*- coding: utf-8 -*-
"""
안경원 자동 정산
  02_카드내역/*.xlsx      신한카드 이용내역
  03_데이터/계좌내역.xlsx  신한은행 입출금 내역
      → 분류 → 원장 → 월별 리포트 + 대시보드

실행: 실행.bat 더블클릭

■ 결제일과 귀속월
  거래처·임대료·인건비는 '전월분'을 다음 달에 낸다.
  그래서 거래마다 날짜를 두 개 붙인다.
    결제월 = 돈이 실제로 오간 달
    귀속월 = 그 물건/서비스/노동이 발생한 달
  순이익은 '귀속 기준'으로 봐야 매출과 시점이 맞는다.

■ 카드대금은 비용이 아니다
  계좌에서 나가는 '신한카드 13,105,280' 같은 건 카드 이용내역에
  건별로 이미 다 잡혀 있다. 비용으로 또 세면 이중계산이므로 제외한다.
"""
import base64
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import warnings

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# 신한카드 엑셀에는 기본 서식이 없어 경고가 뜨는데, 읽는 데는 지장이 없다.
warnings.filterwarnings("ignore", message=".*no default style.*")
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent.parent
CARD_DIR = BASE / "02_카드내역"
DATA_DIR = BASE / "03_데이터"
REPORT_DIR = BASE / "04_리포트"
SALES_DIR = BASE / "06_매출"
MASTER_XLSX = DATA_DIR / "거래처마스터.xlsx"
BANK_XLSX = DATA_DIR / "계좌내역.xlsx"
LEDGER_XLSX = DATA_DIR / "원장.xlsx"
FIXED_XLSX = DATA_DIR / "월고정비.xlsx"
PASSWORD_TXT = DATA_DIR / "비밀번호.txt"
DASHBOARD = BASE / "대시보드.html"

# 대시보드 잠금에 쓰는 값 (브라우저 표준 Web Crypto 와 같은 방식)
PBKDF2_ROUNDS = 250_000

# ── 분류 ───────────────────────────────────────────────────────────────
# 매출은 06_매출 폴더의 일일결산 엑셀에서만 가져온다.
# 통장 입금(카드사 정산)은 매출이 아니라 '돈이 들어온 시점'일 뿐이라
# 매출로 세면 이중계산이 되고 금액도 시차·수수료 때문에 다르다.
INCOME = ["매출", "기타수입"]
REFUND = ["매출환불"]        # 손님에게 돌려준 돈 → 매출에서 뺀다
# 식대 = 매장에서 직원들이 먹는 밥값. 복리후생비 성격이라 사업 비용이 맞다.
# (사장 개인 장보기·쇼핑은 '개인'으로 남겨 순이익에서 뺀다)
COST = ["거래처매입", "인건비", "임대료", "식대", "사업경비", "금융비용"]
COST_REFUND = ["매입환불"]   # 거래처가 돌려준 돈 → 비용에서 뺀다
# 순이익 계산에서 빠지는 것들
#  임대료납부 : 이마트에 실제로 나간 돈. 임대료는 월고정비 750만원으로 잡으므로
#               이걸 또 비용으로 세면 이중계산이 된다. 매달 상품권을 얼마나
#               입금했느냐에 따라 계좌 출금액이 달라져서 그때그때 다르다.
NEUTRAL = ["개인", "카드대금", "정산입금", "임대료납부", "미분류"]
CATEGORIES = INCOME + REFUND + COST + COST_REFUND + NEUTRAL

CYCLES = ["전월분", "당월분"]

# 카드 수수료율.
#   매출은 손님이 낸 총액으로 잡는데, 카드사는 수수료를 떼고 입금한다.
#   그 차액이 비용이므로 '카드매출 × 이 비율' 을 매달 자동으로 뺀다.
#   2.8% 는 7월 실측값(카드매출 31,175,400 vs 통장 입금 30,300,141)에서 나온 어림값.
#   정확한 수수료율을 알게 되면 이 숫자만 고치면 된다.
CARD_FEE_RATE = 0.028
LATE_DAY = 20
DUP_MIN_AMOUNT = 50_000
DUP_CATEGORIES = {"거래처매입", "사업경비", "임대료", "인건비"}

# 카드사 정산 입금 코드 (이 접두어로 시작하면 매출)
CARD_SETTLE_PREFIX = ["KB9", "SHC", "삼성17", "NH1", "현126", "우601",
                      "하나99", "712673342BC", "롯데99", "카카오페이정산"]
# 이용내역 엑셀을 갖고 있는 카드 → 대금 출금은 건별로 이미 잡혀 있으므로 제외
# (02_카드내역 에 다른 카드사 내역이 추가되면 여기에도 넣어야 한다)
STATEMENT_CARDS = ["신한카드"]
# 이용내역이 없는 카드 → 대금 출금 자체를 비용으로 잡아야 한다.
# 제외해 버리면 그 카드로 쓴 돈이 통째로 사라진다.
OTHER_CARDS = ["롯데카드", "우리카드", "삼성카드", "현대카드",
               "국민카드", "하나카드", "BC카드", "농협카드", "카카오뱅크카드"]

# 표기명 키워드 → (분류, 결제주기, 키워드, 적용할 원천)
# 원천이 None 이면 카드·계좌 둘 다에 적용된다.
# ※ '이마트'는 카드로 긁으면 장보기(개인), 계좌에서 나가면 임대료라서
#    원천을 구분하지 않으면 42건짜리 장보기가 통째로 임대료로 잡힌다.
RULES = [
    # 이마트에 나간 돈은 '임대료납부'(중립). 임대료 자체는 월고정비에서 잡는다.
    ("임대료납부", "전월분", ["이마트", "임대", "월세", "관리비"], {"계좌"}),
    ("거래처매입", "전월분", ["옵티칼", "옵티컬", "옵틱", "광학", "렌즈", "아이웨어",
                              "안경", "존슨앤드존슨비전", "바슈롬", "파리미키",
                              "글라스킹", "에실로", "호야", "아큐브", "케미",
                              "니콘", "자이스", "메디칼", "메디컬", "네오다임"], None),
    ("사업경비", "당월분", ["건강보험", "국민연금", "고용보험", "산재",
                            "KT", "SK브로드밴드", "LG유플러스", "통신", "네트웍스",
                            "한국전력", "도시가스", "수도",
                            "지자체세입금", "국세", "세무", "세금",
                            "주유소", "오일뱅크", "칼텍스", "에너지", "주차",
                            "손해보험", "DB손", "S1(", "SMS", "에스원", "기획사"], None),
    # 식대 — 직원 식사. '개인'보다 먼저 걸러야 한다.
    ("식대", "당월분", ["이마트", "푸드", "식당", "김치찌개", "칼국수", "순대",
                        "감자탕", "부대찌개", "떡볶이", "닭발", "국밥", "국수",
                        "곰탕", "짬뽕", "콩나물", "버거킹", "롯데리아", "맘스터치",
                        "빽다방", "카페", "베이커리", "연경", "반미", "더소희",
                        "토스트", "김밥", "돈까스", "우육면", "공차", "디저트",
                        "와플", "커피", "아딸", "죽", "에스씨케이"], None),
    ("개인", "당월분", ["다이소", "약국", "편의점", "GS25", "CU ",
                        "네이버파이낸셜", "쿠팡", "배달", "패션플러스"], None),
]

# 8/1에 급여로 나간 사람들 — 이름만으로는 알 수 없어 직접 지정한다.
# 새 직원이 생기면 거래처마스터.xlsx 에서 분류를 '인건비'로 바꿔주면 된다.
PAYROLL_NAMES = {"이철기", "김상기", "한나영"}


def norm(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip()


def short_name(name: str) -> str:
    s = re.sub(r"주식회사|\(주\)|\(유\)|㈜", "", norm(name))
    s = re.sub(r"\s*(하남|덕풍|풍산|미사)?\s*\S*점$", "", s)
    return re.sub(r"\s+", " ", s).strip() or norm(name)


def prev_month(y, m):
    return (y - 1, 12) if m == 1 else (y, m - 1)


def next_month_key(key: str) -> str:
    y, m = int(key[:4]), int(key[5:7])
    return f"{y + 1:04d}-01" if m == 12 else f"{y:04d}-{m + 1:02d}"


def attrib_month(dt: datetime, cycle: str) -> str:
    if cycle == "전월분":
        y, m = prev_month(dt.year, dt.month)
        return f"{y:04d}-{m:02d}"
    return dt.strftime("%Y-%m")


def parse_date(v):
    if isinstance(v, datetime):
        return v
    s = norm(v)
    for fmt in ("%Y.%m.%d %H:%M", "%Y.%m.%d", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


# ── 자동 분류 ──────────────────────────────────────────────────────────
def guess(source: str, label: str, direction: str, days, _unused=None):
    """(분류, 결제주기, 비고) 를 추정한다."""
    if source == "계좌":
        if direction == "수입":
            if any(label.startswith(p) for p in CARD_SETTLE_PREFIX):
                return "정산입금", "당월분", "카드사 정산 입금 — 매출은 매출엑셀에서 잡으므로 참고용"
            if "캐시백" in label:
                return "기타수입", "당월분", "카드사 캐시백"
            return "정산입금", "당월분", "계좌이체 입금 — 매출은 매출엑셀에서 잡으므로 참고용"
        if any(label.startswith(p) for p in STATEMENT_CARDS):
            return "카드대금", "당월분", "이용내역이 있어 건별로 잡힘 — 중복 방지로 제외"
        if any(label.startswith(p) for p in OTHER_CARDS):
            return ("사업경비", "전월분",
                    "이용내역 없는 카드 대금 — 전월 사용분. 개인용이면 '개인'으로 바꿔주세요")
        if label in PAYROLL_NAMES:
            return "인건비", "전월분", "급여 — 전월 근무분"

    for cat, cycle, keywords, sources in RULES:
        if sources and source not in sources:
            continue
        for kw in keywords:
            if kw.lower() in label.lower():
                # 거래처인데 결제일이 월초·월중이면 즉시결제로 본다
                if cat == "거래처매입" and days and sum(1 for d in days if d >= LATE_DAY) * 2 < len(days):
                    return cat, "당월분", f"결제일이 월초·월중({', '.join(map(str, sorted(set(days))))}일) — 즉시결제로 봄"
                return cat, cycle, "자동추정 — 확인 필요"

    if source == "계좌" and direction == "지출":
        return "미분류", "당월분", "계좌 출금 — 분류를 지정해주세요"
    return "미분류", "당월분", "자동추정 — 확인 필요"


# ── 읽기 ───────────────────────────────────────────────────────────────
def read_cards():
    files = [f for f in sorted(CARD_DIR.glob("*.xls*")) if not f.name.startswith("~$")]
    rows = []
    seen = set()                  # 이미 읽은 거래 (파일 간 중복 제거용)
    dup = defaultdict(int)        # 파일별로 몇 건이나 중복이었는지
    for f in files:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        col = {h: i for i, h in enumerate(norm(c) for c in next(it))}
        if any(c not in col for c in ("거래일", "가맹점명", "금액")):
            print(f"      [!] {f.name}: 신한카드 형식이 아닙니다 — 건너뜁니다")
            wb.close()
            continue
        get = lambda r, n: r[col[n]] if n in col else None
        for r in it:
            dt = parse_date(get(r, "거래일"))
            if dt is None:
                continue
            merchant = norm(get(r, "가맹점명"))
            amount = float(get(r, "금액") or 0)
            approval = norm(get(r, "승인번호"))
            # 기간이 겹치는 파일을 같이 넣어도 한 번만 세도록 거른다.
            # 승인번호는 거래마다 고유하다. (예: '신한카드7월'과 '4~7월' 파일을
            #  둘 다 넣으면 7월이 통째로 두 번 잡혀 비용이 부풀려진다)
            key = (dt, merchant, amount, approval)
            if key in seen:
                dup[f.name] += 1
                continue
            seen.add(key)
            rows.append({
                "원천": "카드", "거래일시": dt, "날짜": dt.strftime("%Y-%m-%d"),
                "결제월": dt.strftime("%Y-%m"), "일": dt.day, "방향": "지출",
                "표기명": merchant, "금액": amount,
                "수단": norm(get(r, "이용카드")), "참조": approval,
                "출처": f.name,
            })
        wb.close()

    for name, n in dup.items():
        print(f"      [중복 제외] {name}: {n}건은 다른 파일에 이미 있어 건너뜀")
    return rows


# 결제수단 열 이름. 한 이름에 여러 표기가 있을 수 있다.
#  · 옛 양식은 '현영', 새 양식은 '현금영수증'
#  · 8월 24일 시트처럼 '현영(상품권 포함)' 으로 적힌 날도 있다
MEANS = ["현금", "현영", "상품권", "카드", "미수"]
MEANS_ALIAS = {"현영": ["현금영수증", "현영"], "현금": ["현금"],
               "상품권": ["상품권"], "카드": ["카드"], "미수": ["미수"]}
# 품목 열.
#  · 새 양식(2026-09~)은 품목마다 칸이 따로 있고, 드롭다운에서 제품을 고른다.
#    고른 것이 곧 1건이다. 한 손님이 테와 누진을 같이 사도 한 줄이다.
#  · 옛 양식은 프레임/안경렌즈/콘택트 세 칸이었다. (콘택트는 이름이 겹친다)
ITEM_COLS = ["안경테", "단초점", "누진", "콘택트"]
OLD_ITEM_COLS = ["프레임", "안경렌즈"]


def find_columns(ws):
    """일일결산 시트에서 (헤더행, 결제수단 열, 품목 열) 을 찾는다.

    옛 양식은 헤더가 두 줄(결제수단 / 품목)로 나뉘어 있고,
    새 양식은 한 줄에 다 있다. 둘 다 읽을 수 있어야 한다.
    """
    for r in range(1, 9):
        vals = [norm(c.value) for c in ws[r]]
        if "현금" not in vals or "카드" not in vals:
            continue
        col = {}
        # 긴 이름(현금영수증)을 먼저 잡아야 '현금' 이 가로채지 않는다
        for key in sorted(MEANS, key=lambda k: -max(len(a) for a in MEANS_ALIAS[k])):
            for alias in sorted(MEANS_ALIAS[key], key=len, reverse=True):
                for i, v in enumerate(vals):
                    if i in col.values():
                        continue
                    if v == alias or (v.startswith(alias) and key != "현금"):
                        col[key] = i
                        break
                if key in col:
                    break
        # 품목 열은 같은 줄이나 바로 아랫줄에 있다
        items = {}
        for rr in (r, r + 1):
            row = [norm(c.value) for c in ws[rr]]
            for i, v in enumerate(row):
                for k in ITEM_COLS + OLD_ITEM_COLS:
                    if k not in items and v == k:
                        items[k] = i
        return r, col, items
    return None, {}, {}


def read_sales():
    """06_매출/*.xlsx (일일결산일지) 에서 월별 매출을 읽는다.

    각 파일은 '1일'~'31일' 시트를 갖고, 결제수단별로 금액이 적혀 있다.
    ※ 미수(외상)는 매출에서 뺀다. 받는 날 다시 기록하므로,
       판 시점에도 매출로 잡으면 나중에 이중으로 계산된다.
    ※ '결산' 시트의 총금액과 일치하는지 대조해봤고 한 원도 안 틀린다.
       다만 결산 시트의 결제수단별 합계 행은 깨져 있어 쓰지 않는다.
    """
    out = {}
    for f in sorted(SALES_DIR.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        # "2026년 7월", "26년 4월", "2026-07" 다 받아준다.
        # 파일명 앞에 상호명 같은 게 붙어 있어도 된다. ("굿미소 26년 4월")
        mm = re.search(r"(\d{2,4})\s*[년\-.]\s*(\d{1,2})\s*월?", f.stem)
        if not mm:
            print(f"      [!] {f.name}: 파일명에서 연·월을 못 읽었습니다 — 건너뜁니다")
            print(f"          '2026년 7월' 처럼 연도와 월이 들어가게 바꿔주세요")
            continue
        year, mon = int(mm.group(1)), int(mm.group(2))
        if year < 100:                       # '26년' → 2026년
            year += 2000
        if not 1 <= mon <= 12:
            print(f"      [!] {f.name}: 월이 {mon} 로 읽혔습니다 — 건너뜁니다")
            continue
        month = f"{year:04d}-{mon:02d}"

        wb = openpyxl.load_workbook(f, data_only=True)
        tot = dict.fromkeys(MEANS, 0.0)
        # 품목별 건수와 '그 품목이 들어간 판매액'.
        # 한 손님이 테와 누진을 같이 사면 둘 다 1건으로 세고 금액은 양쪽에
        # 똑같이 붙는다. 그래서 품목 금액을 다 더하면 총매출보다 크다.
        item_cnt = dict.fromkeys(ITEM_COLS, 0)
        item_amt = dict.fromkeys(ITEM_COLS, 0.0)
        days = 0
        for name in wb.sheetnames:
            if not re.fullmatch(r"\d+일", name):
                continue
            ws = wb[name]
            hdr, col, items = find_columns(ws)
            if hdr is None:
                continue
            # 품목 건수는 새 양식에서만 센다. 옛 양식의 '콘택트' 칸에는
            # 개수가 아니라 제품 이름이 들어가 있어서 세면 엉뚱한 값이 나온다.
            new_form = all(k in items for k in ITEM_COLS)
            day = dict.fromkeys(MEANS, 0.0)
            for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
                if not row or not isinstance(row[0], (int, float)):
                    continue          # 번호가 숫자인 행만 실제 판매 기록
                for k, i in col.items():
                    v = row[i] if i < len(row) else None
                    if isinstance(v, (int, float)):
                        day[k] += v
                paid = sum(row[col[k]] for k in ("현금", "현영", "상품권", "카드")
                           if k in col and col[k] < len(row)
                           and isinstance(row[col[k]], (int, float)))
                if new_form:
                    for k in ITEM_COLS:
                        i = items[k]
                        if i < len(row) and norm(row[i]):
                            item_cnt[k] += 1      # 고른 것이 곧 1건
                            item_amt[k] += paid
            if sum(day.values()):
                days += 1
                for k in MEANS:
                    tot[k] += day[k]
        wb.close()
        tot["총매출"] = sum(tot[k] for k in MEANS if k != "미수")
        tot["영업일"] = days
        tot["품목건수"] = item_cnt
        tot["품목금액"] = item_amt
        tot["출처"] = f.name
        out[month] = tot
    return out


def read_sales_detail():
    """일일결산에서 건별 판매 기록을 뽑는다 (품목·브랜드·객단가용).

    ※ 한 건에 프레임과 렌즈가 같이 있어도 금액은 한 줄로만 적혀 있어
       품목별로 금액을 쪼갤 수 없다. 그래서 '대표 품목' 하나에 건별
       금액 전체를 배정한다. 비중은 그 기준으로 읽어야 한다.
    ※ 원가가 없어 마진율은 계산할 수 없다. 전부 매출 기준이다.
    """
    out = {}
    for f in sorted(SALES_DIR.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        mm = re.search(r"(\d{2,4})\s*[년\-.]\s*(\d{1,2})\s*월?", f.stem)
        if not mm:
            continue
        year, mon = int(mm.group(1)), int(mm.group(2))
        if year < 100:
            year += 2000
        if not 1 <= mon <= 12:
            continue
        month = f"{year:04d}-{mon:02d}"

        wb = openpyxl.load_workbook(f, data_only=True)
        sales = []
        for name in wb.sheetnames:
            dm = re.fullmatch(r"(\d+)일", name)
            if not dm:
                continue
            ws = wb[name]
            hdr, col, items = find_columns(ws)
            if hdr is None:
                continue
            # 품목 열 — 새 양식은 이름으로 찾고, 옛 양식은 결제수단 뒤에 순서대로 있다
            if len(items) < 3:
                base = max(col.values()) + 1
                items = {k: base + n for n, k in enumerate(ITEM_COLS)}
            for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
                if not row or not isinstance(row[0], (int, float)):
                    continue
                paid = sum(row[col[k]] for k in ("현금", "현영", "상품권", "카드")
                           if k in col and col[k] < len(row)
                           and isinstance(row[col[k]], (int, float)))
                if paid <= 0:
                    continue
                get = lambda k: (norm(row[items[k]])
                                 if k in items and items[k] < len(row) else "")
                sales.append({
                    "date": f"{month}-{int(dm.group(1)):02d}",
                    "staff": norm(row[1]) if len(row) > 1 else "",
                    "customer": norm(row[2]) if len(row) > 2 else "",
                    "frame": get("프레임"), "lens": get("안경렌즈"),
                    "contact": get("콘택트"),
                    "amount": paid,
                })
        wb.close()
        out[month] = sales
    return out


def classify_item(s):
    """건별 대표 품목. (금액을 쪼갤 수 없어 한 건 = 한 품목으로 본다)"""
    blob = f"{s['frame']} {s['lens']} {s['contact']}"
    if "선글" in blob or "썬글" in blob:
        return "선글라스"
    if s["contact"]:
        return "콘택트렌즈"
    if s["lens"]:
        return "누진렌즈" if "누진" in s["lens"] else "일반렌즈"
    if s["frame"]:
        return "안경테"
    return "기타"


def mask_name(name: str) -> str:
    """고객 이름 가운데를 가린다."""
    n = norm(name)
    if len(n) <= 1:
        return n or "—"
    return n[0] + "○" * (len(n) - 1)


def read_bank():
    """03_데이터/계좌내역*.xlsx 를 모두 읽는다 (월별로 파일이 나뉘어 있어도 됨)."""
    rows = []
    for f in sorted(DATA_DIR.glob("계좌내역*.xlsx")):
        if f.name.startswith("~$"):
            continue
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        col = {h: i for i, h in enumerate(norm(c) for c in next(it))}
        if any(c not in col for c in ("날짜", "구분", "금액", "적요")):
            print(f"      [!] {f.name}: 계좌내역 형식이 아닙니다 — 건너뜁니다")
            wb.close()
            continue
        for r in it:
            if not r or not r[0]:
                continue
            dt = parse_date(r[col["날짜"]])
            if dt is None:
                continue
            kind = norm(r[col["구분"]])
            rows.append({
                "원천": "계좌", "거래일시": dt, "날짜": dt.strftime("%Y-%m-%d"),
                "결제월": dt.strftime("%Y-%m"), "일": dt.day,
                "방향": "수입" if kind == "입금" else "지출",
                "표기명": norm(r[col["적요"]]), "금액": float(r[col["금액"]] or 0),
                "수단": "계좌", "참조": norm(r[col["시각"]]) if "시각" in col else "",
                "출처": f.name,
            })
        wb.close()
    return rows


# ── 마스터 ─────────────────────────────────────────────────────────────
def load_master():
    master, first_run = {}, not MASTER_XLSX.exists()
    if not first_run:
        wb = openpyxl.load_workbook(MASTER_XLSX, data_only=True)
        ws = wb.active
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or not r[0] or not r[2]:
                continue
            c = list(r) + [None] * 10
            master[(norm(c[0]), norm(c[1]), norm(c[2]))] = {
                "거래처명": norm(c[3]) or short_name(c[2]),
                "분류": norm(c[4]) if norm(c[4]) in CATEGORIES else "미분류",
                "결제주기": norm(c[5]) if norm(c[5]) in CYCLES else "당월분",
                "비고": norm(c[6]),
            }
        wb.close()
    return master, first_run


def update_master(master, rows):
    # 같은 이름이 입금·출금 양쪽에 나올 수 있다 (예: 김상기 — 급여 출금이자 계좌이체 입금).
    # 그래서 방향까지 열쇠에 넣어 따로 관리한다.
    days = defaultdict(list)
    for r in rows:
        days[(r["원천"], r["방향"], r["표기명"])].append(r["일"])

    new = []
    for k in days:
        if k in master:
            continue
        src, direction, label = k
        cat, cycle, note = guess(src, label, direction, days[k], None)
        master[k] = {"거래처명": short_name(label), "분류": cat,
                     "결제주기": cycle, "비고": note}
        new.append(k)
    return new


def save_master(master, rows):
    last, means = {}, defaultdict(set)
    total = defaultdict(float)
    for r in rows:
        k = (r["원천"], r["방향"], r["표기명"])
        if k not in last or r["날짜"] > last[k]:
            last[k] = r["날짜"]
        means[k].add(r["수단"])
        total[k] += r["금액"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "거래처마스터"
    ws.append(["원천", "방향", "표기명(카드 가맹점명 / 계좌 적요)", "거래처명", "분류",
               "결제주기", "비고", "합계금액", "수단", "최근일"])
    order = {c: i for i, c in enumerate(CATEGORIES)}
    for k in sorted(master, key=lambda x: (order.get(master[x]["분류"], 99), -total.get(x, 0))):
        i = master[k]
        ws.append([k[0], k[1], k[2], i["거래처명"], i["분류"], i["결제주기"], i["비고"],
                   total.get(k, 0), ", ".join(sorted(means.get(k, []))), last.get(k, "")])
    style_sheet(ws, widths=[7, 7, 34, 22, 12, 10, 44, 14, 12, 13], money_cols=[8])
    for letter, opts in (("E", CATEGORIES), ("F", CYCLES)):
        dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{letter}2:{letter}{ws.max_row}")
    wb.save(MASTER_XLSX)


def apply_master(rows, master):
    for r in rows:
        i = master[(r["원천"], r["방향"], r["표기명"])]
        r["거래처명"] = i["거래처명"]
        r["분류"] = i["분류"]
        r["결제주기"] = i["결제주기"]
        r["귀속월"] = attrib_month(r["거래일시"], i["결제주기"])


def fixed_rows(months):
    """03_데이터/월고정비.xlsx 의 항목을 달마다 한 줄씩 만든다.

    카드·계좌에 안 찍히지만 매달 발생하는 비용을 여기에 적는다.
    예) 퇴직금 충당금(1년치를 12로 나눠 매달 쌓아두는 몫), 대출 이자.
    시작월·종료월을 비워두면 자료가 있는 모든 달에 적용된다.
    """
    if not FIXED_XLSX.exists():
        return []
    wb = openpyxl.load_workbook(FIXED_XLSX, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    col = {h: i for i, h in enumerate(norm(c) for c in next(it))}
    need = ("항목", "분류", "월금액")
    if any(c not in col for c in need):
        print(f"      [!] {FIXED_XLSX.name}: {need} 열이 필요합니다 — 건너뜁니다")
        wb.close()
        return []

    out = []
    for r in it:
        if not r or not norm(r[col["항목"]]):
            continue
        name = norm(r[col["항목"]])
        cat = norm(r[col["분류"]])
        if cat not in COST:
            print(f"      [!] 월고정비 '{name}': 분류 '{cat}' 를 알 수 없어 건너뜁니다")
            continue
        amount = float(r[col["월금액"]] or 0)
        if not amount:
            continue
        start = norm(r[col["시작월"]]) if "시작월" in col else ""
        end = norm(r[col["종료월"]]) if "종료월" in col else ""
        memo = norm(r[col["메모"]]) if "메모" in col else ""
        for m in months:
            if start and m < start:
                continue
            if end and m > end:
                continue
            dt = datetime(int(m[:4]), int(m[5:7]), 1)
            out.append({
                "원천": "월고정비", "거래일시": dt, "날짜": f"{m}-01",
                "결제월": m, "귀속월": m, "일": 1, "방향": "지출",
                "분류": cat, "결제주기": "당월분", "거래처명": name,
                "표기명": f"{m} {name}", "금액": amount,
                "수단": "월고정비", "참조": memo, "출처": FIXED_XLSX.name,
            })
    wb.close()
    return out


def sales_rows(sales):
    """매출 엑셀에서 시스템에 넣을 줄을 만든다 (매출만).

    상품권은 손님이 낸 결제수단이라 매출에 이미 들어 있고,
    그걸 이마트에 넘겨 임대료를 상계하는 건 '지불 방법'일 뿐이다.
    매달 상품권을 얼마 입금했느냐에 따라 계좌에서 나가는 금액이 달라지므로,
    임대료는 월고정비(월 750만원 평균)에서 잡고 여기서는 건드리지 않는다.
    """
    out = []
    for month, s in sales.items():
        dt = datetime(int(month[:4]), int(month[5:7]), 1)
        base = {"원천": "매출엑셀", "거래일시": dt, "날짜": f"{month}-01",
                "결제월": month, "귀속월": month, "일": 1,
                "결제주기": "당월분", "수단": "매출엑셀", "출처": s["출처"]}
        out.append({**base, "방향": "수입", "분류": "매출",
                    "거래처명": "매출 (일일결산)", "표기명": f"{month} 매출",
                    "금액": s["총매출"],
                    "참조": f"영업일 {s['영업일']}일 · 미수 {s['미수']:,.0f}원 제외"})
        fee = round(s["카드"] * CARD_FEE_RATE)
        if fee:
            out.append({**base, "방향": "지출", "분류": "사업경비",
                        "거래처명": f"카드 수수료 ({CARD_FEE_RATE * 100:.1f}%)",
                        "표기명": f"{month} 카드수수료", "금액": fee,
                        "참조": f"카드매출 {s['카드']:,.0f}원 기준"})
    return out


# ── 이상 탐지 ──────────────────────────────────────────────────────────
def find_issues(rows, master, new_keys, first_run):
    issues = []

    def add(kind, level, who, amount, date, desc):
        issues.append({"유형": kind, "심각도": level, "가맹점명": who,
                       "금액": amount, "날짜": date, "설명": desc})

    if first_run:
        add("첫 실행", "확인", "—", 0, min(r["날짜"] for r in rows),
            f"거래처·적요 {len(new_keys)}건이 자동 분류됐습니다. "
            f"03_데이터/거래처마스터.xlsx 를 열어 분류와 결제주기를 한 번만 확인해주세요.")
    else:
        for src, direction, label in new_keys:
            rel = [r for r in rows if r["원천"] == src and r["방향"] == direction
                   and r["표기명"] == label]
            add("신규", "확인", f"[{src}] {label}", sum(r["금액"] for r in rel),
                min(r["날짜"] for r in rel),
                f"처음 보는 곳입니다 ({len(rel)}건). 거래처마스터에서 분류를 확인해주세요.")

    for r in rows:
        if r["분류"] == "미분류":
            add("미분류", "위험" if r["금액"] >= 100_000 else "확인",
                f"[{r['원천']}] {r['표기명']}", r["금액"], r["날짜"],
                "분류가 정해지지 않아 순이익 계산에서 빠져 있습니다. 거래처마스터에서 지정해주세요.")

    seen = defaultdict(list)
    for r in rows:
        if r["분류"] in DUP_CATEGORIES and r["금액"] >= DUP_MIN_AMOUNT:
            seen[(r["날짜"], r["표기명"], r["금액"])].append(r)
    for (date, label, amt), g in seen.items():
        if len(g) > 1:
            add("중복 결제 의심", "위험", label, amt * len(g), date,
                f"같은 날 같은 금액({amt:,.0f}원)이 {len(g)}건 나갔습니다.")

    order = {"위험": 0, "확인": 1}
    issues.sort(key=lambda x: (order.get(x["심각도"], 9), -x["금액"]))
    return issues


# ── 엑셀 스타일 ────────────────────────────────────────────────────────
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(color="FFFFFF", bold=True, size=10)
THIN = Side(style="thin", color="D9D9D9")


def style_sheet(ws, widths=None, money_cols=(), freeze="A2"):
    for c in ws[1]:
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    for col in money_cols:
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            for c in row:
                c.number_format = "#,##0"
    for row in ws.iter_rows(min_row=1):
        for c in row:
            c.border = Border(bottom=THIN)
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions


# ── 집계 ───────────────────────────────────────────────────────────────
def aggregate(mrows):
    by_cat = defaultdict(float)
    parties = defaultdict(lambda: {"amount": 0.0, "count": 0, "category": "", "dir": ""})
    for r in mrows:
        by_cat[r["분류"]] += r["금액"]
        p = parties[r["거래처명"]]
        p["amount"] += r["금액"]
        p["count"] += 1
        p["category"] = r["분류"]
        p["dir"] = r["방향"]
    # 환불은 비용이 아니라 '안 팔린 것'이므로 매출에서 뺀다.
    # 거래처가 돌려준 돈(매입환불)은 반대로 비용에서 뺀다.
    revenue = sum(by_cat[c] for c in INCOME) - sum(by_cat[c] for c in REFUND)
    cost = sum(by_cat[c] for c in COST) - sum(by_cat[c] for c in COST_REFUND)
    return by_cat, parties, revenue, cost


def month_status(rows, basis):
    """월별로 어떤 자료가 빠졌는지 돌려준다. {월: [빠진 자료 설명, ...]}

    귀속 기준에서 M월이 온전하려면
      · 카드·계좌 : M월분(그 자리에서 결제한 것)과
                    M+1월분(거래처·임대료·인건비처럼 다음 달에 나가는 것) 둘 다
      · 매출자료  : M월분만 있으면 된다
    """
    have = {"카드내역": {r["결제월"] for r in rows if r["원천"] == "카드"},
            "계좌내역": {r["결제월"] for r in rows if r["원천"] == "계좌"},
            "매출자료": {r["귀속월"] for r in rows if r["원천"] == "매출엑셀"}}
    key = "귀속월" if basis == "귀속" else "결제월"
    status = {}
    for month in sorted({r[key] for r in rows}):
        two = {month, next_month_key(month)} if basis == "귀속" else {month}
        need = {"카드내역": two, "계좌내역": two, "매출자료": {month}}
        missing = [f"{mm} {label}"
                   for label in ("매출자료", "카드내역", "계좌내역")
                   for mm in sorted(need[label] - have[label])]
        status[month] = missing
    return status


# ── 리포트 ─────────────────────────────────────────────────────────────
def build_report(month, rows, issues, missing):
    mrows = [r for r in rows if r["귀속월"] == month]
    if not mrows:
        return None
    by_cat, parties, revenue, cost = aggregate(mrows)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.append(["항목", "금액", "비고"])
    if missing:
        ws.append(["※ 자료 부족 — 아래 숫자는 아직 정확하지 않습니다", None,
                   "빠진 자료: " + ", ".join(missing)])
        ws.append(["", None, ""])
    ws.append(["매출", by_cat.get("매출", 0),
               "06_매출 일일결산 합계 (현금·현영·상품권·카드 / 미수는 제외)"
               if by_cat.get("매출")
               else "← 이 달 매출 엑셀이 없어 매출이 안 잡혔습니다"])
    if by_cat.get("기타수입"):
        ws.append(["기타수입", by_cat["기타수입"], "캐시백 등"])
    if by_cat.get("매출환불"):
        ws.append(["매출환불", -by_cat["매출환불"], "손님에게 돌려준 돈 — 매출에서 뺌"])
    if by_cat.get("기타수입") or by_cat.get("매출환불"):
        ws.append(["매출 합계", revenue, "매출 + 기타수입 − 환불"])
    ws.append(["", None, ""])
    for c in COST:
        if by_cat.get(c):
            ws.append([c, by_cat[c], f"매출 대비 {by_cat[c] / revenue * 100:.1f}%" if revenue else ""])
    for c in COST_REFUND:
        if by_cat.get(c):
            ws.append([c, -by_cat[c], "거래처가 돌려준 돈 — 비용에서 뺌"])
    ws.append(["사업비용 합계", cost, " + ".join(c for c in COST if by_cat.get(c))])
    ws.append(["", None, ""])
    if revenue:
        ws.append(["순이익", revenue - cost, "매출 + 기타수입 − 사업비용"])
        ws.append(["이익률", f"{(revenue - cost) / revenue * 100:.1f}%", ""])
    else:
        ws.append(["순이익", None, "매출 자료가 없어 계산할 수 없습니다"])
    ws.append(["", None, ""])
    for c in NEUTRAL:
        if by_cat.get(c):
            note = {"개인": "사업과 무관 — 순이익에서 제외",
                    "카드대금": "이용내역이 있는 카드의 대금 — 건별로 이미 잡혀 중복 제외",
                    "정산입금": "통장에 들어온 돈 — 매출은 매출엑셀로 잡으므로 참고용",
                    "임대료납부": "이마트에 실제로 나간 돈 — 임대료는 월고정비로 잡아 중복 제외",
                    "미분류": "분류 필요 — 지금은 계산에서 빠져 있음"}.get(c, "")
            ws.append([f"(참고) {c}", by_cat[c], note])
    style_sheet(ws, widths=[24, 18, 62], money_cols=[2])

    ws = wb.create_sheet("거래처별")
    ws.append(["거래처명", "분류", "수입/지출", "건수", "금액"])
    for name, p in sorted(parties.items(), key=lambda x: -x[1]["amount"]):
        ws.append([name, p["category"], p["dir"], p["count"], p["amount"]])
    style_sheet(ws, widths=[30, 14, 11, 8, 16], money_cols=[5])

    ws = wb.create_sheet("확인필요")
    ws.append(["심각도", "유형", "대상", "날짜", "금액", "설명"])
    for i in issues:
        ws.append([i["심각도"], i["유형"], i["가맹점명"], i["날짜"], i["금액"], i["설명"]])
    style_sheet(ws, widths=[10, 15, 34, 13, 14, 74], money_cols=[5])

    ws = wb.create_sheet("원장")
    ws.append(["날짜", "원천", "방향", "거래처명", "분류", "결제주기",
               "표기명", "수단", "금액", "참조"])
    for r in sorted(mrows, key=lambda x: x["거래일시"]):
        ws.append([r["날짜"], r["원천"], r["방향"], r["거래처명"], r["분류"],
                   r["결제주기"], r["표기명"], r["수단"], r["금액"], r["참조"]])
    style_sheet(ws, widths=[13, 7, 8, 24, 12, 10, 32, 12, 14, 13], money_cols=[9])

    out = REPORT_DIR / f"{month}_정산리포트.xlsx"
    wb.save(out)
    return out


def save_ledger(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "원장"
    ws.append(["날짜", "결제월", "귀속월", "원천", "방향", "거래처명", "분류",
               "결제주기", "표기명", "수단", "금액", "참조", "출처"])
    for r in sorted(rows, key=lambda x: x["거래일시"]):
        ws.append([r["날짜"], r["결제월"], r["귀속월"], r["원천"], r["방향"],
                   r["거래처명"], r["분류"], r["결제주기"], r["표기명"],
                   r["수단"], r["금액"], r["참조"], r["출처"]])
    style_sheet(ws, widths=[13, 10, 10, 7, 8, 24, 12, 10, 32, 12, 14, 13, 24],
                money_cols=[11])
    wb.save(LEDGER_XLSX)


# ── 대시보드 데이터 ────────────────────────────────────────────────────
def sales_analytics(detail):
    """월별 품목·브랜드·객단가 집계 (전부 매출 기준 — 원가 자료가 없다)."""
    out = {}
    for month, sales in detail.items():
        if not sales:
            continue
        total = sum(s["amount"] for s in sales)
        by_item = defaultdict(lambda: {"amount": 0.0, "count": 0})
        by_brand = defaultdict(lambda: {"amount": 0.0, "count": 0})
        for s in sales:
            it = classify_item(s)
            by_item[it]["amount"] += s["amount"]
            by_item[it]["count"] += 1
            if s["frame"]:
                by_brand[s["frame"]]["amount"] += s["amount"]
                by_brand[s["frame"]]["count"] += 1
        out[month] = {
            "count": len(sales),
            "total": total,
            "avgTicket": round(total / len(sales)),
            "byItem": sorted(({"name": k, **v} for k, v in by_item.items()),
                             key=lambda x: -x["amount"]),
            "byBrand": sorted(({"name": k, **v} for k, v in by_brand.items()),
                              key=lambda x: -x["amount"])[:8],
            "recent": [
                {"date": s["date"], "staff": s["staff"],
                 "customer": mask_name(s["customer"]),
                 "item": classify_item(s),
                 "detail": " / ".join(x for x in (s["frame"], s["lens"], s["contact"]) if x),
                 "amount": s["amount"]}
                for s in sorted(sales, key=lambda x: (x["date"], -x["amount"]),
                                reverse=True)[:40]],
        }
    return out


RECUR_MONTHS = 5          # 이 개월 수 이상 나오면 '매달 나가는 돈'
FIXED_CATS = ["임대료", "인건비", "금융비용"]


def kr(n):
    """1,234,567 → '123만원' 처럼 읽기 쉽게."""
    n = round(n or 0)
    sign = "-" if n < 0 else ""
    n = abs(n)
    if n < 10_000:
        return f"{sign}{n:,}원"
    eok, man = divmod(n, 100_000_000)[0], (n % 100_000_000) // 10_000
    out = f"{eok}억 " if eok else ""
    if man or not eok:
        out += f"{man:,}만"
    return sign + out.strip() + "원"


def build_insights(rows, basis):
    """숫자를 보고 '살펴볼 것' 과 '아낄 수 있는 곳' 을 뽑는다.

    자료가 온전한 달만 평균에 쓴다. 미완성 달은 비용이 덜 잡혀 있어
    평균을 끌어내리고 엉뚱한 경고를 만든다.
    """
    full = [m for m in sorted(basis) if basis[m]["revenue"] > 0 and not basis[m]["missing"]]
    warn, save = [], []
    if not full:
        return warn, save

    def parts(m):
        b = basis[m]
        c = b["byCategory"]
        mat = c.get("거래처매입", 0) - c.get("매입환불", 0)
        fixed = sum(p["amount"] for p in b["parties"]
                    if p["category"] in FIXED_CATS
                    or (p["category"] == "사업경비" and p["months"] >= RECUR_MONTHS))
        return b["revenue"], mat, fixed, b["profit"]

    last = full[-1]
    rev, mat, fixed, profit = parts(last)
    avg_rev = sum(parts(m)[0] for m in full) / len(full)
    avg_mat_rate = sum(parts(m)[1] / parts(m)[0] for m in full) / len(full)
    label = f"{int(last[5:])}월"

    # ── 살펴볼 것 ──────────────────────────────────────────────
    losers = [m for m in full if basis[m]["profit"] < 0]
    if losers:
        worst = min(losers, key=lambda m: basis[m]["profit"])
        r, mt, _, pf = parts(worst)
        warn.append({
            "level": "high", "title": f"{int(worst[5:])}월은 {kr(abs(pf))} 적자였습니다",
            "detail": f"매출 {kr(r)}인데 재료값이 {kr(mt)}(매출의 {mt/r*100:.0f}%)이었습니다. "
                      f"물건을 많이 들여온 달은 이렇게 적자로 보입니다 — "
                      f"그 재고가 다음 달에 팔리면 그 달이 흑자가 됩니다.",
        })

    breakeven = fixed + mat
    if rev < breakeven:
        warn.append({
            "level": "high", "title": f"{label}은 본전에 {kr(breakeven - rev)} 모자랍니다",
            "detail": f"고정비 {kr(fixed)} + 재료값 {kr(mat)} = {kr(breakeven)}은 팔아야 "
                      f"본전인데 {kr(rev)} 팔았습니다.",
        })
    else:
        warn.append({
            "level": "ok", "title": f"{label}은 본전을 {kr(rev - breakeven)} 넘겼습니다",
            "detail": f"본전 매출 {kr(breakeven)} (고정비 {kr(fixed)} + 재료값 {kr(mat)})",
        })

    if rev and mat / rev > avg_mat_rate * 1.15:
        over = mat - rev * avg_mat_rate
        warn.append({
            "level": "mid", "title": "재료값이 평소보다 많이 나갔습니다",
            "detail": f"{label} 재료값은 매출의 {mat/rev*100:.0f}%로, 평소({avg_mat_rate*100:.0f}%)보다 "
                      f"약 {kr(over)} 많습니다. 재고가 쌓이고 있는지 확인해보세요.",
        })

    if rev and fixed / rev > 0.5:
        warn.append({
            "level": "mid", "title": f"고정비가 매출의 {fixed/rev*100:.0f}%나 됩니다",
            "detail": f"월세·월급 같은 고정비가 {kr(fixed)}입니다. 매출이 조금만 줄어도 "
                      f"바로 적자가 됩니다. 매출을 늘리거나 고정비를 줄여야 안전해집니다.",
        })

    if len(full) >= 3:
        recent3 = [parts(m)[0] for m in full[-3:]]
        if recent3[0] > recent3[1] > recent3[2]:
            warn.append({
                "level": "mid", "title": "매출이 3개월째 줄고 있습니다",
                "detail": " → ".join(f"{kr(v)}" for v in recent3) + "원",
            })

    due = defaultdict(float)
    for r in rows:
        if r["원천"] == "매출엑셀":
            continue
    # 미수는 매출엑셀 요약에 있어 여기선 다루지 않는다

    # ── 아낄 수 있는 곳 ────────────────────────────────────────
    fee = next((p["amount"] for p in basis[last]["parties"]
                if p["name"].startswith("카드 수수료")), 0)
    if fee:
        save.append({
            "title": f"카드 수수료 — 월 {kr(fee)}",
            "detail": f"지금 {CARD_FEE_RATE*100:.1f}%로 계산하고 있습니다. 연매출 규모에 따라 "
                      f"영세·중소가맹점 우대수수료율(0.5~1.5%)을 받을 수 있습니다.",
            "action": "여신금융협회 '가맹점 수수료율 조회'에서 우대 대상인지 확인하기",
            "impact": f"1.5%로 낮추면 월 약 {kr(fee * (1 - 1.5/(CARD_FEE_RATE*100)))} 절약",
        })

    unknown = [p for p in basis[last]["parties"]
               if any(p["name"].startswith(c) for c in OTHER_CARDS)]
    if unknown:
        tot = sum(p["amount"] for p in unknown)
        save.append({
            "title": "내역을 모르는 카드값 — 월 " + f"{kr(tot)}",
            "detail": ", ".join(f"{p['name']} {kr(p['amount'])}" for p in unknown) +
                      " — 이용내역이 없어 무엇에 썼는지 알 수 없습니다.",
            "action": "카드사에서 이용내역 엑셀을 받아 02_카드내역 폴더에 넣기",
            "impact": "쓸데없이 나가는 정기결제가 있는지 바로 드러납니다",
        })

    small = sorted([p for p in basis[last]["parties"]
                    if p["category"] == "사업경비" and p["months"] >= RECUR_MONTHS
                    and p["amount"] < 300_000],
                   key=lambda x: -x["amount"])
    if small:
        tot = sum(p["amount"] for p in small)
        save.append({
            "title": f"매달 빠져나가는 자잘한 정기결제 — 월 {kr(tot)}",
            "detail": ", ".join(f"{p['name']} {kr(p['amount'])}" for p in small[:8]),
            "action": "안 쓰는 게 있는지 하나씩 확인하기 (연 " + f"{kr(tot*12)}" + ")",
            "impact": f"하나만 끊어도 연 수십만원",
        })

    meal = basis[last]["byCategory"].get("식대", 0)
    if meal:
        save.append({
            "title": f"밥값 — 월 {kr(meal)}",
            "detail": f"직원 식대로 매달 나가는 돈입니다. 연 {kr(meal*12)}.",
            "action": "식대 한도를 정해두면 관리가 쉬워집니다",
            "impact": f"10%만 줄여도 연 {kr(meal*12*0.1)}",
        })

    priv = basis[last]["byCategory"].get("개인", 0)
    if priv:
        save.append({
            "title": f"사업용 카드로 결제한 개인 지출 — 월 {kr(priv)}",
            "detail": "순이익 계산에서는 빼두었지만, 통장에서는 실제로 나갑니다.",
            "action": "개인 카드와 사업 카드를 나눠 쓰면 정산이 훨씬 깔끔해집니다",
            "impact": "세금 신고 때도 유리합니다",
        })

    vendors = [p for p in basis[last]["parties"] if p["category"] == "거래처매입"]
    if len(vendors) >= 3:
        top = sorted(vendors, key=lambda x: -x["amount"])[:3]
        tot = sum(p["amount"] for p in vendors)
        share = sum(p["amount"] for p in top) / tot * 100 if tot else 0
        save.append({
            "title": f"재료값의 {share:.0f}%가 상위 3곳에 몰려 있습니다",
            "detail": ", ".join(f"{p['name']} {kr(p['amount'])}" for p in top),
            "action": "많이 사는 곳부터 단가를 다시 협의해보세요",
            "impact": f"이 3곳에서 5%만 깎아도 월 "
                      f"{kr(sum(p['amount'] for p in top)*0.05)}",
        })

    return warn, save


def vendor_months(rows):
    """거래처마다 몇 개 달에 나왔는지 — 매달 나가는 고정비인지 가리는 데 쓴다."""
    seen = defaultdict(set)
    for r in rows:
        if r["분류"] in COST:
            seen[r["거래처명"]].add(r["귀속월"])
    return {k: len(v) for k, v in seen.items()}


def basis_data(rows, issues, basis):
    key = "귀속월" if basis == "귀속" else "결제월"
    status = month_status(rows, basis)
    vmonths = vendor_months(rows)
    # 다음 달에 결제되는 분류들 — 그 달 자료가 없으면 통째로 비어 보인다
    later = sorted({r["분류"] for r in rows
                    if r["결제주기"] == "전월분" and r["분류"] in COST})
    out = {}
    for month in sorted({r[key] for r in rows}):
        mrows = [r for r in rows if r[key] == month]
        by_cat, parties, revenue, cost = aggregate(mrows)
        out[month] = {
            "revenue": revenue, "cost": cost, "profit": revenue - cost,
            "count": len(mrows), "complete": not status[month],
            "missing": status[month], "hasRevenue": revenue > 0,
            "nextMonth": next_month_key(month),
            # 이 달 몫이지만 다음 달에 결제돼서 아직 안 잡혔거나 일부만 잡힌 비용 항목
            # (0원인 것만이 아니라 전월분 분류 전부 — 일부만 들어와도 부족하긴 마찬가지)
            "pending": (later
                        if basis == "귀속" and any(next_month_key(month) in x
                                                   for x in status[month]) else []),
            "byCategory": {c: by_cat[c] for c in CATEGORIES if by_cat.get(c)},
            "costs": [{"name": c, "amount": by_cat[c]} for c in COST if by_cat.get(c)],
            "parties": sorted(
                ({"name": k, **v, "months": vmonths.get(k, 0)}
                 for k, v in parties.items() if v["category"] in COST),
                key=lambda x: -x["amount"]),
            "income": sorted(
                ({"name": k, **v} for k, v in parties.items()
                 if v["category"] in INCOME + REFUND),
                key=lambda x: -x["amount"]),
            "issues": [i for i in issues
                       if any(r["날짜"] == i["날짜"] and i["가맹점명"].endswith(r["표기명"])
                              for r in mrows)],
            "ledger": sorted(
                ({"date": r["날짜"], "src": r["원천"], "dir": r["방향"],
                  "party": r["거래처명"], "category": r["분류"], "cycle": r["결제주기"],
                  "label": r["표기명"], "means": r["수단"], "amount": r["금액"]}
                 for r in mrows), key=lambda x: (x["date"], -x["amount"])),
        }

    # ── 아직 안 들어온 고정비를 예상치로 채운다 ──────────────────────────
    # 8월 급여는 9월에 나간다. 9월 계좌 자료가 없으면 8월 인건비가 퇴직금
    # 283,333원만 남아서 "월급이 빠졌다" 처럼 보인다. 실제로 빠진 게 아니라
    # 아직 안 들어온 것이므로, 지난 달들의 중앙값으로 얼마인지 알려준다.
    done = [m for m in out if out[m]["complete"] and out[m]["hasRevenue"]]
    typical = {}
    for c in COST:
        vals = sorted(out[m]["byCategory"].get(c, 0) for m in done)
        if vals:
            typical[c] = vals[len(vals) // 2]          # 중앙값
    # 아직 안 끝난 달은 예상치를 넣으면 안 된다. 매출은 며칠치인데 비용만
    # 한 달치가 붙어서 적자처럼 보인다. (9/6 에 9월을 보면 매출 5일치 vs 월급 한 달치)
    now = datetime.now().strftime("%Y-%m")
    for month, b in out.items():
        b["inProgress"] = month >= now
        est = []
        if not b["inProgress"]:
            for c in b["pending"]:
                exp = typical.get(c, 0)
                act = b["byCategory"].get(c, 0)
                if exp > 0 and act < exp * 0.6:        # 절반도 안 들어왔으면
                    est.append({"category": c, "expected": exp, "actual": act,
                                "shortfall": exp - act})
            est.sort(key=lambda x: -x["shortfall"])
        b["estimate"] = est
        b["estTotal"] = sum(e["shortfall"] for e in est)
        b["profitEst"] = b["profit"] - b["estTotal"]
    return out


# ── 대시보드 잠금 ──────────────────────────────────────────────────────
def make_payload(data):
    """03_데이터/비밀번호.txt 가 있으면 대시보드 내용을 암호로 잠근다.

    AES-256-GCM + PBKDF2(SHA-256). 브라우저의 Web Crypto 표준과 같은 방식이라
    비밀번호를 넣어야만 풀린다. 파일을 열어 소스를 봐도 암호문만 보인다.
    """
    plain = json.dumps(data, ensure_ascii=False)
    if not PASSWORD_TXT.exists():
        return {"enc": False, "data": data}, None

    # utf-8-sig : 메모장으로 저장하면 파일 앞에 보이지 않는 BOM 이 붙는데,
    #             그대로 읽으면 비밀번호가 달라져 버린다. 여기서 떼어낸다.
    password = PASSWORD_TXT.read_text(encoding="utf-8-sig").strip().lstrip("﻿")
    if not password:
        return {"enc": False, "data": data}, None

    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32,
                     salt=salt, iterations=PBKDF2_ROUNDS).derive(password.encode())
    ct = AESGCM(key).encrypt(iv, plain.encode(), None)
    b64 = lambda b: base64.b64encode(b).decode()
    return {"enc": True, "rounds": PBKDF2_ROUNDS,
            "salt": b64(salt), "iv": b64(iv), "ct": b64(ct)}, password


def main():
    print("[1/5] 자료 읽는 중...")
    cards = read_cards()
    bank = read_bank()
    sales = read_sales()
    rows = cards + bank
    if not rows:
        sys.exit("[!] 카드내역도 계좌내역도 없습니다.")
    print(f"      카드 {len(cards)}건 · 계좌 {len(bank)}건 · 매출 {len(sales)}개월")
    for m in sorted(sales):
        s = sales[m]
        print(f"        {m} 매출 {s['총매출']:>12,.0f}원 "
              f"(카드 {s['카드']:,.0f} / 현금·현영 {s['현금'] + s['현영']:,.0f} / "
              f"상품권 {s['상품권']:,.0f}) · {s['영업일']}일 "
              f"· 미수 {s['미수']:,.0f}원 제외")
        cnt = s.get("품목건수") or {}
        if sum(cnt.values()):
            parts = [f"{k} {cnt[k]}건" for k in ITEM_COLS if cnt[k]]
            print(f"          품목 — {' · '.join(parts)}")

    print("[2/5] 분류 · 귀속월 계산 중...")
    master, first_run = load_master()
    new_keys = update_master(master, rows)
    save_master(master, rows)
    apply_master(rows, master)
    rows += sales_rows(sales)          # 매출 줄은 마스터를 거치지 않는다
    fixed = fixed_rows(sorted({r["귀속월"] for r in rows}))
    rows += fixed
    if fixed:
        per = {}
        for f in fixed:
            per[f["거래처명"]] = f["금액"]
        print("      월고정비: " + ", ".join(f"{k} {v:,.0f}원" for k, v in per.items()))
    print(f"      거래처 {len(master)}곳 (신규 {len(new_keys)}곳) · 총 {len(rows)}건")

    print("[3/5] 이상 항목 점검 중...")
    issues = find_issues(rows, master, new_keys, first_run)
    print(f"      확인 필요 {len(issues)}건")

    print("[4/5] 원장 · 리포트 생성 중...")
    save_ledger(rows)
    status = month_status(rows, "귀속")
    for month in sorted({r["귀속월"] for r in rows}):
        out = build_report(month, rows, issues, status[month])
        if out:
            miss = status[month]
            print(f"      {out.name}" +
                  (f"  (부족: {', '.join(miss)})" if miss else "  ✓ 완전"))

    print("[5/5] 대시보드 생성 중...")
    # 대시보드에 보여줄 최근 입출금 (원장에서 뽑는다)
    SKIP = {"정산입금", "카드대금", "임대료납부"}   # 참고용이라 목록에서 뺀다
    recent = [{"date": r["날짜"], "kind": r["분류"], "who": r["거래처명"],
               "amount": r["금액"], "dir": r["방향"]}
              for r in sorted(rows, key=lambda x: x["거래일시"], reverse=True)
              if r["분류"] not in SKIP and r["원천"] != "매출엑셀"][:60]

    by_month = basis_data(rows, issues, "귀속")
    warn, save = build_insights(rows, by_month)
    print(f"      진단 — 살펴볼 것 {len(warn)}건 · 아낄 곳 {len(save)}건")
    data = {"generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "basis": {"귀속": by_month,
                      "결제": basis_data(rows, issues, "결제")},
            "recent": recent, "warn": warn, "save": save}
    payload, password = make_payload(data)
    html = (BASE / "scripts" / "dashboard_template.html").read_text(encoding="utf-8")
    DASHBOARD.write_text(
        html.replace("/*__DATA__*/null", json.dumps(payload, ensure_ascii=False)),
        encoding="utf-8")
    lock = "  [비밀번호 잠김]" if password else "  (잠금 없음)"
    print(f"      {DASHBOARD.name}{lock}")

    print("\n완료. 대시보드를 엽니다...")
    if not password:
        print("     ※ 잠그려면 03_데이터/비밀번호.txt 파일에 비밀번호를 적어두세요.")

    # 대시보드 열기는 파이썬이 직접 한다.
    # 배치 파일에서 한글 파일명을 다루면 글자가 깨져서 열리지 않는다.
    try:
        os.startfile(DASHBOARD)
    except Exception as e:
        print(f"     [!] 자동으로 열지 못했습니다 ({e}). 대시보드.html 을 직접 더블클릭해주세요.")


if __name__ == "__main__":
    main()
