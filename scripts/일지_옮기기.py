# -*- coding: utf-8 -*-
"""
옛 일일결산일지를 새 양식으로 옮긴다.

  python scripts/일지_옮기기.py "C:\\...\\2026년 9월.xlsx" 2026 9

옛 양식은 머리글이 3·4행에 나뉘어 있고 5행부터 판매 기록이다.
  A번호 B담당자 C이름 D현금 E현영 F상품권 G카드 H미수
  I프레임 J안경렌즈 K콘택트 L·M비고

새 양식으로 옮기면서 하는 일은 세 가지뿐이다.
  1. 프레임 → 안경테,  콘택트 → 콘택트 (그대로)
  2. 안경렌즈 → 이름에 '누진·다초점·기능' 이 들어 있으면 누진, 아니면 단초점
  3. 옛 파일에 있던 제품 이름을 설정 시트 드롭다운 목록에 넣어준다

■ 금액은 한 원도 건드리지 않는다.
  옮긴 뒤 원본과 총액·건수를 대조해서 다르면 멈춘다.

■ 한 줄은 한 줄 그대로 옮긴다.
  옛 파일에서 손님 한 분이 두 줄에 걸쳐 적혀 있어도 합치지 않는다.
  이름이 없는 줄이 많아 어느 줄이 같은 손님인지 알 수 없기 때문이다.
  (짐작으로 합치면 손님 수가 틀어진다)
"""
import calendar
import re
import sys
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
import 일일결산_양식만들기 as 양식

BASE = Path(__file__).resolve().parent.parent
SALES_DIR = BASE / "06_매출"

# 옛 양식 열 위치 (1부터). 머리글을 못 읽었을 때만 쓰는 기본값.
OLD = dict(no=1, staff=2, cust=3, cash=4, cre=5, gift=6, card=7, due=8,
           frame=9, lens=10, contact=11, memo=12)
# 머리글 글자로 열을 찾는다. 순서대로 먼저 맞은 것이 이긴다.
#   ※ '현영' 을 '상품권' 보다 먼저 봐야 한다. 18일 시트의 현영 칸에
#     '현영(상품권포함)' 이라고 적혀 있어서, 순서가 바뀌면 상품권으로 잡힌다.
#   ※ 9·11·23일 시트에는 담당자 칸이 아예 없어 한 칸씩 밀려 있다.
#     고정 위치로 읽으면 미수가 카드로, 카드가 상품권으로 들어간다.
HEAD = [("no", ("번호",)), ("staff", ("담당",)), ("cust", ("이름", "고객")),
        ("cre", ("현영", "현금영수")), ("cash", ("현금",)),
        ("gift", ("상품권",)), ("card", ("카드",)), ("due", ("미수",)),
        ("frame", ("프레임", "안경테")), ("lens", ("안경렌즈", "렌즈")),
        ("contact", ("콘택트",)), ("memo", ("비고",))]
MONEY_MAP = [("cash", "현금"), ("cre", "현금영수증"), ("gift", "상품권"),
             ("card", "카드"), ("due", "미수")]
# 렌즈 이름에 이게 들어 있으면 누진으로 본다.
#
# 2026-09-19 에 넓혔다. 9월 일일결산의 누진 8건을 애니아이 판매기록의 같은 날·같은
# 손님과 하나씩 맞춰보니, 사장님은 '대명누진' 처럼 적기도 하지만 '1.6 hd40 11mm'
# 처럼 제품명으로 적는 날이 더 많았다.
#   대명누진  = dx5 1.55 11mm / 1.6 hd40 11mm / 1.74 dx에센셜 11mm / 1.55 hd10
#   호야누진  = 호야 1.60 라이프스타일4 어반 11mm
# 누진대 길이(11mm·13mm)가 가장 확실한 표시다. 누진 렌즈에만 적는다.
PROG = (
    "누진", "다초점", "프로그레", "기능성",
    "내면멀티", "멀티포컬",            # '중멀티' 는 멀티코팅이라 누진이 아니다 (평균 2.5만원)
    "hd10", "hd20", "hd30", "hd40", "hd 40",          # 대명
    "dx3", "dx5", "dx7", "dx9", "dc5", "에센셜",
    "라이프스타일", "언루트",                          # 호야
    "발란시스", "리버티", "수퍼브", "슈퍼브", "퓨어브",   # 에실로 / 자이스
    "바리락스", "피지오", "컴포트맥스", "컴포트 맥스",
    "써미트", "앰플리튜드", "스타터", "홈엔오피스", "홈앤오피스",
)
# 누진대 길이 — 8~25mm 사이면 누진으로 본다 (11mm, 13미리, 12 MM 다 잡는다)
PROG_MM = re.compile(r"(?<![0-9.])(\d{1,2})\s*(?:mm|미리|파이)", re.I)
# 누진인지 단초점인지 이름만으로는 알 수 없는 것들 — 사용자에게 물어봐야 한다
UNSURE = ("드라이브", "오피스", "리딩", "실내")


def find_old_cols(ws):
    """머리글(3·4행)을 읽어 이 시트의 열 위치를 찾는다.

    시트마다 칸이 하나씩 빠져 있는 경우가 있어서 고정 위치로 읽으면 안 된다.
    못 찾은 칸은 기본값(OLD)을 쓴다.
    """
    cols = {}
    for c in range(1, 16):
        txt = ws.cell(4, c).value or ws.cell(3, c).value
        if txt in (None, ""):
            continue
        txt = re.sub(r"\s+", "", str(txt))
        for key, words in HEAD:
            if key in cols:
                continue
            if any(w in txt for w in words):
                cols[key] = c
                break
    if not (cols.get("cust") and cols.get("cash") and cols.get("due")):
        return dict(OLD), False
    # 담당자 머리글이 비어 있는 시트가 많다. 번호와 이름 사이에 칸이 있으면
    # 그 칸이 담당자다. 붙어 있으면(9·11·23일) 담당자 칸이 아예 없는 것이다.
    if not cols.get("staff") and cols["cust"] - cols.get("no", 1) > 1:
        cols["staff"] = cols["cust"] - 1
    # 못 찾은 칸은 '없는 칸' 으로 둔다. 기본값을 끼워 넣으면 밀린 시트에서
    # 담당자 자리에 고객 이름이 들어간다.
    return {k: cols.get(k) for k in OLD}, True


def read_old(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out, sheets, shifted = [], 0, []
    for name in wb.sheetnames:
        mm = re.fullmatch(r"(\d+)일", name)
        if not mm:
            continue
        sheets += 1
        ws = wb[name]
        cols, ok = find_old_cols(ws)
        if ok and any(cols[k] != OLD[k] for k in OLD):
            shifted.append(name)
        for r in ws.iter_rows(min_row=5, max_col=16, values_only=True):
            if not r or not isinstance(r[0], (int, float)):
                continue      # 번호가 숫자인 줄만 판매 기록
            def g(k, cols=cols):
                c = cols.get(k)
                if not c:
                    return None
                return r[c - 1] if c - 1 < len(r) else None
            txt = lambda k: ("" if g(k) in (None, "")
                             else str(g(k)).strip().replace("\n", " "))
            num = lambda k: g(k) if isinstance(g(k), (int, float)) else 0
            rec = dict(day=int(mm.group(1)), staff=txt("staff"), cust=txt("cust"),
                       frame=txt("frame"), lens=txt("lens"), contact=txt("contact"),
                       memo=txt("memo"),
                       **{k: num(k) for k, _ in MONEY_MAP})
            if not any(rec[k] for k, _ in MONEY_MAP) and \
               not any(rec[k] for k in ("frame", "lens", "contact")):
                continue      # 빈 줄
            out.append(rec)
    wb.close()
    return out, sheets, shifted


def read_gyeolsan(path):
    """원본 '결산' 시트의 일자별 총금액(미수 제외). 못 읽으면 None."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if "결산" not in wb.sheetnames:
        wb.close()
        return None
    ws = wb["결산"]
    col = next((c for c in range(1, 13)
                if ws.cell(1, c).value and "총금액" in str(ws.cell(1, c).value)), None)
    if not col:
        wb.close()
        return None
    out = {}
    for r in range(2, 40):
        d, v = ws.cell(r, 1).value, ws.cell(r, col).value
        if isinstance(d, int) and 1 <= d <= 31:
            out[d] = v if isinstance(v, (int, float)) else 0
    wb.close()
    return out


def split_lens(name):
    """렌즈 이름을 (누진, 단초점) 중 하나로 가른다."""
    if not name:
        return None
    low = name.lower()
    m = PROG_MM.search(low)
    if m and 8 <= int(m.group(1)) <= 25:
        return "누진"                      # 누진대 길이가 적혀 있으면 누진이다
    return "누진" if any(k in low for k in PROG) else "단초점"


def main():
    # --out "저장할 파일.xlsx" 를 주면 그 자리에 저장한다. 안 주면 06_매출 폴더.
    argv, out_path = [], None
    it = iter(sys.argv[1:])
    for a in it:
        if a == "--out":
            out_path = Path(next(it, ""))
        else:
            argv.append(a)
    if not argv:
        sys.exit("쓰는 법: python scripts/일지_옮기기.py \"원본.xlsx\" [연] [월] "
                 "[--out \"저장할 파일.xlsx\"]")
    src = Path(argv[0])
    if not src.exists():
        sys.exit(f"[!] 파일이 없습니다: {src}")
    if len(argv) >= 3:
        y, m = int(argv[1]), int(argv[2])
    else:
        mm = re.search(r"(\d{2,4})\s*[년\-.]\s*(\d{1,2})\s*월?", src.stem)
        if not mm:
            sys.exit("[!] 파일명에서 연·월을 못 읽었습니다. 연도와 월을 붙여주세요.")
        y = int(mm.group(1))
        y += 2000 if y < 100 else 0
        m = int(mm.group(2))

    print(f"원본: {src.name}  →  {y}년 {m}월 새 양식\n")
    recs, sheets, shifted = read_old(src)
    if shifted:
        print(f"   [알림] 칸이 다른 시트 {len(shifted)}개 — 머리글을 보고 맞춰 읽었습니다.")
        print(f"          {' · '.join(shifted)}\n")
    days = calendar.monthrange(y, m)[1]
    over = sorted({r["day"] for r in recs if r["day"] > days})
    if over:
        sys.exit(f"[!] {m}월은 {days}일까지인데 {over} 일 시트에 기록이 있습니다.")

    # 품목 배분
    lens_kind = Counter()
    for r in recs:
        r["안경테"] = r["frame"]
        r["콘택트"] = r["contact"]
        r["누진"] = r["단초점"] = ""
        k = split_lens(r["lens"])
        if k:
            r[k] = r["lens"]
            lens_kind[(k, r["lens"])] += 1

    # 옛 파일에 있던 이름을 드롭다운 목록에 넣어준다
    extra = {c: sorted({r[c] for r in recs if r[c]}) for c in 양식.ITEMS}
    wb, products, _ = 양식.make_workbook(y, m, extra)

    A, FIRST = 양식.A, 양식.FIRST
    cap = 양식.ROWS_PER_DAY
    put = Counter()
    warn_rows = []
    for r in sorted(recs, key=lambda x: x["day"]):
        d = r["day"]
        if put[d] >= cap:
            sys.exit(f"[!] {d}일에 {cap}줄이 넘습니다. ROWS_PER_DAY 를 늘려야 합니다.")
        ws = wb[f"{d}일"]
        row = FIRST + put[d]
        put[d] += 1
        ws[f'{A["담당자"]}{row}'] = r["staff"] or None
        ws[f'{A["고객명"]}{row}'] = r["cust"] or None
        for c in 양식.ITEMS:
            ws[f"{A[c]}{row}"] = r[c] or None
        for k, col in MONEY_MAP:
            if r[k]:
                ws[f"{A[col]}{row}"] = r[k]
        note = r["memo"]
        if r["lens"] and any(u in r["lens"] for u in UNSURE):
            note = (note + " / " if note else "") + "← 누진이면 옮겨주세요"
            ws[f'{A["단초점"]}{row}'].fill = PatternFill("solid", fgColor="FCE4EC")
            warn_rows.append((d, row, r["lens"]))
        if note:
            ws[f'{A["비고"]}{row}'] = note

    out = out_path or (SALES_DIR / f"{y}년 {m}월.xlsx")
    if out.exists():
        bak = out.with_name(f"{out.stem} (옮기기전){out.suffix}")
        out.replace(bak)
        print(f"기존 파일은 '{bak.name}' 로 옮겨뒀습니다.")
    wb.save(out)

    # ---- 대조: 옮긴 게 원본과 같은지 ----
    src_money = {col: sum(r[k] for r in recs) for k, col in MONEY_MAP}
    src_sales = sum(v for c, v in src_money.items() if c != "미수")
    chk = openpyxl.load_workbook(out, data_only=False)
    new_rows = 0
    new_money = dict.fromkeys([c for _, c in MONEY_MAP], 0)
    new_item = Counter()
    new_day = Counter()
    for d in range(1, days + 1):
        ws = chk[f"{d}일"]
        for rr in range(FIRST, FIRST + cap):
            vals = {c: ws[f"{A[c]}{rr}"].value for c in 양식.ITEMS}
            mon = {c: ws[f"{A[c]}{rr}"].value for _, c in MONEY_MAP}
            if not any(v for v in vals.values()) and not any(mon.values()):
                continue
            new_rows += 1
            for c, v in mon.items():
                if isinstance(v, (int, float)):
                    new_money[c] += v
                    if c != "미수":
                        new_day[d] += v
            for c, v in vals.items():
                if v:
                    new_item[c] += 1
    chk.close()

    print(f"\n옮긴 결과: {out.name}")
    print(f"   {new_rows}줄 · 영업일 {sum(1 for d in put.values() if d)}일\n")
    print(f"   {'항목':<8}{'원본':>14}{'옮긴 뒤':>14}")
    ok = True
    for _, col in MONEY_MAP:
        same = abs(src_money[col] - new_money[col]) < 0.5
        ok &= same
        print(f"   {col:<8}{src_money[col]:>14,.0f}{new_money[col]:>14,.0f}"
              f"   {'O' if same else 'X'}")
    tot_new = sum(v for c, v in new_money.items() if c != "미수")
    ok &= abs(src_sales - tot_new) < 0.5 and new_rows == len(recs)
    print(f"   {'총매출':<8}{src_sales:>14,.0f}{tot_new:>14,.0f}"
          f"   {'O' if abs(src_sales-tot_new) < 0.5 else 'X'}")
    print(f"   {'줄 수':<8}{len(recs):>14,}{new_rows:>14,}"
          f"   {'O' if new_rows == len(recs) else 'X'}")

    # 원본 '결산' 시트와 날짜별로 대조한다. 2026-09-21 에 이 대조로 11일
    # 시트의 열이 밀린 것을 찾아냈다. (미수 10만원이 매출로 들어가 있었다)
    gs = read_gyeolsan(src)
    if gs:
        diff = [(d, new_day[d], gs.get(d, 0)) for d in range(1, days + 1)
                if abs(new_day[d] - gs.get(d, 0)) >= 0.5]
        gs_tot = sum(gs.get(d, 0) for d in range(1, days + 1))
        print(f"\n   결산 시트 대조 (총금액, 미수 제외)")
        print(f"      결산 시트 {gs_tot:>14,.0f}  /  옮긴 뒤 {tot_new:>14,.0f}"
              f"   {'O' if not diff else 'X'}")
        for d, a, b in diff:
            print(f"      [!] {d}일 — 옮긴 뒤 {a:,.0f} / 결산 시트 {b:,.0f}"
                  f"  (차이 {a-b:+,.0f})")
        if diff:
            print(f"      결산 시트의 수식이 낡았을 수도 있습니다. 그 날을 봐주세요.")

    print(f"\n   품목별 건수")
    for c in 양식.ITEMS:
        print(f"      {c:<5} {new_item[c]:>3}건")

    if lens_kind:
        print(f"\n   렌즈 이름을 이렇게 갈랐습니다")
        for (k, n), cnt in sorted(lens_kind.items(), key=lambda x: (x[0][0], -x[1])):
            print(f"      {k:<4} {n:<18} {cnt}건")
    if warn_rows:
        print(f"\n   [확인 필요] 누진인지 아닌지 이름만으론 모르겠는 것 {len(warn_rows)}건")
        print(f"      단초점 칸에 넣고 분홍색으로 칠해뒀습니다. 누진이면 옮겨주세요.")
        for d, rr, n in warn_rows:
            print(f"      {d}일 {rr}행  {n}")

    print("\n" + ("전부 맞습니다." if ok else "[!] 숫자가 다릅니다 — 확인이 필요합니다."))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
