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

# 옛 양식 열 위치 (1부터)
OLD = dict(no=1, staff=2, cust=3, cash=4, cre=5, gift=6, card=7, due=8,
           frame=9, lens=10, contact=11, memo=12)
MONEY_MAP = [("cash", "현금"), ("cre", "현금영수증"), ("gift", "상품권"),
             ("card", "카드"), ("due", "미수")]
# 렌즈 이름에 이게 들어 있으면 누진으로 본다
PROG = ("누진", "다초점", "기능성", "프로그레")
# 누진인지 단초점인지 이름만으로는 알 수 없는 것들 — 사용자에게 물어봐야 한다
UNSURE = ("라이프스타일", "드라이브", "오피스", "리딩", "실내")


def read_old(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out, sheets = [], 0
    for name in wb.sheetnames:
        mm = re.fullmatch(r"(\d+)일", name)
        if not mm:
            continue
        sheets += 1
        ws = wb[name]
        for r in ws.iter_rows(min_row=5, max_col=14, values_only=True):
            if not r or not isinstance(r[0], (int, float)):
                continue      # 번호가 숫자인 줄만 판매 기록
            def g(k):
                i = OLD[k] - 1
                return r[i] if i < len(r) else None
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
    return out, sheets


def split_lens(name):
    """렌즈 이름을 (누진, 단초점) 중 하나로 가른다."""
    if not name:
        return None
    return "누진" if any(k in name for k in PROG) else "단초점"


def main():
    if len(sys.argv) < 2:
        sys.exit("쓰는 법: python scripts/일지_옮기기.py \"원본.xlsx\" [연] [월]")
    src = Path(sys.argv[1])
    if not src.exists():
        sys.exit(f"[!] 파일이 없습니다: {src}")
    if len(sys.argv) >= 4:
        y, m = int(sys.argv[2]), int(sys.argv[3])
    else:
        mm = re.search(r"(\d{2,4})\s*[년\-.]\s*(\d{1,2})\s*월?", src.stem)
        if not mm:
            sys.exit("[!] 파일명에서 연·월을 못 읽었습니다. 연도와 월을 붙여주세요.")
        y = int(mm.group(1))
        y += 2000 if y < 100 else 0
        m = int(mm.group(2))

    print(f"원본: {src.name}  →  {y}년 {m}월 새 양식\n")
    recs, sheets = read_old(src)
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

    out = SALES_DIR / f"{y}년 {m}월.xlsx"
    if out.exists():
        bak = SALES_DIR / f"{y}년 {m}월 (옮기기전).xlsx"
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
