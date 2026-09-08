# -*- coding: utf-8 -*-
"""
03_데이터/사입단가.xlsx 를 정리된 형식으로 다시 만든다.

  · 구분(렌즈/콘택트/프레임) 열을 붙인다
  · '90p' 처럼 앞 이름이 생략된 행을 펼친다  (모이스트30p → 모이스트90p)
  · 같은 이름이 여러 번 나오면 '중복' 으로 표시한다
  · 매입단가와 소비자가를 나란히 두고, 마진율을 자동 계산한다

원본은 사입단가_원본.xlsx 로 남겨둔다.
"""
import shutil
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DATA = Path(__file__).resolve().parent.parent / "03_데이터"
SRC = DATA / "사입단가.xlsx"
BACKUP = DATA / "사입단가_원본.xlsx"

# 원본 행 범위로 구분을 정한다 (사용자가 그렇게 적어 넣었다)
RANGES = [(2, 18, "렌즈"), (19, 39, "콘택트"), (40, 97, "프레임")]


def main():
    if not SRC.exists():
        print(f"[!] {SRC.name} 이 없습니다.")
        return
    if not BACKUP.exists():
        shutil.copy(SRC, BACKUP)
        print(f"원본을 {BACKUP.name} 로 보관했습니다.")

    ws = openpyxl.load_workbook(BACKUP, data_only=True).active
    raw = list(ws.iter_rows(min_row=1, max_col=3, values_only=True))

    items, prev_stem = [], ""
    for n in range(2, len(raw) + 1):
        name, cost, price = (list(raw[n - 1]) + [None] * 3)[:3]
        name = str(name).strip() if name else ""
        if not name:
            continue
        kind = next((k for a, b, k in RANGES if a <= n <= b), "기타")

        # '90p' / '12p' 처럼 숫자로 시작하면 앞 품목의 다른 용량이다
        full = name
        if kind == "콘택트" and name[0].isdigit() and prev_stem:
            full = prev_stem + name
        elif kind == "콘택트":
            # '모이스트30p' → 앞머리 '모이스트' 를 기억
            stem = name
            for suf in ("30p", "90p", "6p", "12p", "30", "90"):
                if stem.endswith(suf):
                    stem = stem[: -len(suf)]
                    break
            prev_stem = stem.strip()

        items.append({"구분": kind, "품목명": full,
                      "매입단가": cost if isinstance(cost, (int, float)) else None,
                      "소비자가": price if isinstance(price, (int, float)) else None,
                      "원본행": n})

    dup = {k for k, v in Counter(i["품목명"] for i in items).items() if v > 1}

    wb = openpyxl.Workbook()
    ws2 = wb.active
    ws2.title = "사입단가"
    ws2.append(["구분", "품목명", "매입단가", "소비자가", "마진", "마진율", "비고"])
    order = {"프레임": 0, "렌즈": 1, "콘택트": 2, "기타": 3}
    items.sort(key=lambda x: (order.get(x["구분"], 9), x["품목명"]))

    for i in items:
        r = ws2.max_row + 1
        note = []
        if i["품목명"] in dup:
            note.append("이름 중복 — 모델을 구분해 이름을 다르게 적어주세요")
        if i["매입단가"] is None:
            note.append("매입단가 비어 있음")
        if i["소비자가"] is None:
            note.append("소비자가 비어 있음")
        ws2.append([i["구분"], i["품목명"], i["매입단가"], i["소비자가"],
                    f"=IF(AND(C{r}<>\"\",D{r}<>\"\"),D{r}-C{r},\"\")",
                    f"=IF(AND(C{r}<>\"\",D{r}<>\"\",D{r}>0),(D{r}-C{r})/D{r},\"\")",
                    " / ".join(note)])

    head_fill = PatternFill("solid", fgColor="1F3864")
    for c in ws2[1]:
        c.fill = head_fill
        c.font = Font(color="FFFFFF", bold=True, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, w in enumerate([9, 26, 13, 13, 13, 10, 46], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    thin = Side(style="thin", color="D9D9D9")
    for row in ws2.iter_rows(min_row=1):
        for c in row:
            c.border = Border(bottom=thin)
    for row in ws2.iter_rows(min_row=2, min_col=3, max_col=5):
        for c in row:
            c.number_format = "#,##0"
    for row in ws2.iter_rows(min_row=2, min_col=6, max_col=6):
        for c in row:
            c.number_format = "0.0%"

    # 채워야 할 칸을 눈에 띄게 — 노란 칸만 채우면 된다
    todo = PatternFill("solid", fgColor="FFF2CC")
    warn = PatternFill("solid", fgColor="FCE4EC")
    for r in range(2, ws2.max_row + 1):
        if ws2.cell(r, 3).value is None:
            ws2.cell(r, 3).fill = todo
        if ws2.cell(r, 4).value is None:
            ws2.cell(r, 4).fill = todo
        if "중복" in str(ws2.cell(r, 7).value or ""):
            ws2.cell(r, 2).fill = warn

    ws2.freeze_panes = "C2"
    ws2.auto_filter.ref = ws2.dimensions

    # 쓰는 법 안내 시트
    hs = wb.create_sheet("사용법")
    for line in [
        "사입단가표 — 쓰는 법",
        "",
        "1. 노란색 칸을 채우시면 됩니다.",
        "   · 매입단가 = 거래처에서 사올 때 우리가 낸 값",
        "   · 소비자가 = 손님에게 파는 값",
        "   둘 다 있으면 마진과 마진율이 자동으로 계산됩니다.",
        "",
        "2. 분홍색으로 표시된 품목명은 같은 이름이 두 번 이상 있습니다.",
        "   모델이 다르면 이름을 다르게 적어주세요.",
        "   예) 버나드  →  버나드(티타늄) / 버나드(뿔테)",
        "   이름이 같으면 어느 단가를 써야 할지 프로그램이 알 수 없습니다.",
        "",
        "3. 한 번에 다 채우지 않으셔도 됩니다.",
        "   아는 것부터 채우시면, 채운 품목부터 마진이 계산됩니다.",
        "",
        "4. 새 품목이 생기면 맨 아래에 한 줄 추가하시면 됩니다.",
        "   구분은 프레임 / 렌즈 / 콘택트 중 하나로 적어주세요.",
        "",
        "5. 고치신 뒤에는 '실행' 을 눌러주세요. 그때 반영됩니다.",
    ]:
        hs.append([line])
    hs.column_dimensions["A"].width = 76
    hs["A1"].font = Font(bold=True, size=13)

    wb.save(SRC)

    n_dup = sum(1 for i in items if i["품목명"] in dup)
    no_cost = sum(1 for i in items if i["매입단가"] is None)
    no_price = sum(1 for i in items if i["소비자가"] is None)
    print(f"\n정리 완료 — 품목 {len(items)}개")
    for k in ("프레임", "렌즈", "콘택트"):
        print(f"   {k}: {sum(1 for i in items if i['구분'] == k)}개")
    print(f"\n   이름 중복        {n_dup}개  ← 모델을 구분해 이름을 다르게 적어주세요")
    print(f"   매입단가 비어있음 {no_cost}개")
    print(f"   소비자가 비어있음 {no_price}개")
    print(f"\n저장: {SRC.name}  (원본: {BACKUP.name})")


if __name__ == "__main__":
    main()
