# -*- coding: utf-8 -*-
"""애니아이에 친 것을 읽어 일일결산 엑셀을 만든다.

  애니아이일지.bat 을 더블클릭하면 이번 달이 만들어진다.
  다른 달을 보려면:  python scripts/anyeye_journal.py 2026 9

※ 파일 이름이 영문인 이유 — .bat 은 ASCII 만 쓸 수 있어서
  한글 이름의 .py 를 부를 수가 없다.

■ 결과는 .bat 이 놓인 자리 옆 '07_애니아이' 폴더에 나온다.
  .bat 이 --outdir 로 자기 위치를 알려준다. 바탕화면에 두든
  프로그램 폴더에 두든 결과가 늘 .bat 옆에 생기게 하려는 것이다.

■ 애니아이 자료만 읽는다. 손으로 적은 06_매출 파일은 건드리지 않는다.
  결과는 07_애니아이 폴더에 따로 나온다. 06_매출 에 넣으면 run.py 가
  같은 달을 두 번 읽어 매출이 두 배가 된다.

■ 매번 그 달 전체를 다시 만든다.
  애니아이가 원본이라 다시 만들어도 사람이 손댄 것이 날아갈 일이 없다.
  그래서 오늘 누르면 오늘까지, 내일 누르면 내일까지 저절로 이어진다.

■ 손으로 적은 파일이 있으면 자동으로 대조해서 어디가 다른지 보여준다.
  두 방식을 같이 쓰는 동안 이게 안전장치다.

애니아이는 이 PC의 SQL Server(.\\SQLEXPRESS, EyeYesSolution)에 있다.
파이썬용 드라이버가 없어서 PowerShell 로 CSV 를 뽑아 읽는다.
※ 그 PowerShell 조각에는 한글을 쓰지 않는다. PowerShell 5.1 이 BOM 없는
   파일을 ANSI 로 읽어 통째로 깨지기 때문이다.
"""
import calendar
import csv
import re
import subprocess
import sys
import tempfile
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

from openpyxl.styles import PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
import 일일결산_양식만들기 as 양식

BASE = Path(__file__).resolve().parent.parent
OUT_DIR = BASE / "07_애니아이"
SALES_DIR = BASE / "06_매출"

# ── 애니아이를 쓰기 시작한 날 ────────────────────────────────────
# 이 날부터는 애니아이에서 읽는다. 그 전 날은 손으로 친 일일결산
# (06_매출/YYYY년 M월.xlsx) 에서 그대로 가져온다.
#
# 왜 이렇게 하나:
#   2026-09-20 까지는 뜨내기 손님을 애니아이에 안 찍어서 애니아이가
#   실제보다 적다. 그 기간은 손으로 친 엑셀이 맞다.
#   2026-09-21 부터 애니아이에 다 찍기 시작했다.
#
# 10월부터는 신경 쓸 것이 없다. 모든 날이 이 날보다 뒤라서
# 그냥 애니아이만 읽는다. 이 설정이 끼어들지 않는다.
ANYEYE_FROM = date(2026, 9, 21)
CONN = r"Server=.\SQLEXPRESS;Integrated Security=SSPI;Database=EyeYesSolution;Connect Timeout=8"

# 애니아이 분류 → 일일결산 칸
FRAME_G = {"1", "2"}                 # 1.안경테 / [테]수입
LENS_G = {"3", "4"}                  # 2.렌즈 / [렌즈]수입
CONTACT_G = {"5", "6", "7"}          # 소프트 / 일회용 / 하드

# 누진·기능성 판별 — 9월 일일결산과 애니아이를 건별로 맞춰 확인한 규칙
PROG_MM = re.compile(r"(?<![0-9.])(\d{1,2})\s*(?:mm|미리|파이)", re.I)
PROG = ("누진", "다초점", "프로그레", "내면멀티", "멀티포컬",
        "hd10", "hd20", "hd30", "hd40", "hd 40",
        "dx3", "dx5", "dx7", "dx9", "dc5", "에센셜",
        "라이프스타일", "언루트", "발란시스", "리버티",
        "수퍼브", "슈퍼브", "퓨어브", "바리락스", "피지오",
        "컴포트맥스", "컴포트 맥스", "컴포드 맥스",
        "써미트", "앰플리튜드", "스타터")
# 기능성도 누진 칸에 넣는다. 일일결산에 기능성 칸이 따로 없고,
# 옛 양식에서도 '누진*기능성' 한 칸이었다.
FUNC = ("기능성", "스텔리스트", "프리벤시아", "프레벤시아", "마이오", "근시억제",
        "홈앤오피스", "홈엔오피스", "홈 앤 오피스", "릴렉시", "워크스마트",
        "워크스타일", "드라이브세이프", "드라이브 세이프")
# 취소·교환 메모는 판매가 아니다
NOT_PROG = ("누진부적응", "누진 재교정", "다초점변심", "다초점 교환예정", "다초점 반품")


def is_prog(name):
    s = (name or "").strip().lower()
    if not s or any(x in s for x in NOT_PROG):
        return False
    if any(k in s for k in FUNC):
        return True
    m = PROG_MM.search(s)
    if m and 8 <= int(m.group(1)) <= 25:
        return True
    return any(k in s for k in PROG)


def pull(sql, tmp, name):
    """PowerShell 로 애니아이에서 CSV 를 뽑는다 (ASCII 로만 쓴다)."""
    out = tmp / name
    ps = tmp / (name + ".ps1")
    ps.write_text(
        "$ErrorActionPreference='Stop'\n"
        f"$c = New-Object System.Data.SqlClient.SqlConnection '{CONN}'\n"
        "$c.Open()\n"
        "$cmd = $c.CreateCommand()\n"
        f"$cmd.CommandText = @'\n{sql}\n'@\n"
        "$cmd.CommandTimeout = 180\n"
        "$a = New-Object System.Data.SqlClient.SqlDataAdapter $cmd\n"
        "$t = New-Object System.Data.DataTable\n"
        "[void]$a.Fill($t)\n"
        "$c.Close()\n"
        f"$t | Export-Csv -Path '{out}' -NoTypeInformation -Encoding UTF8\n",
        encoding="ascii")
    r = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-File", str(ps)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0 or not out.exists():
        print("[!] 애니아이에서 자료를 못 읽었습니다.")
        print("    애니아이가 깔린 PC 에서 돌려야 하고, SQL Server 가 켜져 있어야 합니다.")
        if r.stderr.strip():
            print("   ", r.stderr.strip()[:500])
        sys.exit(2)
    with open(out, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def read_anyeye(y, m, tmp):
    last = calendar.monthrange(y, m)[1]
    a, b = f"{y:04d}-{m:02d}-01", f"{y:04d}-{m:02d}-{last:02d}"
    optic = pull(f"""SELECT o.Optic_Code AS code,
       ISNULL(o.Cust_Code,0) AS ccode,
       CONVERT(char(10), o.Optic_Date, 120) AS d,
       ISNULL(cu.Cust_Name,'') AS cust,
       ISNULL(NULLIF(LTRIM(RTRIM(o.Optic_EmpManage)),''),
              ISNULL(o.Optic_EmpSell,'')) AS emp,
       ISNULL(o.Optic_CashPrice,0) AS cash,
       ISNULL(o.Optic_CardPrice,0) AS card,
       ISNULL(o.Optic_GiftCardPrice,0) AS gift,
       ISNULL(o.Optic_UncollectPrice,0) AS unpaid,
       ISNULL(o.Optic_CashReceipt,'') AS creceipt,
       ISNULL(o.Optic_Kind,'') AS kind
FROM tblOptic o LEFT JOIN tblCust cu ON cu.Cust_Code = o.Cust_Code
WHERE o.Optic_Date >= '{a}' AND o.Optic_Date <= '{b}'
ORDER BY o.Optic_Date, o.Optic_Code""", tmp, "optic.csv")
    sell = pull(f"""SELECT s.Optic_Code AS code, s.Group_Code AS grp,
       LTRIM(RTRIM(ISNULL(s.Brand_Name,''))) AS brand,
       LTRIM(RTRIM(ISNULL(s.Goods_Name,''))) AS nm
FROM tblSell s
WHERE s.Sell_Date >= '{a}' AND s.Sell_Date <= '{b}'
ORDER BY s.Optic_Code, s.Sell_Code""", tmp, "sell.csv")
    brands = pull("SELECT DISTINCT LTRIM(RTRIM(Brand_Name)) AS b FROM tblBrand "
                  "WHERE ISNULL(Brand_Name,'') <> ''", tmp, "brand.csv")
    known = sorted({_clean_brand(r["b"]) for r in brands} - {""}, key=len, reverse=True)
    return optic, sell, known


# 뜨내기 손님은 애니아이에서 전부 이 고객으로 찍힌다. 합치면 안 된다.
WALKIN = "G_일반판매"
# 메모가 붙은 자리 — 여기서부터는 잘라낸다
MEMO_CUT = re.compile(r"\s*(->|=>|→|\(.*반품|반품|재작업|다시작업|서비스해|교환예정)")
# 이름 끝에 붙은 값(사입가·정가) — '뉴럭스 1.6  260', '모르텐 골드 \\740', '(330)'
TAIL_PRICE = re.compile(r"\s*[\\(]?\s*\d{3,}\s*\)?\s*$")


def _clean_brand(b):
    """tblBrand 이름은 '2.레이밴' 처럼 앞에 숫자가 붙어 있다."""
    return re.sub(r"^\d+\.", "", (b or "").strip()).strip()


def brand_only(brand, nm, known):
    """안경테는 브랜드 이름만 남긴다. '어라운드 누코 31 c2' → '어라운드'"""
    b = _clean_brand(brand)
    if b:
        return b
    s = TAIL_PRICE.sub("", (nm or "").strip()).strip()
    if not s:
        return ""
    low = s.lower()
    # 브랜드 칸이 비었으면 상품명 안에서 아는 브랜드를 찾아본다
    hit = [k for k in known if len(k) >= 2 and k.lower() in low]
    if hit:
        return max(hit, key=len)
    # 그래도 모르면 첫 낱말. 품번뿐이면 그냥 상품명을 둔다.
    first = re.split(r"[\s,]+", s)[0]
    return s if re.fullmatch(r"[0-9\-./]+", first) else first


def lens_name(brand, nm):
    """렌즈는 무엇을 썼는지만. 뒤에 붙은 메모와 값은 잘라낸다."""
    s = MEMO_CUT.split((nm or "").strip(), maxsplit=1)[0].strip()
    s = TAIL_PRICE.sub("", s).strip(" -/,")
    if not s:
        s = (nm or "").strip()
    b = _clean_brand(brand)
    if b and b.lower() not in s.lower():
        s = f"{b} {s}"
    return re.sub(r"\s{2,}", " ", s)


def join(parts):
    """같은 말이 두 번 들어가지 않게 이어붙인다."""
    out = []
    for p in parts:
        p = (p or "").strip()
        if p and p not in out:
            out.append(p)
    return " / ".join(out)


def to_rows(optic, sell, known_brands):
    """애니아이 영수증을 일일결산 줄로 바꾼다.

    같은 날 같은 손님이 영수증 여러 건이면 한 줄로 합친다.
    (안경테 따로, 렌즈 따로 찍는 경우가 있어서 그냥 두면 한 사람이
     두 줄로 나와 금액이 두 번 적힌 것처럼 보인다)
    뜨내기 손님(G_일반판매)은 다 같은 고객번호라 합치지 않는다.
    """
    items = defaultdict(list)
    for s in sell:
        items[s["code"]].append(s)

    # 잘못 넣었다 되돌린 줄은 뺀다.
    # 같은 날·같은 손님에 금액이 정확히 반대인 두 줄은 서로 지워진 것이다.
    # (그냥 두면 줄이 둘 생기고 미수가 부풀어 오른다)
    cancel = set()
    for i, a in enumerate(optic):
        if a["code"] in cancel:
            continue
        pa = float(a["cash"]) + float(a["card"]) + float(a["gift"])
        if pa == 0:
            continue
        for b in optic[i + 1:]:
            if b["code"] in cancel:
                continue
            pb = float(b["cash"]) + float(b["card"]) + float(b["gift"])
            if (a["d"] == b["d"] and a["cust"] == b["cust"] and abs(pa + pb) < 0.5
                    and abs(float(a["unpaid"]) + float(b["unpaid"])) < 0.5):
                cancel.add(a["code"])
                cancel.add(b["code"])
                break
    if cancel:
        print(f"   서로 지워지는 줄   : {len(cancel)}개를 뺐습니다 "
              f"(넣었다가 되돌린 것)")
        optic = [o for o in optic if o["code"] not in cancel]

    merged = {}
    order = []
    for o in optic:
        buckets = {k: [] for k in 양식.ITEMS}
        etc = []
        for s in items.get(o["code"], []):
            nm = (s["nm"] or "").strip()
            br = (s["brand"] or "").strip()
            if not nm and not br:
                continue
            g = s["grp"]
            if g in FRAME_G:
                buckets["안경테"].append(brand_only(br, nm, known_brands))
            elif g in LENS_G:
                buckets["누진" if is_prog(nm) else "단초점"].append(lens_name(br, nm))
            elif g in CONTACT_G:
                buckets["콘택트"].append(lens_name(br, nm))
            else:
                etc.append(lens_name(br, nm))
        cash = float(o["cash"])
        cr = (o["creceipt"] or "").strip() == "발행함"
        unpaid = float(o["unpaid"])
        # 미수가 음수면 '못 받은 돈' 이 아니라 '전에 못 받았던 걸 받았다' 는 뜻이다.
        # 받은 돈은 이미 현금·카드 칸에 들어가 있으니 미수 칸은 비운다.
        # (음수를 그대로 두면 그 달 미수 합계가 깎인다)
        repaid = -unpaid if unpaid < 0 else 0
        money = {"현금": 0 if cr else cash, "현금영수증": cash if cr else 0,
                 "상품권": float(o["gift"]), "카드": float(o["card"]),
                 "미수": unpaid if unpaid > 0 else 0}
        if not any(money.values()) and not any(buckets.values()) and not etc:
            continue
        day = int(o["d"][8:10])
        cust = (o["cust"] or "").strip()
        note = o["kind"] if o["kind"] and o["kind"] not in ("판매", "일반판매") else ""
        if repaid:
            note = (note + " " if note else "") + f"미수 {repaid:,.0f}원 받음"

        # 합칠 열쇠 — 이름 있는 손님만 (뜨내기는 건건이 따로)
        if cust and cust != WALKIN:
            key = (day, o["ccode"], cust)
        else:
            key = (day, "단품", o["code"])

        if key in merged:
            r = merged[key]
            for k in buckets:
                r["_b"][k].extend(buckets[k])
            r["_etc"].extend(etc)
            for k in money:
                r[k] += money[k]
            if not r["emp"]:
                r["emp"] = (o["emp"] or "").strip()
            if note and note not in r["note"]:
                r["note"] = (r["note"] + " / " + note).strip(" /")
            r["_n"] += 1
        else:
            merged[key] = {"day": day, "emp": (o["emp"] or "").strip(),
                           "cust": "" if cust == WALKIN else cust,
                           "_b": buckets, "_etc": etc, "note": note, "_n": 1,
                           **money}
            order.append(key)

    rows = []
    for key in order:
        r = merged[key]
        rows.append({k: v for k, v in r.items() if not k.startswith("_")}
                    | {c: join(r["_b"][c]) for c in 양식.ITEMS}
                    | {"etc": join(r["_etc"]), "merged": r["_n"]})
    rows.sort(key=lambda r: (r["day"], r["cust"]))
    return rows


def build(y, m, rows):
    extra = {c: sorted({r[c] for r in rows if r[c]}) for c in 양식.ITEMS}
    wb, _, _ = 양식.make_workbook(y, m, extra)
    A, FIRST, cap = 양식.A, 양식.FIRST, 양식.ROWS_PER_DAY
    put = Counter()
    over = []
    for r in rows:
        d = r["day"]
        if put[d] >= cap:
            over.append(d)
            continue
        ws = wb[f"{d}일"]
        row = FIRST + put[d]
        put[d] += 1
        ws[f'{A["담당자"]}{row}'] = r["emp"] or None
        ws[f'{A["고객명"]}{row}'] = r["cust"] or None
        for c in 양식.ITEMS:
            ws[f"{A[c]}{row}"] = r[c] or None
        if r["etc"]:
            ws[f'{A["상세"]}{row}'] = r["etc"]
        for k in 양식.MONEY:
            if r[k]:
                ws[f"{A[k]}{row}"] = r[k]
        if not r["emp"]:                       # 담당자가 비었으면 눈에 띄게
            ws[f'{A["담당자"]}{row}'].fill = PatternFill("solid", fgColor="FFF2CC")
        if r["note"]:
            ws[f'{A["비고"]}{row}'] = r["note"]
    return wb, put, sorted(set(over))


# ── 손으로 적은 파일과 대조 ──────────────────────────────────────
def read_hand(y, m):
    """손으로 친 일일결산(06_매출) 을 통째로 읽는다.

    대조에도 쓰고, 애니아이를 쓰기 전 기간의 줄을 그대로 가져올 때도 쓴다.
    """
    import openpyxl
    f = SALES_DIR / f"{y}년 {m}월.xlsx"
    if not f.exists():
        return None
    wb = openpyxl.load_workbook(f, data_only=True)
    norm = lambda v: re.sub(r"\s+", "", str(v)) if v is not None else ""
    out = []
    for name in wb.sheetnames:
        mm = re.fullmatch(r"(\d+)일", name)
        if not mm:
            continue
        ws = wb[name]
        hdr, cm = None, {}
        for r in range(1, 9):
            vals = [norm(c.value) for c in ws[r]]
            if "현금" in vals and "카드" in vals:
                hdr = r
                # 옛 양식은 머리글이 두 줄이다 (결제수단 / 품목)
                for rr in (r, r + 1):
                    for i, v in enumerate([norm(c.value) for c in ws[rr]]):
                        if v and v not in cm:
                            cm[v] = i
                break
        if hdr is None:
            continue
        for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
            if not row or not isinstance(row[0], (int, float)):
                continue
            g = lambda k: (row[cm[k]] if k in cm and cm[k] < len(row) else None)
            num = lambda k: g(k) if isinstance(g(k), (int, float)) else 0
            txt = lambda k: (str(g(k)).strip() if g(k) not in (None, "") else "")
            money = {k: num(k) for k in 양식.MONEY}
            money["현금영수증"] = money["현금영수증"] or num("현영")
            if not any(money.values()):
                continue
            out.append({"day": int(mm.group(1)),
                        "cust": txt("고객명") or txt("이름"),
                        "emp": txt("담당자"),
                        "안경테": txt("안경테") or txt("프레임"),
                        "단초점": txt("단초점"),
                        "누진": txt("누진"),
                        "콘택트": txt("콘택트"),
                        "etc": txt("상세"),
                        "note": txt("비고"),
                        "from_excel": True,
                        **money})
    wb.close()
    return out


def compare(auto, hand, from_day):
    """손으로 적은 것과 대조. 애니아이를 쓰기 시작한 날부터만 본다."""
    if hand is None:
        print("\n   (06_매출 에 손으로 적은 파일이 없어 대조는 건너뜁니다)")
        return
    if from_day > 31:
        print("\n   이 달은 통째로 손으로 친 엑셀을 옮긴 것이라 대조할 것이 없습니다.")
        return
    auto = [r for r in auto if r["day"] >= from_day]
    hand = [r for r in hand if r["day"] >= from_day]
    if not auto and not hand:
        print(f"\n   ({from_day}일부터 대조하는데 아직 자료가 없습니다)")
        return
    days = sorted({r["day"] for r in auto} | {r["day"] for r in hand})
    print(f"\n   손으로 적은 것과 대조  ({from_day}일부터)")
    print(f"      {'날':>4}{'자동줄':>7}{'손줄':>6}{'자동합계':>13}{'손합계':>13}{'차이':>12}")
    ta = th = 0
    for d in days:
        A = [r for r in auto if r["day"] == d]
        H = [r for r in hand if r["day"] == d]
        sa = sum(r[k] for r in A for k in ("현금", "현금영수증", "상품권", "카드"))
        sh = sum(r[k] for r in H for k in ("현금", "현금영수증", "상품권", "카드"))
        ta += sa
        th += sh
        mark = "" if abs(sa - sh) < 0.5 else "  <<<"
        print(f"      {d:>4}{len(A):>7}{len(H):>6}{sa:>13,.0f}{sh:>13,.0f}{sa-sh:>12,.0f}{mark}")
    print(f"      {'합':>4}{len(auto):>7}{len(hand):>6}{ta:>13,.0f}{th:>13,.0f}{ta-th:>12,.0f}")
    if abs(ta - th) < 0.5 and len(auto) == len(hand):
        print("\n   두 방식이 똑같습니다. 이제 엑셀을 손으로 안 쓰셔도 됩니다.")
    else:
        print("\n   아직 다릅니다. 위에서 <<< 표시된 날을 보세요.")
        print("   대개 원인은 셋입니다 — 애니아이에 안 친 뜨내기 손님,")
        print("   상품권·현금영수증을 카드로 뭉뚱그린 것, 미수 칸을 고쳐버린 것.")


def main():
    global OUT_DIR
    argv = sys.argv[1:]
    if "--outdir" in argv:
        i = argv.index("--outdir")
        here = Path(argv[i + 1]) if i + 1 < len(argv) else None
        if here and here.is_dir():
            OUT_DIR = here / "07_애니아이"
        del argv[i:i + 2]
    args = [a for a in argv if not a.startswith("-")]
    if len(args) >= 2:
        y, m = int(args[0]), int(args[1])
    else:
        t = date.today()
        y, m = t.year, t.month

    print(f"\n  애니아이 → 일일결산  ({y}년 {m}월)\n")
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        optic, sell, brands = read_anyeye(y, m, tmp)
    print(f"   애니아이에서 읽음 : 판매 {len(optic)}건 · 품목 {len(sell)}줄")

    rows = to_rows(optic, sell, brands)

    # ── 애니아이를 쓰기 전 기간은 손으로 친 엑셀에서 그대로 가져온다 ──
    hand = read_hand(y, m)
    from_day = 1
    if (y, m) == (ANYEYE_FROM.year, ANYEYE_FROM.month):
        from_day = ANYEYE_FROM.day
    elif (y, m) < (ANYEYE_FROM.year, ANYEYE_FROM.month):
        from_day = 32                      # 그 달은 통째로 손으로 친 것
    borrowed = 0
    if from_day > 1:
        rows = [r for r in rows if r["day"] >= from_day]
        if hand:
            pre = [dict(h, etc=h.get("etc", ""), note=h.get("note", ""),
                        merged=1) for h in hand if h["day"] < from_day]
            borrowed = len(pre)
            rows = pre + rows
            rows.sort(key=lambda r: (r["day"], r["cust"]))
        else:
            print(f"\n   [!] {from_day}일 전 자료를 06_매출 에서 가져와야 하는데")
            print(f"       '{y}년 {m}월.xlsx' 이 없습니다.")

    if not rows:
        print("\n   [!] 그 달에 기록이 없습니다.")
        return 1

    wb, put, over = build(y, m, rows)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / f"{y}년 {m}월 (애니아이).xlsx"
    try:
        wb.save(out)
    except PermissionError:
        # 엑셀이 그 파일을 열어두고 있으면 덮어쓸 수 없다.
        # 매일 누르는 도구라 이게 제일 자주 나는 문제다.
        print("\n   [!] 파일을 저장하지 못했습니다.")
        print(f"       엑셀에서 '{out.name}' 을 열어두고 계십니다.")
        print("       그 엑셀 창을 닫고 다시 눌러주세요.\n")
        print(f"       (파일 위치: {out})")
        return 3

    last = max(r["day"] for r in rows)
    money = {k: sum(r[k] for r in rows) for k in 양식.MONEY}
    total = sum(money[k] for k in ("현금", "현금영수증", "상품권", "카드"))
    noemp = sum(1 for r in rows if not r["emp"])
    joined = sum(r["merged"] - 1 for r in rows if r["merged"] > 1)

    print(f"   만든 줄          : {len(rows)}줄 · {m}/1 ~ {m}/{last} · 영업 {len(put)}일")
    if borrowed and from_day > 31:
        print(f"   이 달은 통째로 손으로 친 엑셀을 옮겼습니다 "
              f"(애니아이를 {ANYEYE_FROM.month}/{ANYEYE_FROM.day} 부터 씁니다)")
    elif borrowed:
        print(f"   그 중 {borrowed}줄은 {m}/1~{m}/{from_day - 1} — 손으로 친 엑셀에서 가져왔습니다")
        print(f"   (그 기간은 애니아이에 뜨내기 손님을 안 찍어서 엑셀이 맞습니다)")
    if joined:
        print(f"   같은 손님 합침    : {joined}건 "
              f"(안경테·렌즈를 따로 찍은 것을 한 줄로 모았습니다)")
    print(f"\n   {'현금':<10}{money['현금']:>12,.0f}")
    print(f"   {'현금영수증':<10}{money['현금영수증']:>12,.0f}")
    print(f"   {'상품권':<10}{money['상품권']:>12,.0f}")
    print(f"   {'카드':<10}{money['카드']:>12,.0f}")
    print(f"   {'미수':<10}{money['미수']:>12,.0f}")
    print(f"   {'합계':<10}{total:>12,.0f}")

    if noemp:
        print(f"\n   [확인] 담당자가 비어 있는 줄 {noemp}개 — 노란색으로 칠해뒀습니다.")
    if over:
        print(f"\n   [!] 하루 {양식.ROWS_PER_DAY}줄이 넘어 빠진 날: {over}")

    compare(rows, hand, from_day if from_day <= 31 else 32)
    print(f"\n   저장: {out}")
    try:
        import os
        os.startfile(out)          # 만든 파일을 바로 띄워준다
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    # 무슨 일이 나든 한글 한 줄은 보이게 한다.
    # 사용자가 영어 오류를 읽기 어려워한다.
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except KeyboardInterrupt:
        print("\n   중간에 멈췄습니다.")
        sys.exit(1)
    except Exception as e:
        print("\n   [!] 예상 못 한 일이 생겼습니다.")
        print(f"       {type(e).__name__}: {e}")
        print("\n       이 화면을 그대로 찍어서 보여주시면 고치겠습니다.")
        sys.exit(9)
