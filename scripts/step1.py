# -*- coding: utf-8 -*-
"""1단계 - 자료 읽기

  애니아이 일지 만들기  ->  장부 만들기  ->  금액 적는 파일 열기

  이 파일은 '1단계 - 자료 읽기.bat' 이 부릅니다.
  경로에 한글이 있어도 파이썬은 문제없이 다룹니다.
  (.bat 은 한글 경로를 못 다뤄서 이렇게 나눴습니다)
"""
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

AI = Path(r"C:\Download\AI")
STORE = Path(r"C:\Download\매장관리")
DASH = Path(r"C:\Download\애니아이엑셀")
OUTDIR = Path(r"C:\Users\admin\Desktop\굿미소 실행")

PY = sys.executable


def line(t=""):
    print(t, flush=True)


def head(n, t):
    line()
    line("=" * 60)
    line("  [%s] %s" % (n, t))
    line("=" * 60)


def run(args, cwd):
    try:
        r = subprocess.run(args, cwd=str(cwd))
        return r.returncode == 0
    except Exception as e:
        line("   실행하지 못했습니다 - %s" % e)
        return False


def main():
    line()
    line("  굿미소안경 - 1단계 : 자료 읽기")
    line("  (애니아이 일지 -> 장부 만들기 -> 금액 적는 파일 열기)")

    # 1. 애니아이 일지
    head(1, "애니아이에서 이번 달 매출 읽기")
    j = AI / "scripts" / "anyeye_journal.py"
    ok = False
    if j.exists():
        ok = run([PY, str(j), "--outdir", str(OUTDIR)], AI)
    else:
        line("   애니아이 일지 프로그램을 찾지 못했습니다. 건너뜁니다.")
    if not ok:
        line()
        line("   * 애니아이를 못 읽었습니다. 매장 컴퓨터가 아니거나")
        line("     애니아이 프로그램이 꺼져 있을 수 있습니다.")
        line("     손으로 적은 일일결산을 쓰신다면 그냥 넘어가도 됩니다.")

    # 2. 장부
    head(2, "장부 만들기")
    m = STORE / "program" / "main.py"
    if not m.exists():
        line("   장부 프로그램을 찾지 못했습니다 : %s" % m)
        return 1
    if not run([PY, str(m)], STORE):
        line()
        line("   * 장부를 만들지 못했습니다. 위 메시지를 봐주세요.")
        line("     엑셀 파일이 열려 있으면 멈춥니다. 전부 닫고 다시 눌러주세요.")
        return 1

    # 3. 금액 적는 파일 열기
    head(3, "금액 적는 파일 열기")
    key = date.today().strftime("%Y-%m")
    f = STORE / "2_확인필요" / ("명세서입력_%s.xlsx" % key)
    if f.exists():
        line("   %s" % f.name)
        os.startfile(str(f))
    else:
        line("   이번 달 파일이 없어 폴더만 엽니다.")
        os.startfile(str(STORE / "2_확인필요"))

    line()
    line("-" * 60)
    line("  여기까지 끝났습니다.")
    line()
    line("  이제 열린 엑셀의 [노란 칸]에 금액을 적어주세요.")
    line("    · 금액(할인 전) = 명세서 맨 아래 합계금액 (부가세 포함)")
    line("    · 색칠된 줄 = 월 청구서 6곳 (호야 일도 한알 한샘 에실로 렌즈아이)")
    line("      -> 청구서 금액을 한 줄만 적으세요. 이게 빠지면 매입이 반토막 납니다.")
    line()
    line("  다 적고 저장(Ctrl+S)한 뒤,")
    line("  [2단계 - 마무리하고 대시보드 보기] 를 눌러주세요.")
    line("-" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
