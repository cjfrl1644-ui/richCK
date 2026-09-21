# -*- coding: utf-8 -*-
"""2단계 - 마무리

  장부 다시 만들기  ->  대시보드 새로고침  ->  대시보드 열기

  이 파일은 '2단계 - 마무리하고 대시보드 보기.bat' 이 부릅니다.
"""
import os
import subprocess
import sys
from pathlib import Path

STORE = Path(r"C:\Download\매장관리")
DASH = Path(r"C:\Download\애니아이엑셀")
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
        return subprocess.run(args, cwd=str(cwd)).returncode == 0
    except Exception as e:
        line("   실행하지 못했습니다 - %s" % e)
        return False


def main():
    line()
    line("  굿미소안경 - 2단계 : 마무리")
    line("  (장부 다시 만들기 -> 대시보드 새로고침 -> 대시보드 열기)")

    head(1, "장부 다시 만들기 (적으신 금액 반영)")
    m = STORE / "program" / "main.py"
    if not m.exists():
        line("   장부 프로그램을 찾지 못했습니다 : %s" % m)
        return 1
    if not run([PY, str(m)], STORE):
        line()
        line("   * 장부를 만들지 못했습니다. 엑셀을 전부 닫고 다시 눌러주세요.")
        return 1

    head(2, "대시보드 새로고침")
    r = DASH / "refresh.py"
    if not r.exists():
        line("   대시보드 프로그램을 찾지 못했습니다 : %s" % r)
        return 1
    if not run([PY, str(r)], DASH):
        line("   * 대시보드를 새로 만들지 못했습니다.")
        return 1

    head(3, "대시보드 열기")
    h = DASH / "안경원_운영_대시보드.html"
    if h.exists():
        os.startfile(str(h))
        line("   %s" % h.name)
    else:
        line("   대시보드 파일이 없습니다 : %s" % h)

    line()
    line("-" * 60)
    line("  끝났습니다.")
    line()
    line("  화면 위 단추 5개")
    line("    사장님용 / 직원용 / 매입·비용 / 순수익(비밀번호) / 인센티브 설정")
    line()
    line("  숫자가 이상하면 [매입·비용] 맨 위")
    line("  '자료가 어디까지 들어와 있나' 를 펼쳐보세요.")
    line("-" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
