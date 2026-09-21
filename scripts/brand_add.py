# -*- coding: utf-8 -*-
"""애니아이 '분류 1(1.안경테)' 에 안경테 브랜드를 등록한다.

  브랜드등록.bat 을 더블클릭하면 된다.

■ 기존 브랜드는 하나도 건드리지 않는다.
  애니아이에 등록된 브랜드 93개 중 92개가 실제로 쓰이고 있다
  (판매 48,151줄 · 상품 1,157개가 그 번호를 참조한다).
  게다가 DB 에 외래키가 없어서 지워도 막아주지 않는다 — 그냥 고아가 된다.
  그래서 '지우기' 는 하지 않는다.

■ 지울 필요도 없다.
  옛 안경테 브랜드 75개는 전부 분류 2([테]수입) 에 있고,
  지금은 분류 1(1.안경테) 만 쓰신다. 분류 1 에는 국산·수입 둘뿐이라
  여기에 넣으면 목록에 새 브랜드만 보인다.

■ 목록은 03_데이터/안경테_브랜드목록.txt 에 있다. 한 줄에 하나.
  빼고 싶으면 그 줄을 지우고 다시 누르면 된다 (이미 있는 건 안 넣는다).

되돌리기: 03_데이터/안경테_브랜드목록.txt 를 보고 애니아이에서 지우거나,
          _백업 폴더의 .bak 으로 복구한다.
"""
import csv
import subprocess
import sys
import tempfile
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
LIST = BASE / "03_데이터" / "안경테_브랜드목록.txt"
CONN = r"Server=.\SQLEXPRESS;Integrated Security=SSPI;Database=EyeYesSolution;Connect Timeout=8"
GROUP = 1        # 1.안경테
COMPANY = 98     # GoodMiso (국산·수입 이 등록된 자리와 같다)


def ps(body, tmp, out=None):
    """PowerShell 조각을 돌린다.
    ※ .ps1 에는 한글을 한 글자도 쓰지 않는다. PowerShell 5.1 이 BOM 없는
      파일을 ANSI 로 읽어서 통째로 깨지기 때문이다.
      한글(브랜드 이름)은 UTF-8 CSV 로 넘긴다."""
    f = tmp / "run.ps1"
    f.write_text("$ErrorActionPreference='Stop'\n"
                 f"$c = New-Object System.Data.SqlClient.SqlConnection '{CONN}'\n"
                 "$c.Open()\n" + body + "\n$c.Close()\n", encoding="ascii")
    r = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-File", str(f)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        print("\n   [!] 애니아이에 연결하지 못했습니다.")
        print("       애니아이가 깔린 컴퓨터에서 돌려야 합니다.")
        if r.stderr.strip():
            print("      ", r.stderr.strip()[:400])
        sys.exit(2)
    if out and out.exists():
        with open(out, encoding="utf-8-sig", newline="") as fh:
            return list(csv.DictReader(fh))
    return r.stdout.strip()


def main():
    if not LIST.exists():
        print(f"\n   [!] 목록 파일이 없습니다: {LIST}")
        return 1
    want = [n.strip() for n in LIST.read_text(encoding="utf-8").splitlines()
            if n.strip() and not n.lstrip().startswith("#")]
    if not want:
        print("\n   [!] 목록이 비어 있습니다.")
        return 1

    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        out = tmp / "have.csv"
        have = ps("$cmd = $c.CreateCommand()\n"
                  f"$cmd.CommandText = 'SELECT Brand_Name FROM tblBrand WHERE Group_Code = {GROUP}'\n"
                  "$a = New-Object System.Data.SqlClient.SqlDataAdapter $cmd\n"
                  "$t = New-Object System.Data.DataTable\n"
                  "[void]$a.Fill($t)\n"
                  f"$t | Export-Csv -Path '{out}' -NoTypeInformation -Encoding UTF8\n",
                  tmp, out)
        already = {r["Brand_Name"].strip() for r in have}
        todo = [n for n in want if n not in already]

        print(f"\n  애니아이 안경테 브랜드 등록\n")
        print(f"   목록 파일    : {LIST.name} ({len(want)}개)")
        print(f"   이미 있는 것 : {len(already)}개  {' · '.join(sorted(already)) or '없음'}")
        print(f"   새로 넣을 것 : {len(todo)}개")
        if not todo:
            print("\n   더 넣을 것이 없습니다.")
            return 0
        for i in range(0, len(todo), 6):
            print("      " + " · ".join(todo[i:i + 6]))

        print(f"\n   분류 {GROUP}(1.안경테) 에 넣습니다. 기존 브랜드는 건드리지 않습니다.")
        ans = input("\n   넣을까요? (y 를 치고 엔터) : ").strip().lower()
        if ans not in ("y", "yes", "ㅛ"):
            print("\n   그만뒀습니다. 아무것도 바뀌지 않았습니다.")
            return 0

        # 한글 이름은 CSV 로 넘긴다 (.ps1 에 직접 쓰지 않는다)
        names_csv = tmp / "names.csv"
        with open(names_csv, "w", encoding="utf-8", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(["name"])
            for n in todo:
                w.writerow([n])

        body = (
            f"$rows = Import-Csv -Path '{names_csv}' -Encoding UTF8\n"
            "$cmd = $c.CreateCommand()\n"
            "$cmd.CommandText = 'INSERT INTO tblBrand (Company_Code, Group_Code, "
            "Brand_Name, MasterBrand_Code, Brand_YN) VALUES (@co, @gr, @nm, 0, 1)'\n"
            "[void]$cmd.Parameters.Add('@co',[System.Data.SqlDbType]::Int)\n"
            "[void]$cmd.Parameters.Add('@gr',[System.Data.SqlDbType]::Int)\n"
            "[void]$cmd.Parameters.Add('@nm',[System.Data.SqlDbType]::VarChar,50)\n"
            f"$cmd.Parameters['@co'].Value = {COMPANY}\n"
            f"$cmd.Parameters['@gr'].Value = {GROUP}\n"
            "$n = 0\n"
            "foreach ($r in $rows) {\n"
            "  $cmd.Parameters['@nm'].Value = $r.name\n"
            "  $n += $cmd.ExecuteNonQuery()\n"
            "}\n"
            "Write-Output $n\n")
        n = ps(body, tmp)
        print(f"\n   {n}개를 넣었습니다.")

        out2 = tmp / "after.csv"
        after = ps("$cmd = $c.CreateCommand()\n"
                   f"$cmd.CommandText = 'SELECT Brand_Code, Brand_Name FROM tblBrand "
                   f"WHERE Group_Code = {GROUP} ORDER BY Brand_Name'\n"
                   "$a = New-Object System.Data.SqlClient.SqlDataAdapter $cmd\n"
                   "$t = New-Object System.Data.DataTable\n"
                   "[void]$a.Fill($t)\n"
                   f"$t | Export-Csv -Path '{out2}' -NoTypeInformation -Encoding UTF8\n",
                   tmp, out2)
        print(f"\n   이제 분류 1(1.안경테) 브랜드는 {len(after)}개입니다:")
        names = [r["Brand_Name"] for r in after]
        for i in range(0, len(names), 6):
            print("      " + " · ".join(names[i:i + 6]))
        print("\n   애니아이를 껐다 켜면 목록에 나옵니다.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except KeyboardInterrupt:
        print("\n   그만뒀습니다.")
        sys.exit(1)
    except Exception as e:
        print("\n   [!] 예상 못 한 일이 생겼습니다.")
        print(f"       {type(e).__name__}: {e}")
        sys.exit(9)
