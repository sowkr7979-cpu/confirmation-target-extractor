# -*- coding: utf-8 -*-
"""E2E 시험을 한 번에 돌린다. `E2E테스트.bat`이 이 파일을 부른다.

  1) 견본 분개장 만들기 (양식이 서로 다른 2개)
  2) 요구사항 테스트 (tests/verify.py)
  3) 실제 저장 (src/app.py — 양식을 복사해 Control Sheet를 새로 만든다)

실제 감사 자료는 쓰지 않는다. 모두 합성 데이터다.
"""
from __future__ import annotations

import os
import subprocess
import sys

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(BASE, "tests", "_out")
JOURNAL = os.path.join(BASE, "samples", "샘플_분개장.xlsx")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass


def step(no: int, title: str) -> None:
    print(f"\n[{no}/3] {title}", flush=True)


def run(args: list[str], log_name: str) -> tuple[int, str]:
    """자식 프로세스를 돌리고 로그를 파일로 남긴다. (종료코드, 로그본문)"""
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.run([sys.executable, *args], cwd=BASE, env=env,
                          capture_output=True, text=True,
                          encoding="utf-8", errors="replace")
    text = (proc.stdout or "") + (proc.stderr or "")
    path = os.path.join(OUT, log_name)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return proc.returncode, text


def excel_running() -> int:
    """열려 있는 Excel 창 수. 시험 중 Excel이 닫히면 COM 호출이 끊긴다."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
                             capture_output=True, text=True, encoding="cp949",
                             errors="replace").stdout
    except OSError:
        return 0
    return sum(1 for line in out.splitlines() if "EXCEL.EXE" in line.upper())


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    print("=" * 62)
    print(" 조회처 추출 — E2E 시험 (실제 감사 자료는 쓰지 않습니다)")
    print("=" * 62)

    n = excel_running()
    if n:
        print()
        print(f"[주의] Excel이 {n}개 떠 있습니다.")
        print("  시험은 Excel을 여러 번 열고 닫습니다. 작업 중인 문서가 있으면 저장 후 닫고")
        print("  다시 실행하세요. 그대로 두면 COM 호출이 끊겨 테스트가 실패할 수 있습니다.")
        print("  (지금 그대로 진행합니다)")

    step(1, "견본 분개장 만들기 — 양식이 서로 다른 2개")
    rc, text = run([os.path.join("tools", "make_samples.py")], "1_견본.txt")
    for line in text.splitlines():
        if line.startswith("만듦"):
            print("  " + line)
    if rc != 0:
        print(text[-1500:])
        return rc

    step(2, "요구사항 테스트 — 전체 로그는 tests/_out/2_테스트.txt")
    rc, text = run([os.path.join("tests", "verify.py")], "2_테스트.txt")
    lines = [l for l in text.strip().splitlines() if l.strip()]
    fails = [l for l in lines if "[   FAIL   ]" in l]
    print("  " + (lines[-1] if lines else "결과 없음"))
    for l in fails:
        print("  " + l)
    if rc != 0:
        return rc

    step(3, "실제 저장 — 양식을 복사해 Control Sheet를 새로 만듭니다")
    rc, text = run([os.path.join("src", "app.py"), "--journal", JOURNAL], "3_저장.txt")
    made = ""
    for line in text.splitlines():
        if line.startswith("출력 파일:"):
            made = line.split(":", 1)[1].strip()
        if line.startswith(("양식 복사", "조회처", "검증 요약", "출력 파일")):
            print("  " + line)
    if rc != 0:
        print(text[-1500:])
        return rc

    print("\n" + "=" * 62)
    print(" 끝났습니다. 다음을 눈으로 확인하세요.")
    print(f"   1) {made or 'samples/output/ 의 새 Control Sheet'}")
    print("      '정보입력&출력&Control sheet' 시트 A59부터 조회처 19건")
    print("   2) H열 YES / C열 온라인 표시 / B·D:G·I:L 공란")
    print("   3) 맨 뒤 '자동화_' 시트 4개 — 제외후보에 4대보험 3건·PG 2건이 근거와 함께 보존")
    print("   4) 양식(template/Control Sheet_표준양식.xlsm)은 그대로인지")
    print("\n 화면까지 보려면 dist/조회처추출.exe 를 켜고 견본 분개장을 고르세요.")
    print("=" * 62)
    return 0


if __name__ == "__main__":
    sys.exit(main())
