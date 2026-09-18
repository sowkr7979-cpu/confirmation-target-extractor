# -*- coding: utf-8 -*-
"""쓰던 Control Sheet에서 개인·회사 정보를 지워 **빈 표준 양식**을 만든다.

    python tools/make_template.py "<쓰던 Control Sheet.xlsm>" [지울이름 ...]

지울 이름(양식 저작자 등)은 tools/sanitize_marks.txt 에 한 줄씩 적거나 인자로 준다.
코드에는 적지 않는다 — 공유본에 이름을 남기지 않기 위해서다.

지우는 것
  - 회계법인명·사업자번호·담당 회계사명·연락처·이메일 (Control sheet 상단)
  - 감사대상법인 정보 (법인명·법인등록번호·사업자등록번호·대표이사명·담당자)
  - 조회처 목록 전체 (A59 이하 — 기관명·주소·담당자 실명·직통번호)
  - '읽어주세요' 시트의 회계법인·회계사·사무실 주소 문구
  - '사업자목록양식' 시트의 입력값

  - 양식 저작자 표시(이름·소속·이메일)
  - 파일 안에만 남는 흔적: 외부 링크 경로, 문서 속성, 저장 경로, 메모 작성자 이름

남기는 것
  - 시트 구성·서식·수식·매크로·도형, 조회서 양식, 기관 주소록(집중처리금융기관)
  - 사용 조건 문구(`Copyleft(무단수정재배포 가능, 단 상업적 이용 금지)`) — 조건 자체는 남긴다

끝나면 남은 개인정보를 스스로 다시 검사하고, 걸리면 실패로 끝낸다.
"""
from __future__ import annotations

import os
import re
import shutil
import sys

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(BASE, "template", "Control Sheet_표준양식.xlsm")
SHEET = "정보입력&출력&Control sheet"
LIST_FIRST_ROW = 59
LIST_LAST_COL = 14          # N열까지

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

# (시트, 지울 범위) — 라벨이 아니라 '입력값' 칸만 고른다.
CLEAR_RANGES = [
    (SHEET, "D9"),    # 회계법인명
    (SHEET, "D11"),   # 회계법인 사업자번호
    (SHEET, "D13:D17"),   # 환급계좌 은행·번호·예금주
    (SHEET, "D19"),   # 담당 회계사명
    (SHEET, "D21:D25"),   # 연락처·fax·이메일
    (SHEET, "D28:D38"),   # 감사대상법인 정보 일체
    ("읽어주세요", "B3"),
    ("읽어주세요", "H3:H5"),
    ("읽어주세요", "C38"),
    ("읽어주세요", "B54"),
    ("사업자목록양식", "B3"),
    ("사업자목록양식", "A5:E25"),
    ("주소록출력", "A1:L62"),      # 지난 실행의 봉투 출력물 — 담당자 실명·휴대번호가 남아 있다
    # 양식 저작자 표시(이름·소속·이메일). 사용 조건 문구(C23)는 남긴다.
    ("안내&설명서", "C3"),
    ("안내&설명서", "C13"),
    ("안내&설명서", "C25"),
    (SHEET, "H3"),
]

# 지울 이름(양식 저작자 등)은 코드에 적지 않는다. 공유본에 남기지 않기 위해서다.
# tools/sanitize_marks.txt 에 한 줄에 하나씩 적으면 읽어 쓴다(이 파일은 공유 대상이 아니다).
MARKS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sanitize_marks.txt")


def _load_marks() -> tuple[str, ...]:
    marks = ["Copyleft by"]          # 이름이 아닌 일반 표지만 기본값으로 둔다
    if os.path.exists(MARKS_FILE):
        with open(MARKS_FILE, encoding="utf-8") as f:
            marks += [line.strip() for line in f
                      if line.strip() and not line.startswith("#")]
    marks += [m for m in sys.argv[2:] if m.strip()]
    return tuple(dict.fromkeys(marks))


AUTHOR_MARKS = _load_marks()

# 세척 뒤 남아 있으면 안 되는 것
FORBIDDEN = [
    re.compile(r"[가-힣]{2,10}회계법인"),
    re.compile(r"01[016-9]-?\d{3,4}-?\d{4}"),
    re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"),
    re.compile(r"\d{3}-\d{2}-\d{5}"),          # 사업자등록번호
]
ALLOW = ("회계법인명", "회계법인 또는 감사반", "회계법인계좌아님", "회계법인 및 감사반")
MAX_SCAN_ROWS = 2000


def strip_file_traces(path: str) -> list[str]:
    """Excel 화면에 안 보이는 흔적을 파일 안에서 직접 지운다.

    셀을 지워도 남는 것이 있다.
      - 외부 링크 Target: 원작자 PC의 파일 경로(다른 회사명 포함)가 그대로 있다.
      - 메모 작성자 이름: `<authors>`에 남는다.
    링크 구조를 지우면 Excel이 파일을 못 연다. 그래서 경로 문자열만 바꾼다.
    """
    import re
    import zipfile

    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
        order = list(z.namelist())

    done: list[str] = []

    # 외부 링크: 원작자 PC의 경로(다른 회사명 포함)가 Target에 그대로 있다.
    # 링크 자체를 지우면 Excel이 파일 열기를 거부한다(실측). 경로만 바꾼다.
    changed = 0
    for key in [k for k in parts if k.startswith("xl/externalLinks/") and k.endswith(".rels")]:
        text = parts[key].decode("utf-8")
        new = re.sub(r'Target="[^"]*"', 'Target="file:///C:/none.xlsx"', text)
        if new != text:
            parts[key] = new.encode("utf-8")
            changed += 1
    if changed:
        done.append(f"외부 링크 경로 치환({changed}개) — 링크 구조는 그대로 둔다")

    # 문서 속성(만든 이·마지막 저장자)
    core = parts.get("docProps/core.xml")
    if core:
        text = core.decode("utf-8")
        new = re.sub(r"(<dc:creator>)[^<]*(</dc:creator>)",
                     lambda m: m.group(1) + "양식" + m.group(2), text)
        new = re.sub(r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)",
                     lambda m: m.group(1) + "양식" + m.group(2), new)
        if new != text:
            parts["docProps/core.xml"] = new.encode("utf-8")
            done.append("문서 속성(만든 이·마지막 저장자) 정리")
    # 저장한 PC의 절대경로(x15ac:absPath) — 사용자 이름이 그대로 들어간다.
    wbx = parts.get("xl/workbook.xml")
    if wbx:
        text = wbx.decode("utf-8")
        new_text = re.sub(r"<mc:AlternateContent[^>]*>.*?</mc:AlternateContent>", "",
                          text, flags=re.S)
        if new_text != text:
            parts["xl/workbook.xml"] = new_text.encode("utf-8")
            done.append("저장 경로(absPath) 제거 — 사용자 이름이 들어간다")

    # 이진 파일 안에 남은 이름. **길이를 그대로 두고** 바꿔야 파일이 깨지지 않는다.
    #   vbaProject.bin : 매크로 모듈 머리말의 작성자 주석(ASCII)
    #   image1.emf     : 그림 안에 그려진 원작자 PC 경로(UTF-16LE)
    for name, encoding in (("xl/vbaProject.bin", "ascii"),
                           ("xl/media/image1.emf", "utf-16-le")):
        blob = parts.get(name)
        if not blob:
            continue
        before = blob
        for mark in AUTHOR_MARKS:
            try:
                token = mark.encode(encoding)
            except UnicodeEncodeError:
                continue
            if token and token in blob:
                filler = (" " * len(mark)).encode(encoding)
                blob = blob.replace(token, filler)
        if blob != before:
            parts[name] = blob
            done.append(name.split("/")[-1] + " 안의 이름 지움(길이 유지)")
    hit = 0
    for name in [k for k in parts if re.match(r"xl/comments\d+\.xml$", k)]:
        text = parts[name].decode("utf-8")
        before = text
        for mark in AUTHOR_MARKS:
            text = text.replace(mark, "양식")
        if text != before:
            hit += 1
            parts[name] = text.encode("utf-8")
    if hit:
        done.append(f"메모 작성자 이름 정리({hit}개 파일)")

    if done:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in order:
                if n in parts:
                    z.writestr(n, parts[n])
    return done


def _com():
    import win32com.client as w32
    app = w32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False          # 매크로 자동실행 차단
    app.AskToUpdateLinks = False
    app.ScreenUpdating = False
    try:
        app.AutomationSecurity = 3
    except Exception:  # noqa: BLE001
        pass
    return app


def scan(wb) -> list[str]:
    """남아 있는 개인·회사 정보를 찾는다. 셀 값과 수식을 모두 본다."""
    hits: list[str] = []
    for ws in wb.Worksheets:
        try:
            ur = ws.UsedRange
            rows, cols = int(ur.Rows.Count), int(ur.Columns.Count)
        except Exception:  # noqa: BLE001
            continue
        if rows > MAX_SCAN_ROWS:      # 빈 행이 UsedRange에 잡힌 경우
            ur = ws.Range(ws.Cells(1, 1), ws.Cells(MAX_SCAN_ROWS, max(cols, 1)))
        try:
            block = ur.Formula
        except Exception:  # noqa: BLE001
            continue
        cells = block if isinstance(block, tuple) else ((block,),)
        for r, row in enumerate(cells, start=int(ur.Row)):
            row = row if isinstance(row, tuple) else (row,)
            for c, v in enumerate(row, start=int(ur.Column)):
                if not isinstance(v, str) or not v.strip():
                    continue
                if any(a in v for a in AUTHOR_MARKS):
                    hits.append(f"{ws.Name}!R{r}C{c}: 저작자 표시 남음 ({v[:40]!r})")
                    continue
                for rx in FORBIDDEN:
                    m = rx.search(v)
                    if m and not any(a in v for a in ALLOW):
                        hits.append(f"{ws.Name}!R{r}C{c}: {m.group(0)!r} ({v[:40]!r})")
    return hits


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    src = os.path.abspath(sys.argv[1])
    if not os.path.exists(src):
        print(f"원본이 없습니다: {src}")
        return 2

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    if os.path.exists(OUT):
        os.remove(OUT)
    shutil.copy2(src, OUT)
    print(f"복사: {os.path.basename(src)} → template/{os.path.basename(OUT)}")

    app = _com()
    try:
        wb = app.Workbooks.Open(OUT, UpdateLinks=0, ReadOnly=False)
        for sheet, addr in CLEAR_RANGES:
            try:
                rng = wb.Worksheets(sheet).Range(addr)
                try:
                    rng.ClearContents()
                except Exception:  # noqa: BLE001
                    # 병합된 셀은 병합 영역 전체로만 지울 수 있다.
                    for cell in rng:
                        cell.MergeArea.ClearContents()
                print(f"  지움: {sheet}!{addr}")
            except Exception as exc:  # noqa: BLE001
                print(f"  건너뜀: {sheet}!{addr} ({exc})")

        ws = wb.Worksheets(SHEET)
        last = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)      # xlUp
        last = max(last, LIST_FIRST_ROW)
        ws.Range(ws.Cells(LIST_FIRST_ROW, 1),
                 ws.Cells(max(last, LIST_FIRST_ROW + 400), LIST_LAST_COL)).ClearContents()
        print(f"  지움: 조회처 목록 A{LIST_FIRST_ROW}:N{max(last, LIST_FIRST_ROW + 400)}")

        # 빈 행이 UsedRange에 잡혀 파일이 무거워지는 것을 막는다.
        tail = LIST_FIRST_ROW + 400
        if last > tail or int(ws.UsedRange.Rows.Count) > tail:
            ws.Rows(f"{tail + 1}:{ws.Rows.Count}").Delete()
            print(f"  정리: {tail + 1}행 이하 빈 행 삭제")

        ws.Activate()
        ws.Range("A1").Select()
        wb.Save()
        wb.Close(SaveChanges=False)

        wb2 = app.Workbooks.Open(OUT, UpdateLinks=0, ReadOnly=True)
        hits = scan(wb2)
        sheets = [s.Name for s in wb2.Sheets]
        wb2.Close(SaveChanges=False)
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass

    for line in strip_file_traces(OUT):
        print(f"  파일 내부: {line}")

    print(f"\n시트 {len(sheets)}개 유지: {', '.join(sheets)}")
    if hits:
        print(f"\n[실패] 개인·회사 정보가 {len(hits)}건 남아 있습니다:")
        for h in hits[:20]:
            print("  " + h)
        print("CLEAR_RANGES에 해당 칸을 추가하고 다시 실행하세요.")
        return 1
    print("\n[확인] 회계법인명·휴대전화·이메일·사업자등록번호 패턴 0건")
    print(f"완성: {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
