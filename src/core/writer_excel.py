# -*- coding: utf-8 -*-
"""Excel COM으로 XLSM에 기록. 백업 → 임시본 작성 → 재열기 검증 → 교체.

검증 결과는 PASS / FAIL / UNVERIFIED 세 가지로 구분한다.
검사하지 못한 항목을 PASS로 올리지 않는다.
"""
from __future__ import annotations

import datetime as _dt
import os
import shutil
import time
from dataclasses import dataclass, field
from typing import Any, Callable, Sequence

LIST_FIRST_COL = 1        # A
LIST_LAST_COL = 12        # L
NAME_COL = 1              # A 조회금융기관 및 지점
ONLINE_COL = 3            # C 주소 영역 오른쪽 칸 — 온라인 발급 가능 표시
FLAG_COL = 8              # H 출력여부
START_ROW = 59

SHEET_DETAIL = "자동화_기관상세"
SHEET_EVIDENCE = "자동화_검색근거"
SHEET_SKIPPED = "자동화_제외후보"
SHEET_RUNLOG = "자동화_실행기록"
AUTO_SHEETS = (SHEET_DETAIL, SHEET_SKIPPED, SHEET_EVIDENCE, SHEET_RUNLOG)

PASS, FAIL, UNVERIFIED = "PASS", "FAIL", "UNVERIFIED"

_XL_UP = -4162
_XL_PASTE_FORMATS = -4122
_XL_VALIDATE_LIST = 3
_XL_ALERT_STOP = 1
_XL_BETWEEN = 1
_XL_CELLTYPE_FORMULAS = -4123
_XL_TEXT_FORMAT = "@"

FORMULA_SCAN_LIMIT = 5000   # 이보다 많으면 해당 시트는 미검증으로 기록


# Excel 연결이 끊겼을 때의 HRESULT. 파일 문제가 아니라 Excel 쪽 문제다.
_COM_DISCONNECTED = (-2147023170,   # 원격 프로시저를 호출하지 못했습니다
                     -2147023174,   # RPC 서버를 사용할 수 없습니다
                     -2147417848,   # 호출된 개체가 클라이언트에서 분리되었습니다
                     -2147417846,   # 호출된 개체가 호출을 거부했습니다
                     -2146959355,   # 서버 실행이 실패했습니다
                     -2147221164,   # 클래스가 등록되지 않았습니다
                     -2147221021)   # 작업을 사용할 수 없습니다
# COM 개체가 죽으면 속성 접근이 AttributeError로 온다. 이름으로 알아본다.
_DEAD_OBJECT_HINTS = ("Open.", ".Worksheets", ".Workbooks", "Excel.Application",
                      "Quit", "DispatchEx")
RETRY_LIMIT = 3
RETRY_WAIT = 5        # 초. Excel이 완전히 내려갈 시간을 준다.

DISCONNECT_GUIDE = (
    "Excel과의 연결이 끊겼습니다. 원본 파일은 그대로입니다.\n"
    "  1) 열려 있는 Excel 창을 모두 저장하고 닫으세요.\n"
    "  2) 작업 관리자에서 남아 있는 'Microsoft Excel'을 끝내세요.\n"
    "  3) 다시 실행하세요.")


class ExcelWriteError(Exception):
    """저장 단계 실패. 원본은 건드리지 않은 상태여야 한다."""

    def __init__(self, message: str, transient: bool = False) -> None:
        super().__init__(message)
        self.transient = transient


def _is_disconnected(exc: BaseException) -> bool:
    """Excel 연결이 끊긴 오류인지.

    두 가지로 알아본다.
      1) COM 오류 코드(HRESULT) — RPC 끊김·서버 실행 실패 등
      2) 죽은 COM 개체에 접근할 때 나는 AttributeError — 코드가 없으므로 이름으로 본다
    """
    codes = [getattr(exc, "hresult", None)]
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        codes.append(args[0])
    if any(c in _COM_DISCONNECTED for c in codes if c is not None):
        return True
    if isinstance(exc, AttributeError):
        text = str(exc)
        return any(h in text for h in _DEAD_OBJECT_HINTS)
    return False


@dataclass
class WriteReport:
    target: str
    backup: str = ""
    temp: str = ""
    written_rows: int = 0
    start_row: int = START_ROW
    end_row: int = START_ROW - 1
    cleared_to: int = 0
    baseline: dict[str, Any] = field(default_factory=dict)
    after: dict[str, Any] = field(default_factory=dict)
    checks: list[tuple[str, str, str]] = field(default_factory=list)  # (항목, 상태, 설명)
    replaced: bool = False

    @property
    def failures(self) -> list[tuple[str, str, str]]:
        return [c for c in self.checks if c[1] == FAIL]

    @property
    def unverified(self) -> list[tuple[str, str, str]]:
        return [c for c in self.checks if c[1] == UNVERIFIED]

    @property
    def passed(self) -> list[tuple[str, str, str]]:
        return [c for c in self.checks if c[1] == PASS]

    @property
    def ok(self) -> bool:
        """FAIL이 없어야 교체한다. UNVERIFIED는 통과로 세지 않고 기록만 한다."""
        return not self.failures

    def summary(self) -> str:
        return (f"통과 {len(self.passed)} / 실패 {len(self.failures)} / "
                f"미검증 {len(self.unverified)}")


def _com():
    try:
        import win32com.client as w32  # noqa: PLC0415
    except ImportError as exc:  # pragma: no cover
        raise ExcelWriteError(
            "pywin32가 설치되어 있지 않습니다. `pip install pywin32` 후 다시 실행하세요."
        ) from exc
    return w32


def _check_lock(path: str) -> None:
    folder, name = os.path.split(path)
    if os.path.exists(os.path.join(folder, "~$" + name)):
        raise ExcelWriteError(
            f"'{name}' 이(가) Excel에서 열려 있는 것으로 보입니다(~$ 잠금 파일). "
            "파일을 닫은 뒤 다시 실행하세요. 원본은 그대로 둡니다.")
    try:
        with open(path, "r+b"):
            pass
    except OSError as exc:
        raise ExcelWriteError(
            f"출력 파일에 쓸 수 없습니다({exc}). 열려 있거나 권한이 없습니다. "
            "원본은 그대로 둡니다.") from exc


# --------------------------------------------------------------------- 상태 수집
def _has_formula_flag(ws) -> tuple[object, str]:
    """UsedRange.HasFormula: False=수식 없음, True=전부 수식, None=섞임.

    SpecialCells와는 다른 경로라, '수식이 정말 없는 것'과 '조회에 실패한 것'을
    가르는 독립 근거로 쓴다. 이 속성 자체가 실패하면 판단 근거가 없다.
    """
    try:
        return ws.UsedRange.HasFormula, ""
    except Exception as exc:  # noqa: BLE001
        return "확인불가", f"UsedRange.HasFormula 조회 실패({exc})"


def _sheet_formulas(ws) -> tuple[dict[str, str], str]:
    """시트의 수식을 {주소: 수식}으로 모은다. 두 번째 값이 비어 있지 않으면 미검증 사유.

    SpecialCells는 수식이 하나도 없을 때도 COM 예외를 던지고, 실제 조회 실패 때도
    예외를 던진다. 오류 메시지나 오류 코드만으로는 둘을 가를 수 없으므로
    UsedRange.HasFormula라는 독립 경로로 교차 확인한다.
    """
    hf, hf_err = _has_formula_flag(ws)

    try:
        rng = ws.Cells.SpecialCells(_XL_CELLTYPE_FORMULAS)
    except Exception as exc:  # noqa: BLE001
        if hf is False:
            return {}, ""     # 독립 경로가 '수식 없음'을 확인해 줌 → 정상
        return {}, (f"수식 목록 조회 실패({exc}); UsedRange.HasFormula={hf!r}"
                    + (f"; {hf_err}" if hf_err else "")
                    + " — 수식이 없는 것인지 읽지 못한 것인지 구분하지 못함")

    try:
        count = int(rng.Count)
    except Exception as exc:  # noqa: BLE001
        return {}, f"수식 셀 수를 읽지 못함({exc}); UsedRange.HasFormula={hf!r}"
    if count > FORMULA_SCAN_LIMIT:
        return {}, f"수식 {count:,}개 — 검사 한도({FORMULA_SCAN_LIMIT:,}) 초과"

    out: dict[str, str] = {}
    try:
        for area in rng.Areas:
            for cell in area:
                out[str(cell.Address)] = str(cell.Formula)
    except Exception as exc:  # noqa: BLE001
        return out, f"수식 {count:,}개 중 {len(out):,}개만 수집 후 오류: {exc} — 부분 수집"
    if len(out) != count:
        return out, f"수식 셀 {count:,}개 중 {len(out):,}개만 수집됨 — 부분 수집"
    if hf is False and out:
        return out, (f"UsedRange.HasFormula=False인데 수식 {len(out)}개가 조회됨 "
                     "— 두 경로가 어긋나 신뢰할 수 없음")
    return out, ""


def _inspect(wb) -> dict[str, Any]:
    """통합문서 전체 상태를 내용 단위로 수집한다."""
    snap: dict[str, Any] = {
        "sheets": [], "formulas": {}, "formula_skips": {},
        "shapes": {}, "ole": {}, "vba": {}, "vba_status": "",
    }
    for ws in wb.Worksheets:
        name = ws.Name
        snap["sheets"].append(name)
        fx, skip = _sheet_formulas(ws)
        snap["formulas"][name] = fx
        if skip:
            snap["formula_skips"][name] = skip
        shapes = []
        try:
            for i in range(1, ws.Shapes.Count + 1):
                s = ws.Shapes(i)
                shapes.append((str(s.Name), int(s.Type)))
        except Exception as exc:  # noqa: BLE001
            shapes.append((f"<수집오류:{exc}>", -1))
        snap["shapes"][name] = shapes
        oles = []
        try:
            for i in range(1, ws.OLEObjects().Count + 1):
                o = ws.OLEObjects(i)
                try:
                    prog = str(o.progID)
                except Exception:  # noqa: BLE001
                    prog = "?"
                oles.append((str(o.Name), prog))
        except Exception as exc:  # noqa: BLE001
            oles.append((f"<수집오류:{exc}>", "?"))
        snap["ole"][name] = oles

    try:
        comps = {}
        for c in wb.VBProject.VBComponents:
            cm = c.CodeModule
            n = int(cm.CountOfLines)
            comps[str(c.Name)] = (int(c.Type), str(cm.Lines(1, n)) if n else "")
        snap["vba"] = comps
        snap["vba_status"] = "ok"
    except Exception as exc:  # noqa: BLE001
        snap["vba"] = {}
        snap["vba_status"] = f"VBA 프로젝트에 접근할 수 없음({exc})"
    return snap


def _last_row(ws, col: int, floor_row: int) -> int:
    return max(ws.Cells(ws.Rows.Count, col).End(_XL_UP).Row, floor_row)


def _read_prev_end(wb) -> int:
    try:
        ws = wb.Worksheets(SHEET_RUNLOG)
        for r in range(1, 80):
            if str(ws.Cells(r, 1).Value or "") == "자동_작성_끝행":
                return int(float(ws.Cells(r, 2).Value or 0))
    except Exception:  # noqa: BLE001
        return 0
    return 0


def _put_block(ws, top: int, left: int, data: Sequence[Sequence[Any]],
               text_cols: Sequence[int] = ()) -> None:
    if not data:
        return
    width = max(len(r) for r in data)
    rows = [list(r) + [None] * (width - len(r)) for r in data]
    rng = ws.Range(ws.Cells(top, left), ws.Cells(top + len(rows) - 1, left + width - 1))
    for c in text_cols:
        ws.Range(ws.Cells(top, left + c),
                 ws.Cells(top + len(rows) - 1, left + c)).NumberFormat = _XL_TEXT_FORMAT
    rng.Value = tuple(tuple(r) for r in rows)


def _fmt(v: Any) -> Any:
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    if v is None or isinstance(v, (int, float)):
        return v
    s = str(v)
    return s if len(s) <= 32767 else s[:32760] + "…(생략)"


def _addr_in_list_area(addr: str, first_row: int, last_row: int) -> bool:
    """'$A$59' 형태 주소가 조회처 출력 영역(A~L열, first~last행)에 드는지."""
    parts = [p for p in addr.replace("$", " ").split() if p]
    if len(parts) != 2 or not parts[1].isdigit():
        return False
    col, row = parts[0].upper(), int(parts[1])
    n = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            return False
        n = n * 26 + (ord(ch) - 64)
    return LIST_FIRST_COL <= n <= LIST_LAST_COL and first_row <= row <= last_row


def _excel_session(w32, rep: WriteReport, sheet_name: str, names, detail_rows,
                   detail_header, evidence_rows, evidence_header, runlog_rows,
                   online_flags, skipped_rows, skipped_header, detail_text_cols,
                   evidence_text_cols, skipped_text_cols, log, start_row: int) -> None:
    """임시본을 열어 목록을 쓰고 보조 시트를 만든 뒤, 다시 열어 검증한다.

    Excel 연결이 끊긴 경우에는 transient 표시가 붙은 ExcelWriteError를 낸다.
    """
    app = None
    pid = None
    try:
        app = w32.DispatchEx("Excel.Application")
        pid = _excel_pid(app)
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False          # 매크로 자동실행(Workbook_Open 등) 차단
        app.AskToUpdateLinks = False
        app.ScreenUpdating = False
        try:
            app.AutomationSecurity = 3     # msoAutomationSecurityForceDisable
        except Exception:  # noqa: BLE001
            log("경고: AutomationSecurity를 설정하지 못했습니다.")

        wb = app.Workbooks.Open(rep.temp, UpdateLinks=0, ReadOnly=False)
        try:
            ws = wb.Worksheets(sheet_name)
        except Exception as exc:  # noqa: BLE001
            raise ExcelWriteError(f"시트 '{sheet_name}'을(를) 찾지 못했습니다.") from exc

        log("기존 상태 수집 중(수식 주소·내용, 도형, ActiveX, VBA 코드)…")
        rep.baseline = _inspect(wb)
        nfx = sum(len(v) for v in rep.baseline["formulas"].values())
        log(f"기준: 시트 {len(rep.baseline['sheets'])}개, 수식 {nfx}개, "
            f"VBA {len(rep.baseline['vba'])}개 ({rep.baseline['vba_status'] or 'ok'})")

        prev_end = _read_prev_end(wb)
        used_end = start_row - 1
        for col in range(LIST_FIRST_COL, LIST_LAST_COL + 1):
            used_end = max(used_end, _last_row(ws, col, start_row - 1))
        clear_end = max(used_end, prev_end, start_row - 1)
        rep.cleared_to = clear_end
        if clear_end >= start_row:
            ws.Range(ws.Cells(start_row, LIST_FIRST_COL),
                     ws.Cells(clear_end, LIST_LAST_COL)).ClearContents()
            log(f"기존 목록 정리: A{start_row}:L{clear_end} (백업에 원본 보존)")

        n = len(names)
        rep.written_rows = n
        rep.end_row = start_row + n - 1 if n else start_row - 1

        if n:
            src = ws.Range(ws.Cells(start_row, LIST_FIRST_COL), ws.Cells(start_row, LIST_LAST_COL))
            dst = ws.Range(ws.Cells(start_row, LIST_FIRST_COL), ws.Cells(rep.end_row, LIST_LAST_COL))
            # 첫 행의 서식을 아래로 편다. **클립보드를 쓰지 않는다.**
            # Copy() + PasteSpecial은 Windows 클립보드를 거치는데, 다른 프로그램이
            # 클립보드를 잡고 있으면 '데이터를 붙여 넣을 수 없습니다'로 실패한다(실측).
            if rep.end_row > start_row:
                try:
                    dst.FillDown()
                except Exception:  # noqa: BLE001
                    # FillDown이 막히면 대상을 직접 지정해 복사한다. 이것도 클립보드를 쓰지 않는다.
                    src.Copy(dst)
                # 서식만 필요하다. 첫 행에 값이 있었다면 아래로 번진 값을 지운다.
                ws.Range(ws.Cells(start_row + 1, LIST_FIRST_COL),
                         ws.Cells(rep.end_row, LIST_LAST_COL)).ClearContents()
            hcol = ws.Range(ws.Cells(start_row, FLAG_COL), ws.Cells(rep.end_row, FLAG_COL))
            try:
                hcol.Validation.Delete()
            except Exception:  # noqa: BLE001
                pass
            hcol.Validation.Add(_XL_VALIDATE_LIST, _XL_ALERT_STOP, _XL_BETWEEN, "YES,NO")

            namecol = ws.Range(ws.Cells(start_row, NAME_COL), ws.Cells(rep.end_row, NAME_COL))
            namecol.NumberFormat = _XL_TEXT_FORMAT
            namecol.Value = tuple((str(x),) for x in names)
            ws.Range(ws.Cells(start_row, FLAG_COL), ws.Cells(rep.end_row, FLAG_COL)).Value = \
                tuple(("YES",) for _ in names)

            # C열: 금융결제원 온라인 발급 가능 기관 표시. 목록에 없으면 비워 둔다.
            flags = list(online_flags)[:n] + [""] * max(0, n - len(online_flags))
            n_online = sum(1 for f in flags if f)
            if n_online:
                ocol = ws.Range(ws.Cells(start_row, ONLINE_COL),
                                ws.Cells(rep.end_row, ONLINE_COL))
                ocol.NumberFormat = _XL_TEXT_FORMAT
                ocol.Value = tuple((f or None,) for f in flags)
                log(f"조회처 {n}건 작성: A{start_row}:A{rep.end_row}, H열 YES, "
                    f"C열 온라인 {n_online}건")
            else:
                log(f"조회처 {n}건 작성: A{start_row}:A{rep.end_row}, H열 YES")
        else:
            log("검색 결과 0건 — 신규 행을 추가하지 않았습니다.")

        _rebuild_sheet(wb, SHEET_DETAIL, detail_header, detail_rows, text_cols=detail_text_cols)
        _rebuild_sheet(wb, SHEET_SKIPPED, skipped_header, skipped_rows,
                       text_cols=skipped_text_cols)
        _rebuild_sheet(wb, SHEET_EVIDENCE, evidence_header, evidence_rows,
                       text_cols=evidence_text_cols)
        _rebuild_sheet(wb, SHEET_RUNLOG, ("항목", "값"),
                       list(runlog_rows) + [("자동_작성_시작행", start_row),
                                            ("자동_작성_끝행", rep.end_row),
                                            ("정리한_범위", f"A{start_row}:L{clear_end}"),
                                            ("백업파일", rep.backup)],
                       text_cols=(1,))
        _move_to_end(wb)
        ws.Activate()
        ws.Range("A1").Select()
        wb.Save()
        wb.Close(SaveChanges=False)
        log("임시본 저장 완료. 다시 열어 검증합니다.")

        wb2 = app.Workbooks.Open(rep.temp, UpdateLinks=0, ReadOnly=True)
        ws2 = wb2.Worksheets(sheet_name)
        rep.after = _inspect(wb2)
        rep.checks = _verify(ws2, names, rep, start_row, sheet_name, clear_end,
                             list(online_flags) + [""] * max(0, n - len(online_flags)))
        wb2.Close(SaveChanges=False)

    except ExcelWriteError:
        _quit(app, pid)
        raise
    except Exception as exc:  # noqa: BLE001
        _quit(app, pid)
        raise ExcelWriteError(f"Excel 저장 중 오류: {exc}",
                              transient=_is_disconnected(exc)) from exc
    finally:
        _quit(app, pid)


# --------------------------------------------------------------------- 쓰기
def write_control_sheet(
    target: str,
    sheet_name: str,
    names: Sequence[str],
    detail_rows: Sequence[Sequence[Any]],
    detail_header: Sequence[str],
    evidence_rows: Sequence[Sequence[Any]],
    evidence_header: Sequence[str],
    runlog_rows: Sequence[Sequence[Any]],
    backup_dir: str,
    online_flags: Sequence[str] = (),
    skipped_rows: Sequence[Sequence[Any]] = (),
    skipped_header: Sequence[str] = ("조회처명(A열에 넣지 않음)",),
    detail_text_cols: Sequence[int] = (),
    evidence_text_cols: Sequence[int] = (),
    skipped_text_cols: Sequence[int] = (),
    log: Callable[[str], None] = lambda _m: None,
    start_row: int = START_ROW,
) -> WriteReport:
    # Excel COM은 상대경로를 Excel 자신의 작업 폴더에서 찾는다. 반드시 절대경로로 넘긴다.
    target = os.path.abspath(target)
    backup_dir = os.path.abspath(backup_dir)
    if not os.path.exists(target):
        raise ExcelWriteError(f"출력 파일이 없습니다: {target}")
    _check_lock(target)

    _sweep_stale_temps(os.path.dirname(target), log)

    rep = WriteReport(target=target, start_row=start_row)
    os.makedirs(backup_dir, exist_ok=True)
    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(target))[0]

    rep.backup = os.path.join(backup_dir, f"{base}_backup_{ts}.xlsm")
    shutil.copy2(target, rep.backup)
    log(f"백업 생성: {rep.backup}")

    w32 = _com()
    for attempt in range(1, RETRY_LIMIT + 1):
        rep.temp = os.path.join(os.path.dirname(target),
                                f"~자동화임시_{base}_{ts}_{attempt}.xlsm"
                                if attempt > 1 else f"~자동화임시_{base}_{ts}.xlsm")
        shutil.copy2(target, rep.temp)
        try:
            _excel_session(w32, rep, sheet_name, names, detail_rows, detail_header,
                           evidence_rows, evidence_header, runlog_rows, online_flags,
                           skipped_rows, skipped_header, detail_text_cols,
                           evidence_text_cols, skipped_text_cols, log, start_row)
            break
        except ExcelWriteError as exc:
            _cleanup(rep.temp)
            if exc.transient and attempt < RETRY_LIMIT:
                log(f"Excel 연결이 끊겼습니다 — {RETRY_WAIT}초 쉬고 다시 시도합니다 "
                    f"({attempt}/{RETRY_LIMIT - 1})")
                time.sleep(RETRY_WAIT)
                continue
            if exc.transient:
                raise ExcelWriteError(
                    f"{RETRY_LIMIT}번 시도했지만 {DISCONNECT_GUIDE}", transient=True) from exc
            raise

    if not rep.ok:
        fails = "; ".join(f"{c[0]}: {c[2]}" for c in rep.failures)
        raise ExcelWriteError(
            f"검증 실패로 대상 파일을 바꾸지 않았습니다 → {fails} (임시본: {rep.temp})")

    try:
        _check_lock(target)
        os.replace(rep.temp, target)
        rep.replaced = True
        log(f"검증 결과 {rep.summary()} → 대상 파일 교체 완료: {target}")
    except OSError as exc:
        raise ExcelWriteError(
            f"검증은 통과했으나 대상 파일 교체에 실패했습니다({exc}). "
            f"원본은 그대로이며 결과는 임시본에 있습니다: {rep.temp}") from exc
    return rep


# --------------------------------------------------------------------- 검증
def _verify(ws2, names: Sequence[str], rep: WriteReport, start_row: int,
            sheet_name: str, clear_end: int,
            online_flags: Sequence[str] = ()) -> list[tuple[str, str, str]]:
    checks: list[tuple[str, str, str]] = []
    n = len(names)

    if n:
        got = [str(ws2.Cells(start_row + i, NAME_COL).Value or "") for i in range(n)]
        want = [str(x) for x in names]
        checks.append((f"A{start_row}부터 {n}건 연속 작성",
                       PASS if got == want else FAIL,
                       "일치" if got == want else f"불일치 (예: {got[:3]})"))
        flags = [str(ws2.Cells(start_row + i, FLAG_COL).Value or "") for i in range(n)]
        okh = all(f == "YES" for f in flags)
        checks.append(("이름 있는 행의 H열 YES", PASS if okh else FAIL,
                       "전부 YES" if okh else f"YES 아님 {n - flags.count('YES')}건"))
        blank_cols = [c for c in range(2, 13) if c not in (FLAG_COL, ONLINE_COL)]
        bad = [f"{chr(64 + c)}{start_row + i}"
               for i in range(n) for c in blank_cols
               if ws2.Cells(start_row + i, c).Value not in (None, "")]
        checks.append(("B · D:G · I:L 공란", PASS if not bad else FAIL,
                       "모두 공란" if not bad else f"{len(bad)}개 값 존재 ({bad[:3]})"))
        want_c = [(online_flags[i] if i < len(online_flags) else "") for i in range(n)]
        got_c = [str(ws2.Cells(start_row + i, ONLINE_COL).Value or "") for i in range(n)]
        okc = got_c == [w or "" for w in want_c]
        checks.append((f"C열 온라인 표시 {sum(1 for w in want_c if w)}건 / 나머지 공란",
                       PASS if okc else FAIL,
                       "일치" if okc else
                       f"불일치: {[(i + start_row, g, w) for i, (g, w) in enumerate(zip(got_c, want_c)) if g != (w or '')][:3]}"))
        nxt = ws2.Cells(rep.end_row + 1, NAME_COL).Value
        checks.append(("이전 결과 잔여행 없음", PASS if nxt in (None, "") else FAIL,
                       "없음" if nxt in (None, "") else f"A{rep.end_row + 1}={nxt!r}"))
        dup = len(set(want)) == n
        checks.append(("A열 중복 없음", PASS if dup else FAIL,
                       "중복 없음" if dup else f"중복 {n - len(set(want))}건"))
    else:
        v = ws2.Cells(start_row, NAME_COL).Value
        checks.append(("0건 — 신규 행 없음", PASS if v in (None, "") else FAIL,
                       f"A{start_row} 공란" if v in (None, "") else f"{v!r}"))

    b, a = rep.baseline, rep.after

    missing = [s for s in b["sheets"] if s not in a["sheets"]]
    added = [s for s in a["sheets"] if s not in b["sheets"]]
    unexpected = [s for s in added if s not in AUTO_SHEETS]
    checks.append(("기존 시트 보존",
                   PASS if not missing and not unexpected else FAIL,
                   f"{len(b['sheets'])}개 → {len(a['sheets'])}개, 추가 {added or '없음'}"
                   + (f", 사라짐 {missing}" if missing else "")))

    # 수식: 시트명·주소·내용 대조
    allowed_last = max(clear_end, rep.end_row)
    diffs: list[str] = []
    compared = 0
    skips = dict(b.get("formula_skips", {}))
    skips.update(a.get("formula_skips", {}))
    for sname in b["sheets"]:
        if sname in AUTO_SHEETS or sname in skips:
            continue
        before = b["formulas"].get(sname, {})
        after = a["formulas"].get(sname, {})
        in_area = (lambda ad: sname == sheet_name
                   and _addr_in_list_area(ad, start_row, allowed_last))
        for addr, f in before.items():
            if in_area(addr):
                continue
            compared += 1
            if addr not in after:
                diffs.append(f"{sname}!{addr} 사라짐")
            elif after[addr] != f:
                diffs.append(f"{sname}!{addr} 변경({f!r}→{after[addr]!r})")
        for addr in after:
            if addr not in before and not in_area(addr):
                diffs.append(f"{sname}!{addr} 신규 수식")
    if diffs:
        status, note = FAIL, f"차이 {len(diffs)}건: {diffs[:5]}"
    elif compared == 0 and skips:
        status, note = UNVERIFIED, (
            f"대조한 수식이 0개이고 시트 {len(skips)}개를 읽지 못했습니다 — 보존을 확인하지 못했습니다")
    else:
        status, note = PASS, (
            f"{compared}개 대조, 차이 없음. 허용 변경범위: "
            f"'{sheet_name}'!A{start_row}:L{allowed_last} 및 자동화 보조 시트"
            + (f". 단, 시트 {len(skips)}개는 대조하지 못했습니다(아래 미검증 항목 참조)"
               if skips else ""))
    checks.append(("기존 수식 보존(시트·주소·내용 대조)", status, note))
    if skips:
        checks.append(("수식 대조를 못 한 시트", UNVERIFIED,
                       "; ".join(f"{s}: {r}" for s, r in skips.items())))

    # VBA: 이름·종류·코드 본문
    if b["vba_status"] != "ok" or a["vba_status"] != "ok":
        checks.append(("VBA 구성요소 보존(이름·종류·코드 본문)", UNVERIFIED,
                       f"기준 {b['vba_status'] or 'ok'} / 저장본 {a['vba_status'] or 'ok'} — "
                       "코드 비교를 하지 못했습니다(XLSM 형식 저장 자체는 유지됨)"))
    else:
        vb, va = b["vba"], a["vba"]
        vdiffs: list[str] = []
        for name, (typ, code) in vb.items():
            if name not in va:
                vdiffs.append(f"{name} 사라짐")
            elif va[name][0] != typ:
                vdiffs.append(f"{name} 종류 변경({typ}→{va[name][0]})")
            elif va[name][1] != code:
                vdiffs.append(f"{name} 코드 변경({len(code)}자→{len(va[name][1])}자)")
        new_comps = [k for k in va if k not in vb]
        new_code = [k for k in new_comps if va[k][1].strip()]
        checks.append(("VBA 구성요소 보존(이름·종류·코드 본문)",
                       PASS if not vdiffs else FAIL,
                       (f"기존 {len(vb)}개의 이름·종류·코드 본문이 모두 동일. "
                        f"추가된 구성요소 {len(new_comps)}개({', '.join(new_comps) or '없음'})는 "
                        f"새 보조 시트의 빈 코드모듈, 그중 코드가 있는 것 {len(new_code)}개")
                       if not vdiffs else f"차이 {len(vdiffs)}건: {vdiffs[:5]}"))

    # 도형 / ActiveX: 이름·종류 대조
    sdiffs = [f"{s}: {b['shapes'].get(s)} → {a['shapes'].get(s)}"
              for s in b["sheets"]
              if s not in AUTO_SHEETS and b["shapes"].get(s, []) != a["shapes"].get(s, [])]
    total_shapes = sum(len(v) for v in b["shapes"].values())
    checks.append(("도형 이름·종류 보존", PASS if not sdiffs else FAIL,
                   f"{total_shapes}개 대조, 동일" if not sdiffs else f"차이: {sdiffs[:3]}"))

    odiffs = [f"{s}: {b['ole'].get(s)} → {a['ole'].get(s)}"
              for s in b["sheets"]
              if s not in AUTO_SHEETS and b["ole"].get(s, []) != a["ole"].get(s, [])]
    total_ole = sum(len(v) for v in b["ole"].values())
    checks.append(("ActiveX 이름·progID 보존", PASS if not odiffs else FAIL,
                   f"{total_ole}개 대조, 동일" if not odiffs else f"차이: {odiffs[:3]}"))

    checks.append(("도형·ActiveX 내부 속성·바이너리 내용 보존", UNVERIFIED,
                   "이름·종류·progID만 대조했습니다. 컨트롤 속성값·이미지 데이터·"
                   "이벤트 바인딩은 비교하지 않았습니다"))
    checks.append(("매크로 동작", UNVERIFIED,
                   "기존 매크로를 실행하지 않았습니다(조회서 출력·인쇄 금지). 동작 여부 미확인"))
    checks.append(("저장본 재열기", PASS, "Excel COM으로 다시 열어 위 항목을 읽었습니다"))
    return checks


def _move_to_end(wb, names: Sequence[str] = AUTO_SHEETS) -> None:
    """자동화 보조 시트를 통합문서 맨 뒤로, 정해진 순서대로 보낸다.

    Worksheets.Add(After=…)만으로는 통합문서에 따라 중간에 끼어든다.
    조서를 열었을 때 원래 시트가 먼저 보여야 하므로 마지막에 한 번 더 옮긴다.
    """
    for name in names:
        try:
            # 이름 있는 인자(After=)는 COM 동적 호출에서 무시될 수 있다. 위치 인자로 넘긴다.
            wb.Worksheets(name).Move(None, wb.Sheets(wb.Sheets.Count))
        except Exception:  # noqa: BLE001
            pass


def _rebuild_sheet(wb, name: str, header: Sequence[str], rows: Sequence[Sequence[Any]],
                   text_cols: Sequence[int] = ()) -> None:
    try:
        wb.Worksheets(name).Delete()
    except Exception:  # noqa: BLE001
        pass
    ws = wb.Worksheets.Add(None, wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    _put_block(ws, 1, 1, [list(header)])
    ws.Range(ws.Cells(1, 1), ws.Cells(1, max(1, len(header)))).Font.Bold = True
    body = [[_fmt(v) for v in r] for r in rows]
    if body:
        chunk = 2000
        for i in range(0, len(body), chunk):
            _put_block(ws, 2 + i, 1, body[i:i + chunk], text_cols=text_cols)
        try:
            ws.Rows(1).AutoFilter()
        except Exception:  # noqa: BLE001
            pass
    try:
        ws.Columns.AutoFit()
    except Exception:  # noqa: BLE001
        pass


def _excel_pid(app) -> int | None:
    """이 Excel 인스턴스의 프로세스 번호. 우리가 띄운 것만 골라 끝내기 위해서다."""
    try:
        import win32process  # noqa: PLC0415
        _tid, pid = win32process.GetWindowThreadProcessId(int(app.Hwnd))
        return int(pid) or None
    except Exception:  # noqa: BLE001
        return None


def _force_kill(pid: int | None) -> None:
    """Quit이 듣지 않는 Excel만 끝낸다. 사용자가 열어 둔 Excel은 건드리지 않는다."""
    if not pid:
        return
    try:
        import subprocess  # noqa: PLC0415
        subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                       capture_output=True, check=False)
    except OSError:
        pass


def _quit(app, pid: int | None = None) -> None:
    if app is None:
        _force_kill(pid)
        return
    try:
        for i in range(app.Workbooks.Count, 0, -1):
            app.Workbooks(i).Close(SaveChanges=False)
    except Exception:  # noqa: BLE001
        pass
    quit_ok = True
    try:
        app.Quit()
    except Exception:  # noqa: BLE001
        quit_ok = False
    if not quit_ok:
        # 연결이 끊긴 Excel은 Quit이 듣지 않는다. 남으면 임시본을 잡고 다음 실행을 막는다.
        _force_kill(pid)


def _sweep_stale_temps(folder: str, log: Callable[[str], None]) -> None:
    """앞선 실행이 남긴 임시본을 지운다. 지워지지 않으면(열려 있으면) 그냥 둔다."""
    try:
        names = os.listdir(folder)
    except OSError:
        return
    gone = 0
    for name in names:
        if name.startswith("~자동화임시_") or name.startswith("~$~자동화임시_"):
            try:
                os.remove(os.path.join(folder, name))
                gone += 1
            except OSError:
                pass
    if gone:
        log(f"이전 실행이 남긴 임시본 {gone}개를 정리했습니다")


def _cleanup(path: str) -> None:
    for _ in range(3):
        try:
            if path and os.path.exists(path):
                os.remove(path)
            return
        except OSError:
            time.sleep(0.4)
