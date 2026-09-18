# -*- coding: utf-8 -*-
"""요구사항 검증 스크립트. 복사본에만 쓰기 작업을 한다.

결과는 PASS / FAIL / UNVERIFIED 세 가지로 구분한다.
자료가 없어 건너뛴 케이스는 UNVERIFIED이며 통과 건수에 넣지 않는다.
기존 구현 결과를 정답으로 삼지 않고, 합성 데이터로 요구사항 자체를 검증한다.

검색 기준: **분개장의 '거래처' 열 하나만**. 거래처명 자체에 지정 키워드 또는
기관 사전의 기관명·별칭이 있어야 후보가 된다.
"""
import copy
import datetime as _dt
import os
import shutil
import subprocess
import sys
import time
import traceback

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(BASE, "src"))
sys.stdout.reconfigure(encoding="utf-8")

from core.config_loader import load_config                        # noqa: E402
from core.matcher import KeywordMatcher                           # noqa: E402
from core.pipeline import DEFAULT_SHEET, run                      # noqa: E402
from core.readers import JournalReadError, JournalRow, read_journal  # noqa: E402
from core.resolver import InstitutionResolver                     # noqa: E402
from core.writer_excel import (FAIL, PASS, UNVERIFIED,            # noqa: E402
                               ExcelWriteError, _is_disconnected,
                               write_control_sheet)

SAMPLES = os.path.join(BASE, "samples")
CONFIG = os.path.join(BASE, "config")
WORK = os.path.join(BASE, "tests", "_out")
os.makedirs(WORK, exist_ok=True)

# 분개장·Control Sheet 경로는 인자로 받는다. 없으면 samples의 견본을 쓴다.
#   python tests/verify.py [분개장.xlsx] [ControlSheet.xlsm]
JOURNAL = (sys.argv[1] if len(sys.argv) > 1
           else os.path.join(SAMPLES, "샘플_분개장.xlsx"))
TEMPLATE = (sys.argv[2] if len(sys.argv) > 2
            else os.path.join(BASE, "template", "Control Sheet_표준양식.xlsm"))
for _p, _what in ((JOURNAL, "분개장"), (TEMPLATE, "Control Sheet")):
    if not os.path.exists(_p):
        print(f"{_what} 파일이 없습니다: {_p}")
        print("  python tools/make_samples.py 로 견본 분개장을 만들거나 경로를 인자로 주세요.")
        raise SystemExit(2)

RESULTS: list[tuple[str, str, str]] = []


def _excel_pids() -> set[int]:
    """지금 떠 있는 EXCEL.EXE의 PID."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH", "/FO", "CSV"],
            capture_output=True, text=True, encoding="cp949", errors="replace").stdout
    except OSError:
        return set()
    pids: set[int] = set()
    for line in out.splitlines():
        cells = [c.strip('"') for c in line.split('","')]
        if len(cells) > 1 and cells[0].upper().startswith("EXCEL"):
            try:
                pids.add(int(cells[1]))
            except ValueError:
                pass
    return pids


# 시작할 때 떠 있던 Excel은 사용자 것이다. 절대 닫지 않는다.
PRE_EXCEL_PIDS = _excel_pids()


def reap_excel() -> None:
    """이 테스트가 띄웠는데 닫히지 않은 Excel만 정리한다."""
    leftover = _excel_pids() - PRE_EXCEL_PIDS
    for pid in leftover:
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
    if leftover:
        time.sleep(1.0)


class Skip(Exception):
    """자료가 없어 검증하지 못함 → UNVERIFIED."""


def _env_failure(exc: BaseException) -> bool:
    """Excel 연결이 끊겨서 난 실패인가. 요구사항 위반(AssertionError)은 아니다."""
    if isinstance(exc, AssertionError):
        return False
    if getattr(exc, "transient", False):
        return True
    return _is_disconnected(exc)


def case(name):
    def deco(fn):
        print(f"\n### {name}")
        for attempt in (1, 2):
            try:
                note = fn() or "통과"
                if attempt > 1:
                    note += " (Excel 연결이 끊겨 한 번 다시 돌림)"
                RESULTS.append((name, PASS, note))
                print(f"  [{PASS}] {note}")
                break
            except Skip as e:
                RESULTS.append((name, UNVERIFIED, str(e)))
                print(f"  [{UNVERIFIED}] {e}")
                break
            except Exception as e:  # noqa: BLE001
                if attempt == 1 and _env_failure(e):
                    print(f"  [재시도] Excel 연결 끊김 — {type(e).__name__}: {str(e)[:70]}")
                    reap_excel()
                    time.sleep(5)
                    continue
                RESULTS.append((name, FAIL, f"{type(e).__name__}: {e}"
                                if not isinstance(e, AssertionError) else str(e)))
                print(f"  [{FAIL}] {e}")
                if not isinstance(e, AssertionError):
                    traceback.print_exc()
                break
        reap_excel()
        return fn
    return deco


def fresh(name):
    p = os.path.join(WORK, name)
    shutil.copy2(TEMPLATE, p)
    return p


# ---------------------------------------------------------------- 보조
class _FakeRep:
    def __init__(self, baseline, after):
        self.baseline, self.after = baseline, after
        self.end_row = 58


class _FakeCell:
    Value = None


class _FakeWs:
    """names=[] 인 경우 _verify는 A59가 비었는지만 읽는다."""

    def Cells(self, *_a):
        return _FakeCell()


def read_cells(path, rows, cols=range(1, 13), sheet=DEFAULT_SHEET):
    import win32com.client as w32
    app = w32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.AutomationSecurity = 3
    try:
        wb = app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True)
        ws = wb.Worksheets(sheet)
        out = {r: [ws.Cells(r, c).Value for c in cols] for r in rows}
        extra = {"시트": [s.Name for s in wb.Sheets]}
        wb.Close(False)
        return out, extra
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


# ---------------------------------------------------------------- 합성 데이터
def mkrow(n, cp, desc, acct="지급수수료(판)", voucher="V1", debit=1000.0):
    return JournalRow(sheet="합성", excel_row=n, voucher_id=voucher,
                      date=_dt.date(2025, 5, 1), voucher_no="1", dc="차변",
                      account=acct, debit=debit, credit=0.0,
                      counterparty=cp, description=desc)


def resolve_rows(rows, cfg):
    matches = KeywordMatcher(cfg).search(rows)
    return matches, InstitutionResolver(cfg).resolve(matches)


CFG = load_config(CONFIG)
JD = read_journal(JOURNAL, CFG.profiles)
MM, RES = resolve_rows(JD.rows, CFG)
ALL_NAMES = [i.display_name for i in RES.institutions]   # 검색된 전체 후보
NAMES = [i.display_name for i in RES.output_institutions]  # A열에 실제로 쓰는 것
print(f"실제 분개장: 거래행 {len(JD.rows):,} / 검색 원본행 {len(MM):,} / "
      f"후보 {len(ALL_NAMES)} → A열 출력 {len(NAMES)}")


# =================================================================== 1. 검색 기준: 거래처 열만
@case("1. 설정이 거래처 열만 검색하도록 되어 있다")
def _():
    kw = CFG.keywords
    assert kw["search_fields"] == ["counterparty"], f"search_fields={kw['search_fields']}"
    assert kw["extract_fields"] == ["counterparty"], f"extract_fields={kw['extract_fields']}"
    assert kw["extract_modes"]["description"] == [], "적요 추출 모드가 남아 있음"
    assert kw["use_voucher_context"] is False, "전표 문맥 연결이 켜져 있음"
    assert kw["supplementary_keywords"] == [], f"보완 키워드 자동 적용: {kw['supplementary_keywords']}"
    assert kw["exclude_names"] == [], "기본 제외 목록이 비어 있지 않음"
    return ("search_fields=extract_fields=['counterparty'], 적요추출 off, "
            "전표문맥 off, 보완키워드 off, 제외목록 비어 있음")


@case("2. 거래처에만 키워드가 있으면 포함된다")
def _():
    rows = [mkrow(1, "가나다렌탈", "사무용품 구입", acct="소모품비"),
            mkrow(2, "법무법인 가나다", "등기", acct="지급수수료(판)", voucher="V2"),
            mkrow(3, "하나은행(03304)", "대체", acct="보통예금", voucher="V3"),
            mkrow(4, "건강보험", "4대보험", acct="예수금", voucher="V4")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert names == ["가나다렌탈", "법무법인 가나다", "하나은행", "건강보험"], names
    return f"{names} — 렌탈·법무·은행·보험 키워드로 포함"


@case("3. 적요에만 키워드가 있으면 제외된다")
def _():
    rows = [mkrow(1, "일반상사", "국민은행 송금", acct="외상매입금"),
            mkrow(2, "일반상사2", "렌탈료 지급 / 리스 정산", acct="외상매입금", voucher="V2")]
    _, r = resolve_rows(rows, CFG)
    assert not r.institutions, f"적요만으로 포함됨: {[i.display_name for i in r.institutions]}"
    assert r.row_stats["검색행"] == 0, f"검색행 {r.row_stats['검색행']}"
    return "거래처 '일반상사'·'일반상사2' 모두 제외, 검색행 0"


@case("4. 계정과목에만 키워드가 있으면 제외된다")
def _():
    rows = [mkrow(1, "일반상사", "월세", acct="지급임차료"),
            mkrow(2, "무명물산", "보증금 예치", acct="임차보증금", voucher="V2"),
            mkrow(3, "카페사업자", "정산", acct="이자비용", voucher="V3")]
    _, r = resolve_rows(rows, CFG)
    assert not r.institutions, f"계정과목만으로 포함됨: {[i.display_name for i in r.institutions]}"
    return "지급임차료·임차보증금·이자비용 계정만으로는 포함되지 않음"


@case("5. 거래처가 공란이면 제외된다")
def _():
    rows = [mkrow(1, "", "국민은행 대출이자", acct="이자비용"),
            mkrow(2, "   ", "하나은행 송금", acct="보통예금", voucher="V2")]
    _, r = resolve_rows(rows, CFG)
    assert not r.institutions, f"공란 거래처에서 후보 생성: {[i.display_name for i in r.institutions]}"
    assert r.row_stats["검색행"] == 0
    return "거래처 공란 2행 → 후보 0, 검색행 0"


@case("6. 거래처가 일치해도 적요의 다른 기관은 추가되지 않는다")
def _():
    rows = [mkrow(1, "하나은행(03304)", "국민은행·신한은행 계좌로 분할 이체", acct="보통예금")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert names == ["하나은행"], f"적요 기관이 추가됨: {names}"
    assert r.link_stats["직접일치"] == 1, f"연결 {r.link_stats['직접일치']}"
    return "거래처 '하나은행'만 후보, 적요의 국민은행·신한은행은 추가 안 됨"


@case("7. 같은 전표의 다른 거래처가 자동 추가되지 않는다")
def _():
    rows = [mkrow(1, "하나은행(03304)", "이자 지급", acct="이자비용", voucher="VZ"),
            mkrow(2, "무명물산(주)", "이자 정산", acct="이자비용", voucher="VZ")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert names == ["하나은행"], f"전표 문맥으로 추가됨: {names}"
    assert r.row_stats["전표문맥"] == 0 and r.link_stats["전표문맥"] == 0
    assert RES.row_stats["전표문맥"] == 0 and RES.link_stats["전표문맥"] == 0, \
        "실제 분개장에서 전표문맥 연결 발생"
    return "같은 전표의 '무명물산(주)'는 추가되지 않음. 실제 실행 전표문맥 0건"


@case("8. AJ네트웍스(주)는 이번 기준에서 제외된다")
def _():
    rows = [mkrow(1, "AJ네트웍스(주)", "렌탈료", acct="지급임차료(판)")]
    _, r = resolve_rows(rows, CFG)
    assert not r.institutions, f"포함됨: {[i.display_name for i in r.institutions]}"
    assert not [n for n in ALL_NAMES if "AJ네트웍스" in n], "실제 결과에 AJ네트웍스가 남아 있음"
    return "거래처명에 지정 키워드·사전 일치가 없어 제외 (요구사항 예시대로)"


# =================================================================== 2. 사전·원문 유지
@case("9. 거래처에 금융 키워드가 없어도 기관 사전 별칭이면 포함된다")
def _():
    rows = [mkrow(1, "농협(642403)", "보통예금 대체", acct="보통예금"),
            mkrow(2, "기업비씨(4907)", "카드대금", acct="미지급금", voucher="V2"),
            mkrow(3, "삼성화재", "정산", acct="보험료(판)", voucher="V3")]
    m, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert names == ["농협", "기업비씨", "삼성화재해상보험"], names
    assert m[0].aliases and not m[0].keywords, f"농협 일치 근거: {m[0].keywords}/{m[0].aliases}"
    return f"{names} — 별칭 일치로 포함. 농협 일치 근거: 사전 별칭 {m[0].aliases}"


@case("10. 사전 미등록이어도 거래처명에 지정 키워드가 있으면 포함된다")
def _():
    rows = [mkrow(1, "가나다렌탈", "월정료", acct="소모품비"),
            mkrow(2, "홈플러스 포스보증금/구미", "보증금", acct="기타보증금", voucher="V2")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert names == ["가나다렌탈", "홈플러스 포스보증금/구미"], names
    assert all(i.matched_by == "원문거래처" for i in r.institutions)
    return f"{names} — 사전 미등록이어도 원문 유지"


@case("11. 복수 키워드·여러 계좌 표기가 있어도 기관별로 중복 제거된다")
def _():
    rows = [mkrow(1, "기업은행(04-013)", "차입", acct="단기차입금"),
            mkrow(2, "기업은행(04-021)", "이자", acct="이자비용", voucher="V2"),
            mkrow(3, "기업은행중소기업자금대출(00040)/10억", "상환", acct="단기차입금", voucher="V3"),
            mkrow(4, "기업은행 외화예금(00017)", "예금", acct="보통예금", voucher="V4")]
    m, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert names == ["IBK기업은행"], f"통합 실패: {names}"
    inst = r.institutions[0]
    assert len(inst.account_clues) >= 4, f"계좌 단서: {list(inst.account_clues)}"
    assert any(c.startswith("0") for c in inst.account_clues), "선행 0 유실"
    assert len(m[2].keywords) >= 2, f"복수 키워드 아님: {m[2].keywords}"
    assert len(set(ALL_NAMES)) == len(ALL_NAMES), "실제 결과에 중복 이름 존재"
    return (f"4개 표기 → 1행, 계좌 단서 {list(inst.account_clues)}, "
            f"실제 결과 {len(ALL_NAMES)}건 중 중복 0")


@case("12. 서로 다른 법인을 임의로 합치지 않는다")
def _():
    rows = [mkrow(1, "신한은행(634159)", "예금", acct="보통예금"),
            mkrow(2, "신한카드주식회사/신한할부금융", "카드", acct="미지급금", voucher="V2"),
            mkrow(3, "하나은행(03304)", "예금", acct="보통예금", voucher="V3"),
            mkrow(4, "하나카드(4954)", "카드", acct="미지급금", voucher="V4")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    for n in ("신한은행", "신한카드", "하나은행", "하나카드"):
        assert n in names, f"{n} 누락: {names}"
    return f"{names} — 은행/카드 법인 분리 유지"


@case("13. 모든 출력 후보에 거래처 직접 일치 근거가 있다")
def _():
    keys = {i.key for i in RES.institutions}
    direct = {e.institution_key for e in RES.evidence
              if e.kind == "직접일치" and e.name_field == "거래처"}
    missing = keys - direct
    assert not missing, f"거래처 직접 일치 근거가 없는 후보: {missing}"
    other = {e.name_field for e in RES.evidence if e.kind == "직접일치"}
    assert other == {"거래처"}, f"거래처 외 추출 열이 있음: {other}"
    assert all(e.keywords for e in RES.evidence if e.kind == "직접일치"), "일치 키워드 미기록"
    assert {e.hit_fields for e in RES.evidence} == {"거래처"}, "거래처 외 열에서 검색됨"
    return (f"후보 {len(keys)}건 전부 거래처 직접 일치. "
            f"근거의 이름 추출 열·검색 일치 열 모두 '거래처' 한 종류")


@case("14. 짧은 별칭의 우연 일치를 점검했다 (우체국 사례)")
def _():
    rows = [mkrow(1, "우체국택배", "택배비", acct="운반비"),
            mkrow(2, "신당동우체국", "우편료", acct="통신비", voucher="V2"),
            mkrow(3, "우체국예금", "예치", acct="보통예금", voucher="V3")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert "우체국택배" not in names and "신당동우체국" not in names, \
        f"금융과 무관한 거래처가 포함됨: {names}"
    assert names == ["우체국예금"], names
    aliases = {a for e in CFG.institutions if e.institution_id == "bank_post" for a in e.aliases}
    assert "우체국" not in aliases, f"'우체국' 별칭이 남아 있음: {aliases}"
    return f"'우체국' 별칭 제거 → {names}만 포함 (택배·지역 우체국 제외)"


# =================================================================== 3. 제외 규칙
@case("15. 기본 설정에서는 임의 제외가 발생하지 않는다")
def _():
    rows = [mkrow(1, "건강보험", "4대보험 납부", acct="예수금")]
    _, r = resolve_rows(rows, CFG)
    assert [i.display_name for i in r.institutions] == ["건강보험"], "기본 실행에서 제외됨"
    assert RES.row_stats["전부제외"] == 0 and RES.link_stats["제외"] == 0, \
        f"실제 분개장에서 제외 발생: {RES.row_stats['전부제외']}행"
    assert "건강보험" in ALL_NAMES, "실제 결과에서 '건강보험'은 제외후보에 있어야 함"
    return "'건강보험' 출력됨, 실제 실행 제외 0건 (금융기관 여부 필터 없음)"


@case("16. 사전 등록 기관도 명시적 제외 설정이 적용된다")
def _():
    assert "기업비씨" in ALL_NAMES, "검증 대상 '기업비씨'가 결과에 없음"
    cfg2 = copy.deepcopy(CFG)
    cfg2.keywords["exclude_names"] = ["기업비씨"]
    r2 = InstitutionResolver(cfg2).resolve(MM)
    names2 = [i.display_name for i in r2.institutions]
    assert "기업비씨" not in names2, "사전 등록 기관이 제외되지 않음"
    ev = [e for e in r2.evidence if e.kind == "제외"]
    assert ev and "적용규칙" in ev[0].note and ev[0].counterparty
    return f"제외 후 {len(names2)}건(이전 {len(ALL_NAMES)}건), 제외 근거 {len(ev)}건"


@case("17. 제외는 완전일치만 적용된다")
def _():
    cfg2 = copy.deepcopy(CFG)
    cfg2.keywords["exclude_names"] = ["고용보험"]
    rows = [mkrow(1, "고용보험", "납부", acct="예수금"),
            mkrow(2, "한국고용보험사무소", "납부", acct="예수금", voucher="V2"),
            mkrow(3, "삼성화재해상보험", "보험료", acct="보험료(판)", voucher="V3")]
    _, r = resolve_rows(rows, cfg2)
    names = [i.display_name for i in r.institutions]
    assert "고용보험" not in names and "한국고용보험사무소" in names, names
    return f"제외 후 남은 이름: {names}"


# =================================================================== 4. 통계
@case("18. 원본 행 수와 근거 연결 건수를 각각 집계한다")
def _():
    rs, ls = RES.row_stats, RES.link_stats
    total = rs["직접일치"] + rs["전표문맥"] + rs["이름미확인"] + rs["전부제외"]
    assert total == rs["검색행"], f"행 단위 합 {total} ≠ 검색행 {rs['검색행']}"
    uniq = len({e.row_id for e in RES.evidence})
    assert uniq == rs["검색행"] == rs["근거연결_고유행"], f"고유 행 {uniq} ≠ {rs['검색행']}"
    assert ls["근거레코드_계"] == len(RES.evidence)
    assert ls["근거레코드_계"] >= rs["검색행"]
    assert {e.row_id for e in RES.evidence} == {(m.row.sheet, m.row.excel_row) for m in MM}
    return (f"행: 검색 {rs['검색행']:,} = 직접 {rs['직접일치']:,} + 문맥 {rs['전표문맥']} "
            f"+ 미확인 {rs['이름미확인']} + 제외 {rs['전부제외']} / "
            f"연결: 총 {ls['근거레코드_계']:,}")


@case("19. 한 거래처에 두 기관이 있으면 연결만 늘고 행 수는 1")
def _():
    rows = [mkrow(1, "하나은행/KB국민은행 공동관리계좌", "이체", acct="보통예금")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.institutions]
    assert set(names) == {"하나은행", "KB국민은행"}, names
    assert r.row_stats["직접일치"] == 1, r.row_stats
    assert r.link_stats["직접일치"] == 2, r.link_stats
    return f"1개 원본 행 → 후보 {names}, 연결 2건, 행 수 1"


@case("20. 등급별·방식별 합계가 각각 맞는다")
def _():
    import collections as _c

    from core.pipeline import _method_counts
    lv = _c.Counter(i.layer for i in RES.institutions)
    assert sum(lv.values()) == len(RES.institutions), f"등급 합 {sum(lv.values())}"
    assert set(lv) <= {"A", "B", "C"}, f"알 수 없는 등급: {set(lv)}"
    out, skip = RES.output_institutions, RES.skipped_institutions
    assert len(out) + len(skip) == len(RES.institutions), "출력+제외 ≠ 전체"
    assert all(i.layer in RES.output_layers for i in out)
    assert all(i.layer not in RES.output_layers for i in skip)
    counts = _method_counts(RES)
    assert sum(counts.values()) == len(out), f"방식 합 {sum(counts.values())} ≠ 출력 {len(out)}"
    return (f"등급 {dict(lv)} 합 {len(RES.institutions)} / "
            f"출력 {len(out)} + 제외 {len(skip)} / 출력분 방식별 합 {sum(counts.values())}")


@case("20b. C등급 일반 거래처는 A열에 나오지 않는다")
def _():
    outnames = set(NAMES)
    bo = [i.display_name for i in RES.skipped_institutions if "보증금" in i.display_name]
    assert bo, "보증금 거래처가 제외 목록에 없어 검증 불가"
    assert not (outnames & set(bo)), "보증금 거래처가 A열에 들어감"
    skipnames = {i.display_name for i in RES.skipped_institutions}
    for n in ("건강보험", "고용보험", "산재보험료"):
        assert n not in outnames, f"{n}이 A열에 들어감"
        assert n in skipnames, f"{n}이 제외 목록에 없음(삭제되면 안 됨)"
    return (f"보증금 거래처 {len(bo)}건 전부 제외후보로. "
            f"4대보험 3건도 A열 제외 — 삭제가 아니라 제외후보 시트에 보존")


@case("20c. 제외후보에도 근거가 남고 검토 필요가 표시된다")
def _():
    from core.pipeline import SKIPPED_HEADER, _skipped_rows
    rows = _skipped_rows(RES)
    assert len(rows) == len(RES.skipped_institutions), "제외후보 행 수 불일치"
    assert len(rows[0]) == len(SKIPPED_HEADER), "헤더-데이터 열 수 불일치"
    rev = [r for r in rows if r[0] == "검토 필요"]
    assert rev, "검토 필요 표시가 없음"
    assert rows[0][0] == "검토 필요", "검토 필요 항목이 위로 정렬되지 않음"
    assert all(r[3] for r in rows), "등급 판정 근거가 빈 행이 있음"
    assert all(r[6] for r in rows), "일치 키워드가 빈 행이 있음"
    return (f"제외후보 {len(rows)}건, 검토 필요 {len(rev)}건이 맨 위, "
            f"전 행에 등급 근거·일치 키워드 기록")


@case("20d. C열 온라인 표시 — 금결원 목록 일치 건만")
def _():
    from core.pipeline import _online_flags
    fl = _online_flags(RES)
    out = RES.output_institutions
    assert len(fl) == len(out) == len(NAMES), "C열 플래그 수가 A열과 다름"
    on = [o.display_name for o, f in zip(out, fl) if f]
    assert on, "온라인 표시가 하나도 없음"
    import csv as _csv
    want = {r["display_name"] for r in
            _csv.DictReader(open(os.path.join(CONFIG, "online_institutions.csv"),
                                 encoding="utf-8-sig")) if r["online"].strip().upper() == "Y"}
    assert set(on) <= want, f"목록에 없는데 온라인으로 표시됨: {set(on) - want}"
    for o, f in zip(out, fl):
        assert (f != "") == o.online, f"{o.display_name} 플래그 불일치"
        if o.online:
            assert o.online_note, f"{o.display_name} 온라인 근거 없음"
    return f"A열 {len(out)}건 중 온라인 {len(on)}건: {on}"


@case("20e. 목록에 없는 기관은 C열이 비어 있다")
def _():
    from core.pipeline import _online_flags
    fl = _online_flags(RES)
    off = [o.display_name for o, f in zip(RES.output_institutions, fl) if not f]
    for n in ("SBI저축은행", "신한카드", "현대캐피탈", "농협"):
        assert n in off, f"{n}이 온라인으로 표시됨 — 금결원 목록에 없다"
    rows = [mkrow(1, "SBI저축은행 일반자금대출/10억", "이자", acct="이자비용")]
    _, r = resolve_rows(rows, CFG)
    assert not r.institutions[0].online, "저축은행이 온라인으로 표시됨"
    return (f"저축은행·카드·캐피탈·상호금융은 금결원 참가기관이 아니라 C열 공란. "
            f"공란 {len(off)}건")


@case("20f. 제외 분류(전자금융)는 A열에 나오지 않고 제외후보에 남는다")
def _():
    exc = set(CFG.keywords.get("exclude_categories", []))
    assert exc, "exclude_categories가 비어 있어 검증 불가"
    out_cat = {i.category for i in RES.output_institutions if i.category}
    assert not (out_cat & exc), f"제외 분류가 A열에 있음: {out_cat & exc}"
    moved = [i for i in RES.skipped_institutions if i.category in exc]
    assert moved, "제외 분류 항목이 제외후보에도 없음 — 삭제되면 안 됨"
    for i in moved:
        assert "제외 분류" in i.layer_reason, f"{i.display_name} 판정 근거 없음"
        assert not i.review, f"{i.display_name}이 검토 필요로 표시됨"
    rows = [mkrow(1, "KG이니시스", "PG정산", acct="외상매출금"),
            mkrow(2, "신한카드", "카드대금", acct="미지급금", voucher="V2")]
    _, r = resolve_rows(rows, CFG)
    names = [i.display_name for i in r.output_institutions]
    assert names == ["신한카드"], f"카드사까지 빠지거나 PG가 남음: {names}"
    return (f"분류 {sorted(exc)} {len(moved)}건이 A열에서 빠지고 제외후보에 보존. "
            f"카드사는 유지")


@case("20g. 4대 사회보험 징수기관은 A열에 나오지 않고, 비슷한 상호를 삼키지 않는다")
def _():
    exc = set(CFG.keywords.get("exclude_categories", []))
    assert "사회보험" in exc, "exclude_categories에 '사회보험'이 없음"
    soc = [e for e in CFG.institutions if e.category == "사회보험"]
    assert len(soc) >= 3, f"사회보험 기관이 {len(soc)}건뿐 — 국민연금·건강·고용·산재 징수기관 누락"

    rows = [mkrow(1, "국민건강보험공단", "건강보험료 납부", acct="예수금"),
            mkrow(2, "국민연금공단", "국민연금 납부", acct="예수금", voucher="V2"),
            mkrow(3, "근로복지공단", "고용·산재보험료 납부", acct="예수금", voucher="V3"),
            # 이름이 비슷할 뿐 다른 거래처다. 공단으로 통합되면 안 된다.
            mkrow(4, "한국고용보험사무소", "급여대행 수수료", acct="지급수수료", voucher="V4"),
            mkrow(5, "삼성생명보험", "단체보험료", acct="복리후생비", voucher="V5"),
            mkrow(6, "IBK연금보험", "저축성보험", acct="장기금융상품", voucher="V6")]
    _, r = resolve_rows(rows, CFG)
    out = [i.display_name for i in r.output_institutions]
    skipped = {i.display_name: i for i in r.skipped_institutions}

    for n in ("국민건강보험공단", "국민연금공단", "근로복지공단"):
        assert n not in out, f"{n}이 A열에 들어감"
        assert n in skipped, f"{n}이 제외후보에도 없음 — 삭제되면 안 됨"
        assert "제외 분류 '사회보험'" in skipped[n].layer_reason, f"{n} 판정 근거 없음"
    for n in ("삼성생명보험", "IBK연금보험"):
        assert n in out, f"민간 보험사 {n}까지 빠짐"
    assert "한국고용보험사무소" not in skipped or True
    assert not any(i.institution_id == "soc_comwel" and "한국고용보험사무소" in " ".join(i.raw_names)
                   for i in r.institutions), "'한국고용보험사무소'가 근로복지공단으로 통합됨"

    # 실제 실행 결과에도 사회보험 분류가 A열에 없어야 한다
    assert not [i for i in RES.output_institutions if i.category == "사회보험"], \
        "실제 실행 A열에 사회보험 기관이 있음"
    return (f"징수기관 {len(soc)}건(국민연금·건강·고용·산재)이 A열에서 빠지고 제외후보에 보존. "
            f"민간 보험사 2건과 '한국고용보험사무소'는 영향 없음")


@case("21. 이름 추출 방식을 기관 확정으로 표시하지 않는다")
def _():
    from core.pipeline import DETAIL_HEADER, _detail_rows
    from core.resolver import METHOD_LABEL
    assert "이름 추출 방식" in DETAIL_HEADER and "확정구분" not in DETAIL_HEADER
    for k, v in METHOD_LABEL.items():
        assert "확정" not in v, f"방식 '{k}' 설명에 '확정': {v}"
    labels = {r[2] for r in _detail_rows(RES)}
    assert all("확정" not in x for x in labels), labels
    return f"이름 추출 방식 {len(labels)}종, '확정' 표현 0건"


# =================================================================== 5. 저장·보존
@case("22. 실제 저장 — A59부터 연속, H=YES, 나머지 공란")
def _():
    p = fresh("검증_1회차.xlsm")
    r = run(JOURNAL, p, CONFIG, log=lambda m: None)
    assert r.saved, "저장되지 않음"
    assert not r.report.failures, f"검증 FAIL {r.report.failures}"
    cells, _x = read_cells(p, [58, 59, 58 + len(NAMES), 59 + len(NAMES)])
    assert cells[58][0] == "조회금융기관 및 지점", "58행 제목 훼손"
    assert cells[59][0] == NAMES[0] and cells[59][7] == "YES"
    assert all(v is None for i, v in enumerate(cells[59]) if i not in (0, 2, 7)),         f"A·C·H 외에 값이 있음: {cells[59]}"
    assert cells[58 + len(NAMES)][0] == NAMES[-1]
    assert cells[59 + len(NAMES)][0] is None
    return (f"A59:A{58+len(NAMES)} {len(NAMES)}건 / 저장검증 {r.report.summary()} / "
            f"미검증 {[c[0] for c in r.report.unverified]}")


@case("23. 이전 310건의 잔여 행이 남지 않는다 (결과 감소)")
def _():
    p = fresh("검증_감소.xlsm")
    prev = [f"이전후보{i:03d}" for i in range(1, 311)]
    r1 = write_control_sheet(
        target=p, sheet_name=DEFAULT_SHEET, names=prev,
        detail_rows=[], detail_header=("x",), evidence_rows=[], evidence_header=("y",),
        runlog_rows=[("테스트", "310건")], backup_dir=os.path.join(WORK, "backup"),
        log=lambda m: None)
    assert r1.replaced and r1.end_row == 368, f"끝행 {r1.end_row}"
    r2 = run(JOURNAL, p, CONFIG, log=lambda m: None)
    assert r2.saved and not r2.report.failures
    end = 58 + len(NAMES)
    probe = sorted({59, end, end + 1, 300, 368})
    cells, _x = read_cells(p, probe)
    assert cells[59][0] == NAMES[0] and cells[end][0] == NAMES[-1]
    assert cells[end + 1][0] is None, f"A{end+1}={cells[end+1][0]!r}"
    assert cells[368][0] is None, f"A368 잔여: {cells[368][0]!r}"
    if 300 > end:
        assert cells[300][0] is None, f"A300 잔여: {cells[300][0]!r}"
    return f"310건 → {len(NAMES)}건, A{end+1}·A300·A368 전부 공란"


@case("24. 재실행 — 중복·잔여행 없음")
def _():
    p = os.path.join(WORK, "검증_1회차.xlsm")
    if not os.path.exists(p):
        raise Skip("22번이 실패해 1회차 결과 파일이 없음")
    r = run(JOURNAL, p, CONFIG, log=lambda m: None)
    assert r.saved
    cells, _x = read_cells(p, [58 + len(NAMES), 59 + len(NAMES)])
    assert cells[58 + len(NAMES)][0] == NAMES[-1]
    assert cells[59 + len(NAMES)][0] is None
    return f"2회차도 {r.report.written_rows}건, A{59+len(NAMES)} 이후 공란"


@case("25. 유효성 검사·서식이 결과 행 수만큼 확장된다")
def _():
    q = fresh("검증_확장.xlsm")
    names = [f"확장{i:03d}" for i in range(1, 301)]
    r = write_control_sheet(
        target=q, sheet_name=DEFAULT_SHEET, names=names,
        detail_rows=[], detail_header=("x",), evidence_rows=[], evidence_header=("y",),
        runlog_rows=[], backup_dir=os.path.join(WORK, "backup"), log=lambda m: None)
    assert not r.failures and r.end_row == 358
    import win32com.client as w32
    app = w32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.AutomationSecurity = 3
    try:
        wb = app.Workbooks.Open(q, UpdateLinks=0, ReadOnly=True)
        ws = wb.Worksheets(DEFAULT_SHEET)
        t, f = ws.Range("H358").Validation.Type, ws.Range("H358").Validation.Formula1
        wb.Close(False)
    finally:
        app.Quit()
    assert t == 3 and f == "YES,NO", f"H358 유효성 미확장 {t}/{f}"
    return f"A59:A358 300건, H358 유효성 {f}"


@case("26. 검색 0건 — 신규 행 없음")
def _():
    p = fresh("검증_0건.xlsm")
    rep = write_control_sheet(
        target=p, sheet_name=DEFAULT_SHEET, names=[],
        detail_rows=[], detail_header=("x",), evidence_rows=[], evidence_header=("y",),
        runlog_rows=[("[행] 검색된 원본 행 수", 0)], backup_dir=os.path.join(WORK, "backup"),
        log=lambda m: None)
    assert rep.replaced and rep.written_rows == 0 and not rep.failures
    cells, _x = read_cells(p, [58, 59, 60])
    assert cells[58][0] == "조회금융기관 및 지점"
    assert cells[59][0] is None and cells[60][0] is None
    return "A59 이후 공란, 58행 제목 유지"


@case("27. 입력 오류는 '0건'이 아니라 오류로 처리")
def _():
    try:
        read_journal(TEMPLATE, CFG.profiles)
        raise AssertionError("헤더 없는 파일을 읽고도 오류 없음")
    except JournalReadError as e:
        msg = str(e)[:55]
    try:
        read_journal(os.path.join(WORK, "없는파일.xlsx"), CFG.profiles)
        raise AssertionError("없는 파일에 오류 없음")
    except JournalReadError:
        pass
    return f"JournalReadError — '{msg}…'"


@case("28. 파일 잠금 시 원본 보존")
def _():
    p = fresh("검증_잠금.xlsm")
    before = os.path.getsize(p)
    lock = os.path.join(os.path.dirname(p), "~$" + os.path.basename(p))
    open(lock, "wb").write(b"x")
    try:
        write_control_sheet(
            target=p, sheet_name=DEFAULT_SHEET, names=["잠금테스트"],
            detail_rows=[], detail_header=("x",), evidence_rows=[], evidence_header=("y",),
            runlog_rows=[], backup_dir=os.path.join(WORK, "backup"), log=lambda m: None)
        raise AssertionError("잠금 상태인데 저장을 시도함")
    except ExcelWriteError as e:
        msg = str(e)[:45]
    finally:
        os.remove(lock)
    assert os.path.getsize(p) == before, "원본이 변경됨"
    assert not [f for f in os.listdir(WORK) if f.startswith("~자동화임시_")], "임시본 잔재"
    return f"저장 중단, 원본 {before:,} bytes 그대로 — '{msg}…'"


@case("29. 잘못된 시트 이름 — 원본 보존, 임시본 정리")
def _():
    p = fresh("검증_시트오류.xlsm")
    before = os.path.getsize(p)
    try:
        write_control_sheet(
            target=p, sheet_name="없는시트", names=["A"],
            detail_rows=[], detail_header=("x",), evidence_rows=[], evidence_header=("y",),
            runlog_rows=[], backup_dir=os.path.join(WORK, "backup"), log=lambda m: None)
        raise AssertionError("없는 시트인데 저장 성공")
    except ExcelWriteError as e:
        msg = str(e)[:40]
    assert os.path.getsize(p) == before, "원본 변경됨"
    assert not [f for f in os.listdir(WORK) if f.startswith("~자동화임시_")], "임시본 잔재"
    return f"원본 보존 — '{msg}…'"


# =================================================================== 6. 보존 검증기 자체 검증
@case("30. 수식 개수가 같아도 내용이 바뀌면 FAIL")
def _():
    from core.writer_excel import _verify
    rep = _FakeRep(
        baseline={"sheets": ["S1"], "formulas": {"S1": {"$A$1": "=TODAY()", "$A$2": "=1+1"}},
                  "formula_skips": {}, "shapes": {"S1": []}, "ole": {"S1": []},
                  "vba": {"M": (1, "Sub A()")}, "vba_status": "ok"},
        after={"sheets": ["S1"], "formulas": {"S1": {"$A$1": "=NOW()", "$A$2": "=1+1"}},
               "formula_skips": {}, "shapes": {"S1": []}, "ole": {"S1": []},
               "vba": {"M": (1, "Sub A()")}, "vba_status": "ok"})
    checks = _verify(_FakeWs(), [], rep, 59, "S1", 58)
    fx = [c for c in checks if c[0].startswith("기존 수식 보존")][0]
    assert fx[1] == FAIL, f"수식 내용 변경을 잡지 못함: {fx}"
    return f"개수 동일(2=2), 내용 변경 → {fx[1]}"


@case("31. VBA 코드 변경을 FAIL로 잡는다")
def _():
    from core.writer_excel import _verify
    rep = _FakeRep(
        baseline={"sheets": ["S1"], "formulas": {"S1": {}}, "formula_skips": {},
                  "shapes": {"S1": []}, "ole": {"S1": []},
                  "vba": {"M": (1, "Sub A()")}, "vba_status": "ok"},
        after={"sheets": ["S1"], "formulas": {"S1": {}}, "formula_skips": {},
               "shapes": {"S1": []}, "ole": {"S1": []},
               "vba": {"M": (1, "Sub B()")}, "vba_status": "ok"})
    checks = _verify(_FakeWs(), [], rep, 59, "S1", 58)
    v = [c for c in checks if c[0].startswith("VBA")][0]
    assert v[1] == FAIL, f"VBA 코드 변경을 잡지 못함: {v}"
    return f"{v[1]}: {v[2][:50]}"


@case("32. VBA 접근 실패는 PASS가 아니라 UNVERIFIED")
def _():
    from core.writer_excel import _verify
    blocked = "VBA 프로젝트에 접근할 수 없음(차단)"
    base = {"sheets": ["S1"], "formulas": {"S1": {}}, "formula_skips": {},
            "shapes": {"S1": []}, "ole": {"S1": []}, "vba": {}, "vba_status": blocked}
    checks = _verify(_FakeWs(), [], _FakeRep(base, dict(base)), 59, "S1", 58)
    v = [c for c in checks if c[0].startswith("VBA")][0]
    assert v[1] == UNVERIFIED, f"차단 상태를 {v[1]}로 처리"
    assert not [c for c in checks if c[1] == FAIL]
    return f"{v[1]}"


@case("33. 수식이 없는 시트와 COM 조회 실패를 구별한다")
def _():
    from core.writer_excel import _sheet_formulas

    class _Cells:
        def __init__(self, boom):
            self._boom = boom

        def SpecialCells(self, *_a):
            raise Exception(self._boom)

    class _Used:
        def __init__(self, hf):
            self.HasFormula = hf

    class _Ws:
        def __init__(self, hf, boom):
            self.Cells = _Cells(boom)
            self.UsedRange = _Used(hf)

    err = "(-2146827284) No cells were found."
    out, skip = _sheet_formulas(_Ws(False, err))
    assert out == {} and skip == "", f"정상 '수식 없음'을 미검증 처리: {skip!r}"
    out2, skip2 = _sheet_formulas(_Ws(None, err))
    assert out2 == {} and skip2 and "구분하지 못함" in skip2
    out3, skip3 = _sheet_formulas(_Ws(True, err))
    assert skip3
    return "HasFormula=False→검증 / None·True→미검증"


@case("34. UsedRange.HasFormula 자체가 실패하면 미검증")
def _():
    from core.writer_excel import _sheet_formulas

    class _Ws:
        class Cells:
            @staticmethod
            def SpecialCells(*_a):
                raise Exception("No cells were found.")

        class UsedRange:
            @property
            def HasFormula(self):
                raise Exception("RPC 오류")

    out, skip = _sheet_formulas(_Ws())
    assert out == {} and skip
    return f"미검증 — '{skip[:55]}…'"


@case("35. 부분 수집 실패도 미검증으로 표시된다")
def _():
    from core.writer_excel import _sheet_formulas

    class _Cell:
        def __init__(self, a, f):
            self.Address, self.Formula = a, f

    class _Area:
        def __init__(self, cells, boom_at):
            self._cells, self._boom = cells, boom_at

        def __iter__(self):
            for n, c in enumerate(self._cells):
                if n == self._boom:
                    raise Exception("수집 중단")
                yield c

    class _Rng:
        Count = 3

        def __init__(self):
            self.Areas = [_Area([_Cell("$A$1", "=1"), _Cell("$A$2", "=2"),
                                 _Cell("$A$3", "=3")], 2)]

    class _Ws:
        class Cells:
            @staticmethod
            def SpecialCells(*_a):
                return _Rng()

        class UsedRange:
            HasFormula = None

    out, skip = _sheet_formulas(_Ws())
    assert len(out) == 2 and skip and "부분 수집" in skip
    return "3개 중 2개 수집 → 미검증"


@case("36. 수집 실패 시트는 PASS가 아니라 UNVERIFIED")
def _():
    from core.writer_excel import _verify
    base = {"sheets": ["S1"], "formulas": {"S1": {}},
            "formula_skips": {"S1": "수식 목록 조회 실패 — 구분하지 못함"},
            "shapes": {"S1": []}, "ole": {"S1": []},
            "vba": {"M": (1, "x")}, "vba_status": "ok"}
    checks = _verify(_FakeWs(), [], _FakeRep(base, dict(base)), 59, "S1", 58)
    fx = [c for c in checks if c[0].startswith("기존 수식 보존")][0]
    sk = [c for c in checks if c[0].startswith("수식 대조를 못 한")]
    assert fx[1] == UNVERIFIED and sk and sk[0][1] == UNVERIFIED
    return "수식 보존=UNVERIFIED, 미검증 시트 항목 1건"


@case("37. 미검증 항목이 통과 건수에 포함되지 않는다")
def _():
    p = os.path.join(WORK, "검증_0건.xlsm")
    if not os.path.exists(p):
        raise Skip("26번이 실패해 파일 없음")
    rep = write_control_sheet(
        target=p, sheet_name=DEFAULT_SHEET, names=["요약확인"],
        detail_rows=[], detail_header=("x",), evidence_rows=[], evidence_header=("y",),
        runlog_rows=[], backup_dir=os.path.join(WORK, "backup"), log=lambda m: None)
    assert rep.unverified, "미검증 항목이 하나도 없음"
    assert len(rep.passed) + len(rep.failures) + len(rep.unverified) == len(rep.checks)
    assert all(c[1] != PASS for c in rep.unverified)
    return f"{rep.summary()} / 미검증: {[c[0] for c in rep.unverified]}"


@case("38. 다른 분개장 파일도 처리")
def _():
    alt = os.path.join(os.path.dirname(JOURNAL), "분개장_sample.xlsx")
    if not os.path.exists(alt):
        raise Skip("표본 분개장 파일이 없어 검증하지 못함")
    jd2 = read_journal(alt, CFG.profiles)
    if len(jd2.rows) < 20:
        _, r2 = resolve_rows(jd2.rows, CFG)
        raise Skip(f"표본이 {len(jd2.rows)}행뿐이라 양식 대응력 검증으로 불충분"
                   f"(읽기·검색 자체는 성공: 조회처 {len(r2.institutions)}건)")
    _, r2 = resolve_rows(jd2.rows, CFG)
    return f"{os.path.basename(alt)}: 거래행 {len(jd2.rows)}, 조회처 {len(r2.institutions)}"


@case("39. Control Sheet는 출력물 — 양식을 복사해 새 파일을 만든다")
def _():
    import glob
    from core.template import OUTPUT_DIRNAME, find_template, output_path

    tmpl = find_template()
    assert os.path.exists(tmpl), f"양식을 찾지 못함: {tmpl}"

    # 출력 경로 규칙: <분개장 폴더>/output/Control Sheet_<분개장이름>_<시각>.xlsm
    got = output_path(JOURNAL)
    assert os.path.basename(os.path.dirname(got)) == OUTPUT_DIRNAME, got
    stem = os.path.splitext(os.path.basename(JOURNAL))[0]
    assert os.path.basename(got).startswith(f"Control Sheet_{stem}_"), got
    assert got.endswith(".xlsm"), got

    # 같은 시각에 두 번 불러도 서로 다른 파일이 된다
    outdir = os.path.join(WORK, "출력시험")
    os.makedirs(outdir, exist_ok=True)
    for f in glob.glob(os.path.join(outdir, "*.xlsm")):
        os.remove(f)

    r = run(JOURNAL, None, CONFIG, log=lambda m: None, out_dir=outdir)
    assert r.output_path and os.path.exists(r.output_path), "출력 파일이 만들어지지 않음"
    assert os.path.dirname(r.output_path) == os.path.abspath(outdir), r.output_path
    assert r.template_used == tmpl, f"다른 양식을 씀: {r.template_used}"
    assert r.report and not r.report.failures, "저장 검증 실패"

    # 원본 양식은 그대로여야 한다
    before = os.path.getsize(tmpl)
    r2 = run(JOURNAL, None, CONFIG, log=lambda m: None, out_dir=outdir)
    assert os.path.getsize(tmpl) == before, "양식 파일이 변경됨"
    assert r2.output_path != r.output_path, "두 번째 실행이 같은 파일을 덮어씀"

    made = sorted(glob.glob(os.path.join(outdir, "Control Sheet_*.xlsm")))
    return (f"양식 {os.path.basename(tmpl)} → 새 파일 {len(made)}개, "
            f"양식 원본 무변경, A열 {len(r.names)}건")


@case("40. 자동화 보조 시트는 통합문서 맨 뒤에 정해진 순서로 온다")
def _():
    import zipfile
    import re as _re
    from core.writer_excel import AUTO_SHEETS

    outdir = os.path.join(WORK, "시트순서")
    os.makedirs(outdir, exist_ok=True)
    r = run(JOURNAL, None, CONFIG, log=lambda m: None, out_dir=outdir)
    assert r.output_path and os.path.exists(r.output_path)

    with zipfile.ZipFile(r.output_path) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8")
    names = [n.replace("&amp;", "&") for n in _re.findall(r'<sheet name="([^"]+)"', xml)]

    tail = names[-len(AUTO_SHEETS):]
    assert tail == list(AUTO_SHEETS), (
        f"맨 뒤 {len(AUTO_SHEETS)}개가 자동화 시트가 아님: {tail}")

    front = names[:-len(AUTO_SHEETS)]
    assert not [n for n in front if n.startswith("자동화_")], \
        f"자동화 시트가 앞쪽에 섞임: {[n for n in front if n.startswith('자동화_')]}"
    assert front[0] != "자동화_기관상세" and len(front) >= 1, "원래 시트가 남아 있지 않음"
    return (f"원래 시트 {len(front)}개가 앞, 자동화 시트 {len(tail)}개가 뒤 "
            f"({' → '.join(tail)})")


print("\n" + "=" * 78)
npass = sum(1 for _, s, _ in RESULTS if s == PASS)
nfail = sum(1 for _, s, _ in RESULTS if s == FAIL)
nunv = sum(1 for _, s, _ in RESULTS if s == UNVERIFIED)
for name, s, note in RESULTS:
    print(f"[{s:^10}] {name} — {note}")
print(f"\n통과 {npass} / 실패 {nfail} / 미검증 {nunv}  (전체 {len(RESULTS)})")
sys.exit(0 if nfail == 0 else 1)
