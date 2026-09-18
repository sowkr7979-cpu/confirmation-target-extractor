# -*- coding: utf-8 -*-
"""검색 → 통합 → Excel 저장을 잇는 실행 흐름. 검색 로직과 저장 로직은 분리되어 있다."""
from __future__ import annotations

import datetime as _dt
import os
from dataclasses import dataclass, field
from typing import Any, Callable

from .config_loader import AppConfig, load_config
from .matcher import KeywordMatcher
from .readers import JournalData, JournalReadError, read_journal
from .resolver import (METHOD_LABEL, Institution, InstitutionResolver,
                       ResolveResult)
from .template import TemplateError, create_output
from .writer_excel import ExcelWriteError, WriteReport, write_control_sheet

DEFAULT_SHEET = "정보입력&출력&Control sheet"
START_ROW = 59

DETAIL_HEADER = (
    "후보ID", "조회처명(A열 출력)", "등급", "온라인 발급", "이름 추출 방식", "사전 분류",
    "원문 거래처·적요 명칭",
    "지점·사용처 단서", "계좌 단서(선행 0 유지, 계좌번호 단정 아님)", "관련 계정과목",
    "검색 키워드", "직접일치 행수", "전표문맥 행수", "고유 전표수",
    "최초 거래일", "최종 거래일", "차변 합계(원문)", "대변 합계(원문)", "통합 근거",
)
DETAIL_TEXT_COLS = (0, 6, 7, 8)

EVIDENCE_HEADER = (
    "구분", "후보ID", "조회처명", "원본 시트", "원본 행", "전표 고유번호", "전표일자",
    "계정과목", "거래처(원문)", "적요(원문)", "차변", "대변",
    "검색 키워드", "일치한 열", "이름 추출 열", "지점 단서", "계좌 단서", "비고",
    "적요 미채택 조각(기관명으로 쓰지 않음)",
)
EVIDENCE_TEXT_COLS = (1, 5, 8, 9, 15, 16, 18)

SKIPPED_HEADER = (
    "검토 필요", "조회처명(A열에 넣지 않음)", "등급", "등급 판정 근거",
    "이름 추출 방식", "원문 거래처명", "일치한 검색 키워드", "관련 계정과목",
    "직접일치 행수", "고유 전표수", "최초 거래일", "최종 거래일",
    "차변 합계(원문)", "대변 합계(원문)",
)
SKIPPED_TEXT_COLS = (1, 5)


@dataclass
class RunResult:
    journal: JournalData
    resolve: ResolveResult
    names: list[str]
    report: WriteReport | None = None
    runlog: list[tuple[str, Any]] = field(default_factory=list)
    saved: bool = False
    message: str = ""
    output_path: str = ""       # 만들어진 Control Sheet
    template_used: str = ""     # 어떤 양식을 복사했는지


def _method_counts(res: ResolveResult) -> dict[str, int]:
    """이름을 얻은 방식별 건수. 방식은 서로 배타적이므로 합계 = 전체 조회처 수."""
    out: dict[str, int] = {}
    for i in res.output_institutions:
        out[i.matched_by] = out.get(i.matched_by, 0) + 1
    return out


def _detail_rows(res: ResolveResult) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for inst in res.output_institutions:
        rows.append([
            inst.key, inst.display_name, inst.layer,
            ("온라인" if inst.online else ""),
            METHOD_LABEL.get(inst.matched_by, inst.matched_by),
            inst.category or "",
            " | ".join(inst.raw_names),
            " | ".join(inst.branch_clues),
            " | ".join(inst.account_clues),
            " | ".join(inst.accounts),
            ", ".join(inst.keywords),
            inst.direct_rows, inst.context_rows, len(inst.vouchers),
            inst.date_min, inst.date_max, inst.debit_sum, inst.credit_sum,
            f"{inst.merge_note} / 등급 근거: {inst.layer_reason}"
            + (f" / 온라인 근거: {inst.online_note}" if inst.online else ""),
        ])
    return rows


def _online_flags(res: ResolveResult, label: str = "온라인") -> list[str]:
    """A열에 쓰는 후보와 같은 순서로 C열에 넣을 값."""
    return [(label if i.online else "") for i in res.output_institutions]


def _skipped_rows(res: ResolveResult) -> list[list[Any]]:
    """A열에 넣지 않은 후보. 검토 필요 항목을 위로 올린다."""
    items = sorted(res.skipped_institutions, key=lambda i: (not i.review, i.order))
    return [[
        "검토 필요" if i.review else "", i.display_name, i.layer, i.layer_reason,
        METHOD_LABEL.get(i.matched_by, i.matched_by),
        " | ".join(i.raw_names), ", ".join(i.keywords), " | ".join(i.accounts),
        i.direct_rows, len(i.vouchers), i.date_min, i.date_max,
        i.debit_sum, i.credit_sum,
    ] for i in items]


def _evidence_rows(res: ResolveResult) -> list[list[Any]]:
    return [[
        e.kind, e.institution_key, e.institution_name, e.sheet, e.excel_row,
        e.voucher_id, e.date, e.account, e.counterparty, e.description,
        e.debit, e.credit, e.keywords, e.hit_fields, e.name_field,
        e.branch_clue, e.account_clue, e.note, e.desc_fragment,
    ] for e in res.evidence]


def analyze(journal_path: str, config: AppConfig,
            profile_name: str | None = None,
            log: Callable[[str], None] = lambda _m: None) -> tuple[JournalData, ResolveResult]:
    log(f"분개장 읽는 중: {os.path.basename(journal_path)}")
    jd = read_journal(journal_path, config.profiles, profile_name)
    log(f"양식 '{jd.profile_name}' / 시트 '{jd.sheet}' / 헤더 {jd.header_row}행 / "
        f"거래행 {len(jd.rows):,}건 (제외 {jd.skipped})")
    if jd.missing_fields:
        log(f"주의: 연결되지 않은 열 {jd.missing_fields}")
    log(f"거래 기간 {jd.date_min} ~ {jd.date_max}")

    matches = KeywordMatcher(config).search(jd.rows)
    log(f"키워드 검색 — 원본 행 {len(matches):,}건")
    res = InstitutionResolver(config).resolve(matches)
    log(f"원본 행 단위: {res.row_stats}")
    log(f"후보-근거 연결 단위: {res.link_stats}")
    log(f"후보 {len(res.institutions)}건 → A열 출력 {len(res.output_institutions)}건"
        f"(A {res.row_stats.get('등급A_후보',0)} / B {res.row_stats.get('등급B_후보',0)}), "
        f"제외후보 {len(res.skipped_institutions)}건"
        f"(검토 필요 {res.row_stats.get('등급C_검토필요',0)}) / "
        f"온라인 표시 {res.row_stats.get('온라인발급_후보',0)}건")
    return jd, res


def run(journal_path: str, target_path: str | None, config_dir: str,
        sheet_name: str = DEFAULT_SHEET, profile_name: str | None = None,
        start_row: int = START_ROW, save: bool = True,
        log: Callable[[str], None] = lambda _m: None,
        out_dir: str | None = None, template_path: str | None = None) -> RunResult:
    """target_path를 주지 않으면 양식을 복사해 **새 Control Sheet를 만들어** 거기에 쓴다."""
    cfg = load_config(config_dir)
    log(f"설정 버전: {cfg.versions}")
    jd, res = analyze(journal_path, cfg, profile_name, log)
    names = [i.display_name for i in res.output_institutions]

    runlog: list[tuple[str, Any]] = [
        ("실행일시", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("분개장 파일", journal_path),
        ("분개장 시트", jd.sheet),
        ("입력 프로파일", jd.profile_name),
        ("헤더 행", jd.header_row),
        ("거래행 수", len(jd.rows)),
        ("제외한 행", str(jd.skipped)),
        ("거래 기간", f"{jd.date_min} ~ {jd.date_max}"),
        ("[행] 검색된 원본 행 수", res.row_stats.get("검색행", 0)),
        ("[행] 후보 직접연결 행 수", res.row_stats.get("직접일치", 0)),
        ("[행] 전표문맥 행 수", res.row_stats.get("전표문맥", 0)),
        ("[행] 이름미확인 행 수", res.row_stats.get("이름미확인", 0)),
        ("[행] 전부 제외된 행 수", res.row_stats.get("전부제외", 0)),
        ("[행] 근거에 연결된 고유 행 수", res.row_stats.get("근거연결_고유행", 0)),
        ("[연결] 직접일치 근거 레코드", res.link_stats.get("직접일치", 0)),
        ("[연결] 전표문맥 근거 레코드", res.link_stats.get("전표문맥", 0)),
        ("[연결] 이름미확인 근거 레코드", res.link_stats.get("이름미확인", 0)),
        ("[연결] 제외 근거 레코드", res.link_stats.get("제외", 0)),
        ("[연결] 근거 레코드 합계", res.link_stats.get("근거레코드_계", 0)),
        ("주의", "행 단위 수치와 연결 단위 수치는 서로 더하지 않는다"),
        ("조회처 건수", len(names)),
        ("[등급] A 사전 일치", res.row_stats.get("등급A_후보", 0)),
        ("[등급] B 금융·법률 명칭 형태", res.row_stats.get("등급B_후보", 0)),
        ("[등급] C 일반 거래처(A열 제외)", res.row_stats.get("등급C_후보", 0)),
        ("[등급] C 중 검토 필요", res.row_stats.get("등급C_검토필요", 0)),
        ("[온라인] C열 표시 건수", res.row_stats.get("온라인발급_후보", 0)),
        ("A열 출력 등급", ", ".join(sorted(res.output_layers))),
        ("주의2", "등급과 추출 방식은 이름의 형태일 뿐 금융기관·변호사 확정이 아니다"),
        *[(f"[이름 추출 방식] {k}", v) for k, v in _method_counts(res).items()],
        ("[이름 추출 방식] 합계", sum(_method_counts(res).values())),
        ("출력 시트", sheet_name),
        ("keywords.json", cfg.versions["keywords.json"]),
        ("institutions.csv", cfg.versions["institutions.csv"]),
        ("input_profiles.json", cfg.versions["input_profiles.json"]),
        ("online_institutions.csv", cfg.versions.get("online_institutions.csv", "없음")),
    ]

    result = RunResult(journal=jd, resolve=res, names=names, runlog=runlog)
    if not save:
        result.message = "분석만 수행했습니다(저장 안 함)."
        return result

    if target_path:
        log(f"기존 Control Sheet에 씁니다: {target_path}")
    else:
        target_path, used = create_output(journal_path, out_dir, template_path)
        result.output_path, result.template_used = target_path, used
        log(f"양식 복사: {os.path.basename(used)} → {target_path}")
    runlog.append(("출력 파일", target_path))
    if result.template_used:
        runlog.append(("사용한 양식", result.template_used))

    made_here = bool(result.output_path)
    try:
        rep = write_control_sheet(
            target=target_path, sheet_name=sheet_name, names=names,
            detail_rows=_detail_rows(res), detail_header=DETAIL_HEADER,
            evidence_rows=_evidence_rows(res), evidence_header=EVIDENCE_HEADER,
            runlog_rows=runlog,
            backup_dir=os.path.join(os.path.dirname(target_path), "backup"),
            online_flags=_online_flags(res, cfg.keywords.get("online_label", "온라인")),
            skipped_rows=_skipped_rows(res), skipped_header=SKIPPED_HEADER,
            detail_text_cols=DETAIL_TEXT_COLS, evidence_text_cols=EVIDENCE_TEXT_COLS,
            skipped_text_cols=SKIPPED_TEXT_COLS,
            log=log, start_row=start_row,
        )
    except Exception:
        # 우리가 만든 빈 사본을 남기지 않는다. 사용자가 결과물로 착각한다.
        if made_here and os.path.exists(target_path):
            try:
                os.remove(target_path)
                log(f"만들다 만 파일을 지웠습니다: {os.path.basename(target_path)}")
            except OSError:
                pass
        raise
    result.report = rep
    result.saved = rep.replaced
    result.output_path = result.output_path or target_path
    where = os.path.basename(target_path)
    result.message = (
        f"조회처 {len(names)}건을 '{where}'의 A{rep.start_row}:A{rep.end_row}에 작성했습니다."
        if names else f"검색 결과 0건 — '{where}'에 목록 없이 실행기록만 남겼습니다.")
    return result


__all__ = ["run", "analyze", "RunResult", "JournalReadError", "ExcelWriteError",
           "TemplateError", "DEFAULT_SHEET", "START_ROW", "Institution"]
