# -*- coding: utf-8 -*-
"""금융기관·변호사 조회처 추출 자동화 — 진입점.

화면  : python src/app.py
명령줄 : python src/app.py --journal <분개장.xlsx> --target <Control Sheet.xlsm> [--dry-run]
"""
from __future__ import annotations

import argparse
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 한글 콘솔(cp949)에서 '—' 같은 글자에 걸려 실행이 끊기지 않게 한다.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from core.paths import config_dir  # noqa: E402
from core.pipeline import DEFAULT_SHEET, START_ROW, run  # noqa: E402
from core.readers import JournalReadError  # noqa: E402
from core.template import TemplateError  # noqa: E402
from core.writer_excel import ExcelWriteError  # noqa: E402


def run_cli(args: argparse.Namespace) -> int:
    try:
        res = run(args.journal, args.target, config_dir(), sheet_name=args.sheet,
                  profile_name=args.profile, start_row=args.start_row,
                  save=not args.dry_run, log=lambda m: print(m, flush=True),
                  out_dir=args.outdir, template_path=args.template)
    except JournalReadError as exc:
        print(f"[입력 오류] {exc}", file=sys.stderr)
        return 2
    except TemplateError as exc:
        print(f"[양식 오류] {exc}", file=sys.stderr)
        return 4
    except ExcelWriteError as exc:
        print(f"[저장 실패] {exc}", file=sys.stderr)
        return 3
    except Exception as exc:  # noqa: BLE001
        print(f"[오류] {exc}", file=sys.stderr)
        traceback.print_exc()
        return 1

    print("\n=== 조회처 목록 ===")
    for i, name in enumerate(res.names, 1):
        print(f"{i:3d}. {name}")
    if res.report:
        print("\n=== 검증 ===")
        for name, state, note in res.report.checks:
            print(f"  [{state:^10}] {name} — {note}")
        print(f"\n검증 요약: {res.report.summary()}")
        print(f"백업: {res.report.backup}")
    print(f"\n{res.message}")
    return 0


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="분개장에서 금융기관·변호사 조회처 이름을 추출해 Control Sheet A59부터 작성")
    ap.add_argument("--journal", help="분개장 파일 경로")
    ap.add_argument("--target", default=None,
                    help="이미 있는 Control Sheet에 쓸 때만 지정. 기본은 새로 만든다")
    ap.add_argument("--outdir", default=None,
                    help="출력 폴더. 기본은 분개장 옆 output/")
    ap.add_argument("--template", default=None,
                    help="쓸 양식(.xlsm). 기본은 template 폴더 또는 내장 표준양식")
    ap.add_argument("--sheet", default=DEFAULT_SHEET, help="출력 시트 이름")
    ap.add_argument("--profile", default=None, help="config/input_profiles.json의 프로파일 이름")
    ap.add_argument("--start-row", type=int, default=START_ROW, help="출력 시작 행(기본 59)")
    ap.add_argument("--dry-run", action="store_true", help="저장하지 않고 검색 결과만 출력")
    ap.add_argument("--gui", action="store_true", help="화면 실행")
    return ap.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.journal and not args.gui:
        return run_cli(args)
    from ui import launch  # 화면을 쓸 때만 tkinter를 불러온다
    return launch()


if __name__ == "__main__":
    raise SystemExit(main())
