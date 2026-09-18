# -*- coding: utf-8 -*-
"""Control Sheet는 **출력물**이다. 빈 표준 양식을 복사해 새 파일을 만들어 준다.

양식을 찾는 순서
  1) 사용자가 지정한 파일(`--template`)
  2) 실행 파일 옆 `template/` 폴더의 첫 번째 .xlsm  ← 자기 양식을 쓰려면 여기 넣는다
  3) 프로그램에 함께 넣어 둔 표준 양식

출력 파일은 **분개장 옆 `output/` 폴더**에 만든다.
  <분개장 폴더>/output/Control Sheet_<분개장이름>_<날짜시각>.xlsm
"""
from __future__ import annotations

import datetime as _dt
import os
import re
import shutil

from .paths import app_dir, bundle_dir

TEMPLATE_DIRNAME = "template"
OUTPUT_DIRNAME = "output"
STANDARD_NAME = "Control Sheet_표준양식.xlsm"


class TemplateError(Exception):
    """양식을 찾지 못했거나 출력 파일을 만들지 못했을 때."""


def template_dirs() -> list[str]:
    """양식을 찾을 폴더. 앞에 있는 것이 우선이다."""
    out = [os.path.join(app_dir(), TEMPLATE_DIRNAME)]
    bundled = os.path.join(bundle_dir(), TEMPLATE_DIRNAME)
    if bundled not in out:
        out.append(bundled)
    return out


def find_template(explicit: str | None = None) -> str:
    if explicit:
        path = os.path.abspath(explicit)
        if not os.path.exists(path):
            raise TemplateError(f"지정한 양식 파일이 없습니다: {path}")
        return path

    for folder in template_dirs():
        if not os.path.isdir(folder):
            continue
        standard = os.path.join(folder, STANDARD_NAME)
        if os.path.exists(standard):
            return standard
        cands = sorted(f for f in os.listdir(folder)
                       if f.lower().endswith(".xlsm") and not f.startswith("~$"))
        if cands:
            return os.path.join(folder, cands[0])

    raise TemplateError(
        "Control Sheet 양식을 찾지 못했습니다. "
        f"실행 파일 옆 '{TEMPLATE_DIRNAME}' 폴더에 .xlsm 양식을 넣으세요. "
        f"(찾아본 곳: {', '.join(template_dirs())})")


def _safe(name: str) -> str:
    """파일 이름에 쓸 수 없는 글자를 뺀다."""
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip() or "분개장"


def output_path(journal_path: str, out_dir: str | None = None,
                stamp: str | None = None) -> str:
    """새로 만들 Control Sheet 경로. 같은 이름이 있으면 뒤에 번호를 붙인다."""
    journal_path = os.path.abspath(journal_path)
    folder = os.path.abspath(out_dir) if out_dir else \
        os.path.join(os.path.dirname(journal_path), OUTPUT_DIRNAME)
    stem = _safe(os.path.splitext(os.path.basename(journal_path))[0])
    stamp = stamp or _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.join(folder, f"Control Sheet_{stem}_{stamp}")
    path = base + ".xlsm"
    n = 2
    while os.path.exists(path):
        path = f"{base}({n}).xlsm"
        n += 1
    return path


def create_output(journal_path: str, out_dir: str | None = None,
                  template: str | None = None) -> tuple[str, str]:
    """양식을 복사해 새 Control Sheet를 만든다. (출력파일, 쓴 양식) 반환."""
    src = find_template(template)
    dest = output_path(journal_path, out_dir)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    try:
        shutil.copy2(src, dest)
    except OSError as exc:
        raise TemplateError(f"출력 파일을 만들지 못했습니다: {exc}") from exc
    return dest, src
