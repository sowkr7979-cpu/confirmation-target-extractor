# -*- coding: utf-8 -*-
"""빌드가 끝난 뒤 안내를 한글로 찍는다. 배치 파일에는 한글을 두지 않는다.

배치 파일은 cmd가 시스템 코드페이지로 읽어서 한글이 깨진다. 그래서 안내문은 여기 둔다.
"""
from __future__ import annotations

import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
EXE = os.path.join(BASE, "dist", "조회처추출.exe")

size = f"{os.path.getsize(EXE) / 1024 / 1024:.1f}MB" if os.path.exists(EXE) else "없음"
print()
print("=" * 62)
print(f" 빌드 완료: dist\\조회처추출.exe ({size})")
print("=" * 62)
print(" - dist 폴더를 통째로 건네면 Python 없이 실행됩니다.")
print(" - config는 exe 옆 폴더를 먼저 읽습니다. 기관 사전을 고칠 때")
print("   exe를 다시 만들지 않아도 됩니다.")
print(" - 받는 PC에도 Microsoft Excel(데스크톱)이 있어야 저장이 됩니다.")
print("=" * 62)
