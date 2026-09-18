# -*- coding: utf-8 -*-
"""화면 색·글자·여백 값 한 곳. apple.com/kr/store 의 표기 규칙을 따른다.

색과 크기는 이 파일에서만 정한다. 위젯 코드에 색 문자열을 직접 쓰지 않는다.
"""
from __future__ import annotations

import tkinter.font as tkfont

# --------------------------------------------------------------------- 색
WHITE = "#ffffff"
CANVAS = "#f5f5f7"          # 페이지 바탕
TEXT = "#1d1d1f"            # 본문
TEXT_SUB = "#6e6e73"        # 보조 설명
TEXT_MUTED = "#86868b"      # 더 약한 설명
BLUE = "#0071e3"            # 기본 단추
BLUE_HOVER = "#0077ed"
BLUE_PRESS = "#0068d1"
LINK = "#0066cc"
BORDER = "#d2d2d7"
FILL_SOFT = "#e8e8ed"       # 보조 단추 바탕
FILL_SOFT_HOVER = "#dededf"
GREEN = "#008009"           # 완료
RED = "#d70015"             # 실패
DISABLED_FILL = "#e8e8ed"
DISABLED_TEXT = "#aeaeb2"

# --------------------------------------------------------------------- 크기
RADIUS_CARD = 18
RADIUS_FIELD = 12
PAD_PAGE = 28
PAD_CARD = 18
GAP = 12

_FAMILY_PREF = ("Apple SD Gothic Neo", "Malgun Gothic", "맑은 고딕", "Segoe UI")
_MONO_PREF = ("D2Coding", "Consolas", "Courier New")


def _pick(prefs: tuple[str, ...], fallback: str) -> str:
    try:
        installed = set(tkfont.families())
    except Exception:  # noqa: BLE001  — Tk 미초기화 등
        return fallback
    for name in prefs:
        if name in installed:
            return name
    return fallback


class Fonts:
    """Tk 루트를 만든 뒤 한 번 생성한다."""

    def __init__(self) -> None:
        fam = _pick(_FAMILY_PREF, "Malgun Gothic")
        mono = _pick(_MONO_PREF, "Consolas")
        self.title = (fam, 24, "bold")
        self.subtitle = (fam, 11)
        self.section = (fam, 13, "bold")
        self.body = (fam, 10)
        self.body_bold = (fam, 10, "bold")
        self.caption = (fam, 9)
        self.button = (fam, 10, "bold")
        self.mono = (mono, 9)
