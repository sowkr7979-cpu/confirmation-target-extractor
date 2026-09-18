# -*- coding: utf-8 -*-
"""문자열 정규화. 검색·대조에만 쓰고 원문은 바꾸지 않는다.

세 단계만 있다.
  clean : 전각→반각(NFKC), ㈜→(주), 공백 정리. 대소문자는 그대로.
  upper : clean + 대문자 — 키워드 검색용.
  key   : clean + 공백 제거 + 대문자 — 이름 대조·통합용.
"""
from __future__ import annotations

import re
import unicodedata

_SPACE = re.compile(r"\s+")


def clean(text: object) -> str:
    """표시용 정규화. 빈 값은 빈 문자열."""
    if not text:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    s = s.replace("（", "(").replace("）", ")").replace("㈜", "(주)")
    return _SPACE.sub(" ", s).strip()


def upper(text: object) -> str:
    """검색용 정규화. 공백은 한 칸으로 남긴다."""
    return clean(text).upper()


def key(text: object) -> str:
    """대조용 키. 공백을 모두 지우고 대문자로."""
    return _SPACE.sub("", clean(text)).upper()


def nospace(text: str) -> str:
    return _SPACE.sub("", text)
