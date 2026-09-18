# -*- coding: utf-8 -*-
"""실행 경로. 스크립트로 돌 때와 exe로 돌 때의 폴더가 다르다.

- app_dir()    : 사용자가 보는 폴더. exe가 놓인 자리(또는 프로젝트 최상위).
- config_dir() : 설정 폴더. exe 옆의 config를 먼저 쓰고, 없으면 exe 안에 넣어 둔 사본을 쓴다.
  옆에 두면 사용자가 institutions.csv·keywords.json을 직접 고칠 수 있다.
"""
from __future__ import annotations

import os
import sys

_SRC = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PROJECT = os.path.dirname(_SRC)


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_dir() -> str:
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return _PROJECT


def bundle_dir() -> str:
    """exe 안에 포함된 자원이 풀리는 임시 폴더. 스크립트일 때는 프로젝트 최상위."""
    if is_frozen():
        return getattr(sys, "_MEIPASS", app_dir())
    return _PROJECT


def config_dir() -> str:
    external = os.path.join(app_dir(), "config")
    if os.path.isdir(external):
        return external
    return os.path.join(bundle_dir(), "config")
