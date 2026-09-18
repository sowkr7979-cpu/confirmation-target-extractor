# -*- coding: utf-8 -*-
"""설정 파일(키워드/기관사전/입력프로파일) 로더."""
from __future__ import annotations

import csv
import json
import os
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from .textnorm import key as _norm_key


def _mtime(path: str) -> str:
    try:
        return datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
    except OSError:
        return "알수없음"


@dataclass
class InstitutionEntry:
    institution_id: str
    display_name: str
    category: str
    aliases: list[str]
    note: str = ""


@dataclass
class AppConfig:
    config_dir: str
    keywords: dict[str, Any]
    institutions: list[InstitutionEntry]
    profiles: dict[str, Any]
    versions: dict[str, str] = field(default_factory=dict)
    online: dict[str, str] = field(default_factory=dict)
    """온라인 발급 가능 기관. 키는 institution_id 또는 정규화한 이름, 값은 근거 메모."""

    @property
    def all_keywords(self) -> list[str]:
        out: list[str] = []
        for key in ("finance_keywords", "legal_keywords", "supplementary_keywords"):
            for kw in self.keywords.get(key, []):
                if kw and kw not in out:
                    out.append(kw)
        return out

    def keyword_group(self, kw: str) -> str:
        if kw in self.keywords.get("finance_keywords", []):
            return "금융"
        if kw in self.keywords.get("legal_keywords", []):
            return "법률"
        if kw in self.keywords.get("supplementary_keywords", []):
            return "보완"
        return "패턴"


def load_config(config_dir: str) -> AppConfig:
    kw_path = os.path.join(config_dir, "keywords.json")
    inst_path = os.path.join(config_dir, "institutions.csv")
    prof_path = os.path.join(config_dir, "input_profiles.json")

    for p in (kw_path, inst_path, prof_path):
        if not os.path.exists(p):
            raise FileNotFoundError(f"설정 파일이 없습니다: {p}")

    with open(kw_path, encoding="utf-8") as f:
        keywords = json.load(f)
    with open(prof_path, encoding="utf-8") as f:
        profiles = json.load(f)

    institutions: list[InstitutionEntry] = []
    seen_ids: set[str] = set()
    with open(inst_path, encoding="utf-8-sig", newline="") as f:
        for lineno, row in enumerate(csv.DictReader(f), start=2):
            iid = (row.get("institution_id") or "").strip()
            name = (row.get("display_name") or "").strip()
            if not iid or not name:
                continue
            if iid in seen_ids:
                raise ValueError(f"institutions.csv {lineno}행: institution_id 중복 '{iid}'")
            seen_ids.add(iid)
            raw_alias = (row.get("aliases") or "").strip()
            aliases = [a.strip() for a in raw_alias.split("|") if a.strip()]
            if name not in aliases:
                aliases.append(name)
            institutions.append(
                InstitutionEntry(
                    institution_id=iid,
                    display_name=name,
                    category=(row.get("category") or "").strip(),
                    aliases=aliases,
                    note=(row.get("note") or "").strip(),
                )
            )

    online: dict[str, str] = {}
    online_path = os.path.join(config_dir, "online_institutions.csv")
    if os.path.exists(online_path):
        with open(online_path, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                flag = (row.get("online") or "").strip().upper()
                if flag != "Y":
                    continue
                note = (row.get("source_note") or "").strip()
                for col in ("institution_id", "display_name", "aliases"):
                    val = (row.get(col) or "").strip()
                    if not val:
                        continue
                    for one in val.split("|"):
                        k = _norm_key(one)
                        if k:
                            online.setdefault(k, note)

    versions = {
        "keywords.json": f"{keywords.get('version', '-')} (수정 {_mtime(kw_path)})",
        "institutions.csv": f"{len(institutions)}건 (수정 {_mtime(inst_path)})",
        "input_profiles.json": f"{profiles.get('version', '-')} (수정 {_mtime(prof_path)})",
        "online_institutions.csv": (f"{len(online)}키 (수정 {_mtime(online_path)})"
                                    if os.path.exists(online_path) else "없음"),
    }
    return AppConfig(config_dir, keywords, institutions, profiles, versions, online)
