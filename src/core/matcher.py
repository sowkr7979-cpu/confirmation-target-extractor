# -*- coding: utf-8 -*-
"""정규화와 검색. 원문은 건드리지 않고 검색용 문자열만 정리한다.

검색 대상 열은 `search_fields`로 정한다. 기본값은 거래처·적요 두 열이다.
일치 조건은 두 가지다.
  1) 지정 키워드(금융·법률·정규식)
  2) 기관 사전(institutions.csv)의 기관명·별칭
둘 중 하나라도 검색 열에 들어 있어야 후보가 된다.

열마다 쓰는 검색어의 범위가 다르다(`field_scope`). 거래처는 전부 쓰고, 적요는
법률 키워드·법률 정규식·법률 분류 기관 별칭만 쓴다. 적요는 문장이라 '보증보험료',
'이자' 같은 금융 낱말이 거래처와 무관하게 자주 나오기 때문이다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .config_loader import AppConfig
from .readers import JournalRow
from .textnorm import nospace
from .textnorm import upper as normalize

FIELD_LABEL = {"counterparty": "거래처", "description": "적요", "account": "계정과목"}


@dataclass
class RowMatch:
    row: JournalRow
    keywords: list[str] = field(default_factory=list)          # 맞은 키워드
    aliases: list[str] = field(default_factory=list)           # 맞은 기관명·별칭
    fields: list[str] = field(default_factory=list)            # 맞은 열(한글 이름)
    norm: dict[str, str] = field(default_factory=dict)         # 정규화 텍스트
    field_hits: dict[str, list[str]] = field(default_factory=dict)  # 열 키 → 맞은 검색어

    def hits_in(self, field_key: str) -> list[str]:
        """특정 열(counterparty/description)에서 맞은 검색어. 사전 별칭은 '사전:' 접두."""
        return list(self.field_hits.get(field_key, []))

    @property
    def keyword_text(self) -> str:
        parts = list(self.keywords) + [f"사전:{a}" for a in self.aliases]
        return ", ".join(parts)

    @property
    def field_text(self) -> str:
        return ", ".join(self.fields)

    @property
    def terms(self) -> list[str]:
        return list(self.keywords) + list(self.aliases)


class KeywordMatcher:
    def __init__(self, config: AppConfig):
        self.config = config
        kw = config.keywords
        self.search_fields = list(kw.get("search_fields", ["counterparty", "description"]))
        scope_cfg = kw.get("field_scope", {}) or {}

        # 키워드를 그룹별로 나눠 둔다. 열마다 쓰는 그룹이 다르다.
        self.group_plain: dict[str, list[tuple[str, str]]] = {}
        for grp, key in (("finance", "finance_keywords"), ("legal", "legal_keywords"),
                         ("supplementary", "supplementary_keywords")):
            items = []
            for src in kw.get(key, []):
                if src and (src, normalize(src)) not in items:
                    items.append((src, normalize(src)))
            self.group_plain[grp] = items
        self.group_regex: dict[str, list[tuple[str, re.Pattern[str]]]] = {}
        for item in kw.get("regex_keywords", []):
            grp = str(item.get("group", "finance"))
            self.group_regex.setdefault(grp, []).append(
                (item["label"], re.compile(item["pattern"], re.IGNORECASE)))

        # 기관 사전의 기관명·별칭도 검색어로 쓴다. 분류를 함께 두어 열별로 거른다.
        self.alias_rx: list[tuple[re.Pattern[str], str, int, str]] = []
        if kw.get("match_institution_aliases", True):
            seen: set[str] = set()
            for ent in config.institutions:
                for alias in ent.aliases:
                    a = nospace(normalize(alias))
                    if not a or a in seen:
                        continue
                    seen.add(a)
                    pat = r"\s*".join(re.escape(ch) for ch in a)
                    self.alias_rx.append((re.compile(pat, re.IGNORECASE), alias, len(a),
                                          ent.category))
        self.short_min = int(kw.get("short_alias_min_len", 3))

        # 열별 검색 범위. 지정이 없는 열은 전부 쓴다.
        all_groups = ["finance", "legal", "supplementary"]
        self.scope: dict[str, dict[str, list[str] | None]] = {}
        for f in self.search_fields:
            c = scope_cfg.get(f) or {}
            self.scope[f] = {
                # keywords가 있으면 그 목록만 쓰고 keyword_groups는 무시한다.
                "keywords": ([(k, normalize(k)) for k in c["keywords"] if k]
                             if c.get("keywords") is not None else None),
                "keyword_groups": list(c.get("keyword_groups", all_groups)),
                "regex_groups": list(c.get("regex_groups", all_groups)),
                # None 이면 모든 분류의 별칭을 쓴다.
                "alias_categories": (list(c["alias_categories"])
                                     if c.get("alias_categories") is not None else None),
            }

    # 하위 호환: 예전 코드가 참조하던 속성
    @property
    def plain_src(self) -> list[str]:
        return [s for items in self.group_plain.values() for s, _ in items]

    @property
    def regexes(self) -> list[tuple[str, re.Pattern[str]]]:
        return [r for items in self.group_regex.values() for r in items]

    def _alias_hits(self, text: str, categories: list[str] | None) -> list[str]:
        """짧은 약칭이 무관한 거래처명 일부와 우연히 맞는 것을 막는다."""
        out: list[str] = []
        for rx, alias, alen, cat in self.alias_rx:
            if categories is not None and cat not in categories:
                continue
            for m in rx.finditer(text):
                if alen < self.short_min:
                    before = text[m.start() - 1] if m.start() > 0 else ""
                    after = text[m.end()] if m.end() < len(text) else ""
                    if re.match(r"[0-9A-Za-z가-힣]", before or "") or \
                       re.match(r"[A-Za-z가-힣]", after or ""):
                        continue
                if alias not in out:
                    out.append(alias)
                break
        return out

    def match_row(self, row: JournalRow) -> RowMatch | None:
        norms = {f: normalize(row.text_of(f)) for f in self.search_fields}
        kws: list[str] = []
        als: list[str] = []
        fields: list[str] = []
        field_hits: dict[str, list[str]] = {}
        for f, text in norms.items():
            if not text:
                continue
            sc = self.scope[f]
            hits: list[str] = []
            if sc["keywords"] is not None:
                plain_items = sc["keywords"]
            else:
                plain_items = [it for grp in (sc["keyword_groups"] or [])
                               for it in self.group_plain.get(grp, [])]
            for src, plain in plain_items:
                if plain and plain in text:
                    hits.append(src)
                    if src not in kws:
                        kws.append(src)
            for grp in sc["regex_groups"] or []:
                for label, rx in self.group_regex.get(grp, []):
                    if rx.search(text):
                        hits.append(label)
                        if label not in kws:
                            kws.append(label)
            for alias in self._alias_hits(text, sc["alias_categories"]):
                hits.append(f"사전:{alias}")
                if alias not in als:
                    als.append(alias)
            if hits:
                fields.append(FIELD_LABEL.get(f, f))
                field_hits[f] = hits
        if not kws and not als:
            return None
        return RowMatch(row=row, keywords=kws, aliases=als, fields=fields, norm=norms,
                        field_hits=field_hits)

    def search(self, rows: list[JournalRow]) -> list[RowMatch]:
        out: list[RowMatch] = []
        for r in rows:
            m = self.match_row(r)
            if m:
                out.append(m)
        return out
