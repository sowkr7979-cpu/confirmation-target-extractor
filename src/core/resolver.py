# -*- coding: utf-8 -*-
"""기관명 추출 · 별칭 통합 · 원문 거래처 유지 · 중복 제거 · 근거 구성.

원칙
- 사전·명칭패턴·접미사로 기관명을 확정할 수 있으면 그 이름으로 통합한다.
- 확정할 수 없어도 **원문 거래처명이 있으면 그 이름을 그대로 후보로 쓴다.**
  사전 미등록·유형 불명확을 이유로 제외하지 않는다.
- 거래처가 비어 있으면 적요 → 전표 문맥 순으로 찾고, 그래도 없으면 '이름미확인'.
- 제외는 사용자가 exclude_names에 명시한 이름에만, 정규화 완전 일치로 적용한다.
"""
from __future__ import annotations

import re
from collections import OrderedDict
from dataclasses import dataclass, field

from .config_loader import AppConfig
from .matcher import RowMatch
from .textnorm import clean as norm_display
from .textnorm import key as _key
from .textnorm import upper as norm_upper

_ACCOUNT_CLUE = re.compile(r"(?<!\d)(\d[\d\-]{2,}\d)(?!\d)")

FIELD_LABEL = {"counterparty": "거래처", "description": "적요"}

# 후보 확정 방식
BY_DICT = "사전"
BY_LEGAL = "명칭패턴"
BY_SUFFIX = "접미사"
BY_LEGAL_SUFFIX = "법률접미사"
BY_RAW = "원문거래처"

# 이름을 '얻은 방식'만 나타낸다. 실제 금융기관·변호사인지, 발송 대상인지는 판정하지 않는다.
METHOD_LABEL = {
    "사전": "기관사전 별칭 일치(사전 등재 사실만 확인)",
    "명칭패턴": "법률기관 명칭패턴 일치(기관 여부 미확인)",
    "접미사": "거래처 접미사 규칙으로 잘라낸 이름(기관 여부 미확인)",
    "법률접미사": "법률기관 접미사 규칙으로 잘라낸 이름(기관 여부 미확인)",
    "원문거래처": "원문 거래처명 그대로 유지(기관 여부 미확인)",
}


@dataclass
class Mention:
    name: str                 # A열에 쓸 이름 (원문 후보면 원문 그대로)
    key: str                  # 통합 키
    institution_id: str       # 사전 확정 시 ID, 아니면 ""
    category: str
    matched_by: str
    raw_name: str             # 이름을 뽑아낸 원문 필드 전체
    field: str                # '거래처' | '적요'
    leftover: str             # 기관명을 뺀 나머지 원문 (원문 후보면 "")
    alias: str = ""           # 사전에서 실제로 맞은 별칭

    @property
    def method_label(self) -> str:
        """이름을 '어떻게 얻었는지'만 말한다. 기관 여부를 확정하지 않는다."""
        return METHOD_LABEL.get(self.matched_by, self.matched_by)


@dataclass
class Institution:
    key: str
    display_name: str
    institution_id: str
    category: str
    matched_by: str
    order: int
    raw_names: "OrderedDict[str, None]" = field(default_factory=OrderedDict)
    branch_clues: "OrderedDict[str, None]" = field(default_factory=OrderedDict)
    account_clues: "OrderedDict[str, None]" = field(default_factory=OrderedDict)
    keywords: "OrderedDict[str, None]" = field(default_factory=OrderedDict)
    desc_hits: "OrderedDict[str, None]" = field(default_factory=OrderedDict)
    """연결된 거래행의 '적요'에서 맞은 검색어(법률 낱말·법률 기관 별칭)."""
    accounts: "OrderedDict[str, None]" = field(default_factory=OrderedDict)
    vouchers: set[str] = field(default_factory=set)
    direct_rows: int = 0       # 이 후보에 직접 연결된 고유 원본 행 수
    context_rows: int = 0      # 전표문맥으로 연결된 고유 원본 행 수
    debit_sum: float = 0.0
    credit_sum: float = 0.0
    date_min: object = None
    date_max: object = None
    merge_note: str = ""
    layer: str = ""            # A(사전 일치) / B(금융·법률 명칭 형태) / C(그 밖)
    layer_reason: str = ""     # 등급 판정 근거
    review: bool = False       # C등급 중 먼저 볼 것
    online: bool = False       # 금융결제원 온라인 발급 가능 기관인지
    online_note: str = ""      # 온라인 판정 근거


@dataclass
class Evidence:
    kind: str                # 직접일치 / 전표문맥 / 이름미확인 / 제외
    institution_key: str
    institution_name: str
    sheet: str
    excel_row: int
    voucher_id: str
    date: object
    account: str
    counterparty: str
    description: str
    debit: float
    credit: float
    keywords: str
    hit_fields: str
    name_field: str
    branch_clue: str
    account_clue: str
    note: str
    desc_fragment: str = ""   # 적요에서 접미사에 걸렸지만 기관명으로 채택하지 않은 조각

    @property
    def row_id(self) -> tuple[str, int]:
        return (self.sheet, self.excel_row)


@dataclass
class ResolveResult:
    institutions: list[Institution]
    evidence: list[Evidence]
    row_stats: dict[str, int]    # 중복 없는 '원본 행' 단위
    link_stats: dict[str, int]   # '후보-근거 연결' 단위 (한 행이 여러 건일 수 있음)
    output_layers: frozenset[str] = frozenset({"A", "B"})

    @property
    def output_institutions(self) -> list["Institution"]:
        """A열에 쓸 후보(기본 A·B등급)."""
        return [i for i in self.institutions if i.layer in self.output_layers]

    @property
    def skipped_institutions(self) -> list["Institution"]:
        """A열에서 뺀 후보(기본 C등급). 제외후보 시트에만 남긴다."""
        return [i for i in self.institutions if i.layer not in self.output_layers]

class InstitutionResolver:
    def __init__(self, config: AppConfig):
        self.config = config
        kw = config.keywords
        self.extract_fields = kw.get("extract_fields", ["counterparty", "description"])
        self.exclude = {_key(x) for x in kw.get("exclude_names", []) if str(x).strip()}
        self.exclude_src = {_key(x): str(x) for x in kw.get("exclude_names", []) if str(x).strip()}
        self.standalone = set(kw.get("standalone_suffixes", []))
        max_pre = int(kw.get("suffix_prefix_max_len", 12))

        def _suffix_rx(items):
            items = sorted({s for s in items if s}, key=len, reverse=True)
            if not items:
                return None
            return re.compile(
                r"(?P<pre>[0-9A-Za-z가-힣]{0,%d}?)(?P<suf>%s)"
                % (max_pre, "|".join(re.escape(s) for s in items)))

        all_suffixes = kw.get("institution_suffixes", [])
        self.legal_suffixes = set(kw.get("legal_suffixes", []))
        self.suffix_rx = _suffix_rx(all_suffixes)
        # 적요 전용: 법률기관 접미사만. 금융 접미사는 적요 문장에서 쓰지 않는다.
        self.legal_suffix_rx = _suffix_rx(self.legal_suffixes & set(all_suffixes)
                                          or self.legal_suffixes)
        modes = kw.get("extract_modes", {})
        self.modes = {
            "counterparty": list(modes.get("counterparty",
                                           ["사전", "명칭패턴", "접미사", "원문유지"])),
            "description": list(modes.get("description", ["사전", "명칭패턴", "법률접미사"])),
        }
        self.use_voucher_context = bool(kw.get("use_voucher_context", False))
        self.layer_b_terms = [t for t in kw.get("layer_b_terms", []) if t]
        # 영문 낱말은 앞뒤에 영문자가 없을 때만 맞춘다('REFUND'의 FUND, 'ENTRUST'의 TRUST 제외).
        self.layer_b_rx = [(t, self._term_rx(t)) for t in self.layer_b_terms]
        scope_cfg = kw.get("field_scope", {}) or {}
        desc_scope = scope_cfg.get("description") or {}
        self.desc_alias_categories = (set(desc_scope["alias_categories"])
                                      if desc_scope.get("alias_categories") is not None
                                      else None)
        self.layer_c_review_terms = [t for t in kw.get("layer_c_review_terms", []) if t]
        self.output_layers = set(kw.get("output_layers", ["A", "B"]))
        self.exclude_categories = {str(c).strip() for c in kw.get("exclude_categories", [])
                                   if str(c).strip()}
        self.online_map = dict(getattr(config, "online", {}) or {})
        self.online_label = kw.get("online_label", "온라인")
        self.legal_rx = [re.compile(p) for p in kw.get("legal_entity_patterns", [])]
        self.reject_rx = [re.compile(p) for p in kw.get("reject_name_patterns", []) if p]

        self.alias_rx: list[tuple[re.Pattern[str], str, str, str, str]] = []
        self.alias_exact: dict[str, tuple[str, str, str, str]] = {}
        for ent in config.institutions:
            for alias in ent.aliases:
                a = norm_display(alias)
                if not a:
                    continue
                pat = r"\s*".join(re.escape(ch) for ch in re.sub(r"\s+", "", a))
                self.alias_rx.append(
                    (re.compile(pat, re.IGNORECASE), ent.institution_id,
                     ent.display_name, ent.category, a))
                self.alias_exact.setdefault(
                    _key(a), (ent.institution_id, ent.display_name, ent.category, a))
        self.short_min = int(kw.get("short_alias_min_len", 3))

    @staticmethod
    def _term_rx(term: str) -> "re.Pattern[str] | None":
        """B등급 낱말 정규식. 대상 문자열은 공백을 한 칸으로 남긴 대문자(norm_upper).

        한글 낱말은 부분 일치 그대로(공백은 있어도 없어도 맞게).
        영문 낱말은 앞에 영문자가 없고, 뒤에는 복수형·-ING·-IAL 정도만 허용한다.
        'ABC BANK LTD'·'BANKING'은 맞고 'EMBANKMENT'·'REFUND'·'ENTRUST'는 안 맞는다.
        """
        t = re.sub(r"\s+", " ", str(term or "")).strip().upper()
        if not t:
            return None
        if re.fullmatch(r"[A-Z0-9&.\- ]+", t):
            body = r"\s*".join(re.escape(w) for w in t.split(" "))
            return re.compile(r"(?<![A-Z])" + body + r"(?:S|ES|ING|IAL)?(?![A-Z])")
        body = r"\s*".join(re.escape(ch) for ch in t.replace(" ", ""))
        return re.compile(body)

    # ------------------------------------------------------------ 제외 판정
    def exclusion_reason(self, name: str, institution_id: str = "", alias: str = "") -> str:
        """사용자가 명시한 제외 이름에만, 정규화 완전 일치로 적용한다.

        사전 등록 기관과 미등록 원문 거래처에 동일하게 적용된다.
        반환값은 적용된 규칙 설명이며, 제외 대상이 아니면 빈 문자열.
        """
        for label, value in (("조회처명", name), ("기관ID", institution_id), ("별칭", alias)):
            if not value:
                continue
            k = _key(value)
            if k in self.exclude:
                return f"exclude_names 완전일치({label}='{value}' = '{self.exclude_src[k]}')"
        for rx in self.reject_rx:
            if rx.search(name):
                return f"reject_name_patterns 일치(pattern='{rx.pattern}')"
        return ""

    # ------------------------------------------------------------------ 온라인
    def classify_online(self, inst: "Institution") -> None:
        """금융결제원 온라인 발급 가능 기관인지 표시한다.

        config/online_institutions.csv의 institution_id·기관명·별칭 중
        하나라도 정규화 완전 일치하면 온라인으로 본다. 추측하지 않는다.
        """
        cands = [inst.institution_id, inst.display_name]
        for c in cands:
            if not c:
                continue
            note = self.online_map.get(_key(c))
            if note is not None:
                inst.online = True
                inst.online_note = note or "online_institutions.csv 일치"
                return
        inst.online = False
        inst.online_note = ""

    # ------------------------------------------------------------------ 등급
    def classify_layer(self, inst: "Institution") -> None:
        """거래처명이 금융·법률 기관 명칭 형태인지로 등급을 매긴다.

        A: 기관 사전의 기관명·별칭과 일치 (institution_id 있음)
        B: 사전에는 없지만 거래처명에 금융·법률 기관 낱말이 들어 있음
        C: 그 밖 — 일반 거래처. A열에 출력하지 않고 제외후보 시트에만 남긴다.

        등급은 '이름을 어떻게 얻었는가'가 아니라 '거래처명이 어떤 형태인가'로 정한다.
        어느 등급도 실제 금융기관·변호사라는 확정이 아니다.
        """
        if inst.category and inst.category in self.exclude_categories:
            inst.layer = "C"
            inst.review = False
            inst.layer_reason = (f"제외 분류 '{inst.category}' — exclude_categories 설정으로 "
                                 "A열에서 뺌(기관 자체를 부정하는 것이 아님)")
            return
        if inst.institution_id:
            inst.layer = "A"
            inst.layer_reason = f"기관사전 일치(institution_id={inst.institution_id})"
            return
        hay = norm_upper(inst.display_name)
        hit = [t for t, rx in self.layer_b_rx if rx is not None and rx.search(hay)]
        if hit:
            inst.layer = "B"
            inst.layer_reason = "거래처명에 금융·법률 기관 낱말: " + ", ".join(hit[:3])
            return
        if inst.desc_hits:
            # 거래처명만 보면 일반 거래처지만, 그 행의 적요에 법률 낱말이 있다.
            # 변호사 개인·법무법인 약칭 등 이름에 낱말이 없는 조회처를 놓치지 않기 위해 A열에 넣는다.
            inst.layer = "B"
            inst.layer_reason = ("적요에 법률 낱말: "
                                 + ", ".join(list(inst.desc_hits)[:3]))
            return
        inst.layer = "C"
        hay_key = _key(inst.display_name)
        rev = [t for t in self.layer_c_review_terms if _key(t) in hay_key]
        inst.review = bool(rev)
        inst.layer_reason = ("일반 거래처 — 금융·법률 기관 낱말 없음"
                             + (f" / 검토 필요 낱말: {', '.join(rev[:3])}" if rev else ""))

    # ------------------------------------------------------------------ 추출
    def extract_named(self, text: str, field_label: str,
                      modes: list[str] | None = None,
                      fragments: list[str] | None = None) -> list[Mention]:
        """사전·명칭패턴·접미사로 기관명을 뽑는다. 원문 대체(fallback)는 하지 않는다.

        modes로 어떤 방식을 쓸지 정한다. fragments 리스트를 주면, 금융 접미사에는
        걸렸지만 이 필드에서 채택하지 않기로 한 조각을 거기에 담는다.
        """
        disp = norm_display(text)
        if not disp:
            return []
        if modes is None:
            modes = self.modes.get(
                "counterparty" if field_label == FIELD_LABEL["counterparty"] else "description",
                ["사전", "명칭패턴", "접미사"])

        cands: list[tuple[int, int, str, str, str, str, str]] = []  # s,e,id,name,cat,by,alias

        is_desc = field_label == FIELD_LABEL["description"]
        if "사전" in modes:
            for rx, iid, name, cat, alias in self.alias_rx:
                if (is_desc and self.desc_alias_categories is not None
                        and cat not in self.desc_alias_categories):
                    continue
                for m in rx.finditer(disp):
                    if len(re.sub(r"\s+", "", alias)) < self.short_min:
                        before = disp[m.start() - 1] if m.start() > 0 else ""
                        after = disp[m.end()] if m.end() < len(disp) else ""
                        if re.match(r"[0-9A-Za-z가-힣]", before or "") or \
                           re.match(r"[A-Za-z가-힣]", after or ""):
                            continue
                    cands.append((m.start(), m.end(), iid, name, cat, BY_DICT, alias))

        if "명칭패턴" in modes:
            for rx in self.legal_rx:
                for m in rx.finditer(disp):
                    whole = re.sub(r"\s+", " ", m.group(0)).strip()
                    cands.append((m.start(), m.end(), "", whole, "법률", BY_LEGAL, ""))

        use_all = "접미사" in modes
        use_legal_only = (not use_all) and "법률접미사" in modes
        rx_sfx = self.suffix_rx if use_all else (self.legal_suffix_rx if use_legal_only else None)
        if rx_sfx is not None:
            by = BY_SUFFIX if use_all else BY_LEGAL_SUFFIX
            for m in rx_sfx.finditer(disp):
                pre, suf = m.group("pre"), m.group("suf")
                if not pre and suf not in self.standalone:
                    continue
                whole = (pre + suf).strip()
                if len(whole) < 2:
                    continue
                cands.append((m.start(), m.end(), "", whole, "", by, ""))

        # 이 필드에서 금융 접미사를 쓰지 않기로 했다면, 걸린 조각만 근거용으로 모은다.
        if fragments is not None and not use_all and self.suffix_rx is not None:
            for m in self.suffix_rx.finditer(disp):
                pre, suf = m.group("pre"), m.group("suf")
                if not pre and suf not in self.standalone:
                    continue
                whole = (pre + suf).strip()
                if len(whole) < 2 or suf in self.legal_suffixes:
                    continue
                if _key(whole) in self.alias_exact:
                    continue      # 사전에 있는 이름이면 위에서 이미 후보로 잡혔다
                if whole not in fragments:
                    fragments.append(whole)

        # 접미사·패턴으로 뽑은 이름이 사전 별칭과 완전히 같으면 사전 기관으로 통합
        fixed = []
        for s, e, iid, name, cat, by, alias in cands:
            if not iid:
                hit = self.alias_exact.get(_key(name))
                if hit:
                    iid, name, cat, alias, by = hit[0], hit[1], hit[2], hit[3], BY_DICT
            fixed.append((s, e, iid, name, cat, by, alias))

        fixed.sort(key=lambda c: (-(c[1] - c[0]), 0 if c[5] == BY_DICT else 1, c[0]))
        taken: list[tuple[int, int, str, str, str, str, str]] = []
        for c in fixed:
            if any(not (c[1] <= t[0] or c[0] >= t[1]) for t in taken):
                continue
            taken.append(c)
        taken.sort(key=lambda c: c[0])

        leftover = disp
        for s, e in sorted(((t[0], t[1]) for t in taken), reverse=True):
            leftover = leftover[:s] + " " + leftover[e:]
        leftover = re.sub(r"\s+", " ", leftover).strip(" /,-")

        return [
            Mention(name=name, key=(iid if iid else "원문:" + _key(name)),
                    institution_id=iid, category=cat, matched_by=by,
                    raw_name=disp, field=field_label, leftover=leftover, alias=alias)
            for _s, _e, iid, name, cat, by, alias in taken
        ]

    def row_mentions(self, row) -> tuple[list[Mention], list[str]]:
        """한 거래행에서 후보 이름을 뽑는다.

        1) 거래처에서 기관명을 찾는다. 못 찾았고 거래처가 비어 있지 않으면
           **원문 거래처명 전체**를 후보로 쓴다.
        2) 적요에서 찾은 기관명은 별도 후보로 함께 연결한다.
        3) 거래처가 비어 있고 적요에서도 기관명을 못 찾으면, 적요 원문은 만들지 않는다
           (적요는 문장이라 조회처명이 아니다). 호출부가 전표 문맥/이름미확인으로 처리한다.

        반환값은 (후보 목록, 적요 미채택 조각 목록)이다. 미채택 조각은 근거에만 남는다.
        """
        out: "OrderedDict[str, Mention]" = OrderedDict()
        fragments: list[str] = []

        cp_raw = (row.counterparty or "").strip()
        cp_named = self.extract_named(cp_raw, FIELD_LABEL["counterparty"]) if cp_raw else []
        for m in cp_named:
            out.setdefault(m.key, m)

        if cp_raw and not cp_named and "원문유지" in self.modes["counterparty"]:
            # 기관으로 확정할 수 없는 거래처 — 원문 이름을 그대로 유지한다.
            out.setdefault("원문:" + _key(cp_raw), Mention(
                name=cp_raw, key="원문:" + _key(cp_raw), institution_id="", category="",
                matched_by=BY_RAW, raw_name=norm_display(cp_raw),
                field=FIELD_LABEL["counterparty"], leftover="", alias=""))

        if "description" in self.extract_fields and self.modes.get("description"):
            for m in self.extract_named(row.description or "", FIELD_LABEL["description"],
                                        fragments=fragments):
                out.setdefault(m.key, m)

        # 이미 후보가 된 이름과 겹치는 조각은 근거에 남길 필요가 없다.
        taken = {_key(m.name) for m in out.values()}
        fragments = [f for f in fragments if _key(f) not in taken]
        return list(out.values()), fragments

    @staticmethod
    def _clues(leftover: str) -> tuple[list[str], list[str]]:
        accounts = [m.group(1) for m in _ACCOUNT_CLUE.finditer(leftover)]
        rest = _ACCOUNT_CLUE.sub(" ", leftover)
        branches = [t.strip(" ()/-,") for t in re.split(r"[/,()]", rest)]
        branches = [b for b in branches if b and not b.isdigit() and len(b) >= 2]
        return branches, accounts

    # ------------------------------------------------------------------ 통합
    def resolve(self, matches: list[RowMatch]) -> ResolveResult:
        insts: "OrderedDict[str, Institution]" = OrderedDict()
        evidence: list[Evidence] = []
        pending: list[RowMatch] = []                      # 후보를 못 만든 행
        voucher_keys: dict[str, "OrderedDict[str, None]"] = {}

        row_stats = {"검색행": len(matches), "직접일치": 0, "전표문맥": 0,
                     "이름미확인": 0, "전부제외": 0}
        link_stats = {"직접일치": 0, "전표문맥": 0, "이름미확인": 0, "제외": 0}

        for rm in matches:
            row = rm.row
            kept: "OrderedDict[str, tuple[Mention, list[str]]]" = OrderedDict()
            dropped: list[tuple[Mention, str]] = []

            mentions, fragments = self.row_mentions(row)
            frag_text = " | ".join(fragments)
            for men in mentions:
                reason = self.exclusion_reason(men.name, men.institution_id, men.alias)
                if reason:
                    dropped.append((men, reason))
                    continue
                if men.key in kept:
                    if men.field not in kept[men.key][1]:
                        kept[men.key][1].append(men.field)
                else:
                    kept[men.key] = (men, [men.field])

            for men, reason in dropped:
                link_stats["제외"] += 1
                evidence.append(self._ev(
                    "제외", "", men.name, rm, men.field, "", "",
                    f"원문='{men.raw_name}' / 적용규칙: {reason}", frag_text))

            if not kept:
                if dropped:
                    row_stats["전부제외"] += 1
                else:
                    pending.append(rm)
                continue

            row_stats["직접일치"] += 1
            vk = voucher_keys.setdefault(row.voucher_id, OrderedDict())
            for k, (men, labels) in kept.items():
                inst = insts.get(k)
                if inst is None:
                    inst = Institution(
                        key=k, display_name=men.name, institution_id=men.institution_id,
                        category=men.category, matched_by=men.matched_by, order=len(insts) + 1,
                        merge_note=men.method_label,
                    )
                    insts[k] = inst
                inst.raw_names[men.raw_name] = None
                # 일치 키워드와 사전 별칭을 모두 기록한다.
                # 별칭만으로 포함된 기관(예: 농협·기업비씨)도 근거가 남아야 한다.
                inst.keywords.update({kw: None for kw in rm.keywords})
                inst.keywords.update({f"사전:{a}": None for a in rm.aliases})
                inst.desc_hits.update({h: None for h in rm.hits_in("description")})
                if row.account:
                    inst.accounts[row.account] = None
                inst.vouchers.add(row.voucher_id)
                inst.direct_rows += 1
                inst.debit_sum += row.debit
                inst.credit_sum += row.credit
                if row.date:
                    inst.date_min = row.date if inst.date_min is None or row.date < inst.date_min else inst.date_min
                    inst.date_max = row.date if inst.date_max is None or row.date > inst.date_max else inst.date_max
                branches, accounts = self._clues(men.leftover)
                for b in branches:
                    inst.branch_clues[b] = None
                for a in accounts:
                    inst.account_clues[a] = None
                vk[k] = None
                link_stats["직접일치"] += 1
                evidence.append(self._ev(
                    "직접일치", k, inst.display_name, rm, ", ".join(labels),
                    "; ".join(branches), "; ".join(accounts),
                    f"이름 추출 방식: {men.method_label}", frag_text))

        # 후보를 못 만든 행 → 같은 전표의 확정 후보에 '문맥'으로만 연결(신규 후보 생성 없음)
        for rm in pending:
            keys = voucher_keys.get(rm.row.voucher_id) if self.use_voucher_context else None
            if keys:
                row_stats["전표문맥"] += 1
                for k in keys:
                    insts[k].context_rows += 1
                    link_stats["전표문맥"] += 1
                    evidence.append(self._ev(
                        "전표문맥", k, insts[k].display_name, rm, "", "", "",
                        "거래처 공란·적요에서 이름 없음. 같은 전표에 있을 뿐 기관 관계는 확인되지 않음 — A열 신규 생성 없음"))
            else:
                row_stats["이름미확인"] += 1
                link_stats["이름미확인"] += 1
                evidence.append(self._ev(
                    "이름미확인", "", "", rm, "", "", "",
                    ("거래처에서 이름을 얻지 못함 — A열에 임의 명칭을 만들지 않음. "
                     "적요·계정과목·전표문맥으로 구제하지 않는다"
                     if not self.use_voucher_context else
                     "거래처가 공란이고 적요·전표문맥에서도 이름을 찾지 못함 — A열에 임의 명칭을 만들지 않음")))

        out = list(insts.values())
        for inst in out:
            self.classify_layer(inst)
            self.classify_online(inst)
        link_stats["근거레코드_계"] = len(evidence)
        row_stats["근거연결_고유행"] = len({e.row_id for e in evidence})
        for lv in ("A", "B", "C"):
            row_stats[f"등급{lv}_후보"] = sum(1 for i in out if i.layer == lv)
        row_stats["등급C_검토필요"] = sum(1 for i in out if i.layer == "C" and i.review)
        row_stats["온라인발급_후보"] = sum(
            1 for i in out if i.online and i.layer in self.output_layers)
        return ResolveResult(out, evidence, row_stats, link_stats,
                             frozenset(self.output_layers))

    @staticmethod
    def _ev(kind, key, name, rm: RowMatch, name_field, branch, acct, note,
            desc_fragment: str = "") -> Evidence:
        r = rm.row
        return Evidence(
            kind=kind, institution_key=key, institution_name=name,
            sheet=r.sheet, excel_row=r.excel_row, voucher_id=r.voucher_id,
            date=r.date, account=r.account, counterparty=r.counterparty,
            description=r.description, debit=r.debit, credit=r.credit,
            keywords=rm.keyword_text, hit_fields=rm.field_text,
            name_field=name_field, branch_clue=branch, account_clue=acct, note=note,
            desc_fragment=desc_fragment,
        )
