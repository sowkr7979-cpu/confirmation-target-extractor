# -*- coding: utf-8 -*-
"""수집한 업종별 CSV와 기존 사전을 병합해 config/institutions.csv를 만든다.

- display_name 정규화 키로 중복을 합친다(별칭은 합집합, 먼저 나온 ID·note 우선).
- 별칭이 다른 기관의 display_name과 충돌하면 그 별칭을 버리고 기록한다.
- 짧은 별칭(2글자 이하)은 버린다. 무관한 거래처와 오탐이 난다.
"""
import csv
import os
import re
import sys
import unicodedata
from collections import OrderedDict

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.stdout.reconfigure(encoding="utf-8")

DB_DIR = os.path.join(BASE, "tools", "db")
OUT = os.path.join(BASE, "config", "institutions.csv")
KEEP = os.path.join(DB_DIR, "00_기존유지.csv")

SOURCES = ["00_기존유지.csv", "01_은행_저축은행.csv", "02_보험_보증_공제.csv",
           "03_증권_운용_신탁_선물.csv", "04_여신전문_전자금융.csv",
           # 2026-09-18 전수 수집분. 앞 파일과 같은 이름이면 ID·note는 앞 파일 것을 쓰고 별칭만 합친다.
           "05_자산운용_자문일임_전수.csv", "06_여신_신기술_전자금융_P2P_전수.csv",
           "07_대부_전수.csv", "08_은행_보험_증권_보증_보완.csv"]


def key(s: str) -> str:
    s = unicodedata.normalize("NFKC", str(s or ""))
    s = s.replace("（", "(").replace("）", ")").replace("㈜", "(주)")
    return re.sub(r"\s+", "", s).upper()


rows = OrderedDict()          # key(display_name) -> dict
id_seen = {}
dropped_short = []
dropped_conflict = []
merged_dup = []

for src in SOURCES:
    p = os.path.join(DB_DIR, src)
    if not os.path.exists(p):
        print(f"  (없음) {src}")
        continue
    n = 0
    with open(p, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            iid = (row.get("institution_id") or "").strip()
            name = (row.get("display_name") or "").strip()
            if not iid or not name:
                continue
            n += 1
            al = [a.strip() for a in (row.get("aliases") or "").split("|") if a.strip()]
            if name not in al:
                al.insert(0, name)
            k = key(name)
            if k in rows:
                tgt = rows[k]
                add = [a for a in al if key(a) not in {key(x) for x in tgt["aliases"]}]
                tgt["aliases"].extend(add)
                merged_dup.append((name, tgt["institution_id"], iid))
                continue
            if iid in id_seen:
                iid = iid + "_2"
            id_seen[iid] = name
            rows[k] = {"institution_id": iid, "display_name": name,
                       "category": (row.get("category") or "").strip(),
                       "aliases": al, "note": (row.get("note") or "").strip()}
    print(f"  {src}: {n}행 읽음")

# 별칭 정리: 짧은 것 제거, 다른 기관의 정식명칭과 충돌하면 제거
names = {key(r["display_name"]): r["institution_id"] for r in rows.values()}
for r in rows.values():
    keep = []
    for a in r["aliases"]:
        ka = key(a)
        if len(ka) <= 2:
            dropped_short.append((r["display_name"], a))
            continue
        if ka in names and names[ka] != r["institution_id"]:
            dropped_conflict.append((r["display_name"], a, names[ka]))
            continue
        if ka not in {key(x) for x in keep}:
            keep.append(a)
    r["aliases"] = keep

# 같은 별칭이 둘 이상의 기관에 붙어 있으면 전부에서 제거(어느 쪽인지 알 수 없다)
owner = {}
for r in rows.values():
    for a in r["aliases"]:
        owner.setdefault(key(a), []).append(r["institution_id"])
ambiguous = {k for k, v in owner.items() if len(set(v)) > 1}
for r in rows.values():
    before = list(r["aliases"])
    r["aliases"] = [a for a in r["aliases"]
                    if key(a) not in ambiguous or key(a) == key(r["display_name"])]
    for a in before:
        if a not in r["aliases"]:
            dropped_conflict.append((r["display_name"], a, "중복 별칭"))

with open(OUT, "w", encoding="utf-8", newline="") as f:
    w = csv.writer(f)
    w.writerow(["institution_id", "display_name", "category", "aliases", "note"])
    for r in rows.values():
        w.writerow([r["institution_id"], r["display_name"], r["category"],
                    "|".join(r["aliases"]), r["note"]])

import collections  # noqa: E402
print(f"\n최종 {len(rows)}건 → {OUT}")
print("분류별:", dict(collections.Counter(r["category"] for r in rows.values())))
print(f"\n같은 이름이라 합친 것 {len(merged_dup)}건:")
for name, a, b in merged_dup[:15]:
    print(f"   {name}: {b} → {a}")
print(f"\n짧아서 버린 별칭 {len(dropped_short)}건: {dropped_short[:10]}")
print(f"\n충돌로 버린 별칭 {len(dropped_conflict)}건:")
for x in dropped_conflict[:15]:
    print("   ", x)
