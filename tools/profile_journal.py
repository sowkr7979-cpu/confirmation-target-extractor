"""Read-only keyword profile to support the automation design."""
import collections
import json
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
PATH = Path(sys.argv[1] if len(sys.argv) > 1 else "")  # 분석할 분개장 경로를 인자로 준다
FIN = "은행 보험 해상 화재 이자 금융 보증 부보 저당 질권 차입 대출 예금 적금 사용제한 대여 상환 캐피탈 투자 증권 배당 계좌 L/C 신용장 파생 어음 수표 견질 신용 당좌 리스 임차 렌탈 농협 신협".split()
LAW = "법무 법률 변호사 김앤장 소송 배상 성공보수 로펌".split()
SUPPLEMENT = "고문료 자문료 법무법인 법률사무소 손해보험 생명보험 새마을금고 수협 산업은행 수출입은행 저축은행 파이낸셜".split()
def norm(value):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value or ""))).casefold()

def main():
    wb = openpyxl.load_workbook(PATH, read_only=True, data_only=True)
    ws = wb["분개장"]
    stats = collections.Counter()
    names = collections.Counter()
    law_names = collections.Counter()
    extras = collections.Counter()
    keywords = collections.Counter()
    examples = {}
    dates = []
    vouchers = set()
    total_debit = total_credit = 0
    excluded = []
    for rownum, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if not any(v is not None for v in row):
            continue
        stats["nonempty_rows"] += 1
        try:
            date = datetime.strptime(str(row[2])[:10], "%Y-%m-%d").date()
        except (ValueError, TypeError):
            excluded.append({"row": rownum, "date_cell": str(row[2])})
            continue
        stats["transaction_rows"] += 1
        dates.append(date.isoformat())
        if row[1]: vouchers.add(str(row[1]))
        if isinstance(row[7], (int, float)): total_debit += row[7]
        if isinstance(row[8], (int, float)): total_credit += row[8]
        texts = {"계정과목": norm(row[6]), "거래처": norm(row[9]), "적요": norm(row[11])}
        hits = {col: [k for k in FIN + LAW if norm(k) in text] for col, text in texts.items()}
        primary = bool(hits["거래처"] or hits["적요"])
        any_hit = any(hits.values())
        stats["counterparty_or_description_hit_rows"] += int(primary)
        stats["including_account_hit_rows"] += int(any_hit)
        stats["account_only_hit_rows"] += int(any_hit and not primary)
        combined = " ".join(texts.values())
        matched = [k for k in FIN + LAW if norm(k) in combined]
        keywords.update(matched)
        party = str(row[9] or "[거래처 없음]")
        if any_hit:
            names[party] += 1
            if party not in examples:
                examples[party] = {"row": rownum, "account": row[6], "description": row[11], "hits": hits}
        if any(norm(k) in combined for k in LAW):
            law_names[party] += 1
        if any(norm(k) in combined for k in SUPPLEMENT) and not any_hit:
            extras[party] += 1
    result = {"file": str(PATH), "sheet": "분개장", "stats": dict(stats), "excluded": excluded, "date_min": min(dates), "date_max": max(dates), "unique_voucher_ids": len(vouchers), "total_debit": total_debit, "total_credit": total_credit, "keyword_counts": dict(keywords), "distinct_hit_counterparties": len(names), "top_counterparties": names.most_common(60), "law_counterparties": law_names.most_common(), "supplement_only_counterparties": extras.most_common(30), "examples": examples}
    Path("work/inspection/journal_profile.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k:v for k,v in result.items() if k not in ["examples", "top_counterparties", "keyword_counts"]}, ensure_ascii=False, indent=2))
    print("LEGAL_EXAMPLES", json.dumps({k:v for k,v in examples.items() if "법무" in k or k == "한한손해보험"}, ensure_ascii=False))
    wb.close()

if __name__ == "__main__":
    main()
