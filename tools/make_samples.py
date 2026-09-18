# -*- coding: utf-8 -*-
"""E2E 시험용 견본 두 개를 만든다. 실제 감사 자료는 일절 쓰지 않는다.

  samples/샘플_분개장.xlsx     — 합성 거래 + 합계·빈행·반복헤더
  samples/분개장_sample.xlsx   — 양식이 다른 분개장(시트·헤더위치·열이름이 모두 다르다)

거래처 이름은 지어낸 것이거나 공개된 금융기관명이다. 금액·날짜도 지어낸 값이다.
Control Sheet는 견본을 만들지 않는다 — 프로그램이 template 폴더의 양식을 복사해
**출력물로 새로 만든다.**

    python tools/make_samples.py
"""
from __future__ import annotations

import datetime as _dt
import os
import sys

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SAMPLES = os.path.join(BASE, "samples")
# (거래처, 적요, 계정과목, 차변, 대변) — 무엇을 확인하려는 행인지 주석에 적는다.
ROWS = [
    ("KB국민은행 강남지점", "보통예금 이자", "보통예금", 0, 12_500),        # 사전 일치 + 지점 단서
    ("KB국민은행 강남지점", "계좌 00123-45-678901 대체", "보통예금", 0, 4_000),  # 계좌 단서
    ("(주)하나은행", "차입금 이자", "이자비용", 1_250_000, 0),               # 접두 (주) 제거
    ("하나은행(00000)", "정기예금 예치", "단기금융상품", 30_000_000, 0),     # 같은 기관으로 통합
    ("중소기업은행", "시설자금 대출 실행", "장기차입금", 0, 500_000_000),
    ("우리은행", "수수료", "지급수수료", 3_000, 0),
    ("신한은행 여의도", "외화송금", "외화예금", 8_800_000, 0),
    ("삼성화재해상보험", "화재보험료", "보험료", 1_200_000, 0),
    ("DB손해보험", "자동차보험료", "차량유지비", 780_000, 0),
    ("서울보증보험", "이행보증보험료", "지급수수료", 240_000, 0),
    ("신용보증기금", "보증료", "지급수수료", 1_100_000, 0),
    ("신한카드", "법인카드 결제", "미지급금", 2_430_000, 0),
    ("비씨카드", "법인카드 결제", "미지급금", 310_000, 0),
    ("SBI저축은행", "차입금 상환", "단기차입금", 20_000_000, 0),
    ("현대캐피탈", "리스료", "지급임차료", 1_650_000, 0),
    ("농협(000000)", "대체", "보통예금", 0, 5_000_000),                    # 짧은 별칭
    ("법무법인 가나다", "등기 자문료", "지급수수료", 3_300_000, 0),          # 명칭패턴 B
    ("마바사 법무사사무소", "법인등기 대행", "지급수수료", 550_000, 0),       # 법률 접미사 B
    ("가나다렌탈", "복합기 렌탈료", "지급임차료", 132_000, 0),               # 접미사 '렌탈'
    ("토스페이먼츠", "PG 정산수수료", "지급수수료", 88_000, 0),              # 전자금융 → 제외 분류
    ("나이스페이먼츠", "PG 정산수수료", "지급수수료", 45_000, 0),            # 전자금융 → 제외 분류
    ("마나마나(주)", "사무용품", "소모품비", 150_000, 0),                   # 키워드 없음 → 검색 안 됨
    ("두리두리상사", "원재료 매입", "원재료", 12_000_000, 0),               # 키워드 없음 → 검색 안 됨
    ("", "국민은행 대출이자 이체", "이자비용", 450_000, 0),                 # 거래처 공란 → 이름미확인
    ("한한손해보험", "보험료", "보험료", 90_000, 0),                        # 오타형 이름도 고치지 않는다
    ("기업비씨", "법인카드 결제", "미지급금", 1_780_000, 0),                 # 지정 키워드 없이 별칭만 일치
    ("건강보험", "4대보험 납부", "예수금", 2_100_000, 0),                   # 사회보험 → 제외 분류
    ("고용보험", "4대보험 납부", "예수금", 340_000, 0),                     # 사회보험 → 제외 분류
    ("산재보험료", "4대보험 납부", "예수금", 180_000, 0),                   # 사회보험 → 제외 분류
    ("국민연금공단", "4대보험 납부", "예수금", 1_450_000, 0),               # 사회보험 → 제외 분류
    ("국민건강보험공단", "4대보험 납부", "예수금", 2_050_000, 0),            # 사회보험 → 제외 분류
    ("근로복지공단", "4대보험 납부", "예수금", 520_000, 0),                 # 사회보험 → 제외 분류
    ("한국고용보험사무소", "급여 대행 수수료", "지급수수료", 300_000, 0),      # 이름이 비슷한 다른 거래처
    ("삼성생명보험", "단체보험료", "복리후생비", 600_000, 0),                # 민간 보험사는 유지된다
    ("행복마트 포스보증금/구미", "포스 보증금", "보증금", 500_000, 0),       # C등급 보증금 거래처
    ("기업시설자금대출(00033)", "대출 원금 상환", "장기차입금", 10_000_000, 0),  # C등급 '검토 필요'
]


def make_journal(path: str) -> int:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "분개장"
    ws.append(["전표일자", "전표번호", "구분", "계정과목", "거래처", "적요", "차변", "대변"])

    day = _dt.date(2025, 1, 6)
    n = 0
    for i, (cp, desc, acct, debit, credit) in enumerate(ROWS, start=1):
        d = day + _dt.timedelta(days=i * 11)
        ws.append([d.strftime("%Y-%m-%d"), str(i), "차변" if debit else "대변",
                   acct, cp, desc, debit or None, credit or None])
        n += 1
        if i == 8:                                   # 빈 행
            ws.append([None] * 8)
        if i == 14:                                  # 반복 헤더
            ws.append(["전표일자", "전표번호", "구분", "계정과목", "거래처", "적요", "차변", "대변"])

    ws.append(["합계", "", "", "", "", "",
               sum(r[3] for r in ROWS), sum(r[4] for r in ROWS)])
    for col, width in zip("ABCDEFGH", (12, 9, 7, 16, 26, 30, 14, 14)):
        ws.column_dimensions[col].width = width
    wb.save(path)
    return n


def make_alt_journal(path: str) -> int:
    """양식이 다른 분개장. 시트 이름·헤더 위치·열 이름이 모두 다르다.

    프로그램이 `input_profiles.json`의 열 이름 후보로 스스로 맞추는지 본다.
    """
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"                      # 시트 이름이 '분개장'이 아니다
    ws["A1"] = "회계프로그램 내보내기"        # 헤더 앞에 머리글 3줄
    ws["A2"] = "기간: 2025-01-01 ~ 2025-12-31"
    ws["A3"] = ""
    ws.append([])                            # 4행부터 헤더
    ws.append(["일자", "번호", "계정명", "거래처명", "비고", "차변금액", "대변금액"])

    day = _dt.date(2025, 2, 3)
    n = 0
    for i, (cp, desc, acct, debit, credit) in enumerate(ROWS, start=1):
        d = day + _dt.timedelta(days=i * 9)
        ws.append([d.strftime("%Y/%m/%d"), i, acct, cp, desc, debit or None, credit or None])
        n += 1
    for col, width in zip("ABCDEFG", (12, 7, 16, 26, 30, 14, 14)):
        ws.column_dimensions[col].width = width
    wb.save(path)
    return n


def main() -> int:
    os.makedirs(SAMPLES, exist_ok=True)
    journal = os.path.join(SAMPLES, "샘플_분개장.xlsx")
    alt = os.path.join(SAMPLES, "분개장_sample.xlsx")

    rows = make_journal(journal)
    print(f"만듦: {journal} (거래 {rows}행 + 합계·빈행·반복헤더)")
    rows_alt = make_alt_journal(alt)
    print(f"만듦: {alt} (거래 {rows_alt}행, 다른 양식 — 시트·헤더위치·열이름이 다르다)")
    print("\n이제 이렇게 시험한다:")
    print("  python src/app.py --journal samples/샘플_분개장.xlsx")
    print("  → samples/output/ 에 Control Sheet가 새로 만들어진다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
