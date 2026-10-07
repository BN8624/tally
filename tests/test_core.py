# 매입·매출 분류와 불공 차감 및 검산 계산을 검증합니다.
from datetime import date
from decimal import Decimal

import pandas as pd
import pytest

from tally.core import classify_purchase_account, process_transactions
from tally.settings import CompanySettings


def transaction(
    row_id: str,
    division: str,
    original_type: str,
    code: str,
    account_name: str,
    supply: int,
    tax: int,
    *,
    item: str = "",
    month: str = "2026-04",
) -> dict[str, object]:
    return {
        "row_id": row_id,
        "sheet": "Sheet1",
        "source_row": int(row_id[1:]),
        "division": division,
        "date": date(2026, int(month[-2:]), 1),
        "month": month,
        "vendor": "거래처",
        "item": item,
        "supply_amount": Decimal(supply),
        "tax_amount": Decimal(tax),
        "total_amount": Decimal(supply + tax),
        "original_type": original_type,
        "account_code": code,
        "account_name": account_name,
        "card_company": "",
        "card_number": "",
    }


def test_account_classification_uses_only_canon_rules_and_company_priority() -> None:
    settings = CompanySettings(
        name="업체",
        account_146_label="음식재료",
        fixed_asset_codes={"210"},
        account_overrides={"899": "제조경비"},
    )
    assert classify_purchase_account("146", settings) == "음식재료"
    assert classify_purchase_account("156", settings) == "원재료(도급)"
    assert classify_purchase_account("210", settings) == "고정"
    assert classify_purchase_account("212", settings) == "고정"
    assert classify_purchase_account("512", settings) == "제조경비"
    assert classify_purchase_account("612", settings) == "도급경비"
    assert classify_purchase_account("813", settings) == "기타"
    assert classify_purchase_account("899", settings) == "제조경비"
    assert classify_purchase_account("499", settings) == "미분류"


def test_explicit_account_override_is_not_replaced_by_tobacco_rule() -> None:
    rows = [
        transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
        transaction("r2", "매입", "과세", "813", "소모품", 500, 50),
    ]
    for row in rows:
        row["vendor"] = "KT&G"
    overridden = process_transactions(
        pd.DataFrame(rows),
        CompanySettings(name="업체", account_overrides={"146": "상품", "813": "상품"}),
    )
    default = process_transactions(pd.DataFrame(rows), CompanySettings(name="업체"))

    assert overridden.transactions["account_category"].tolist() == ["상품", "상품"]
    assert default.transactions["account_category"].tolist() == ["담배", "기타"]


def test_processing_keeps_nondeductible_in_tax_aggregate_then_subtracts_it() -> None:
    data = pd.DataFrame(
        [
            transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
            transaction("r2", "매입", "불공", "813", "접대비", 200, 20),
            transaction("r3", "매입", "카과", "", "", 300, 30),
            transaction("r4", "매입", "현과", "", "", 400, 40),
            transaction("r5", "매입", "면세", "", "", 500, 0),
            transaction("r9", "매입", "공통", "813", "공통매입", 150, 15),
            transaction("r10", "매입", "의제매입세액", "", "", 50, 5),
            transaction("r6", "매출", "과세", "401", "상품매출", 600, 60),
            transaction("r7", "매출", "카과", "401", "상품매출", 700, 70),
            transaction("r8", "매출", "현과", "401", "상품매출", -100, -10),
        ]
    )
    result = process_transactions(data, CompanySettings(name="업체"))
    category_total = result.purchase_by_category["supply_amount"].sum()
    assert category_total == Decimal(1350)

    summary = result.purchase_summary.set_index("item")
    assert summary.loc["카드외", "supply_amount"] == Decimal(700)
    assert summary.loc["의제매입세액", "supply_amount"] == Decimal(50)
    assert summary.loc["그 밖의 공제매입세액", "supply_amount"] == Decimal(750)
    assert summary.loc["매입세액 합계", "supply_amount"] == Decimal(2100)
    assert summary.loc["불공", "supply_amount"] == Decimal(200)
    assert summary.loc["공통", "supply_amount"] == Decimal(150)
    assert summary.loc["차감계", "supply_amount"] == Decimal(1750)
    assert summary.loc["과매계", "supply_amount"] == Decimal(1750)
    assert summary.loc["면세 매입", "supply_amount"] == Decimal(500)

    sales = result.sales_summary.set_index("item")
    assert sales.loc["과세매출 총계", "supply_amount"] == Decimal(1200)
    assert sales.loc["카드매출", "total_amount"] == Decimal(770)
    assert sales.loc["현영매출", "total_amount"] == Decimal(-110)
    assert not result.transactions.loc[result.transactions["division"].eq("매출"), "review_status"].any()
    assert result.validation_passed


def test_candidate_is_not_auto_confirmed_and_unclassified_fails_closed() -> None:
    data = pd.DataFrame(
        [
            transaction("r1", "매입", "과세", "499", "기타비용", 100, 10, item="개인사용"),
        ]
    )
    result = process_transactions(data, CompanySettings(name="업체"))
    assert result.review.iloc[0]["review_status"] == "판단 보류"
    failures = set(result.validation.loc[result.validation["status"].eq("실패"), "check"])
    assert failures == {
        "과세·불공·공통 계정분류 건수",
        "과세·불공·공통 계정분류 공급가액",
        "과세·불공·공통 계정분류 세액",
        "미분류 건수",
        "불공 판단 보류 건수",
    }


def test_candidate_decision_can_be_applied_and_negative_transaction_counts_as_one() -> None:
    data = pd.DataFrame(
        [transaction("r1", "매입", "과세", "822", "차량유지비", -100, -10, item="차량 수리")]
    )
    result = process_transactions(
        data,
        CompanySettings(name="업체"),
        decisions={"r1": {"decision": "불공으로 변경", "reason": "비영업용 차량"}},
    )
    assert result.review.iloc[0]["final_type"] == "불공"
    assert result.purchase_summary.set_index("item").loc["불공", "count"] == 1
    assert result.purchase_summary.set_index("item").loc["불공", "supply_amount"] == Decimal(-100)
    assert result.validation_passed


def test_unaggregated_type_fails_validation_instead_of_vanishing() -> None:
    data = pd.DataFrame(
        [
            transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
            transaction("r2", "매입", "수입", "146", "상품", 5000, 500),
            transaction("r3", "매출", "영세", "401", "상품매출", 8000, 0),
        ]
    )
    result = process_transactions(data, CompanySettings(name="업체"))
    validation = result.validation.set_index("check")

    assert validation.loc["미집계 유형 건수", "actual"] == 2
    assert "수입" in validation.loc["미집계 유형 건수", "detail"]
    assert "영세" in validation.loc["미집계 유형 건수", "detail"]
    assert validation.loc["매입 집계 보존 공급가액", "expected"] == Decimal(6000)
    assert validation.loc["매입 집계 보존 공급가액", "actual"] == Decimal(1000)
    assert validation.loc["매출 집계 보존 건수", "expected"] == 1
    assert validation.loc["매출 집계 보존 건수", "actual"] == 0
    assert not result.validation_passed


def test_empty_input_is_not_reported_as_verified() -> None:
    columns = list(transaction("r1", "매입", "과세", "146", "상품", 0, 0))
    result = process_transactions(pd.DataFrame([], columns=columns), CompanySettings(name="업체"))
    validation = result.validation.set_index("check")

    assert validation.loc["상세 거래 없음", "actual"] == 1
    assert not result.validation_passed


def test_every_sample_type_stays_inside_an_aggregate() -> None:
    data = pd.DataFrame(
        [
            transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
            transaction("r2", "매입", "불공", "813", "접대비", 200, 20),
            transaction("r3", "매입", "카과", "", "", 300, 30),
            transaction("r4", "매입", "현과", "", "", 400, 40),
            transaction("r5", "매입", "카면", "", "", 500, 0),
            transaction("r6", "매입", "면세", "", "", 600, 0),
            transaction("r7", "매입", "영세", "", "", 700, 0),
            transaction("r8", "매출", "과세", "401", "상품매출", 800, 80),
            transaction("r9", "매출", "건별", "401", "상품매출", 900, 90),
            transaction("r10", "매출", "카과", "401", "상품매출", 1000, 100),
            transaction("r11", "매출", "현과", "401", "상품매출", 1100, 110),
            transaction("r12", "매출", "면세", "401", "상품매출", 1200, 0),
            transaction("r13", "매출", "카면", "401", "상품매출", 1300, 0),
            transaction("r14", "매출", "현면", "401", "상품매출", 1400, 0),
            transaction("r15", "매출", "면건", "401", "상품매출", 1500, 0),
        ]
    )
    result = process_transactions(data, CompanySettings(name="업체"))
    validation = result.validation.set_index("check")

    assert validation.loc["미집계 유형 건수", "actual"] == 0
    assert validation.loc["매입 집계 보존 건수", "actual"] == 7
    assert validation.loc["매출 집계 보존 건수", "actual"] == 8
    assert result.validation_passed


def test_information_rows_show_totals_without_deciding_validation() -> None:
    data = pd.DataFrame(
        [
            transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
            transaction("r2", "매출", "과세", "401", "상품매출", 2000, 200),
        ]
    )
    result = process_transactions(data, CompanySettings(name="업체"))
    validation = result.validation.set_index("check")

    assert validation.loc["상세 거래 건수", "status"] == "정보"
    assert validation.loc["상세 거래 건수", "actual"] == 2
    assert validation.loc["공급가액 합계", "actual"] == Decimal(3000)
    assert validation.loc["2026-04 합계금액", "status"] == "정보"
    assert set(result.validation["status"]) == {"통과", "정보"}
    assert result.validation_passed


def test_checks_that_could_never_fail_are_gone() -> None:
    data = pd.DataFrame([transaction("r1", "매입", "과세", "146", "상품", 1000, 100)])
    checks = set(process_transactions(data, CompanySettings(name="업체")).validation["check"])

    assert checks.isdisjoint(
        {
            "중복 거래",
            "세금계산서 매입계 관계",
            "카드매입 관계",
            "과세 매입 총계 관계",
            "과세매출 공급가액",
            "카드매출 보조표",
            "현영매출 보조표",
        }
    )


def test_duplicated_row_id_stops_processing_outright() -> None:
    data = pd.DataFrame(
        [
            transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
            transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
        ]
    )
    with pytest.raises(ValueError, match="row_id가 중복"):
        process_transactions(data, CompanySettings(name="업체"))


def test_zero_pay_detection_allows_intermediate_card_brand_text() -> None:
    row = transaction("r1", "매출", "카과", "401", "상품매출", 200, 20)
    row["vendor"] = "제로(온누리)페이"

    result = process_transactions(pd.DataFrame([row]), CompanySettings(name="업체"))
    sales = result.sales_summary.set_index("item")

    assert sales.loc["제로페이", "total_amount"] == Decimal(220)
    assert "카드매출" not in sales.index


def test_purchase_exempt_payment_types_stay_visible_but_are_not_deducted() -> None:
    data = pd.DataFrame(
        [
            transaction("r1", "매입", "과세", "146", "상품", 1000, 100),
            transaction("r2", "매입", "카과", "", "", 300, 30),
            transaction("r3", "매입", "현과", "", "", 400, 40),
            transaction("r4", "매입", "카면", "", "", 500, 0),
            transaction("r5", "매입", "카영", "", "", 100, 0),
            transaction("r6", "매입", "현면", "", "", 200, 0),
            transaction("r7", "매입", "현영", "", "", 50, 0),
        ]
    )

    result = process_transactions(data, CompanySettings(name="업체"))
    summary = result.purchase_summary.set_index("item")

    assert summary.loc["카드외", "supply_amount"] == Decimal(700)
    assert summary.loc["차감계", "supply_amount"] == Decimal(1700)
    assert summary.loc["카면", "supply_amount"] == Decimal(500)
    assert summary.loc["카영", "supply_amount"] == Decimal(100)
    assert summary.loc["현면", "supply_amount"] == Decimal(200)
    assert summary.loc["현영", "supply_amount"] == Decimal(50)
    assert "의제매입세액" not in summary.index
    assert "공통" not in summary.index
    assert result.validation_passed
