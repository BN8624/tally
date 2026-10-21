# 업종 틀 배치·과세기간·결제 채널·요약 계산·검산을 검증합니다.
from datetime import date
from decimal import Decimal

import pandas as pd
import pytest

from tally import CompanySettings, Period, available_periods, default_period, process_transactions
from tally.core import CARD, EXEMPT, NONDEDUCTIBLE, PAYMENT, PURCHASE, SALES, cell_total


def transaction(
    number: int,
    division: str,
    original_type: str,
    supply: int,
    tax: int,
    *,
    code: str = "",
    name: str = "",
    vendor: str = "거래처",
    item: str = "품목",
    month: int = 7,
    total: int | None = None,
    electronic: str = "",
) -> dict[str, object]:
    return {
        "row_id": f"Sheet1:{number}",
        "sheet": "Sheet1",
        "source_row": number,
        "division": division,
        "date": date(2026, month, 1),
        "month": f"2026-{month:02d}",
        "vendor": vendor,
        "item": item,
        "supply_amount": Decimal(supply),
        "tax_amount": Decimal(tax),
        "total_amount": Decimal(supply + tax if total is None else total),
        "original_type": original_type,
        "account_code": code,
        "account_name": name,
        "card_company": "",
        "card_number": "",
        "electronic": electronic,
    }


def process(rows, industry="기본", **settings):
    return process_transactions(pd.DataFrame(rows), CompanySettings(name="업체", industry=industry, **settings))


def labels(result, block):
    return [column.label for column in result.layout[block]]


def columns_of(result):
    return dict(zip(result.transactions["row_id"], result.transactions["column"]))


def check(result, name):
    return result.validation.set_index("check").loc[name]


def canon_example_rows():
    return [
        transaction(1, "매입", "과세", 1005, 100, code="146", name="원재료"),
        transaction(2, "매입", "불공", 200, 20, code="813", name="접대비"),
        transaction(3, "매입", "과세", 1000, 99, code="208", name="비품"),
        transaction(4, "매입", "카과", 300, 30, code="830", name="소모품"),
        transaction(5, "매출", "과세", 3000, 300, code="401", name="상품매출"),
        transaction(6, "매출", "카과", 2000, 200, code="401", name="상품매출"),
        transaction(7, "매출", "현과", 1000, 100, code="401", name="상품매출"),
        transaction(8, "매출", "건별", 500, 50, code="401", name="상품매출"),
    ]


def test_construction_purchase_and_exempt_columns_follow_canon_rules() -> None:
    result = process(
        [
            transaction(1, "매입", "과세", 100, 10, code="156"),
            transaction(2, "매입", "과세", 100, 10, code="612"),
            transaction(3, "매입", "과세", 100, 10, code="208"),
            transaction(4, "매입", "과세", 100, 10, code="146"),
            transaction(5, "매입", "과세", 100, 10, code="813"),
            transaction(6, "매입", "과세", 100, 10, code=""),
            transaction(7, "매입", "면세", 300, 0, code="153"),
            transaction(8, "매입", "면세", 300, 0, code="156"),
            transaction(9, "매입", "의제", 300, 0, code="146"),
        ],
        industry="건설",
    )
    assert labels(result, PURCHASE) == ["원재료(도급)", "도급", "고정", "기타", "미분류", "계"]
    assert labels(result, EXEMPT) == ["원재료", "기타", "면세 계"]
    assert list(columns_of(result).values()) == ["원재료(도급)", "도급", "고정", "기타", "기타", "미분류", "원재료", "기타", "기타"]
    assert check(result, "계정 미분류 건수")["status"] == "실패"
    assert not result.validation_passed


def test_manufacturing_splits_raw_material_manufacturing_cost_other_and_fixed() -> None:
    result = process(
        [
            transaction(1, "매입", "과세", 100, 10, code="153"),
            transaction(2, "매입", "과세", 100, 10, code="146"),
            transaction(3, "매입", "과세", 100, 10, code="156"),
            transaction(4, "매입", "과세", 100, 10, code="511"),
            transaction(5, "매입", "과세", 100, 10, code="208"),
            transaction(6, "매입", "과세", 100, 10, code="813"),
            transaction(7, "매입", "면세", 500, 0, code="153"),
            transaction(8, "매입", "면세", 200, 0, code="530"),
            transaction(9, "매입", "면세", 70, 0, code="146"),
        ],
        industry="제조",
    )
    assert labels(result, PURCHASE) == ["원재료", "제조경비", "기타", "고정", "계"]
    assert labels(result, EXEMPT) == ["원재료", "제조경비", "기타", "면세 계"]
    assert list(columns_of(result).values()) == [
        "원재료", "기타", "기타", "제조경비", "고정", "기타", "원재료", "제조경비", "기타",
    ]
    assert result.summary_value("deemed", "supply") == 500
    assert result.validation_passed


def test_deemed_supply_follows_exempt_first_column_only_where_deemed_input_exists() -> None:
    rows = [
        transaction(1, "매입", "과세", 1000, 100, code="156"),
        transaction(2, "매입", "면세", 500, 0, code="153"),
        transaction(3, "매입", "면세", 200, 0, code="813"),
    ]
    construction = process(rows, industry="건설")
    assert (construction.summary_value("deemed", "supply"), construction.summary_value("deemed")) == (500, 0)
    assert construction.summary_value("purchase", "supply") == 1500
    keys = [row.key for row in construction.summary]
    assert keys[keys.index("card") + 1] == "deemed"
    assert process(rows, industry="기본").summary_value("deemed", "supply") == 500
    restaurant = process([*rows, transaction(4, "매입", "면세", 70, 0, code="146")], industry="음식점")
    assert restaurant.summary_value("deemed", "supply") == 70


def test_retail_card_purchases_split_goods_and_add_exempt_column() -> None:
    result = process(
        [
            transaction(1, "매입", "카과", 100, 10, code="146"),
            transaction(2, "매입", "카과", 100, 10, code="830"),
            transaction(3, "매입", "카면", 50, 0, code="830"),
        ],
        industry="도소매",
    )
    assert list(columns_of(result).values()) == ["카드상품", "카드기타", "카드면세"]
    assert labels(result, CARD) == ["카드상품", "카드기타", "현영상품", "현영기타", "카드면세", "공제 계"]
    assert "카드외 매입 · 카드면세" in check(result, "틀에 없어 추가한 열·블록")["detail"]
    assert result.summary_value("card", "supply") == Decimal(200)


def test_restaurant_sales_are_split_into_payment_channels() -> None:
    result = process(
        [
            transaction(1, "매출", "카과", 100, 10, vendor="(주)우아한형제들"),
            transaction(2, "매출", "카과", 100, 10, vendor="위대한상상"),
            transaction(3, "매출", "카과", 100, 10, vendor="제로페이"),
            transaction(4, "매출", "카과", 100, 10, vendor="일반손님"),
            transaction(5, "매출", "현과", 100, 10),
            transaction(6, "매출", "건별", 100, 10),
        ],
        industry="음식점",
    )
    assert list(columns_of(result).values()) == ["배민", "요기요", "제로페이", "카드", "현영", "현금"]
    assert labels(result, PAYMENT) == ["카드", "현영", "배민", "요기요", "쿠팡이츠", "제로페이", "과세 계", "현금"]


def test_period_choices_default_and_exclusion() -> None:
    months = [f"2026-{month:02d}" for month in range(1, 11)]
    assert default_period(months) == Period(2026, 2, "예정")
    assert default_period(["2026-07", "2026-08"]) == Period(2026, 2, "예정")
    assert default_period(["2026-04", "2026-06"]) == Period(2026, 1, "확정")
    assert [period.label for period in available_periods(["2026-05", "2026-08"])] == [
        "2026년 1기 확정",
        "2026년 1기 확정(6개월)",
        "2026년 2기 예정",
        "2026년 2기 확정(6개월)",
    ]
    assert Period(2026, 2, "확정6").months == [f"2026-{month:02d}" for month in range(7, 13)]

    rows = [transaction(number, "매입", "과세", 100, 10, code="146", month=number) for number in range(1, 11)]
    result = process(rows)
    assert result.period == Period(2026, 2, "예정")
    assert result.transactions["in_period"].tolist() == [False] * 6 + [True] * 3 + [False]
    assert check(result, "기간 밖 제외 건수")["actual"] == 7
    assert cell_total(result.cells, PURCHASE, ["원재료"], "supply") == Decimal(300)

    six_months = process_transactions(pd.DataFrame(rows), CompanySettings(name="업체"), Period(2026, 2, "확정6"))
    assert cell_total(six_months.cells, PURCHASE, ["원재료"], "count") == 4


def test_summary_matches_canon_example() -> None:
    result = process(canon_example_rows())
    expected = {
        "invoice": (2205, 220),
        "fixed": (1000, 99),
        "general": (1205, 121),
        "card": (300, 30),
        "nondeductible": (200, 20),
        "purchase": (2305, 230),
        "sales_invoice": (3000, 300),
        "card_out": (3000, 300),
        "other": (500, 50),
        "taxable": (6500, 650),
    }
    for key, (supply, tax) in expected.items():
        assert (result.summary_value(key, "supply"), result.summary_value(key)) == (supply, tax), key
    assert result.summary_value("payable") == 420
    assert result.summary_value("card_deduction") == 42
    assert result.summary_value("final") == 378
    assert result.validation_passed


@pytest.mark.parametrize(
    ("purchase_supply", "eligible", "payable", "deduction"),
    [(1000, True, 1100, 171), (11000, True, 100, 100), (20000, True, -800, 0), (1000, False, 1100, 0)],
)
def test_card_deduction_is_capped_by_payable(purchase_supply, eligible, payable, deduction) -> None:
    result = process(
        [
            transaction(1, "매입", "과세", purchase_supply, purchase_supply // 10, code="146"),
            transaction(2, "매출", "카과", 10000, 1000),
            transaction(3, "매출", "현과", 1000, 100),
            transaction(4, "매출", "카과", 1000, 100, vendor="제로페이"),
            transaction(5, "매출", "카면", 5000, 0),
        ],
        card_deduction_eligible=eligible,
    )
    assert result.summary_value("card_out", "supply") == 12000
    assert result.summary_value("exempt_payment", "supply") == 5000
    assert (result.summary_value("payable"), result.summary_value("card_deduction")) == (payable, deduction)


@pytest.mark.parametrize(
    ("rows", "payable", "relief", "shares", "final"),
    [
        ([transaction(1, "매출", "과세", 12345670, 1234567, code="416", name="운송수입")], 1234567, 1222221, (1111111, 61728, 49382), 12346),
        (canon_example_rows(), 420, 415, (378, 21, 16), 5),
        (
            [
                transaction(1, "매입", "과세", 10000000, 1000000, code="522", name="차량유지비"),
                transaction(2, "매출", "과세", 1000000, 100000, code="416", name="운송수입"),
            ],
            -900000,
            0,
            (0, 0, 0),
            -900000,
        ),
    ],
)
def test_taxi_relief_and_allocation_follow_payable(rows, payable, relief, shares, final) -> None:
    result = process(rows, industry="택시", card_deduction_eligible=False)
    basic = process(rows, card_deduction_eligible=False)
    assert result.layout == basic.layout
    assert result.summary_value("payable") == payable
    assert result.summary_value("taxi_relief") == relief
    assert tuple(result.summary_value(key) for key in ("taxi_drivers", "taxi_reduction", "taxi_welfare")) == shares
    assert result.summary_value("final") == final
    assert not {"taxi_relief", "taxi_drivers"} & {row.key for row in basic.summary}


def test_other_sales_are_reversed_from_total_amount() -> None:
    result = process(
        [
            transaction(1, "매출", "건별", 1008420, 100849),
            transaction(2, "매출", "건별", -500, -50),
        ]
    )
    assert (result.summary_value("other", "supply"), result.summary_value("other")) == (1007927, 100792)


def test_construction_adds_payment_block_only_when_ledger_has_payment_sales() -> None:
    without = process([transaction(1, "매출", "과세", 1000, 100, code="404", name="공사수입금")], industry="건설")
    assert PAYMENT not in without.layout
    with_card = process(
        [
            transaction(1, "매출", "과세", 1000, 100, code="404", name="공사수입금"),
            transaction(2, "매출", "카과", 2000, 200),
        ],
        industry="건설",
    )
    assert labels(with_card, PAYMENT) == ["카드", "과세 계"]
    assert "결제 채널 매출 블록" in check(with_card, "틀에 없어 추가한 열·블록")["detail"]
    assert with_card.summary_value("card_out", "supply") == 2000


def test_sales_columns_follow_account_codes_and_fixed_asset_sales() -> None:
    result = process(
        [
            transaction(1, "매출", "과세", 100, 10, code="404", name="제품매출"),
            transaction(2, "매출", "과세", 100, 10, code="401", name="회사설정계정과목"),
            transaction(3, "매출", "과세", 100, 10, code="208", name="차량운반구"),
            transaction(4, "매출", "영세", 100, 0, code="401", name="상품매출"),
            transaction(5, "매출", "면세", 100, 0, code="401", name="상품매출"),
        ]
    )
    assert labels(result, SALES) == ["수입금액", "제품매출", "고정", "계", "영세", "면세"]
    assert result.summary_value("sales_invoice", "supply") == 300
    assert result.summary_value("zero_rated", "supply") == 100
    assert result.summary_value("exempt_invoice", "supply") == 100


def test_nondeductible_rows_stay_in_invoice_total_and_show_separately() -> None:
    result = process(
        [
            transaction(1, "매입", "과세", 1000, 100, code="146"),
            transaction(2, "매입", "불공", 200, 20, code="813"),
        ]
    )
    assert cell_total(result.cells, PURCHASE, ["기타"], "supply") == 200
    assert cell_total(result.cells, NONDEDUCTIBLE, ["불공"], "tax") == 20
    assert check(result, "매입 집계 보존 건수")["actual"] == 2
    assert result.summary_value("purchase") == 100


def test_exempt_and_zero_rated_purchases_go_to_exempt_block() -> None:
    result = process(
        [
            transaction(1, "매입", "면세", 500, 0, code="146"),
            transaction(2, "매입", "의제매입세액", 300, 0, code="830"),
            transaction(3, "매입", "영세", 200, 0, code="146"),
        ]
    )
    assert list(columns_of(result).values()) == ["원재료", "기타", "영세"]
    assert labels(result, EXEMPT) == ["원재료", "기타", "면세 계", "영세"]
    assert result.validation_passed


def test_override_outside_profile_goes_to_other_with_note() -> None:
    result = process(
        [
            transaction(1, "매입", "과세", 100, 10, code="899"),
            transaction(2, "매입", "과세", 100, 10, code="146"),
            transaction(3, "매입", "과세", 100, 10, code="157"),
            transaction(4, "매입", "면세", 100, 0, code="157"),
        ],
        industry="건설",
        account_overrides={"899": "도급", "146": "운송경비", "157": "원재료"},
    )
    assert list(columns_of(result).values()) == ["도급", "기타", "기타", "원재료"]
    assert check(result, "틀에 없어 기타로 넣은 계정 예외")["detail"] == "146=운송경비, 157=원재료"


def test_candidates_are_listed_without_changing_aggregates() -> None:
    result = process(
        [
            transaction(1, "매입", "과세", 100, 10, code="822", name="차량유지비", item="주유"),
            transaction(2, "매입", "카과", 100, 10, code="537", name="LPG", item="주유"),
            transaction(3, "매입", "과세", 100, 10, code="813", name="접대비"),
            transaction(4, "매입", "과세", 100, 10, code="146", name="원재료"),
        ]
    )
    assert result.candidates["row_id"].tolist() == ["Sheet1:1", "Sheet1:2", "Sheet1:3"]
    assert result.summary_value("purchase") == 40
    assert result.validation_passed


def test_unsupported_type_and_empty_period_fail_validation() -> None:
    result = process([transaction(1, "매입", "과세", 100, 10, code="146"), transaction(2, "매입", "수입", 500, 50, code="146")])
    assert check(result, "미집계 유형 건수")["actual"] == 1
    assert check(result, "매입 집계 보존 공급가액")["status"] == "실패"
    assert not result.validation_passed

    empty = process_transactions(pd.DataFrame([transaction(1, "매입", "과세", 100, 10, month=3)]), CompanySettings(name="업체"), Period(2026, 2, "예정"))
    assert check(empty, "기간 안 상세 거래 없음")["actual"] == 1
    assert not empty.validation_passed


def test_duplicated_row_id_stops_processing() -> None:
    rows = [transaction(1, "매입", "과세", 100, 10), transaction(1, "매입", "과세", 100, 10)]
    with pytest.raises(ValueError, match="row_id가 중복"):
        process(rows)


def test_layout_keeps_every_month_of_the_period() -> None:
    result = process([transaction(1, "매입", "과세", 100, 10, code="146", month=8)])
    assert result.period.months == ["2026-07", "2026-08", "2026-09"]
    assert set(result.cells["month"]) == {"2026-08"}
