# 상세 거래를 정본 규칙에 따라 분류·집계하고 원본 대비 검산합니다.
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Mapping

import pandas as pd

from .settings import CompanySettings


ZERO = Decimal("0")
PURCHASE_TAX_TYPES = {"과세", "불공"}
PURCHASE_CARD_TYPES = {"카과", "현과"}
SALES_TAX_TYPES = {"과세", "건별", "카과", "현과"}
CATEGORY_ORDER = ["상품", "음식재료", "원재료(도급)", "제조경비", "도급경비", "기타", "고정"]


@dataclass(slots=True)
class ProcessingResult:
    transactions: pd.DataFrame
    review: pd.DataFrame
    purchase_by_category: pd.DataFrame
    purchase_summary: pd.DataFrame
    sales_by_account: pd.DataFrame
    sales_summary: pd.DataFrame
    validation: pd.DataFrame
    validation_passed: bool


def classify_purchase_account(code: str, settings: CompanySettings) -> str:
    code = str(code).strip()
    if code in settings.account_overrides:
        return settings.account_overrides[code]
    if code in settings.fixed_asset_codes:
        return "고정"
    if code == "146":
        return settings.account_146_label
    if code == "156":
        return "원재료(도급)"
    if len(code) == 3 and code.isdigit():
        if code.startswith("5"):
            return "제조경비"
        if code.startswith("6"):
            return "도급경비"
        if code.startswith("8"):
            return "기타"
    return "미분류"


def find_nondeductible_candidate(row: pd.Series, settings: CompanySettings) -> str:
    if row["division"] != "매입" or row["original_type"] not in PURCHASE_TAX_TYPES:
        return ""
    if row["original_type"] == "불공":
        return "원본 유형이 불공"

    account = f"{row['account_code']} {row['account_name']}".casefold()
    context = f"{row['vendor']} {row['item']} {row['account_name']}".casefold()
    reasons: list[str] = []
    if any(keyword.casefold() in account for keyword in ("접대비", "기업업무추진비")):
        reasons.append(f"계정과목 {row['account_code']} {row['account_name']}")
    vehicle_hits = [keyword for keyword in settings.vehicle_keywords if keyword.casefold() in context]
    if vehicle_hits:
        reasons.append(f"차량 관련 값 감지: {', '.join(vehicle_hits)}")
    vendor_hits = [keyword for keyword in settings.vendor_keywords if keyword.casefold() in context]
    if vendor_hits:
        reasons.append(f"업체 규칙 감지: {', '.join(vendor_hits)}")
    personal_hits = [keyword for keyword in settings.personal_keywords if keyword.casefold() in context]
    if personal_hits:
        reasons.append(f"사업 무관 가능성: {', '.join(personal_hits)}")
    return "; ".join(reasons)


def _decision_values(original_type: str, candidate_reason: str, decision: Mapping[str, str]) -> tuple[str, str, str, str]:
    requested = decision.get("decision", "").strip()
    reason = decision.get("reason", "").strip()
    memo = decision.get("memo", "").strip()
    if not requested:
        if original_type == "불공":
            return "불공 유지", "불공", reason or "원본에서 불공 지정", memo
        if candidate_reason:
            return "판단 보류", original_type, reason, memo
        return "과세 유지", original_type, reason, memo

    valid = {"불공 유지", "과세로 변경", "과세 유지", "불공으로 변경", "판단 보류"}
    if requested not in valid:
        raise ValueError(f"지원하지 않는 불공 판정입니다. 판정={requested}")
    if requested in {"불공 유지", "불공으로 변경"}:
        return requested, "불공", reason or "사용자 불공 확정", memo
    if requested in {"과세 유지", "과세로 변경"}:
        return requested, "과세", reason, memo
    return requested, original_type, reason, memo


def _aggregate(frame: pd.DataFrame, groups: list[str]) -> pd.DataFrame:
    columns = groups + ["count", "supply_amount", "tax_amount", "total_amount"]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    result = (
        frame.groupby(groups, sort=True, dropna=False)
        .agg(
            count=("row_id", "size"),
            supply_amount=("supply_amount", "sum"),
            tax_amount=("tax_amount", "sum"),
            total_amount=("total_amount", "sum"),
        )
        .reset_index()
    )
    return result[columns]


def _summary_rows(frames: Mapping[str, pd.DataFrame]) -> pd.DataFrame:
    rows: list[pd.DataFrame] = []
    for item, frame in frames.items():
        aggregate = _aggregate(frame, ["month"])
        aggregate.insert(0, "item", item)
        rows.append(aggregate)
    if not rows:
        return pd.DataFrame(columns=["item", "month", "count", "supply_amount", "tax_amount", "total_amount"])
    return pd.concat(rows, ignore_index=True)


def _numeric_total(frame: pd.DataFrame, column: str) -> Decimal:
    if frame.empty:
        return ZERO
    value = frame[column].sum()
    return value if isinstance(value, Decimal) else Decimal(str(value))


def _validation_row(check: str, expected: object, actual: object, *, detail: str = "") -> dict[str, object]:
    difference: object = ""
    if isinstance(expected, (int, Decimal)) and isinstance(actual, (int, Decimal)):
        difference = actual - expected
    passed = expected == actual
    return {
        "check": check,
        "status": "통과" if passed else "실패",
        "expected": expected,
        "actual": actual,
        "difference": difference,
        "detail": detail,
    }


def process_transactions(
    transactions: pd.DataFrame,
    settings: CompanySettings,
    decisions: Mapping[str, Mapping[str, str]] | None = None,
) -> ProcessingResult:
    decisions = decisions or {}
    data = transactions.copy()
    if data["row_id"].duplicated().any():
        raise ValueError("입력 상세 거래의 row_id가 중복되었습니다.")

    categories: list[str] = []
    candidates: list[str] = []
    review_statuses: list[str] = []
    final_types: list[str] = []
    nondeductible_reasons: list[str] = []
    memos: list[str] = []
    for _, row in data.iterrows():
        category = ""
        if row["division"] == "매입" and row["original_type"] in PURCHASE_TAX_TYPES:
            category = classify_purchase_account(row["account_code"], settings)
        candidate = find_nondeductible_candidate(row, settings)
        if row["division"] == "매입" and row["original_type"] in PURCHASE_TAX_TYPES:
            status, final_type, reason, memo = _decision_values(
                row["original_type"], candidate, decisions.get(row["row_id"], {})
            )
        else:
            status, final_type, reason, memo = "", row["original_type"], "", ""
        categories.append(category)
        candidates.append(candidate)
        review_statuses.append(status)
        final_types.append(final_type)
        nondeductible_reasons.append(reason)
        memos.append(memo)

    data["account_category"] = categories
    data["candidate_reason"] = candidates
    data["review_status"] = review_statuses
    data["final_type"] = final_types
    data["nondeductible"] = data["final_type"].eq("불공")
    data["nondeductible_reason"] = nondeductible_reasons
    data["review_memo"] = memos

    purchase = data[data["division"].eq("매입")]
    purchase_tax = purchase[purchase["original_type"].isin(PURCHASE_TAX_TYPES)]
    purchase_known = purchase_tax[~purchase_tax["account_category"].eq("미분류")]
    purchase_by_category = _aggregate(purchase_known, ["account_category", "month"])

    general = purchase_known[~purchase_known["account_category"].eq("고정")]
    fixed = purchase_known[purchase_known["account_category"].eq("고정")]
    card_tax = purchase[purchase["original_type"].eq("카과")]
    cash_tax = purchase[purchase["original_type"].eq("현과")]
    card_total = purchase[purchase["original_type"].isin(PURCHASE_CARD_TYPES)]
    taxable_purchase = pd.concat([purchase_known, card_total], ignore_index=True)
    nondeductible = purchase_tax[purchase_tax["final_type"].eq("불공")]
    deductible = taxable_purchase.copy()

    purchase_summary = _summary_rows(
        {
            "일반매입": general,
            "고정": fixed,
            "세금계산서 매입계": purchase_known,
            "카과": card_tax,
            "현과": cash_tax,
            "카드매입": card_total,
            "과세 매입 총계": taxable_purchase,
            "불공": nondeductible,
            "면세 매입": purchase[purchase["original_type"].eq("면세")],
        }
    )
    deductible_summary = _summary_rows({"과매계": deductible})
    if not deductible_summary.empty:
        for month in deductible_summary["month"].unique():
            mask = deductible_summary["month"].eq(month)
            nd = nondeductible[nondeductible["month"].eq(month)]
            deductible_summary.loc[mask, "count"] -= len(nd)
            for column in ("supply_amount", "tax_amount", "total_amount"):
                deductible_summary.loc[mask, column] = deductible_summary.loc[mask, column].map(
                    lambda value, amount=_numeric_total(nd, column): value - amount
                )
        purchase_summary = pd.concat([purchase_summary, deductible_summary], ignore_index=True)

    sales = data[data["division"].eq("매출")]
    taxable_sales = sales[sales["original_type"].isin(SALES_TAX_TYPES)]
    sales_by_account = _aggregate(taxable_sales, ["account_name", "month"])
    sales_summary = _summary_rows(
        {
            "과세매출 총계": taxable_sales,
            "면세 매출": sales[sales["original_type"].eq("면세")],
            "카드매출": sales[sales["original_type"].eq("카과")],
            "현영매출": sales[sales["original_type"].eq("현과")],
        }
    )

    review = data[
        (data["division"].eq("매입"))
        & (data["original_type"].isin(PURCHASE_TAX_TYPES))
        & (
            data["original_type"].eq("불공")
            | data["candidate_reason"].ne("")
            | data["review_status"].eq("판단 보류")
        )
    ].copy()

    rows = [
        _validation_row("상세 거래 건수", len(transactions), len(data)),
        _validation_row("중복 거래", 0, int(data["row_id"].duplicated().sum())),
        _validation_row("공급가액 합계", _numeric_total(transactions, "supply_amount"), _numeric_total(data, "supply_amount")),
        _validation_row("세액 합계", _numeric_total(transactions, "tax_amount"), _numeric_total(data, "tax_amount")),
        _validation_row("합계금액 합계", _numeric_total(transactions, "total_amount"), _numeric_total(data, "total_amount")),
        _validation_row("과세·불공 계정분류 건수", len(purchase_tax), len(purchase_known)),
        _validation_row(
            "과세·불공 계정분류 공급가액",
            _numeric_total(purchase_tax, "supply_amount"),
            _numeric_total(purchase_known, "supply_amount"),
            detail="미분류가 있으면 실패합니다.",
        ),
        _validation_row(
            "과세·불공 계정분류 세액",
            _numeric_total(purchase_tax, "tax_amount"),
            _numeric_total(purchase_known, "tax_amount"),
        ),
        _validation_row(
            "세금계산서 매입계 관계",
            _numeric_total(general, "total_amount") + _numeric_total(fixed, "total_amount"),
            _numeric_total(purchase_known, "total_amount"),
            detail="일반매입 + 고정",
        ),
        _validation_row(
            "카드매입 관계",
            _numeric_total(card_tax, "total_amount") + _numeric_total(cash_tax, "total_amount"),
            _numeric_total(card_total, "total_amount"),
            detail="카과 + 현과",
        ),
        _validation_row(
            "과세 매입 총계 관계",
            _numeric_total(purchase_known, "total_amount") + _numeric_total(card_total, "total_amount"),
            _numeric_total(taxable_purchase, "total_amount"),
            detail="세금계산서 매입계 + 카과 + 현과",
        ),
        _validation_row(
            "과매계 관계",
            _numeric_total(taxable_purchase, "total_amount") - _numeric_total(nondeductible, "total_amount"),
            _numeric_total(deductible_summary, "total_amount"),
            detail="과세 매입 총계 - 최종 불공",
        ),
        _validation_row(
            "과세매출 공급가액",
            _numeric_total(sales[sales["original_type"].isin(SALES_TAX_TYPES)], "supply_amount"),
            _numeric_total(taxable_sales, "supply_amount"),
        ),
        _validation_row(
            "카드매출 보조표",
            _numeric_total(sales[sales["original_type"].eq("카과")], "total_amount"),
            _numeric_total(
                sales_summary[sales_summary["item"].eq("카드매출")], "total_amount"
            ),
        ),
        _validation_row(
            "현영매출 보조표",
            _numeric_total(sales[sales["original_type"].eq("현과")], "total_amount"),
            _numeric_total(
                sales_summary[sales_summary["item"].eq("현영매출")], "total_amount"
            ),
        ),
        _validation_row("미분류 건수", 0, int(purchase_tax["account_category"].eq("미분류").sum())),
        _validation_row("불공 판단 보류 건수", 0, int(review["review_status"].eq("판단 보류").sum())),
    ]
    for month, source_month in data.groupby("month", sort=True):
        rows.append(
            _validation_row(
                f"{month} 월별 합계금액",
                _numeric_total(transactions[transactions["month"].eq(month)], "total_amount"),
                _numeric_total(source_month, "total_amount"),
            )
        )
    validation = pd.DataFrame(rows)
    return ProcessingResult(
        transactions=data,
        review=review,
        purchase_by_category=purchase_by_category,
        purchase_summary=purchase_summary,
        sales_by_account=sales_by_account,
        sales_summary=sales_summary,
        validation=validation,
        validation_passed=bool(validation["status"].eq("통과").all()),
    )
