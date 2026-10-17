# 상세 거래를 업종 틀의 블록·열에 월별로 배치하고 요약 값과 검산을 계산합니다.
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, Decimal
from typing import Iterable

import pandas as pd

from .profiles import EXEMPT_CARD_COLUMNS, EXEMPT_PAYMENT, Profile, profile_for
from .settings import CompanySettings

ZERO = Decimal(0)
PURCHASE_INVOICE_TYPES = {"과세", "불공", "공통", "공통매입", "공통매입세액"}
PURCHASE_EXEMPT_TYPES = {"면세", "의제", "의제매입", "의제매입세액"}
PURCHASE_CARD_TYPES = {"카과", "현과", "카면", "카영", "현면", "현영"}
SALES_PAYMENT_TYPES = {"카과", "현과", "건별", "카면", "카영", "현면", "현영", "면건"}
CANDIDATE_TYPES = {"과세", "카과", "현과"}

PURCHASE = "purchase"
EXEMPT = "exempt"
NONDEDUCTIBLE = "nondeductible"
CARD = "card"
SALES = "sales"
PAYMENT = "payment"
BLOCK_TITLES = {
    PURCHASE: "매입 세금계산서",
    EXEMPT: "면세매입",
    NONDEDUCTIBLE: "불공",
    CARD: "카드외 매입",
    SALES: "매출 세금계산서",
    PAYMENT: "결제 채널 매출",
}
DEFAULT_PAYMENT_ORDER = ("카드", "현영", "제로페이", "현금", *EXEMPT_PAYMENT)


@dataclass(frozen=True)
class Period:
    year: int
    half: int
    kind: str

    @property
    def months(self) -> list[str]:
        start = 1 if self.half == 1 else 7
        if self.kind == "확정":
            start += 3
        count = 6 if self.kind == "확정6" else 3
        return [f"{self.year}-{month:02d}" for month in range(start, start + count)]

    @property
    def label(self) -> str:
        return f"{self.year}년 {self.half}기 {'확정(6개월)' if self.kind == '확정6' else self.kind}"


def _quarter_period(month: str) -> Period:
    year, number = int(month[:4]), int(month[5:7])
    return Period(year, 1 if number <= 6 else 2, "예정" if number in {1, 2, 3, 7, 8, 9} else "확정")


def available_periods(months: Iterable[str]) -> list[Period]:
    found: set[Period] = set()
    for month in months:
        quarter = _quarter_period(month)
        found.update({quarter, Period(quarter.year, quarter.half, "확정6")})
    order = {"예정": 0, "확정": 1, "확정6": 2}
    return sorted(found, key=lambda period: (period.year, period.half, order[period.kind]))


def default_period(months: Iterable[str]) -> Period | None:
    """마지막 거래 월까지 끝난 가장 최근 분기, 없으면 마지막 거래 월의 분기를 고릅니다."""
    ordered = sorted(months)
    if not ordered:
        return None
    last = ordered[-1]
    completed = [
        period
        for period in available_periods(ordered)
        if period.kind != "확정6" and period.months[-1] <= last
    ]
    return completed[-1] if completed else _quarter_period(last)


@dataclass(frozen=True)
class Column:
    label: str
    kind: str
    sum_of: tuple[str, ...] = ()
    extra: bool = False


@dataclass(slots=True)
class SummaryRow:
    key: str
    label: str
    supply: Decimal | None
    tax: Decimal | None


@dataclass(slots=True)
class ProcessingResult:
    transactions: pd.DataFrame
    settings: CompanySettings
    profile: Profile
    period: Period | None
    periods: list[Period]
    layout: dict[str, list[Column]]
    cells: pd.DataFrame
    summary: list[SummaryRow]
    candidates: pd.DataFrame
    unclassified: pd.DataFrame
    validation: pd.DataFrame
    validation_passed: bool
    notes: list[str] = field(default_factory=list)

    def summary_value(self, key: str, field_name: str = "tax") -> Decimal:
        row = next(row for row in self.summary if row.key == key)
        value = row.supply if field_name == "supply" else row.tax
        return value if value is not None else ZERO


def d0(value: Decimal) -> Decimal:
    return Decimal(int(value.to_integral_value(rounding=ROUND_DOWN)))


def u0(value: Decimal) -> Decimal:
    return Decimal(int(value.to_integral_value(rounding=ROUND_CEILING if value >= 0 else ROUND_FLOOR)))


def _matches(value: object, keywords: Iterable[str]) -> bool:
    normalized = str(value or "").casefold().replace(" ", "")
    return any(keyword.casefold().replace(" ", "") in normalized for keyword in keywords)


def _is_zero_pay(row: pd.Series) -> bool:
    context = " ".join(str(row.get(column, "")) for column in ("vendor", "item", "card_company", "card_number"))
    normalized = context.casefold().replace(" ", "")
    return "제로페이" in normalized or ("제로" in normalized and "페이" in normalized)


def find_nondeductible_candidate(row: pd.Series, settings: CompanySettings) -> str:
    if row["division"] != "매입" or row["original_type"] not in CANDIDATE_TYPES:
        return ""
    account = f"{row['account_code']} {row['account_name']}".casefold()
    context = f"{row['vendor']} {row['item']} {row['account_name']}".casefold()
    reasons: list[str] = []
    if any(keyword in account for keyword in ("접대비", "기업업무추진비")):
        reasons.append(f"계정과목 {row['account_code']} {row['account_name']}")
    for label, keywords in (
        ("차량 관련 값", settings.vehicle_keywords),
        ("거래처 후보", settings.vendor_keywords),
        ("사업 무관 가능성", settings.personal_keywords),
    ):
        hits = [keyword for keyword in keywords if keyword.casefold() in context]
        if hits:
            reasons.append(f"{label}: {', '.join(hits)}")
    return "; ".join(reasons)


def _payment_column(row: pd.Series, profile: Profile) -> str:
    original = row["original_type"]
    if original == "카과":
        if _is_zero_pay(row):
            return "제로페이"
        for label, keywords in profile.channels:
            if any(_matches(row.get(column, ""), keywords) for column in ("vendor", "item", "card_company")):
                return label
        return "카드"
    return {"현과": "현영", "건별": "현금", "카면": "카드면세", "카영": "카드면세", "현면": "현영면세", "현영": "현영면세", "면건": "면세현금"}[original]


def _sales_column(row: pd.Series) -> str:
    original = row["original_type"]
    if original == "영세":
        return "영세"
    if original == "면세":
        return "면세"
    code = str(row["account_code"]).strip()
    if len(code) == 3 and code.isdigit() and code.startswith("2"):
        return "고정"
    name = str(row["account_name"]).strip() or "세금계산서"
    return "수입금액" if name.replace(" ", "") == "회사설정계정과목" else name


def _place(row: pd.Series, settings: CompanySettings, profile: Profile) -> tuple[str, str, str]:
    division, original = row["division"], row["original_type"]
    code = row["account_code"]
    if division == "매입":
        if original in PURCHASE_INVOICE_TYPES:
            column, note = profile.purchase_column(code, settings.fixed_asset_codes, settings.account_overrides)
            return PURCHASE, column, note
        if original in PURCHASE_EXEMPT_TYPES:
            return EXEMPT, profile.exempt_column(code, settings.account_overrides), ""
        if original == "영세":
            return EXEMPT, "영세", ""
        if original in PURCHASE_CARD_TYPES:
            return CARD, profile.card_column(original, code), ""
    elif division == "매출":
        if original in {"과세", "영세", "면세"}:
            return SALES, _sales_column(row), ""
        if original in SALES_PAYMENT_TYPES:
            return PAYMENT, _payment_column(row, profile), ""
    return "", "", ""


def _ordered_extras(present: Iterable[str], fixed: Iterable[str]) -> list[str]:
    fixed_set = set(fixed)
    return sorted({label for label in present if label and label not in fixed_set})


def _build_layout(placed: pd.DataFrame, profile: Profile) -> tuple[dict[str, list[Column]], list[str]]:
    present = {block: set(group["column"]) for block, group in placed.groupby("block")}
    layout: dict[str, list[Column]] = {}
    added: list[str] = []

    purchase_extras = _ordered_extras(present.get(PURCHASE, ()), profile.purchase_columns)
    purchase_labels = [*profile.purchase_columns, *purchase_extras]
    layout[PURCHASE] = [
        *(Column(label, "tax", extra=label in purchase_extras) for label in purchase_labels),
        Column("계", "tax", tuple(purchase_labels)),
    ]

    exempt_present = present.get(EXEMPT, set())
    exempt_extras = _ordered_extras(exempt_present - {"영세"}, profile.exempt_columns)
    exempt_labels = [*profile.exempt_columns, *exempt_extras]
    layout[EXEMPT] = [
        *(Column(label, "exempt", extra=label in exempt_extras) for label in exempt_labels),
        Column("면세 계", "exempt", tuple(exempt_labels)),
        *([Column("영세", "exempt", extra=True)] if "영세" in exempt_present else []),
    ]
    layout[NONDEDUCTIBLE] = [Column("불공", "tax")]

    card_extras = _ordered_extras(present.get(CARD, ()), profile.card_columns)
    card_labels = [*profile.card_columns, *card_extras]
    taxable_cards = tuple(label for label in card_labels if label not in EXEMPT_CARD_COLUMNS)
    layout[CARD] = [
        *(
            Column(label, "exempt" if label in EXEMPT_CARD_COLUMNS else "tax", extra=label in card_extras)
            for label in card_labels
        ),
        Column("공제 계", "tax", taxable_cards),
    ]

    sales_present = present.get(SALES, set())
    accounts = sorted(
        label for label in sales_present if label not in {"고정", "영세", "면세"}
    ) or ["세금계산서"]
    accounts = _account_order(placed, accounts)
    taxable_sales = [*accounts, "고정"]
    layout[SALES] = [
        *(Column(label, "tax") for label in taxable_sales),
        Column("계", "tax", tuple(taxable_sales)),
        *([Column("영세", "exempt", extra=True)] if "영세" in sales_present else []),
        *(
            [Column("면세", "exempt", extra=not profile.sales_exempt_fixed)]
            if profile.sales_exempt_fixed or "면세" in sales_present
            else []
        ),
    ]

    payment_present = present.get(PAYMENT, set())
    if profile.payment_columns or payment_present:
        fixed = list(profile.payment_columns) or [label for label in DEFAULT_PAYMENT_ORDER if label in payment_present]
        extras = _ordered_extras(payment_present, fixed)
        labels = [*fixed, *extras]
        taxable = [label for label in labels if label not in {"현금", *EXEMPT_PAYMENT}]
        exempt = [label for label in labels if label in EXEMPT_PAYMENT]
        extra_set = set(extras) if profile.payment_columns else set(labels)
        layout[PAYMENT] = [
            *(Column(label, "amount", extra=label in extra_set) for label in taxable),
            Column("과세 계", "amount", tuple(taxable)),
            *([Column("현금", "amount", extra="현금" in extra_set)] if "현금" in labels else []),
            *(Column(label, "amount", extra=label in extra_set) for label in exempt),
        ]
        if not profile.payment_columns:
            added.append(f"{BLOCK_TITLES[PAYMENT]} 블록")

    for block, columns in layout.items():
        added.extend(f"{BLOCK_TITLES[block]} · {column.label}" for column in columns if column.extra)
    return layout, added


def _account_order(placed: pd.DataFrame, accounts: list[str]) -> list[str]:
    sales = placed[(placed["block"] == SALES) & placed["column"].isin(accounts)]
    first_code = sales.groupby("column")["account_code"].min().to_dict() if not sales.empty else {}
    return sorted(accounts, key=lambda label: (str(first_code.get(label, "999")), label))


def _aggregate(placed: pd.DataFrame) -> pd.DataFrame:
    columns = ["block", "column", "month", "count", "supply", "tax", "total"]
    if placed.empty:
        return pd.DataFrame(columns=columns)
    grouped = (
        placed.groupby(["block", "column", "month"], sort=True)
        .agg(
            count=("row_id", "size"),
            supply=("supply_amount", "sum"),
            tax=("tax_amount", "sum"),
            total=("total_amount", "sum"),
        )
        .reset_index()
    )
    return grouped[columns]


def cell_total(cells: pd.DataFrame, block: str, labels: Iterable[str], field_name: str) -> Decimal:
    labels = set(labels)
    rows = cells[(cells["block"] == block) & cells["column"].isin(labels)]
    total = sum((Decimal(str(value)) for value in rows[field_name]), ZERO)
    return Decimal(int(total)) if total == total.to_integral_value() else total


def _summary(cells: pd.DataFrame, layout: dict[str, list[Column]], settings: CompanySettings) -> list[SummaryRow]:
    def data_labels(block: str, kind: str | None = None) -> list[str]:
        return [column.label for column in layout.get(block, []) if not column.sum_of and (kind is None or column.kind == kind)]

    def sum_labels(block: str, sum_label: str) -> tuple[str, ...]:
        return next((column.sum_of for column in layout.get(block, []) if column.label == sum_label), ())

    invoice_labels = sum_labels(PURCHASE, "계")
    invoice_supply = cell_total(cells, PURCHASE, invoice_labels, "supply")
    invoice_tax = d0(invoice_supply * Decimal("0.1"))
    fixed_supply = cell_total(cells, PURCHASE, ["고정"], "supply")
    fixed_tax = cell_total(cells, PURCHASE, ["고정"], "tax")
    card_labels = sum_labels(CARD, "공제 계")
    card_supply = cell_total(cells, CARD, card_labels, "supply")
    card_tax = cell_total(cells, CARD, card_labels, "tax")
    nondeductible_supply = cell_total(cells, NONDEDUCTIBLE, ["불공"], "supply")
    nondeductible_tax = cell_total(cells, NONDEDUCTIBLE, ["불공"], "tax")
    purchase_supply = invoice_supply + card_supply - nondeductible_supply
    purchase_tax = invoice_tax + card_tax - nondeductible_tax

    sales_labels = sum_labels(SALES, "계")
    sales_supply = cell_total(cells, SALES, sales_labels, "supply")
    sales_tax = d0(sales_supply * Decimal("0.1"))
    zero_rated = cell_total(cells, SALES, ["영세"], "supply")
    payment_taxable = cell_total(cells, PAYMENT, sum_labels(PAYMENT, "과세 계"), "total")
    card_out_supply = u0(payment_taxable / Decimal("1.1"))
    card_out_tax = payment_taxable - card_out_supply
    cash_total = cell_total(cells, PAYMENT, ["현금"], "total")
    other_supply = u0(cash_total / Decimal("1.1"))
    other_tax = cash_total - other_supply
    taxable_supply = sales_supply + zero_rated + card_out_supply + other_supply
    taxable_tax = sales_tax + card_out_tax + other_tax
    exempt_invoice = cell_total(cells, SALES, ["면세"], "supply")
    exempt_payment = cell_total(cells, PAYMENT, EXEMPT_PAYMENT, "total")
    exempt_total = exempt_invoice + exempt_payment
    payable = taxable_tax - purchase_tax
    card_deduction = (
        max(ZERO, min(d0(payment_taxable * Decimal(13) / Decimal(1000)), payable))
        if settings.card_deduction_eligible
        else ZERO
    )
    return [
        SummaryRow("general", "일반", invoice_supply - fixed_supply, invoice_tax - fixed_tax),
        SummaryRow("fixed", "고정", fixed_supply, fixed_tax),
        SummaryRow("invoice", "세금계산서 계", invoice_supply, invoice_tax),
        SummaryRow("card", "카드", card_supply, card_tax),
        SummaryRow("nondeductible", "불공", nondeductible_supply, nondeductible_tax),
        SummaryRow("purchase", "매입계", purchase_supply, purchase_tax),
        SummaryRow("sales_invoice", "세금계산서", sales_supply, sales_tax),
        SummaryRow("zero_rated", "영세", zero_rated, ZERO),
        SummaryRow("card_out", "카드외", card_out_supply, card_out_tax),
        SummaryRow("other", "기타(건별)", other_supply, other_tax),
        SummaryRow("taxable", "과세계", taxable_supply, taxable_tax),
        SummaryRow("exempt_invoice", "면세 계산서", exempt_invoice, None),
        SummaryRow("exempt_payment", "면세 결제", exempt_payment, None),
        SummaryRow("exempt", "면세계", exempt_total, None),
        SummaryRow("total_sales", "총매출", taxable_supply + exempt_total, None),
        SummaryRow("payable", "납부(환급)세액", None, payable),
        SummaryRow("card_deduction", "카드공제", None, card_deduction),
        SummaryRow("final", "차감 납부할 세액", None, payable - card_deduction),
    ]


def _validation_row(check: str, expected: object, actual: object, *, detail: str = "") -> dict[str, object]:
    difference: object = ""
    if isinstance(expected, (int, Decimal)) and isinstance(actual, (int, Decimal)):
        difference = actual - expected
    return {
        "check": check,
        "status": "통과" if expected == actual else "실패",
        "expected": expected,
        "actual": actual,
        "difference": difference,
        "detail": detail,
    }


def _information_row(check: str, value: object, *, detail: str = "") -> dict[str, object]:
    return {"check": check, "status": "정보", "expected": "", "actual": value, "difference": "", "detail": detail}


def _decimal_sum(frame: pd.DataFrame, column: str) -> Decimal:
    return sum((Decimal(str(value)) for value in frame[column]), ZERO)


def process_transactions(
    transactions: pd.DataFrame,
    settings: CompanySettings,
    period: Period | None = None,
) -> ProcessingResult:
    data = transactions.copy()
    if "electronic" not in data:
        data["electronic"] = ""
    if data["row_id"].duplicated().any():
        raise ValueError("입력 상세 거래의 row_id가 중복되었습니다.")
    profile = profile_for(settings.industry)
    months = sorted(data["month"].dropna().astype(str).unique())
    periods = available_periods(months)
    period = period or default_period(months)
    period_months = set(period.months) if period else set()

    data["in_period"] = data["month"].isin(period_months)
    placements = [_place(row, settings, profile) if row["in_period"] else ("", "", "") for _, row in data.iterrows()]
    data["block"] = [block for block, _, _ in placements]
    data["column"] = [column for _, column, _ in placements]
    data["override_note"] = [note for _, _, note in placements]
    data["candidate_reason"] = [
        find_nondeductible_candidate(row, settings) if row["in_period"] else "" for _, row in data.iterrows()
    ]

    placed = data[data["block"].ne("")]
    nondeductible = placed[(placed["block"] == PURCHASE) & placed["original_type"].eq("불공")].assign(
        block=NONDEDUCTIBLE, column="불공"
    )
    layout, added = _build_layout(placed, profile)
    cells = _aggregate(pd.concat([placed, nondeductible], ignore_index=True))
    summary = _summary(cells, layout, settings)

    in_period = data[data["in_period"]]
    unsupported = in_period[in_period["block"].eq("")]
    unclassified = in_period[in_period["column"].eq("미분류")]
    candidates = in_period[in_period["candidate_reason"].ne("")]
    excluded = data[~data["in_period"]]
    override_notes = sorted(set(note for note in data["override_note"] if note))

    rows = [
        _validation_row("기간 안 상세 거래 없음", 0, int(in_period.empty), detail="선택한 과세기간에 상세 거래가 없으면 실패합니다."),
        _validation_row(
            "미집계 유형 건수",
            0,
            len(unsupported),
            detail=(
                f"집계할 수 없는 유형: {', '.join(sorted(set(unsupported['division'] + ' ' + unsupported['original_type'])))}"
                if not unsupported.empty
                else "기간 안 모든 유형이 집계 대상입니다."
            ),
        ),
    ]
    for division in ("매입", "매출"):
        source = in_period[in_period["division"] == division]
        kept = placed[placed["division"] == division]
        rows.extend(
            [
                _validation_row(f"{division} 집계 보존 건수", len(source), len(kept)),
                _validation_row(f"{division} 집계 보존 공급가액", _decimal_sum(source, "supply_amount"), _decimal_sum(kept, "supply_amount")),
                _validation_row(f"{division} 집계 보존 세액", _decimal_sum(source, "tax_amount"), _decimal_sum(kept, "tax_amount")),
            ]
        )
    rows.append(
        _validation_row("계정 미분류 건수", 0, len(unclassified), detail="계정코드가 비었거나 세 자리 숫자가 아닌 매입입니다.")
    )
    rows.extend(
        [
            _information_row("과세기간", period.label if period else "없음", detail=", ".join(period.months) if period else ""),
            _information_row("기간 안 상세 거래 건수", len(in_period)),
            _information_row(
                "기간 밖 제외 건수",
                len(excluded),
                detail=", ".join(sorted(excluded["month"].astype(str).unique())),
            ),
            _information_row("틀에 없어 기타로 넣은 계정 예외", len(override_notes), detail=", ".join(override_notes)),
            _information_row("틀에 없어 추가한 열·블록", len(added), detail=", ".join(added)),
            _information_row("불공 후보 건수", len(candidates), detail="판정 없이 불공후보 시트에 목록만 표시합니다."),
            _information_row("기간 안 공급가액 합계", _decimal_sum(in_period, "supply_amount")),
            _information_row("기간 안 세액 합계", _decimal_sum(in_period, "tax_amount")),
            _information_row("기간 안 합계금액 합계", _decimal_sum(in_period, "total_amount")),
        ]
    )
    for month in period.months if period else []:
        rows.append(_information_row(f"{month} 합계금액", _decimal_sum(in_period[in_period["month"] == month], "total_amount")))
    validation = pd.DataFrame(rows)
    return ProcessingResult(
        transactions=data,
        settings=settings,
        profile=profile,
        period=period,
        periods=periods,
        layout=layout,
        cells=cells,
        summary=summary,
        candidates=candidates,
        unclassified=unclassified,
        validation=validation,
        validation_passed=bool(validation.loc[validation["status"].ne("정보"), "status"].eq("통과").all()),
        notes=added,
    )
