# 업종별 집계표 틀의 열 구성·계정 규칙·결제 채널·입력 칸을 정의합니다.
from __future__ import annotations

from dataclasses import dataclass, replace

TAXABLE_PAYMENT_DEFAULT = ("카드", "현영", "제로페이")
EXEMPT_PAYMENT = ("카드면세", "현영면세", "면세현금")
INPUT_DEEMED = "의제"
INPUT_COMMON = "공통매입 면세사업분"
INPUT_CARD_EXEMPT = "카드매출 면세분"
# 의제매입세액 계산을 위해 업종마다 세금계산서·면세 매입에 하나씩만 쓰는 상품·원재료 계정입니다.
REVIEW_ACCOUNTS = frozenset({"146", "153", "156"})
EXEMPT_REVIEW_TYPES = frozenset({"면세", "의제", "의제매입", "의제매입세액", "카면", "카영", "현면", "현영"})


@dataclass(frozen=True)
class Profile:
    name: str
    purchase_columns: tuple[str, ...]
    first_label: str
    first_codes: frozenset[str]
    exempt_label: str
    exempt_codes: frozenset[str]
    contractor_label: str | None
    card_columns: tuple[str, ...]
    payment_columns: tuple[str, ...]
    channels: tuple[tuple[str, tuple[str, ...]], ...]
    sales_exempt_fixed: bool
    inputs: frozenset[str]
    taxi_relief: bool = False
    contractor_prefix: str = ""
    exempt_contractor: bool = False
    invoice_account: str = ""
    exempt_account: str = ""

    @property
    def reviews_accounts(self) -> bool:
        return bool(self.invoice_account)

    def needs_review(self, original_type: str, code: str) -> bool:
        """146·153·156 중 이 업종의 세금계산서(카드 과세 포함)·면세 매입에 맞지 않는 계정인지 봅니다."""
        code = str(code).strip()
        if not self.reviews_accounts or code not in REVIEW_ACCOUNTS:
            return False
        expected = self.exempt_account if original_type in EXEMPT_REVIEW_TYPES else self.invoice_account
        return code != expected

    @property
    def exempt_columns(self) -> tuple[str, ...]:
        if self.exempt_contractor and self.contractor_label:
            return (self.exempt_label, self.contractor_label, "기타")
        return (self.exempt_label, "기타")

    def purchase_column(self, code: str, fixed_codes: set[str], overrides: dict[str, str]) -> tuple[str, str]:
        """세금계산서 매입의 열과, 틀에 없는 예외로 기타에 넣었다면 그 설명을 돌려줍니다."""
        code = str(code).strip()
        if code in overrides:
            target = overrides[code]
            if target in self.purchase_columns:
                return target, ""
            return "기타", f"{code}={target}"
        if code in fixed_codes:
            return "고정", ""
        if not (len(code) == 3 and code.isdigit()):
            return "미분류", ""
        if code.startswith("2"):
            return "고정", ""
        if code in self.first_codes:
            return self.first_label, ""
        if self.contractor_label and code.startswith(self.contractor_prefix):
            return self.contractor_label, ""
        return "기타", ""

    def exempt_column(self, code: str, overrides: dict[str, str]) -> str:
        code = str(code).strip()
        if overrides.get(code) in self.exempt_columns:
            return overrides[code]
        if code in self.exempt_codes:
            return self.exempt_label
        if not (len(code) == 3 and code.isdigit()):
            return "미분류"
        if self.exempt_contractor and self.contractor_label and code.startswith(self.contractor_prefix):
            return self.contractor_label
        return "기타"

    def card_column(self, original_type: str, code: str) -> str:
        taxable_card = original_type == "카과"
        taxable_cash = original_type == "현과"
        exempt_card = original_type in {"카면", "카영"}
        if self.name == "건설":
            if taxable_card:
                return "카드"
            if taxable_cash:
                return "현영"
            return "카면" if exempt_card else "현면"
        if self.name == "도소매":
            goods = str(code).strip() in self.first_codes
            if taxable_card:
                return "카드상품" if goods else "카드기타"
            if taxable_cash:
                return "현영상품" if goods else "현영기타"
            return "카드면세" if exempt_card else "현영면세"
        if taxable_card:
            return "카과"
        if taxable_cash:
            return "현과"
        return "카면" if exempt_card else "현면"


EXEMPT_CARD_COLUMNS = frozenset({"카면", "현면", "카드면세", "현영면세"})

PROFILES: dict[str, Profile] = {
    "기본": Profile(
        name="기본",
        purchase_columns=("원재료", "고정", "기타"),
        first_label="원재료",
        first_codes=frozenset({"146", "153", "156"}),
        exempt_label="원재료",
        exempt_codes=frozenset({"146", "153", "156"}),
        contractor_label=None,
        card_columns=("카과", "카면", "현과", "현면"),
        payment_columns=(*TAXABLE_PAYMENT_DEFAULT, "현금", *EXEMPT_PAYMENT),
        channels=(),
        sales_exempt_fixed=True,
        inputs=frozenset({INPUT_DEEMED, INPUT_COMMON, INPUT_CARD_EXEMPT}),
    ),
    "건설": Profile(
        name="건설",
        purchase_columns=("원재료(도급)", "도급", "고정", "기타"),
        first_label="원재료(도급)",
        first_codes=frozenset({"156"}),
        exempt_label="원재료",
        exempt_codes=frozenset({"153"}),
        contractor_label="도급",
        contractor_prefix="6",
        invoice_account="156",
        exempt_account="153",
        card_columns=("카드", "현영"),
        payment_columns=(),
        channels=(),
        sales_exempt_fixed=False,
        inputs=frozenset({INPUT_DEEMED, INPUT_COMMON}),
    ),
    "도소매": Profile(
        name="도소매",
        purchase_columns=("상품", "고정", "기타"),
        first_label="상품",
        first_codes=frozenset({"146"}),
        exempt_label="상품",
        exempt_codes=frozenset({"146"}),
        contractor_label=None,
        invoice_account="146",
        exempt_account="146",
        card_columns=("카드상품", "카드기타", "현영상품", "현영기타"),
        payment_columns=("카드", "현영", "오픈마켓", "제로페이", "현금", *EXEMPT_PAYMENT),
        channels=(
            ("오픈마켓", ("쿠팡", "네이버", "스마트스토어", "11번가", "G마켓", "지마켓", "옥션", "위메프", "티몬", "인터파크")),
        ),
        sales_exempt_fixed=True,
        inputs=frozenset({INPUT_DEEMED, INPUT_COMMON, INPUT_CARD_EXEMPT}),
    ),
    "음식점": Profile(
        name="음식점",
        purchase_columns=("음식재료", "고정", "기타"),
        first_label="음식재료",
        first_codes=frozenset({"146"}),
        exempt_label="원재료",
        exempt_codes=frozenset({"153"}),
        contractor_label=None,
        invoice_account="146",
        exempt_account="153",
        card_columns=("카과", "카면", "현과", "현면"),
        payment_columns=("카드", "현영", "배민", "요기요", "쿠팡이츠", "제로페이", "현금"),
        channels=(
            ("배민", ("배달의민족", "배민", "우아한형제들")),
            ("요기요", ("요기요", "위대한상상")),
            ("쿠팡이츠", ("쿠팡이츠",)),
        ),
        sales_exempt_fixed=False,
        inputs=frozenset({INPUT_DEEMED}),
    ),
    "제조": Profile(
        name="제조",
        purchase_columns=("원재료", "제조경비", "고정", "기타"),
        first_label="원재료",
        first_codes=frozenset({"153"}),
        exempt_label="원재료",
        exempt_codes=frozenset({"153"}),
        contractor_label="제조경비",
        card_columns=("카과", "카면", "현과", "현면"),
        payment_columns=(*TAXABLE_PAYMENT_DEFAULT, "현금", *EXEMPT_PAYMENT),
        channels=(),
        sales_exempt_fixed=True,
        inputs=frozenset({INPUT_DEEMED, INPUT_COMMON, INPUT_CARD_EXEMPT}),
        contractor_prefix="5",
        exempt_contractor=True,
        invoice_account="153",
        exempt_account="153",
    ),
}
# 택시는 기본 틀을 그대로 쓰고 요약에 택시 경감세액과 그 배분만 더합니다.
PROFILES["택시"] = replace(PROFILES["기본"], name="택시", taxi_relief=True)

INDUSTRIES = tuple(PROFILES)


def profile_for(industry: str) -> Profile:
    return PROFILES.get(industry, PROFILES["기본"])
