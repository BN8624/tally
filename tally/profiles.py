# 업종별 집계표 틀의 열 구성·계정 규칙·결제 채널·입력 칸을 정의합니다.
from __future__ import annotations

from dataclasses import dataclass

TAXABLE_PAYMENT_DEFAULT = ("카드", "현영", "제로페이")
EXEMPT_PAYMENT = ("카드면세", "현영면세", "면세현금")
INPUT_DEEMED = "의제"
INPUT_COMMON = "공통매입 면세사업분"
INPUT_CARD_EXEMPT = "카드매출 면세분"


@dataclass(frozen=True)
class Profile:
    name: str
    purchase_columns: tuple[str, ...]
    first_label: str
    first_codes: frozenset[str]
    contractor_label: str | None
    card_columns: tuple[str, ...]
    payment_columns: tuple[str, ...]
    channels: tuple[tuple[str, tuple[str, ...]], ...]
    sales_exempt_fixed: bool
    inputs: frozenset[str]

    @property
    def exempt_columns(self) -> tuple[str, ...]:
        return (self.first_label, "기타")

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
        if self.contractor_label and code.startswith("6"):
            return self.contractor_label, ""
        return "기타", ""

    def exempt_column(self, code: str, overrides: dict[str, str]) -> str:
        code = str(code).strip()
        if overrides.get(code) == self.first_label or code in self.first_codes:
            return self.first_label
        if not (len(code) == 3 and code.isdigit()):
            return "미분류"
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
        contractor_label=None,
        card_columns=("카과", "카면", "현과", "현면"),
        payment_columns=(*TAXABLE_PAYMENT_DEFAULT, "현금", *EXEMPT_PAYMENT),
        channels=(),
        sales_exempt_fixed=True,
        inputs=frozenset({INPUT_DEEMED, INPUT_COMMON, INPUT_CARD_EXEMPT}),
    ),
    "건설": Profile(
        name="건설",
        purchase_columns=("원재료", "도급", "고정", "기타"),
        first_label="원재료",
        first_codes=frozenset({"156"}),
        contractor_label="도급",
        card_columns=("카드", "현영"),
        payment_columns=(),
        channels=(),
        sales_exempt_fixed=False,
        inputs=frozenset(),
    ),
    "도소매": Profile(
        name="도소매",
        purchase_columns=("상품", "고정", "기타"),
        first_label="상품",
        first_codes=frozenset({"146"}),
        contractor_label=None,
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
        contractor_label=None,
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
}

INDUSTRIES = tuple(PROFILES)


def profile_for(industry: str) -> Profile:
    return PROFILES.get(industry, PROFILES["기본"])
