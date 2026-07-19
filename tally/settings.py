# 업체별 계정 분류와 불공 후보 규칙을 로컬 JSON에 저장합니다.
from __future__ import annotations

from dataclasses import asdict, dataclass, field
import json
from pathlib import Path


@dataclass(slots=True)
class CompanySettings:
    name: str
    account_146_label: str = "상품"
    prior_period_credit: int = 0
    card_sales_deduction: int = 0
    fixed_asset_codes: set[str] = field(default_factory=set)
    account_overrides: dict[str, str] = field(default_factory=dict)
    tobacco_vendor_keywords: list[str] = field(
        default_factory=lambda: [
            "삼양인터내셔널",
            "케이티앤지",
            "KT&G",
            "제이티인터내셔널",
            "JTI",
            "로스만스파이스트",
            "로스만스",
        ]
    )
    vehicle_keywords: list[str] = field(
        default_factory=lambda: ["차량유지비", "주유", "수리", "타이어", "세차", "자동차"]
    )
    vendor_keywords: list[str] = field(default_factory=list)
    personal_keywords: list[str] = field(
        default_factory=lambda: ["개인사용", "가사용", "사적", "가사"]
    )

    def __post_init__(self) -> None:
        self.name = self.name.strip()
        if not self.name:
            raise ValueError("업체명은 비워둘 수 없습니다.")
        if self.account_146_label not in {"상품", "음식재료"}:
            raise ValueError("146번 표시 명칭은 상품 또는 음식재료여야 합니다.")
        self.prior_period_credit = int(self.prior_period_credit)
        self.card_sales_deduction = int(self.card_sales_deduction)
        if self.prior_period_credit < 0 or self.card_sales_deduction < 0:
            raise ValueError("예정미환급세액과 카드매출 세액공제는 0 이상이어야 합니다.")
        self.fixed_asset_codes = {str(code).strip() for code in self.fixed_asset_codes if str(code).strip()}
        self.account_overrides = {
            str(code).strip(): str(category).strip()
            for code, category in self.account_overrides.items()
            if str(code).strip() and str(category).strip()
        }

    def to_dict(self) -> dict[str, object]:
        data = asdict(self)
        data["fixed_asset_codes"] = sorted(self.fixed_asset_codes)
        return data

    @classmethod
    def from_dict(cls, data: dict[str, object]) -> "CompanySettings":
        return cls(
            name=str(data["name"]),
            account_146_label=str(data.get("account_146_label", "상품")),
            prior_period_credit=int(data.get("prior_period_credit", 0)),
            card_sales_deduction=int(data.get("card_sales_deduction", 0)),
            fixed_asset_codes=set(data.get("fixed_asset_codes", [])),
            account_overrides=dict(data.get("account_overrides", {})),
            tobacco_vendor_keywords=list(
                data.get(
                    "tobacco_vendor_keywords",
                    [
                        "삼양인터내셔널",
                        "케이티앤지",
                        "KT&G",
                        "제이티인터내셔널",
                        "JTI",
                        "로스만스파이스트",
                        "로스만스",
                    ],
                )
            ),
            vehicle_keywords=list(data.get("vehicle_keywords", [])),
            vendor_keywords=list(data.get("vendor_keywords", [])),
            personal_keywords=list(data.get("personal_keywords", [])),
        )


class SettingsStore:
    def __init__(self, path: str | Path = "data/companies.json") -> None:
        self.path = Path(path)

    def load_all(self) -> dict[str, CompanySettings]:
        if not self.path.exists():
            return {}
        raw = json.loads(self.path.read_text(encoding="utf-8"))
        return {name: CompanySettings.from_dict(value) for name, value in raw.items()}

    def save(self, settings: CompanySettings) -> None:
        companies = self.load_all()
        companies[settings.name] = settings
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {name: company.to_dict() for name, company in sorted(companies.items())}
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def get(self, name: str) -> CompanySettings:
        try:
            return self.load_all()[name]
        except KeyError as exc:
            raise KeyError(f"업체 설정을 찾을 수 없습니다. 업체명={name}") from exc
