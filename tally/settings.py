# 업체별 업종·계정 예외·불공 후보 키워드·카드공제 대상 여부를 로컬 JSON에 저장합니다.
from __future__ import annotations

from dataclasses import asdict, dataclass, field
import json
from pathlib import Path

from .profiles import INDUSTRIES

DEFAULT_VEHICLE_KEYWORDS = ["차량유지비", "주유", "수리", "타이어", "세차", "자동차"]
DEFAULT_PERSONAL_KEYWORDS = ["개인사용", "가사용", "사적", "가사"]


@dataclass(slots=True)
class CompanySettings:
    name: str
    industry: str = "기본"
    fixed_asset_codes: set[str] = field(default_factory=set)
    account_overrides: dict[str, str] = field(default_factory=dict)
    vehicle_keywords: list[str] = field(default_factory=lambda: list(DEFAULT_VEHICLE_KEYWORDS))
    vendor_keywords: list[str] = field(default_factory=list)
    personal_keywords: list[str] = field(default_factory=lambda: list(DEFAULT_PERSONAL_KEYWORDS))
    card_deduction_eligible: bool = True

    def __post_init__(self) -> None:
        self.name = self.name.strip()
        if not self.name:
            raise ValueError("업체명은 비워둘 수 없습니다.")
        if self.industry not in INDUSTRIES:
            raise ValueError(f"업종은 {', '.join(INDUSTRIES)} 중 하나여야 합니다.")
        self.fixed_asset_codes = {str(code).strip() for code in self.fixed_asset_codes if str(code).strip()}
        self.account_overrides = {
            str(code).strip(): str(category).strip()
            for code, category in self.account_overrides.items()
            if str(code).strip() and str(category).strip()
        }
        self.card_deduction_eligible = bool(self.card_deduction_eligible)

    def to_dict(self) -> dict[str, object]:
        data = asdict(self)
        data["fixed_asset_codes"] = sorted(self.fixed_asset_codes)
        return data

    @classmethod
    def from_dict(cls, data: dict[str, object]) -> "CompanySettings":
        return cls(
            name=str(data["name"]),
            industry=str(data.get("industry", "기본")),
            fixed_asset_codes=set(data.get("fixed_asset_codes", [])),
            account_overrides=dict(data.get("account_overrides", {})),
            vehicle_keywords=list(data.get("vehicle_keywords", DEFAULT_VEHICLE_KEYWORDS)),
            vendor_keywords=list(data.get("vendor_keywords", [])),
            personal_keywords=list(data.get("personal_keywords", DEFAULT_PERSONAL_KEYWORDS)),
            card_deduction_eligible=bool(data.get("card_deduction_eligible", True)),
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
