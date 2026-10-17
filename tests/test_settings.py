# 업체 설정의 업종·예외·키워드·카드공제 대상이 저장되고 이전 설정 파일도 읽히는지 검증합니다.
import pytest

from tally.settings import CompanySettings, SettingsStore


def test_company_settings_round_trip(tmp_path) -> None:
    store = SettingsStore(tmp_path / "companies.json")
    store.save(
        CompanySettings(
            name="테스트상사",
            industry="음식점",
            fixed_asset_codes={"201", "202"},
            account_overrides={"899": "기타"},
            vendor_keywords=["특정주유소"],
            card_deduction_eligible=False,
        )
    )
    actual = store.get("테스트상사")
    assert actual.industry == "음식점"
    assert actual.fixed_asset_codes == {"201", "202"}
    assert actual.account_overrides == {"899": "기타"}
    assert actual.vendor_keywords == ["특정주유소"]
    assert actual.card_deduction_eligible is False


def test_legacy_settings_are_read_with_defaults() -> None:
    legacy = CompanySettings.from_dict(
        {
            "name": "기존업체",
            "account_146_label": "음식재료",
            "prior_period_credit": 100,
            "card_sales_deduction": 20,
            "tobacco_vendor_keywords": ["KT&G"],
        }
    )
    assert legacy.industry == "기본"
    assert legacy.card_deduction_eligible is True
    assert legacy.vehicle_keywords == ["차량유지비", "주유", "수리", "타이어", "세차", "자동차"]
    assert set(legacy.to_dict()) == {
        "name",
        "industry",
        "fixed_asset_codes",
        "account_overrides",
        "vehicle_keywords",
        "vendor_keywords",
        "personal_keywords",
        "card_deduction_eligible",
    }


def test_unknown_industry_is_rejected() -> None:
    with pytest.raises(ValueError, match="업종은"):
        CompanySettings(name="업체", industry="숙박")
