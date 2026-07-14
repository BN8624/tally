# 업체 설정의 우선순위 정보가 로컬에 저장되고 재사용되는지 검증합니다.
from tally.settings import CompanySettings, SettingsStore


def test_company_settings_round_trip(tmp_path) -> None:
    store = SettingsStore(tmp_path / "companies.json")
    expected = CompanySettings(
        name="테스트상사",
        account_146_label="음식재료",
        fixed_asset_codes={"201", "202"},
        account_overrides={"899": "제조경비"},
        vendor_keywords=["특정주유소"],
    )
    store.save(expected)
    actual = store.get("테스트상사")
    assert actual.account_146_label == "음식재료"
    assert actual.fixed_asset_codes == {"201", "202"}
    assert actual.account_overrides == {"899": "제조경비"}
    assert actual.vendor_keywords == ["특정주유소"]

