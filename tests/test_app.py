# 앱이 처리·판정·저장 오류를 대화상자로 바꾸고 창을 유지하는지 검증합니다.
from datetime import date
from decimal import Decimal
from types import SimpleNamespace

import pandas as pd
import pytest

import app as tally_app
from tally import CompanySettings, SettingsStore


class _Recorder:
    def __init__(self) -> None:
        self.errors: list[tuple[str, str]] = []
        self.warnings: list[tuple[str, str]] = []

    def showerror(self, title: str, message: str, parent: object = None) -> None:
        self.errors.append((title, message))

    def showwarning(self, title: str, message: str, parent: object = None) -> None:
        self.warnings.append((title, message))


class _Var:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class _Widget:
    def configure(self, **_options) -> None:
        pass

    def delete(self, *_args) -> None:
        pass


class _AppStub:
    _guarded = tally_app.TallyApp._guarded
    _process_file = tally_app.TallyApp._process_file
    _clear_result = tally_app.TallyApp._clear_result
    _result_is_current = tally_app.TallyApp._result_is_current
    _require_current_result = tally_app.TallyApp._require_current_result
    _apply_selected = tally_app.TallyApp._apply_selected
    _apply_same_condition = tally_app.TallyApp._apply_same_condition
    _save_output = tally_app.TallyApp._save_output

    def __init__(self, store: SettingsStore, *, file_path: str, company: str) -> None:
        self.store = store
        self.file_var = _Var(file_path)
        self.company_var = _Var(company)
        self.status_var = _Var()
        self.result_text = _Widget()
        self.result_banner = _Widget()
        self.result = None
        self.result_settings = None
        self.result_context = None
        self.source_data = None
        self.decisions: dict[str, dict[str, str]] = {}

    def _refresh_review_tree(self) -> None:
        pass

    def _selected_row_id(self) -> str:
        raise AssertionError("재처리 전에는 판정 대상을 고르지 않아야 합니다.")


def _store_with(tmp_path, *companies: CompanySettings) -> SettingsStore:
    store = SettingsStore(tmp_path / "companies.json")
    for company in companies:
        store.save(company)
    return store


def test_result_matches_only_the_processed_file_company_and_settings(tmp_path) -> None:
    processed = CompanySettings(name="A")
    store = _store_with(tmp_path, processed, CompanySettings(name="B"))
    context = ("a.xlsx", "A", processed.to_dict())

    assert tally_app._result_matches(context, "a.xlsx", "A", store)
    assert not tally_app._result_matches(context, "b.xlsx", "A", store)
    assert not tally_app._result_matches(context, "a.xlsx", "B", store)
    assert not tally_app._result_matches(context, "a.xlsx", "없음", store)
    assert not tally_app._result_matches(None, "a.xlsx", "A", store)
    store.save(CompanySettings(name="A", fixed_asset_codes={"212"}))
    assert not tally_app._result_matches(context, "a.xlsx", "A", store)


def test_stale_result_blocks_decisions_and_saving(tmp_path, monkeypatch) -> None:
    recorder = _Recorder()
    monkeypatch.setattr(tally_app, "messagebox", recorder)
    monkeypatch.setattr(tally_app, "filedialog", None)
    processed = CompanySettings(name="A")
    app = _AppStub(_store_with(tmp_path, processed, CompanySettings(name="B")), file_path="a.xlsx", company="B")
    app.result = object()
    app.result_settings = processed
    app.result_context = ("a.xlsx", "A", processed.to_dict())

    app._apply_selected()
    app._apply_same_condition()
    app._save_output()

    assert [title for title, _ in recorder.warnings] == ["재처리 필요"] * 3
    assert app.decisions == {}


def test_failed_processing_clears_previous_result(tmp_path, monkeypatch) -> None:
    recorder = _Recorder()
    monkeypatch.setattr(tally_app, "messagebox", recorder)

    def failing_parse(_path: str) -> None:
        raise tally_app.InputWorkbookError("필수 열을 찾지 못했습니다.")

    monkeypatch.setattr(tally_app, "parse_workbook", failing_parse)
    processed = CompanySettings(name="A")
    app = _AppStub(_store_with(tmp_path, processed), file_path="new.xlsx", company="A")
    app.result = object()
    app.result_settings = processed
    app.result_context = ("old.xlsx", "A", processed.to_dict())
    app.decisions = {"Sheet1:2": {"decision": "판단 보류"}}

    app._process_file()

    assert recorder.errors == [("처리 실패", "필수 열을 찾지 못했습니다.")]
    assert (app.result, app.result_context, app.result_settings, app.decisions) == (None, None, None, {})
    assert app.status_var.get().startswith("처리 실패")


def test_decisions_recalculate_with_settings_used_for_processing() -> None:
    source = pd.DataFrame(
        [
            {
                "row_id": "Sheet1:2",
                "sheet": "Sheet1",
                "source_row": 2,
                "division": "매입",
                "date": date(2026, 4, 1),
                "month": "2026-04",
                "vendor": "거래처",
                "item": "품목",
                "supply_amount": Decimal(1000),
                "tax_amount": Decimal(100),
                "total_amount": Decimal(1100),
                "original_type": "과세",
                "account_code": "813",
                "account_name": "비품",
                "card_company": "",
                "card_number": "",
            }
        ]
    )
    app = SimpleNamespace(
        source_data=source,
        result_settings=CompanySettings(name="A", fixed_asset_codes={"813"}),
        decisions={},
        result=None,
        _refresh_views=lambda: None,
    )

    tally_app.TallyApp._recalculate(app)

    assert app.result.transactions["account_category"].tolist() == ["고정"]


def test_error_text_strips_key_error_quoting() -> None:
    missing = KeyError("업체 설정을 찾을 수 없습니다. 업체명=없음")
    assert tally_app._error_text(missing) == "업체 설정을 찾을 수 없습니다. 업체명=없음"
    assert tally_app._error_text(ValueError("금액 형식 오류")) == "금액 형식 오류"


@pytest.mark.parametrize(
    "failure",
    [
        KeyError("업체 설정을 찾을 수 없습니다. 업체명=없음"),
        ValueError("지원하지 않는 불공 판정입니다."),
        OSError("엑셀 파일을 열 수 없습니다."),
        tally_app.InputWorkbookError("필수 열을 찾지 못했습니다."),
    ],
)
def test_guarded_reports_failure_instead_of_raising(monkeypatch, failure) -> None:
    recorder = _Recorder()
    monkeypatch.setattr(tally_app, "messagebox", recorder)

    def action() -> None:
        raise failure

    assert tally_app.TallyApp._guarded(object(), "처리 실패", action) is False
    assert recorder.errors == [("처리 실패", tally_app._error_text(failure))]


def test_guarded_passes_through_on_success(monkeypatch) -> None:
    recorder = _Recorder()
    monkeypatch.setattr(tally_app, "messagebox", recorder)

    assert tally_app.TallyApp._guarded(object(), "처리 실패", lambda: None) is True
    assert recorder.errors == []


def test_key_error_from_missing_company_is_a_handled_failure() -> None:
    assert isinstance(KeyError(), tally_app.PROCESSING_ERRORS)
    assert isinstance(PermissionError(), tally_app.PROCESSING_ERRORS)
