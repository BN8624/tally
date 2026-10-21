# 앱이 오류를 대화상자로 바꾸고, 업체 설정 파일을 앱 폴더에 두며, 처리 결과를 파일·업체·설정에 묶어 기간 변경과 저장을 관리하는지 검증합니다.
from datetime import date
from decimal import Decimal
import importlib
from pathlib import Path
import sys

import pandas as pd
import pytest

import app as tally_app
from tally import CompanySettings, Period, SettingsStore, process_transactions


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


class _Widget(dict):
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
    _change_period = tally_app.TallyApp._change_period
    _save_output = tally_app.TallyApp._save_output
    _warn_review = tally_app.TallyApp._warn_review

    def __init__(self, store: SettingsStore, *, file_path: str, company: str) -> None:
        self.store = store
        self.file_var = _Var(file_path)
        self.company_var = _Var(company)
        self.period_var = _Var()
        self.status_var = _Var()
        self.result_text = _Widget()
        self.result_banner = _Widget()
        self.period_combo = _Widget()
        self.result = None
        self.result_settings = None
        self.result_context = None
        self.source_data = None
        self.refreshed = 0

    def _refresh_views(self) -> None:
        self.refreshed += 1

    def _refresh_periods(self) -> None:
        pass


def _store_with(tmp_path, *companies: CompanySettings) -> SettingsStore:
    store = SettingsStore(tmp_path / "companies.json")
    for company in companies:
        store.save(company)
    return store


def _source() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "row_id": f"Sheet1:{month}",
                "sheet": "Sheet1",
                "source_row": month,
                "division": "매입",
                "date": date(2026, month, 1),
                "month": f"2026-{month:02d}",
                "vendor": "거래처",
                "item": "품목",
                "supply_amount": Decimal(1000),
                "tax_amount": Decimal(100),
                "total_amount": Decimal(1100),
                "original_type": "과세",
                "account_code": "208",
                "account_name": "비품",
                "card_company": "",
                "card_number": "",
                "electronic": "",
            }
            for month in (5, 8)
        ]
    )


def test_error_text_strips_key_error_quoting() -> None:
    assert tally_app._error_text(KeyError("업체 설정을 찾을 수 없습니다. 업체명=없음")) == "업체 설정을 찾을 수 없습니다. 업체명=없음"
    assert tally_app._error_text(ValueError("금액 형식 오류")) == "금액 형식 오류"


def test_settings_file_stays_in_app_folder_whatever_the_working_directory(tmp_path, monkeypatch) -> None:
    monkeypatch.chdir(tmp_path)
    reloaded = importlib.reload(tally_app)
    assert reloaded.SETTINGS_PATH == Path(reloaded.__file__).resolve().parent / "data" / "companies.json"
    assert tmp_path not in reloaded.SETTINGS_PATH.parents


def test_executable_build_keeps_settings_next_to_the_exe(tmp_path, monkeypatch) -> None:
    executable = tmp_path / "Tally" / "Tally.exe"
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(executable))
    try:
        assert importlib.reload(tally_app).SETTINGS_PATH == executable.resolve().parent / "data" / "companies.json"
    finally:
        monkeypatch.undo()
        importlib.reload(tally_app)


@pytest.mark.parametrize(
    "failure",
    [
        KeyError("업체 설정을 찾을 수 없습니다. 업체명=없음"),
        ValueError("입력 상세 거래의 row_id가 중복되었습니다."),
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
    assert isinstance(PermissionError(), tally_app.PROCESSING_ERRORS)


def test_result_matches_only_the_processed_file_company_and_settings(tmp_path) -> None:
    processed = CompanySettings(name="A")
    store = _store_with(tmp_path, processed, CompanySettings(name="B"))
    context = ("a.xlsx", "A", processed.to_dict())

    assert tally_app._result_matches(context, "a.xlsx", "A", store)
    assert not tally_app._result_matches(context, "b.xlsx", "A", store)
    assert not tally_app._result_matches(context, "a.xlsx", "B", store)
    assert not tally_app._result_matches(context, "a.xlsx", "없음", store)
    assert not tally_app._result_matches(None, "a.xlsx", "A", store)
    store.save(CompanySettings(name="A", industry="건설"))
    assert not tally_app._result_matches(context, "a.xlsx", "A", store)


def test_stale_result_blocks_period_change_and_saving(tmp_path, monkeypatch) -> None:
    recorder = _Recorder()
    monkeypatch.setattr(tally_app, "messagebox", recorder)
    monkeypatch.setattr(tally_app, "filedialog", None)
    processed = CompanySettings(name="A")
    app = _AppStub(_store_with(tmp_path, processed, CompanySettings(name="B")), file_path="a.xlsx", company="B")
    app.result = object()
    app.result_settings = processed
    app.result_context = ("a.xlsx", "A", processed.to_dict())

    app._change_period()
    app._save_output()

    assert [title for title, _ in recorder.warnings] == ["재처리 필요"] * 2
    assert app.refreshed == 0


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

    app._process_file()

    assert recorder.errors == [("처리 실패", "필수 열을 찾지 못했습니다.")]
    assert (app.result, app.result_context, app.result_settings) == (None, None, None)
    assert app.status_var.get().startswith("처리 실패")


def test_period_change_recalculates_with_processed_settings(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(tally_app, "messagebox", _Recorder())
    processed = CompanySettings(name="A", industry="건설")
    app = _AppStub(_store_with(tmp_path, processed), file_path="a.xlsx", company="A")
    app.source_data = _source()
    app.result_settings = processed
    app.result_context = ("a.xlsx", "A", processed.to_dict())
    app.result = process_transactions(app.source_data, processed)
    assert app.result.period == Period(2026, 1, "확정")

    app.period_var.set("2026년 2기 예정")
    app._change_period()

    assert app.result.period == Period(2026, 2, "예정")
    assert app.result.profile.name == "건설"
    assert app.refreshed == 1


def test_manufacturing_processing_warns_about_review_purchases(tmp_path, monkeypatch) -> None:
    recorder = _Recorder()
    monkeypatch.setattr(tally_app, "messagebox", recorder)
    source = _source()
    source.loc[0, "account_code"] = "146"
    source.loc[1, "account_code"] = "156"
    monkeypatch.setattr(tally_app, "parse_workbook", lambda _path: source)
    processed = CompanySettings(name="A", industry="제조")
    app = _AppStub(_store_with(tmp_path, processed), file_path="a.xlsx", company="A")

    app._process_file()
    app.period_var.set("2026년 2기 예정")
    app._change_period()

    assert [title for title, _ in recorder.errors] == ["구분요망", "구분요망"]
    assert "계정 146 매입이 1건" in recorder.errors[0][1]
    assert "계정 156 매입이 1건" in recorder.errors[1][1]


def test_processing_binds_result_to_file_company_and_settings(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(tally_app, "messagebox", _Recorder())
    monkeypatch.setattr(tally_app, "parse_workbook", lambda _path: _source())
    processed = CompanySettings(name="A", industry="도소매")
    app = _AppStub(_store_with(tmp_path, processed), file_path="a.xlsx", company="A")

    app._process_file()

    assert app.result_context == ("a.xlsx", "A", processed.to_dict())
    assert app.result.period == Period(2026, 1, "확정")
    assert app._result_is_current()
