# 앱이 처리·판정·저장 오류를 대화상자로 바꾸고 창을 유지하는지 검증합니다.
import pytest

import app as tally_app


class _Recorder:
    def __init__(self) -> None:
        self.errors: list[tuple[str, str]] = []

    def showerror(self, title: str, message: str, parent: object = None) -> None:
        self.errors.append((title, message))


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
