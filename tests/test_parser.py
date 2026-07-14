# 입력 엑셀 파서의 상세행 판정과 오류 처리를 검증합니다.
from io import BytesIO

from openpyxl import Workbook
import pytest

from tally.parser import InputWorkbookError, parse_workbook


HEADERS = [
    "구분",
    "전표일자",
    "거래처",
    "품명",
    "공급가액",
    "부가세",
    "합계",
    "매입/매출\n유형",
    "Code",
    "계정과목",
    "Code",
    "부서명",
    "Code",
    "카드사명",
    "카드번호",
]


def workbook_bytes(rows: list[list[object]], headers: list[str] | None = None) -> BytesIO:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "전체매입매출"
    sheet.append(headers or HEADERS)
    for row in rows:
        sheet.append(row)
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def test_extracts_only_real_date_rows_and_uses_account_code_next_to_account_name() -> None:
    source = workbook_bytes(
        [
            ["매입", "2026-04-01", "상사", "재료", 1000, 100, 1100, "51.과세", "146", "상품", "D01", "부서", "C01", "카드", "1234"],
            ["매입", "월       계", "", "1건", 1000, 100, 1100, "", "", "", "", "", "", "", ""],
            ["매출", "2026-04-02", "고객", "매출", -500, -50, -550, "17.카과", "401", "상품매출", "", "", "", "국민", "9999"],
        ]
    )
    result = parse_workbook(source)
    assert len(result) == 2
    assert result["account_code"].tolist() == ["146", "401"]
    assert result["original_type"].tolist() == ["과세", "카과"]
    assert result.iloc[1]["supply_amount"] == -500


def test_reports_missing_columns_with_sheet_and_recognized_columns() -> None:
    with pytest.raises(InputWorkbookError) as error:
        parse_workbook(workbook_bytes([], headers=["구분", "전표일자", "거래처"]))
    message = str(error.value)
    assert "찾지 못한 열" in message
    assert "인식한 열" in message
    assert "전체매입매출" in message


def test_rejects_malformed_amount_without_guessing() -> None:
    source = workbook_bytes(
        [["매입", "2026-04-01", "상사", "재료", "천원", 100, 1100, "51.과세", "146", "상품", "", "", "", "", ""]]
    )
    with pytest.raises(InputWorkbookError, match="금액 형식 오류"):
        parse_workbook(source)

