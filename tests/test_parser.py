# 입력 엑셀 파서의 상세행 판정과 오류 처리를 검증합니다.
from io import BytesIO
import re
import zipfile

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


def with_sheet_dimension(source: BytesIO, dimension: str | None) -> BytesIO:
    original = zipfile.ZipFile(source)
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as rewritten:
        for item in original.infolist():
            data = original.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                replacement = "" if dimension is None else f'<dimension ref="{dimension}"/>'
                data = re.sub(r"<dimension [^>]*/>", replacement, data.decode("utf-8")).encode("utf-8")
            rewritten.writestr(item, data)
    output.seek(0)
    return output


@pytest.mark.parametrize("dimension", [None, "A1"])
def test_reads_rows_when_sheet_dimension_is_missing_or_wrong(dimension: str | None) -> None:
    source = workbook_bytes(
        [
            ["매입", "2026-04-01", "상사", "재료", 1000, 100, 1100, "51.과세", "146", "상품"],
            [],
            ["매출", "2026-04-02", "고객", "매출", 500, 50, 550, "17.카과", "401", "상품매출", "", "", "", "국민", "9999"],
        ]
    )
    result = parse_workbook(with_sheet_dimension(source, dimension))
    assert result["source_row"].tolist() == [2, 4]
    assert result["card_number"].tolist() == ["", "9999"]
    assert result.iloc[0]["supply_amount"] == 1000


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


def test_rejects_detail_row_whose_date_cannot_be_read() -> None:
    source = workbook_bytes(
        [
            ["매입", "2026-04-01", "상사", "재료", 1000, 100, 1100, "51.과세", "146", "상품", "", "", "", "", ""],
            ["매입", "2026년 4월 2일", "상사", "재료", 9999, 999, 10998, "51.과세", "146", "상품", "", "", "", "", ""],
        ]
    )
    with pytest.raises(InputWorkbookError, match="전표일자 형식 오류") as error:
        parse_workbook(source)
    assert "행=3" in str(error.value)


@pytest.mark.parametrize(
    ("detail", "message"),
    [
        (["매입", None, "상사", "재료", 9999, 999, 10998, "51.과세", "146", "상품"], "전표일자 누락"),
        (["", "2026-13-40", "상사", "재료", 9999, 999, 10998, "51.과세", "146", "상품"], "전표일자 형식 오류"),
        (["", None, "상사", "재료", 9999, 999, 10998], "전표일자 누락"),
    ],
)
def test_rejects_detail_row_without_readable_date(detail: list[object], message: str) -> None:
    source = workbook_bytes(
        [
            ["매입", "2026-04-01", "상사", "재료", 1000, 100, 1100, "51.과세", "146", "상품"],
            detail,
        ]
    )
    with pytest.raises(InputWorkbookError, match=message) as error:
        parse_workbook(source)
    assert "행=3" in str(error.value)


def test_keeps_skipping_aggregate_and_title_rows() -> None:
    source = workbook_bytes(
        [
            ["전체 매입매출장", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["매입", "2026-04-01", "상사", "재료", 1000, 100, 1100, "51.과세", "146", "상품", "", "", "", "", ""],
            ["매입", "월       계", "", "1건", 1000, 100, 1100, "", "", "", "", "", "", "", ""],
            ["매입", "누   계", "", "1건", 1000, 100, 1100, "", "", "", "", "", "", "", ""],
            ["매입", "분기 누계", "", "1건", 1000, 100, 1100, "", "", "", "", "", "", "", ""],
            ["매입", "합       계", "", "1건", 1000, 100, 1100, "", "", "", "", "", "", "", ""],
            ["매입", None, "", "1건", 1000, 100, 1100],
            ["", None, "[ 4월 합계 ]", "", 1000, 100, 1100],
            ["", None, "", "", 1000, 100, 1100],
            HEADERS,
            ["출력일 2026-07-01", None],
            ["", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ]
    )
    result = parse_workbook(source)
    assert len(result) == 1
    assert result.iloc[0]["supply_amount"] == 1000


def test_stops_when_more_than_one_sheet_has_a_ledger_header() -> None:
    workbook = Workbook()
    first = workbook.active
    first.title = "상반기"
    second = workbook.create_sheet("하반기")
    for sheet in (first, second):
        sheet.append(HEADERS)
        sheet.append(["매입", "2026-04-01", "상사", "재료", 1000, 100, 1100, "51.과세", "146", "상품"])
    workbook.create_sheet("메모").append(["참고"])
    source = BytesIO()
    workbook.save(source)
    source.seek(0)

    with pytest.raises(InputWorkbookError, match="원장 시트가 여러 개") as error:
        parse_workbook(source)
    assert "상반기" in str(error.value)
    assert "하반기" in str(error.value)


def test_rejects_malformed_amount_without_guessing() -> None:
    source = workbook_bytes(
        [["매입", "2026-04-01", "상사", "재료", "천원", 100, 1100, "51.과세", "146", "상품", "", "", "", "", ""]]
    )
    with pytest.raises(InputWorkbookError, match="금액 형식 오류"):
        parse_workbook(source)

