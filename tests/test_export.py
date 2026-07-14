# 결과 엑셀의 필수 시트와 핵심 집계 셀 및 수식을 검증합니다.
from datetime import date
from decimal import Decimal

from openpyxl import load_workbook
import pandas as pd

from tally import CompanySettings, export_workbook, process_transactions


def test_export_creates_required_sheets_and_summary_formulas(tmp_path) -> None:
    rows = [
        {
            "row_id": "Sheet1:2",
            "sheet": "Sheet1",
            "source_row": 2,
            "division": "매입",
            "date": date(2026, 4, 1),
            "month": "2026-04",
            "vendor": "상사",
            "item": "재료",
            "supply_amount": Decimal(1000),
            "tax_amount": Decimal(100),
            "total_amount": Decimal(1100),
            "original_type": "과세",
            "account_code": "146",
            "account_name": "상품",
            "card_company": "",
            "card_number": "",
        },
        {
            "row_id": "Sheet1:3",
            "sheet": "Sheet1",
            "source_row": 3,
            "division": "매출",
            "date": date(2026, 4, 2),
            "month": "2026-04",
            "vendor": "고객",
            "item": "매출",
            "supply_amount": Decimal(2000),
            "tax_amount": Decimal(200),
            "total_amount": Decimal(2200),
            "original_type": "카과",
            "account_code": "401",
            "account_name": "상품매출",
            "card_company": "국민",
            "card_number": "1234",
        },
    ]
    settings = CompanySettings(name="테스트상사")
    result = process_transactions(pd.DataFrame(rows), settings)
    output = export_workbook(result, settings, tmp_path / "result.xlsx")

    workbook = load_workbook(output, data_only=False)
    assert workbook.sheetnames == ["집계표", "불공검토", "거래상세", "미분류", "검산"]
    summary = workbook["집계표"]
    assert summary["A1"].value == "테스트상사 부가세 월별 집계표"
    assert summary["A5"].value == "상품"
    assert [summary.cell(6, column).value for column in range(1, 6)] == [
        "월",
        "매수",
        "공급가액",
        "세액",
        "합계금액",
    ]
    assert summary["A7"].value == "4월"
    assert summary["B7"].value == 1
    assert summary["B7"].number_format == '"("0")"'
    assert summary["A8"].value == "계"
    assert summary["B8"].value == "=SUM(B7:B7)"
    assert summary.page_setup.orientation == "portrait"
    assert summary.page_setup.paperWidth == "170mm"
    assert summary.page_setup.paperHeight == "240mm"
    assert summary.page_setup.fitToWidth == 1
    assert summary.page_setup.fitToHeight == 1
    assert any(
        isinstance(cell.value, str) and cell.value.startswith("=")
        for row in summary.iter_rows()
        for cell in row
    )
    assert workbook["거래상세"].max_row == 5
    assert workbook["검산"]["A1"].value.startswith("검산 결과")
