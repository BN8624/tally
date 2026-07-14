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
            "division": "매입",
            "date": date(2026, 4, 2),
            "month": "2026-04",
            "vendor": "상사",
            "item": "접대",
            "supply_amount": Decimal(500),
            "tax_amount": Decimal(50),
            "total_amount": Decimal(550),
            "original_type": "불공",
            "account_code": "813",
            "account_name": "접대비",
            "card_company": "",
            "card_number": "",
        },
        {
            "row_id": "Sheet1:4",
            "sheet": "Sheet1",
            "source_row": 4,
            "division": "매출",
            "date": date(2026, 4, 3),
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
    assert summary["A1"].value == "2026년 1기 확정"
    assert summary["E1"].value is None
    assert summary["I1"].value == "부가세 집계표"
    assert summary["A3"].value == "①"
    assert summary["B3"].value == "상품"
    assert summary["E3"].value == "기타"
    assert summary["H3"].value == "세금계산서 매입계"
    assert summary["A4"].value == "4월"
    assert summary["B4"].value == 1
    assert summary["C4"].value == 1000
    assert summary["D4"].value == 100
    assert summary["E4"].value == 1
    assert summary["F4"].value == 500
    assert summary["G4"].value == 50
    assert summary["H4"].value == 2
    assert summary["I4"].value == 1500
    assert summary["J4"].value == 150
    assert summary["B4"].number_format == '"("0")"'
    assert summary["A5"].value == "계"
    assert summary["B5"].value == "=SUM(B4:B4)"
    assert summary["E6"].value == "계"
    assert summary["F6"].value == "=I5"
    assert summary["G6"].value == "=J5"
    assert summary["E7"].value == "불공"
    assert summary["F7"].value == 500
    assert summary["E8"].value == "차감계"
    assert summary["F8"].value == "=F6-F7"
    assert summary["B12"].value == "불공"
    assert summary["B13"].value == 1
    assert summary["A20"].value == "③  상품매출"
    assert summary["C23"].value == "=C22"
    assert summary["D23"].value == "=ROUNDDOWN(C23*0.1,0)"
    assert summary["B24"].value == "카드외"
    assert summary["C24"].value == "=ROUND((B31)/1.1,0)"
    assert summary["B29"].value == "카드"
    assert summary["B30"].value == 2200
    assert summary["B31"].value == "=SUM(B30:B30)"
    assert not any(
        cell.value
        in {"원재료(도급)", "제조경비", "도급경비", "고정"}
        for row in summary.iter_rows()
        for cell in row
    )
    assert summary["A3"].fill.fill_type == "solid"
    assert summary["A3"].fill.fgColor.rgb.endswith("FFFFFF")
    assert summary.print_area == "'집계표'!$A$1:$J$31"
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
    assert workbook["거래상세"].max_row == 6
    assert workbook["검산"]["A1"].value.startswith("검산 결과")


def test_export_places_other_input_tax_and_deductions_before_sales(tmp_path) -> None:
    def row(
        number: int,
        division: str,
        original_type: str,
        supply: int,
        tax: int,
        code: str = "",
        account_name: str = "",
    ) -> dict[str, object]:
        return {
            "row_id": f"Sheet1:{number}",
            "sheet": "Sheet1",
            "source_row": number,
            "division": division,
            "date": date(2026, 4, number),
            "month": "2026-04",
            "vendor": "거래처",
            "item": "품목",
            "supply_amount": Decimal(supply),
            "tax_amount": Decimal(tax),
            "total_amount": Decimal(supply + tax),
            "original_type": original_type,
            "account_code": code,
            "account_name": account_name,
            "card_company": "",
            "card_number": "",
        }

    rows = [
        row(1, "매입", "과세", 1000, 100, "146", "상품"),
        row(2, "매입", "불공", 200, 20, "813", "접대비"),
        row(3, "매입", "공통", 150, 15, "813", "공통매입"),
        row(4, "매입", "카과", 300, 30),
        row(5, "매입", "현과", 400, 40),
        row(6, "매입", "의제매입세액", 50, 5),
        row(7, "매출", "과세", 1000, 100, "401", "상품매출"),
    ]
    settings = CompanySettings(name="테스트상사")
    result = process_transactions(pd.DataFrame(rows), settings)
    output = export_workbook(result, settings, tmp_path / "vat-layout.xlsx")

    summary = load_workbook(output, data_only=False)["집계표"]
    assert summary["H3"].value == "세금계산서 매입계"
    assert summary["E6"].value == "카드외"
    assert summary["F6"].value == 700
    assert summary["E7"].value == "의제매입세액"
    assert summary["F7"].value == 50
    assert summary["E8"].value == "계"
    assert summary["F8"].value == "=I5+F6+F7"
    assert summary["E9"].value == "불공"
    assert summary["F9"].value == 200
    assert summary["E10"].value == "공통"
    assert summary["F10"].value == 150
    assert summary["E11"].value == "차감계"
    assert summary["F11"].value == "=F8-F9-F10"
    assert summary["A15"].value == "②"
    assert summary["B15"].value == "카과"
    assert summary["E15"].value == "현과"
    assert summary["H15"].value == "의제매입세액"
    assert summary["B19"].value == "불공"
    assert summary["E19"].value == "공통"
    assert summary["A27"].value == "③  상품매출"


def test_export_splits_fixed_assets_after_rounding_adjustment(tmp_path) -> None:
    rows = [
        {
            "row_id": "Sheet1:2",
            "sheet": "Sheet1",
            "source_row": 2,
            "division": "매입",
            "date": date(2026, 4, 1),
            "month": "2026-04",
            "vendor": "일반거래처",
            "item": "일반",
            "supply_amount": Decimal(1005),
            "tax_amount": Decimal(100),
            "total_amount": Decimal(1105),
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
            "division": "매입",
            "date": date(2026, 4, 2),
            "month": "2026-04",
            "vendor": "고정거래처",
            "item": "비품",
            "supply_amount": Decimal(1005),
            "tax_amount": Decimal(100),
            "total_amount": Decimal(1105),
            "original_type": "과세",
            "account_code": "212",
            "account_name": "비품",
            "card_company": "",
            "card_number": "",
        },
    ]
    settings = CompanySettings(name="고정테스트")
    result = process_transactions(pd.DataFrame(rows), settings)
    output = export_workbook(result, settings, tmp_path / "fixed-layout.xlsx")

    summary = load_workbook(output, data_only=False)["집계표"]
    assert summary["E3"].value == "고정"
    assert summary["H3"].value == "세금계산서 매입계"
    assert summary["E6"].value == "단수차이 조정"
    assert summary["F6"].value == "=I5"
    assert summary["G6"].value == "=ROUNDDOWN(F6*0.1,0)"
    assert summary["E7"].value == "일반"
    assert summary["F7"].value == 1005
    assert summary["G7"].value == "=G6-G8"
    assert summary["E8"].value == "고정"
    assert summary["F8"].value == 1005
    assert summary["G8"].value == 100
    assert summary["E9"].value == "계"
    assert summary["F9"].value == "=F7+F8"
    assert summary["G9"].value == "=G7+G8"
    assert summary["E10"].value == "계"
    assert summary["F10"].value == "=F9"
    assert summary["G10"].value == "=G9"
