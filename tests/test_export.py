# 집계표 틀 배치, 요약 수식 계산값, 입력 칸 반영, 전자·종이 시트, 셀 충돌을 검증합니다.
from datetime import date
from decimal import Decimal
import math
import random
import re

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd
import pytest

from tally import CompanySettings, Period, export_workbook, process_transactions
from tally.export import _build_summary_sheet

SUMMARY_KEYS = {
    "일반": "general",
    "고정": "fixed",
    "세금계산서 계": "invoice",
    "카드": "card",
    "불공": "nondeductible",
    "매입계": "purchase",
    "세금계산서": "sales_invoice",
    "영세": "zero_rated",
    "카드외": "card_out",
    "기타(건별)": "other",
    "과세계": "taxable",
    "면세 계산서": "exempt_invoice",
    "면세 결제": "exempt_payment",
    "면세계": "exempt",
    "총매출": "total_sales",
    "납부(환급)세액": "payable",
    "카드공제": "card_deduction",
    "카드공제(비대상)": "card_deduction",
    "택시 경감세액": "taxi_relief",
    "차감 납부할 세액": "final",
    "운수종사자 지급(90%)": "taxi_drivers",
    "감차 보상(5%)": "taxi_reduction",
    "운수종사자 복지(4%)": "taxi_welfare",
}


def _excel_round(value: float) -> float:
    return float(f"{value:.15g}")


def _formula_value(sheet, reference: str) -> float:
    cell = sheet[reference]
    if cell.data_type != "f":
        return cell.value or 0

    def expand_sum(match: re.Match) -> str:
        min_col, min_row, max_col, max_row = range_boundaries(match.group(1))
        return "(" + "+".join(
            f"{get_column_letter(column)}{row}" for row in range(min_row, max_row + 1) for column in range(min_col, max_col + 1)
        ) + ")"

    expression = re.sub(r"SUM\(([A-Z]+\d+:[A-Z]+\d+)\)", expand_sum, cell.value[1:])
    expression = re.sub(r"(?<![A-Za-z\"])([A-Z]{1,3}\d+)(?![\w\"])", r'V("\1")', expression)
    return eval(
        expression,
        {
            "V": lambda ref: _formula_value(sheet, ref),
            "ROUNDDOWN": lambda value, _digits: math.trunc(_excel_round(value)),
            "ROUNDUP": lambda value, _digits: math.copysign(math.ceil(abs(_excel_round(value))), value),
            "MIN": min,
            "MAX": max,
        },
    )


def transaction(number, division, original_type, supply, tax, *, code="", name="", vendor="거래처", month=7, electronic=""):
    return {
        "row_id": f"Sheet1:{number}",
        "sheet": "Sheet1",
        "source_row": number,
        "division": division,
        "date": date(2026, month, 1),
        "month": f"2026-{month:02d}",
        "vendor": vendor,
        "item": "품목",
        "supply_amount": Decimal(supply),
        "tax_amount": Decimal(tax),
        "total_amount": Decimal(supply + tax),
        "original_type": original_type,
        "account_code": code,
        "account_name": name,
        "card_company": "",
        "card_number": "",
        "electronic": electronic,
    }


def canon_example_rows():
    return [
        transaction(1, "매입", "과세", 1005, 100, code="146", name="원재료", electronic="전자"),
        transaction(2, "매입", "불공", 200, 20, code="813", name="접대비", electronic="종이"),
        transaction(3, "매입", "과세", 1000, 99, code="208", name="비품", electronic="전자"),
        transaction(4, "매입", "카과", 300, 30, code="830", name="소모품"),
        transaction(5, "매출", "과세", 3000, 300, code="401", name="상품매출", electronic="전자"),
        transaction(6, "매출", "카과", 2000, 200, code="401", name="상품매출"),
        transaction(7, "매출", "현과", 1000, 100, code="401", name="상품매출"),
        transaction(8, "매출", "건별", 500, 50, code="401", name="상품매출"),
    ]


def exported(tmp_path, rows, industry="기본", period=None, **settings):
    result = process_transactions(pd.DataFrame(rows), CompanySettings(name="검증상사", industry=industry, **settings), period)
    workbook = load_workbook(export_workbook(result, tmp_path / f"{industry}.xlsx"))
    return result, workbook


def summary_cells(sheet) -> dict[str, tuple[int, int]]:
    """요약 블록의 항목 이름별 (행, 공급가액 열)을 찾습니다."""
    anchor = next(cell for row in sheet.iter_rows() for cell in row if cell.value == "요약")
    rows: dict[str, tuple[int, int]] = {}
    for row in range(anchor.row + 1, sheet.max_row + 1):
        label = sheet.cell(row, anchor.column).value
        if isinstance(label, str) and label not in rows:
            rows[label] = (row, anchor.column + 1)
    return rows


def summary_value(sheet, label: str, field_name: str = "tax") -> float:
    row, supply_column = summary_cells(sheet)[label]
    column = supply_column if field_name == "supply" else supply_column + 1
    return _formula_value(sheet, f"{get_column_letter(column)}{row}")


def test_summary_formulas_match_python_summary(tmp_path) -> None:
    result, workbook = exported(tmp_path, canon_example_rows())
    sheet = workbook["집계표"]
    for label, (row, supply_column) in summary_cells(sheet).items():
        key = SUMMARY_KEYS.get(label)
        if key is None:
            continue
        expected = next(item for item in result.summary if item.key == key)
        if expected.supply is not None:
            assert _formula_value(sheet, f"{get_column_letter(supply_column)}{row}") == expected.supply, label
        if expected.tax is not None:
            assert _formula_value(sheet, f"{get_column_letter(supply_column + 1)}{row}") == expected.tax, label
    assert summary_value(sheet, "차감 납부할 세액") == 378


def test_input_cells_flow_into_final_amount(tmp_path) -> None:
    _, workbook = exported(tmp_path, canon_example_rows())
    sheet = workbook["집계표"]
    cells = summary_cells(sheet)
    extra_row, supply_column = cells["추가 불공"]
    penalty_row, _ = cells["가산세"]
    tax_letter = get_column_letter(supply_column + 1)
    sheet[f"{tax_letter}{extra_row}"] = 10
    sheet[f"{tax_letter}{penalty_row}"] = 5
    assert summary_value(sheet, "매입계") == 220
    assert summary_value(sheet, "납부(환급)세액") == 430
    assert summary_value(sheet, "카드공제") == 42
    assert summary_value(sheet, "차감 납부할 세액") == 393


def test_card_exempt_input_moves_mixed_card_sales_to_exempt(tmp_path) -> None:
    rows = [
        transaction(1, "매입", "과세", 1000, 100, code="146"),
        transaction(2, "매출", "카과", 2000, 200),
        transaction(3, "매출", "현과", 1000, 100),
    ]
    _, workbook = exported(tmp_path, rows, industry="도소매")
    sheet = workbook["집계표"]
    row, supply_column = summary_cells(sheet)["카드매출 면세분(합계금액)"]
    sheet[f"{get_column_letter(supply_column)}{row}"] = 1100
    assert summary_value(sheet, "카드외", "supply") == 2000
    assert summary_value(sheet, "카드외") == 200
    assert summary_value(sheet, "면세계", "supply") == 1100
    assert summary_value(sheet, "카드공제") == 28


@pytest.mark.parametrize(
    ("industry", "purchase_headers", "has_payment", "inputs"),
    [
        ("기본", ["원재료", "고정", "기타", "계"], True, {"의제", "공통매입 면세사업분", "카드매출 면세분(합계금액)"}),
        ("건설", ["원재료", "도급", "고정", "기타", "계"], False, set()),
        ("도소매", ["상품", "고정", "기타", "계"], True, {"의제", "공통매입 면세사업분", "카드매출 면세분(합계금액)"}),
        ("음식점", ["음식재료", "고정", "기타", "계"], True, {"의제"}),
        ("택시", ["원재료", "고정", "기타", "계"], True, {"의제", "공통매입 면세사업분", "카드매출 면세분(합계금액)"}),
    ],
)
def test_layout_follows_industry_frame(tmp_path, industry, purchase_headers, has_payment, inputs) -> None:
    rows = [transaction(1, "매입", "과세", 1000, 100, code="146", month=8), transaction(2, "매출", "과세", 2000, 200, code="401", name="상품매출")]
    result, workbook = exported(tmp_path, rows, industry=industry)
    sheet = workbook["집계표"]
    assert sheet["A3"].value == "매입 세금계산서"
    headers = [sheet.cell(3, column).value for column in range(2, 20) if sheet.cell(3, column).value]
    assert headers[: len(purchase_headers)] == purchase_headers
    assert [sheet.cell(row, 1).value for row in range(4, 8)] == ["7월", "8월", "9월", "계"]
    titles = {sheet.cell(row, 1).value for row in range(1, sheet.max_row + 1)}
    assert ("결제 채널 매출" in titles) == has_payment
    labels = set(summary_cells(sheet))
    optional = {"의제", "공통매입 면세사업분", "카드매출 면세분(합계금액)"}
    assert labels & optional == inputs
    assert {"추가 불공", "경감·공제세액", "예정 미환급세액", "예정 고지세액", "매입자 납부특례 기납부세액", "가산세"} <= labels
    assert ("택시 경감세액" in labels) == (industry == "택시")
    assert workbook.sheetnames == ["집계표", "전자·종이", "불공후보", "거래상세", "미분류", "검산"]


def test_input_cells_are_green_and_empty(tmp_path) -> None:
    _, workbook = exported(tmp_path, canon_example_rows())
    sheet = workbook["집계표"]
    row, supply_column = summary_cells(sheet)["예정 미환급세액"]
    tax_cell = sheet.cell(row, supply_column + 1)
    assert tax_cell.value is None
    assert tax_cell.fill.fgColor.rgb.endswith("C6EFCE")


def test_taxi_relief_is_prefilled_and_allocated_below_summary(tmp_path) -> None:
    rows = [transaction(1, "매출", "과세", 12345670, 1234567, code="416", name="운송수입")]
    _, workbook = exported(tmp_path, rows, industry="택시", card_deduction_eligible=False)
    sheet = workbook["집계표"]
    cells = summary_cells(sheet)
    relief_row, supply_column = cells["택시 경감세액"]
    relief_cell = sheet.cell(relief_row, supply_column + 1)
    assert relief_cell.value.startswith("=MAX(0,ROUNDDOWN(")
    assert relief_cell.fill.fgColor.rgb.endswith("C6EFCE")
    assert summary_value(sheet, "택시 경감세액") == 1222221
    assert summary_value(sheet, "차감 납부할 세액") == 12346
    shares = ["운수종사자 지급(90%)", "감차 보상(5%)", "운수종사자 복지(4%)"]
    assert [summary_value(sheet, label) for label in shares] == [1111111, 61728, 49382]
    assert min(cells[label][0] for label in shares) > cells["차감 납부할 세액"][0]
    relief_cell.value = 1000000
    assert summary_value(sheet, "차감 납부할 세액") == 234567


def test_electronic_sheet_splits_invoices(tmp_path) -> None:
    rows = [
        transaction(1, "매입", "과세", 1000, 100, code="146", electronic="전자"),
        transaction(2, "매입", "과세", 500, 50, code="146", electronic="종이"),
    ]
    _, workbook = exported(tmp_path, rows)
    sheet = workbook["전자·종이"]
    headers = {sheet.cell(3, column).value: column for column in range(2, 30) if sheet.cell(3, column).value}
    electronic, paper = headers["원재료 전자"], headers["원재료 종이"]
    assert _formula_value(sheet, f"{get_column_letter(electronic + 1)}7") == 1000
    assert _formula_value(sheet, f"{get_column_letter(paper + 1)}7") == 500
    assert _formula_value(workbook["집계표"], "L7") == 1500

    _, without = exported(tmp_path, [transaction(1, "매입", "과세", 1000, 100, code="146")])
    assert "전자세금 열이 없어" in without["전자·종이"]["A3"].value


def test_extra_payment_block_feeds_card_out_for_construction(tmp_path) -> None:
    rows = [transaction(1, "매출", "과세", 1000, 100, code="404", name="공사수입금"), transaction(2, "매출", "카과", 2000, 200)]
    _, workbook = exported(tmp_path, rows, industry="건설")
    sheet = workbook["집계표"]
    assert "결제 채널 매출" in {sheet.cell(row, 1).value for row in range(1, sheet.max_row + 1)}
    assert summary_value(sheet, "카드외", "supply") == 2000


def test_text_starting_with_equals_is_not_saved_as_formula(tmp_path) -> None:
    rows = [
        transaction(1, "매입", "과세", 1000, 100, code="146", vendor="=1+2"),
        transaction(2, "매출", "과세", 2000, 200, code="401", name="=매출"),
    ]
    result = process_transactions(pd.DataFrame(rows), CompanySettings(name="=업체"))
    workbook = load_workbook(export_workbook(result, tmp_path / "text.xlsx"))
    found = {
        cell.value: cell.data_type
        for sheet in workbook.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if cell.value in {"=1+2", "=매출", "=업체"}
    }
    assert set(found) == {"=1+2", "=매출", "=업체"}
    assert set(found.values()) == {"s"}


def test_print_area_covers_summary_and_is_landscape(tmp_path) -> None:
    _, workbook = exported(tmp_path, canon_example_rows())
    sheet = workbook["집계표"]
    _, _, last_column, last_row = range_boundaries(sheet.print_area.split("!")[1].replace("$", ""))
    anchor = next(cell for row in sheet.iter_rows() for cell in row if cell.value == "요약")
    assert last_column >= anchor.column + 2
    assert last_row >= max(row for row, _ in summary_cells(sheet).values())
    assert sheet.page_setup.orientation == "landscape"
    assert sheet.page_setup.fitToWidth == 1


def _random_rows(rng: random.Random) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []

    def add(division, original_type, supply, tax, code="", name="", vendor="거래처"):
        month = rng.choice([7, 8, 9])
        rows.append(transaction(len(rows) + 1, division, original_type, supply, tax, code=code, name=name, vendor=vendor, month=month, electronic=rng.choice(["전자", "종이"])))

    for _ in range(rng.randint(0, 8)):
        add("매입", rng.choice(["과세", "과세", "불공", "공통"]), rng.randint(-50, 900) * 10 + rng.randint(0, 9), rng.randint(0, 90), code=rng.choice(["146", "153", "156", "208", "512", "612", "813", "899", "", "4X"]))
    for original_type in rng.sample(["면세", "의제", "영세", "카과", "현과", "카면", "카영", "현면", "현영"], rng.randint(0, 6)):
        add("매입", original_type, rng.randint(1, 900), 0 if original_type not in {"카과", "현과"} else rng.randint(0, 90), code=rng.choice(["146", "830"]))
    for _ in range(rng.randint(0, 4)):
        add("매출", "과세", rng.randint(1, 9000), rng.randint(0, 900), code=rng.choice(["401", "404", "208"]), name=rng.choice(["상품매출", "제품매출", "차량운반구"]))
    for original_type in rng.sample(["영세", "면세", "카과", "카과", "현과", "건별", "카면", "카영", "현면", "현영", "면건"], rng.randint(0, 7)):
        add("매출", original_type, rng.randint(1, 9000), rng.randint(0, 900), code="401", name="상품매출", vendor=rng.choice(["거래처", "제로페이", "우아한형제들", "쿠팡"]))
    return rows


def test_random_ledgers_render_without_overlap_and_formulas_match(monkeypatch) -> None:
    writes: dict[tuple[int, int], list[object]] = {}
    merges: list[tuple[int, int, int, int]] = []
    original_cell = Worksheet.cell
    original_merge = Worksheet.merge_cells

    def recording_cell(self, row, column, value=None):
        if value is not None and self.title == "집계표":
            writes.setdefault((row, column), []).append(value)
        return original_cell(self, row, column, value)

    def recording_merge(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        if self.title == "집계표":
            merges.append((start_row, start_column, end_row, end_column))
        return original_merge(self, range_string, start_row, start_column, end_row, end_column)

    monkeypatch.setattr(Worksheet, "cell", recording_cell)
    monkeypatch.setattr(Worksheet, "merge_cells", recording_merge)
    for seed in range(40):
        rng = random.Random(seed)
        rows = _random_rows(rng)
        if not rows:
            continue
        writes.clear()
        merges.clear()
        settings = CompanySettings(name="배치검사", industry=rng.choice(["기본", "건설", "도소매", "음식점", "택시"]), card_deduction_eligible=rng.random() < 0.5)
        result = process_transactions(pd.DataFrame(rows), settings, Period(2026, 2, "예정"))
        workbook = Workbook()
        _build_summary_sheet(workbook, result)
        sheet = workbook["집계표"]

        overwritten = [key for key, values in writes.items() if len(values) > 1]
        merged_away = [
            key
            for key in writes
            for top, left, bottom, right in merges
            if top <= key[0] <= bottom and left <= key[1] <= right and key != (top, left)
        ]
        _, _, last_column, last_row = range_boundaries(sheet.print_area.split("!")[1].replace("$", ""))
        outside = [key for key in writes if key[0] > last_row or key[1] > last_column]
        assert (seed, overwritten, merged_away, outside) == (seed, [], [], [])
        for label, (row, supply_column) in summary_cells(sheet).items():
            key = SUMMARY_KEYS.get(label)
            if key is None:
                continue
            expected = next(item for item in result.summary if item.key == key)
            for column, value in ((supply_column, expected.supply), (supply_column + 1, expected.tax)):
                if value is not None:
                    assert _formula_value(sheet, f"{get_column_letter(column)}{row}") == value, (seed, label)
