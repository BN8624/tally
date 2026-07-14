# 업체 설정부터 불공 검토와 결과 저장까지 제공하는 로컬 데스크톱 앱입니다.
from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from tally import (
    CompanySettings,
    InputWorkbookError,
    SettingsStore,
    export_workbook,
    parse_workbook,
    process_transactions,
)


class CompanyDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.title("새 업체 설정")
        self.resizable(False, False)
        self.result: CompanySettings | None = None
        self.transient(parent)
        self.grab_set()

        self.name_var = tk.StringVar()
        self.label_var = tk.StringVar(value="상품")
        self.fixed_var = tk.StringVar()
        self.overrides_var = tk.StringVar()
        self.vehicle_var = tk.StringVar(value="차량유지비, 주유, 수리, 타이어, 세차, 자동차")
        self.vendor_var = tk.StringVar()
        self.personal_var = tk.StringVar(value="개인사용, 가사용, 사적, 가사")

        fields = [
            ("업체명", self.name_var),
            ("고정자산 계정코드 · 쉼표 구분", self.fixed_var),
            ("계정 예외 · 예: 899=제조경비", self.overrides_var),
            ("차량 후보 키워드 · 쉼표 구분", self.vehicle_var),
            ("거래처 후보 키워드 · 쉼표 구분", self.vendor_var),
            ("사업 무관 후보 키워드 · 쉼표 구분", self.personal_var),
        ]
        frame = ttk.Frame(self, padding=16)
        frame.grid(sticky="nsew")
        row = 0
        ttk.Label(frame, text="146번 표시 명칭").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Combobox(
            frame,
            textvariable=self.label_var,
            values=("상품", "음식재료"),
            state="readonly",
            width=35,
        ).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1
        for label, variable in fields:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4, padx=(0, 12))
            ttk.Entry(frame, textvariable=variable, width=48).grid(row=row, column=1, sticky="ew", pady=4)
            row += 1

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(button_frame, text="취소", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(button_frame, text="저장", command=self._save).pack(side="right")
        self.bind("<Return>", lambda _: self._save())
        self.wait_visibility()
        self.focus_set()

    @staticmethod
    def _split(value: str) -> list[str]:
        return [item.strip() for item in value.split(",") if item.strip()]

    def _save(self) -> None:
        try:
            overrides: dict[str, str] = {}
            for item in self._split(self.overrides_var.get()):
                if "=" not in item:
                    raise ValueError(f"계정 예외는 코드=분류 형식이어야 합니다. 입력={item}")
                code, category = item.split("=", 1)
                overrides[code.strip()] = category.strip()
            self.result = CompanySettings(
                name=self.name_var.get(),
                account_146_label=self.label_var.get(),
                fixed_asset_codes=set(self._split(self.fixed_var.get())),
                account_overrides=overrides,
                vehicle_keywords=self._split(self.vehicle_var.get()),
                vendor_keywords=self._split(self.vendor_var.get()),
                personal_keywords=self._split(self.personal_var.get()),
            )
        except ValueError as exc:
            messagebox.showerror("업체 설정 오류", str(exc), parent=self)
            return
        self.destroy()


class TallyApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Tally · 부가세 월별 집계")
        self.geometry("1480x900")
        self.minsize(1100, 700)

        self.store = SettingsStore()
        self.source_data: pd.DataFrame | None = None
        self.result = None
        self.decisions: dict[str, dict[str, str]] = {}
        self.file_var = tk.StringVar()
        self.company_var = tk.StringVar()
        self.filter_var = tk.StringVar(value="전체")
        self.decision_var = tk.StringVar(value="과세 유지")
        self.reason_var = tk.StringVar()
        self.memo_var = tk.StringVar()
        self.status_var = tk.StringVar(value="업체와 엑셀 파일을 선택하세요.")

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("맑은 고딕", 16, "bold"))
        style.configure("Status.TLabel", font=("맑은 고딕", 10, "bold"))

        self._build_ui()
        self._refresh_companies()

    def _build_ui(self) -> None:
        header = ttk.Frame(self, padding=(16, 12))
        header.pack(fill="x")
        ttk.Label(header, text="Tally 부가세 월별 집계", style="Title.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.status_var, style="Status.TLabel").pack(side="right")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.setup_tab = ttk.Frame(self.notebook, padding=20)
        self.review_tab = ttk.Frame(self.notebook, padding=12)
        self.result_tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.setup_tab, text="1. 파일 선택")
        self.notebook.add(self.review_tab, text="2. 불공 검토")
        self.notebook.add(self.result_tab, text="3. 집계 결과")
        self._build_setup_tab()
        self._build_review_tab()
        self._build_result_tab()

    def _build_setup_tab(self) -> None:
        form = ttk.LabelFrame(self.setup_tab, text="입력", padding=18)
        form.pack(fill="x", anchor="n")
        ttk.Label(form, text="업체").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=8)
        self.company_combo = ttk.Combobox(form, textvariable=self.company_var, state="readonly", width=45)
        self.company_combo.grid(row=0, column=1, sticky="ew", pady=8)
        ttk.Button(form, text="새 업체 설정 추가", command=self._add_company).grid(row=0, column=2, padx=(10, 0))

        ttk.Label(form, text="전체 매입매출장").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=8)
        ttk.Entry(form, textvariable=self.file_var, state="readonly").grid(row=1, column=1, sticky="ew", pady=8)
        ttk.Button(form, text="엑셀 선택", command=self._select_file).grid(row=1, column=2, padx=(10, 0))
        form.columnconfigure(1, weight=1)
        ttk.Button(form, text="처리 시작", command=self._process_file).grid(
            row=2, column=0, columnspan=3, sticky="ew", pady=(18, 0)
        )

        guide = (
            "원본 파일은 이 PC에서만 처리되며 서버나 외부 API로 전송되지 않습니다.\n"
            "전표일자가 실제 날짜인 상세 거래만 집계하고 월계·누계·합계 행은 자동 제외합니다.\n"
            "필수 열이나 금액 형식이 잘못되면 추정하지 않고 오류 위치를 표시합니다."
        )
        ttk.Label(self.setup_tab, text=guide, justify="left", padding=(4, 20)).pack(anchor="w")

    def _build_review_tab(self) -> None:
        controls = ttk.Frame(self.review_tab)
        controls.pack(fill="x", pady=(0, 8))
        ttk.Label(controls, text="필터").pack(side="left")
        filter_combo = ttk.Combobox(
            controls,
            textvariable=self.filter_var,
            values=("전체", "판단 보류", "원본 불공", "신규 불공 후보"),
            state="readonly",
            width=18,
        )
        filter_combo.pack(side="left", padx=(6, 16))
        filter_combo.bind("<<ComboboxSelected>>", lambda _: self._refresh_review_tree())
        ttk.Button(controls, text="집계 결과 보기", command=lambda: self.notebook.select(self.result_tab)).pack(side="right")

        columns = (
            "date",
            "vendor",
            "item",
            "account_code",
            "account_name",
            "supply",
            "tax",
            "original",
            "candidate",
            "decision",
            "memo",
        )
        labels = (
            "날짜",
            "거래처",
            "품명",
            "계정코드",
            "계정과목",
            "공급가액",
            "세액",
            "원본 유형",
            "후보 사유",
            "최종 판정",
            "메모",
        )
        tree_frame = ttk.Frame(self.review_tab)
        tree_frame.pack(fill="both", expand=True)
        self.review_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        for column, label in zip(columns, labels):
            self.review_tree.heading(column, text=label)
            width = 105
            if column in {"vendor", "item"}:
                width = 150
            if column == "candidate":
                width = 260
            self.review_tree.column(column, width=width, anchor="e" if column in {"supply", "tax"} else "w")
        vertical = ttk.Scrollbar(tree_frame, orient="vertical", command=self.review_tree.yview)
        horizontal = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.review_tree.xview)
        self.review_tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.review_tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.review_tree.bind("<<TreeviewSelect>>", self._load_selected_review)

        editor = ttk.LabelFrame(self.review_tab, text="선택 거래 판정", padding=10)
        editor.pack(fill="x", pady=(10, 0))
        ttk.Label(editor, text="판정").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            editor,
            textvariable=self.decision_var,
            values=("불공 유지", "과세로 변경", "과세 유지", "불공으로 변경", "판단 보류"),
            state="readonly",
            width=18,
        ).grid(row=0, column=1, sticky="w", padx=(6, 16))
        ttk.Label(editor, text="불공 사유").grid(row=0, column=2, sticky="w")
        ttk.Entry(editor, textvariable=self.reason_var, width=34).grid(row=0, column=3, sticky="ew", padx=(6, 16))
        ttk.Label(editor, text="메모").grid(row=0, column=4, sticky="w")
        ttk.Entry(editor, textvariable=self.memo_var, width=30).grid(row=0, column=5, sticky="ew", padx=(6, 16))
        ttk.Button(editor, text="선택 거래 적용", command=self._apply_selected).grid(row=0, column=6, padx=(4, 0))
        ttk.Button(editor, text="같은 조건에 일괄 적용", command=self._apply_same_condition).grid(row=0, column=7, padx=(8, 0))
        editor.columnconfigure(3, weight=1)
        editor.columnconfigure(5, weight=1)

    def _build_result_tab(self) -> None:
        controls = ttk.Frame(self.result_tab)
        controls.pack(fill="x", pady=(0, 8))
        self.result_banner = ttk.Label(controls, text="아직 처리한 결과가 없습니다.", style="Status.TLabel")
        self.result_banner.pack(side="left")
        ttk.Button(controls, text="엑셀 저장", command=self._save_output).pack(side="right")
        ttk.Button(controls, text="불공 검토로 이동", command=lambda: self.notebook.select(self.review_tab)).pack(
            side="right", padx=(0, 8)
        )
        text_frame = ttk.Frame(self.result_tab)
        text_frame.pack(fill="both", expand=True)
        self.result_text = tk.Text(text_frame, wrap="none", font=("Consolas", 10), state="disabled")
        y_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.result_text.yview)
        x_scroll = ttk.Scrollbar(text_frame, orient="horizontal", command=self.result_text.xview)
        self.result_text.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.result_text.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

    def _refresh_companies(self) -> None:
        companies = list(self.store.load_all())
        self.company_combo["values"] = companies
        if companies and self.company_var.get() not in companies:
            self.company_var.set(companies[0])

    def _add_company(self) -> None:
        dialog = CompanyDialog(self)
        self.wait_window(dialog)
        if dialog.result is None:
            return
        self.store.save(dialog.result)
        self._refresh_companies()
        self.company_var.set(dialog.result.name)

    def _select_file(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="전체 매입매출장 선택",
            filetypes=(("Excel 통합 문서", "*.xlsx *.xlsm"), ("모든 파일", "*.*")),
        )
        if path:
            self.file_var.set(path)

    def _process_file(self) -> None:
        if not self.company_var.get():
            messagebox.showwarning("업체 선택", "업체 설정을 먼저 추가하고 선택하세요.", parent=self)
            return
        if not self.file_var.get():
            messagebox.showwarning("파일 선택", "전체 매입매출장 엑셀을 선택하세요.", parent=self)
            return
        try:
            self.source_data = parse_workbook(self.file_var.get())
            self.decisions = {}
            self._recalculate()
        except (InputWorkbookError, OSError, ValueError) as exc:
            messagebox.showerror("처리 실패", str(exc), parent=self)
            return
        self.notebook.select(self.review_tab if not self.result.review.empty else self.result_tab)

    def _recalculate(self) -> None:
        if self.source_data is None:
            return
        settings = self.store.get(self.company_var.get())
        self.result = process_transactions(self.source_data, settings, self.decisions)
        self._refresh_review_tree()
        self._refresh_result_text()
        pending = int(self.result.review["review_status"].eq("판단 보류").sum())
        unclassified = int(self.result.transactions["account_category"].eq("미분류").sum())
        self.status_var.set(
            f"상세 {len(self.result.transactions):,}건 · 불공 검토 {len(self.result.review):,}건 · "
            f"보류 {pending:,}건 · 미분류 {unclassified:,}건"
        )

    def _filtered_review(self) -> pd.DataFrame:
        if self.result is None:
            return pd.DataFrame()
        frame = self.result.review
        selected = self.filter_var.get()
        if selected == "판단 보류":
            return frame[frame["review_status"].eq("판단 보류")]
        if selected == "원본 불공":
            return frame[frame["original_type"].eq("불공")]
        if selected == "신규 불공 후보":
            return frame[frame["original_type"].eq("과세") & frame["candidate_reason"].ne("")]
        return frame

    def _refresh_review_tree(self) -> None:
        for item in self.review_tree.get_children():
            self.review_tree.delete(item)
        frame = self._filtered_review()
        for _, row in frame.iterrows():
            self.review_tree.insert(
                "",
                "end",
                iid=row["row_id"],
                values=(
                    row["date"].isoformat(),
                    row["vendor"],
                    row["item"],
                    row["account_code"],
                    row["account_name"],
                    f"{row['supply_amount']:,}",
                    f"{row['tax_amount']:,}",
                    row["original_type"],
                    row["candidate_reason"],
                    row["review_status"],
                    row["review_memo"],
                ),
            )

    def _selected_row_id(self) -> str | None:
        selection = self.review_tree.selection()
        if not selection:
            messagebox.showwarning("거래 선택", "판정할 거래를 선택하세요.", parent=self)
            return None
        return selection[0]

    def _load_selected_review(self, _event=None) -> None:
        selection = self.review_tree.selection()
        if not selection or self.result is None:
            return
        row = self.result.review[self.result.review["row_id"].eq(selection[0])].iloc[0]
        self.decision_var.set(row["review_status"])
        self.reason_var.set(row["nondeductible_reason"])
        self.memo_var.set(row["review_memo"])

    def _decision_payload(self) -> dict[str, str] | None:
        decision = self.decision_var.get()
        reason = self.reason_var.get().strip()
        if decision in {"불공 유지", "불공으로 변경"} and not reason:
            messagebox.showwarning("불공 사유", "불공으로 확정할 때는 불공 사유를 입력하세요.", parent=self)
            return None
        return {"decision": decision, "reason": reason, "memo": self.memo_var.get().strip()}

    def _apply_selected(self) -> None:
        row_id = self._selected_row_id()
        payload = self._decision_payload()
        if row_id is None or payload is None:
            return
        self.decisions[row_id] = payload
        self._recalculate()

    def _apply_same_condition(self) -> None:
        row_id = self._selected_row_id()
        payload = self._decision_payload()
        if row_id is None or payload is None or self.result is None:
            return
        selected = self.result.review[self.result.review["row_id"].eq(row_id)].iloc[0]
        matching = self.result.review[
            self.result.review["account_code"].eq(selected["account_code"])
            & self.result.review["candidate_reason"].eq(selected["candidate_reason"])
        ]
        if not messagebox.askyesno(
            "같은 조건 일괄 적용",
            f"계정코드와 후보 근거가 같은 {len(matching):,}건에 판정을 적용할까요?",
            parent=self,
        ):
            return
        for matching_id in matching["row_id"]:
            self.decisions[matching_id] = dict(payload)
        self._recalculate()

    @staticmethod
    def _format_frame(frame: pd.DataFrame) -> str:
        if frame.empty:
            return "해당 거래 없음"
        display = frame.copy()
        for column in ("supply_amount", "tax_amount", "total_amount", "expected", "actual", "difference"):
            if column in display:
                display[column] = display[column].map(
                    lambda value: f"{value:,}" if value != "" and pd.notna(value) else ""
                )
        return display.to_string(index=False)

    def _refresh_result_text(self) -> None:
        if self.result is None:
            return
        state = "검산 완료" if self.result.validation_passed else "검산 실패"
        self.result_banner.configure(text=state, foreground="#38761D" if self.result.validation_passed else "#9C0006")
        sections = [
            ("1. 매입 계정 분류별 월별 집계", self.result.purchase_by_category),
            ("2~7. 일반·고정·세금계산서·카과·현과·카드·과세총계·불공·과매계", self.result.purchase_summary),
            ("8. 과세 매출 계정별 월별 집계", self.result.sales_by_account),
            ("9~11. 면세 매출·카드매출·현영매출", self.result.sales_summary),
            ("12. 검산 결과", self.result.validation),
        ]
        text = [f"[{state}]", ""]
        for title, frame in sections:
            text.extend((title, self._format_frame(frame), ""))
        pending = self.result.review[self.result.review["review_status"].eq("판단 보류")]
        unclassified = self.result.transactions[self.result.transactions["account_category"].eq("미분류")]
        text.extend(
            (
                "13. 엑셀 저장 전 확인",
                f"불공 판단 보류 {len(pending):,}건 · 계정 미분류 {len(unclassified):,}건",
            )
        )
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", "\n".join(text))
        self.result_text.configure(state="disabled")

    def _save_output(self) -> None:
        if self.result is None:
            messagebox.showwarning("처리 결과", "먼저 엑셀 파일을 처리하세요.", parent=self)
            return
        if not self.result.validation_passed:
            if not messagebox.askyesno(
                "검산 실패",
                "미분류·판단 보류 또는 금액 차이가 있습니다. 실패 상태를 포함한 검토용 엑셀을 저장할까요?",
                parent=self,
            ):
                return
        suggested = f"{self.company_var.get()}_부가세_월별_집계.xlsx"
        path = filedialog.asksaveasfilename(
            parent=self,
            title="결과 엑셀 저장",
            defaultextension=".xlsx",
            initialfile=suggested,
            filetypes=(("Excel 통합 문서", "*.xlsx"),),
        )
        if not path:
            return
        export_workbook(self.result, self.store.get(self.company_var.get()), path)
        messagebox.showinfo("저장 완료", f"결과 엑셀을 저장했습니다.\n{Path(path)}", parent=self)


def main() -> None:
    TallyApp().mainloop()


if __name__ == "__main__":
    main()
