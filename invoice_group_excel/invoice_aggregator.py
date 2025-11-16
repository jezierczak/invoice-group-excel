from typing import Any, Callable

from openpyxl.styles import Font, PatternFill, Alignment

from invoice_group_excel.config import CFG
from invoice_group_excel.kmexcel.excel.manager import ExcelManager
from invoice_group_excel.kmexcel.excel.types.formula import FormulaTemplate


class InvoiceAggregator(ExcelManager):
    def __init__(self, input_data_file: str, input_data_sheets: list[str]) -> None:
        super().__init__(input_data_file)
        self.input_data_sheets = input_data_sheets
        self._input_data: dict[str, list[dict[str, Any]]] = {}
        for data_sheet in input_data_sheets:
            self._input_data[data_sheet] = self.read_sheet(data_sheet,{'Description':['SUMMARY']})
            self.create_source_sheet(data_sheet)

    def update_input_data(self) -> None:
        for sheet in self.get_sheet_names():
            self._input_data[sheet] = self.read_sheet(sheet,{'Description': ['SUMMARY']})

    def create_source_sheet(self,source_sheet_name: str) -> None:

        value_sheet = self._input_data[source_sheet_name]

        self.add_sheet("SOURCE",value_sheet,formula_templates=self.default_formula_templates(value_sheet))

    def get_input_data(self) -> dict[str,list[dict[str, Any]]]:
        return self._input_data

    def get_sheet_input_data(self,sheet_name: str) -> list[dict[str, Any]]:
        return self._input_data.get(sheet_name)

    def get_sheet_input_data_with_conditions(self,sheet_name: str, condition_fn: Callable[[dict[str,Any]],bool]) -> list[dict[str, Any]]:
        return [output for output in self.get_sheet_input_data(sheet_name) if condition_fn(output) ]
            #for key,value in output.items() if condition_fn(key,value)]

    # def group_sheets_by_key(self,sheet_name: str, key_to_group: str) -> dict[str,list[dict[str, Any]]]:
    #     output = {}
    #
    #     set_of_values: set[str] = set()
    #     input_data = self.get_sheet_input_data(sheet_name)
    #
    #     for row in input_data:
    #         set_of_values.add(row[key_to_group])
    #
    #     for value in set_of_values:
    #         output[value] = self.get_sheet_input_data_with_conditions(sheet_name, lambda v: v[key_to_group] == value)
    #
    #     return output

    def group_sheets_by_keys(self, source_sheet_name: str, keys_to_group: list[str]) -> dict[str, list[dict[str, Any]]]:

        output = {}

        set_of_values: set[str] = set()
        input_data = self.get_sheet_input_data(source_sheet_name)

        for row in input_data:
            set_of_values.add('.'.join([row[key] for key in keys_to_group]))#[row['Task'],row['TaskV2']]) )

        for value in set_of_values:
            output[value] = self.get_sheet_input_data_with_conditions(source_sheet_name,
                lambda x: all([x[key_to_group] == v for key_to_group,v in zip(keys_to_group,value.split('.'))]))

        return output


    def add_sheets_grouped_by_key(self, source_sheet_name: str, key_to_group: list[str], sort_by: str | None = None) -> None:
        sheets_grouped = self.group_sheets_by_keys(source_sheet_name, key_to_group)


        for key_sheet, value_sheet in sheets_grouped.items():
            if sort_by:
                value_sheet.sort(key=lambda x: x[sort_by])
            self.add_sheet(key_sheet, value_sheet, formula_templates=self.default_formula_templates(value_sheet))

    def add_sheet_with_contitions(self,
                                  source_sheet_name: str,
                                  output_sheet_name: str ,
                                  condition_fn: Callable[[dict[str,Any]],bool],
                                  sort_by: str | None = None
                                  ) -> None:
        data_to_add = self.get_sheet_input_data_with_conditions(source_sheet_name, condition_fn)
        if sort_by:
            data_to_add.sort(key=lambda x: x[sort_by])
        self.add_sheet(output_sheet_name,data_to_add,formula_templates=self.default_formula_templates(data_to_add))


    def save_and_style(self,filepath: str | None = None) -> None:

        for key_sheet in self.get_sheet_names():
            self.style_all(
                key_sheet,
                style={
                    "font": Font(name="Calibri", size=11),
                    "alignment": Alignment(horizontal="center"),
                    "border_sides": {
                        "top": {"style": "thin", "color": "999999"},
                        "bottom": {"style": "thin", "color": "999999"},
                        "left": {"style": "thin", "color": "999999"},
                        "right": {"style": "thin", "color": "999999"},
                    },
                },
            )
            self.style_row(
                sheet_name=key_sheet,
                row_index=1,
                start_col_letter="A",
                style={
                    "font": Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
                    "fill": PatternFill(fill_type="solid", fgColor="4F81BD"),
                    "alignment": Alignment(horizontal="center", vertical="center"),
                },
            )
            for date_column in CFG.date_columns:
                self.set_column_format(key_sheet, date_column, CFG.date_format)
            self.autofit_column_widths(key_sheet, offset_dim=2)
        self.save(filepath=filepath)

    @staticmethod
    def default_formula_templates(value_sheet:  list[dict[str, Any]]) -> list[FormulaTemplate]:
        formula_templates: list[FormulaTemplate] = []

        for key, value in CFG.summary_description.items():
            formula_templates.append({"column": key, "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
             "template": value})

        #
        #     {"column": "D", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
        #      "template": "=SUM(D2:D{last})"},
        #     {"column": "E", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
        #      "template": "=SUM(E2:E{last})"},
        #     {"column": "F", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
        #      "template": "=SUM(F2:F{last})"},
        #     {"column": "G", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
        #      "template": "=SUM(G2:G{last})"},
        # ]
        return formula_templates



