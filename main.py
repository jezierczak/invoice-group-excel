from typing import Any
from openpyxl.styles import Font, PatternFill, Alignment

from invoice_group_excel.config import CFG
from invoice_group_excel.invoice_aggregator import InvoiceAggregator
from invoice_group_excel.kmexcel.excel.types.formula import FormulaTemplate
from invoice_group_excel.kmexcel.excel.manager import ExcelManager
from invoice_group_excel.kmexcel.pivot_tables.manager import PivotTableManager


def main() -> None:
    # pivot_manager = PivotTableManager(CFG.input_data_file)
    #
    # pivot_manager.create_pivot(
    #     name='AgregacjaKosztowRodzajowych',
    #     source_sheet_name='Faktury',
    #     index=['Invoice'],
    #     # columns=['Nazwazadania'],
    #     values=['Amount','Gross','Net','NotEvid'],
    #     agr_func='sum'
    #
    # )
    # pivot_agregacja = pivot_manager.get_pivot('AgregacjaKosztowRodzajowych')
    # print(pivot_agregacja)

    # excel_manager = ExcelManager(CFG.input_data_file)
    # input_data: dict[str,list[dict[str, Any]]] = excel_manager.read_sheet('Tabela1')

    agregator = InvoiceAggregator(CFG.input_data_file,['Tabela1'])
    # print(agregator.get_sheet_input_data_with_conditions(
    #     'Tabela1',
    #     lambda x: x['Task']=='wyburzenia' and x['TaskV2']=='smieci'))

    sheets_grouped = agregator.group_sheets_by_keys('Tabela1',['Task','TaskV2'])


    for key_sheet, value_sheet in sheets_grouped.items():

        formula_templates: list[FormulaTemplate] = [
            {"column": "D", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2, "template": "=SUM(D2:D{last})"},
            {"column": "E", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
             "template": "=SUM(E2:E{last})"},
            {"column": "F", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
             "template": "=SUM(F2:F{last})"},
            {"column": "G", "start_row": len(value_sheet) + 2, "end_row": len(value_sheet) + 2,
             "template": "=SUM(G2:G{last})"},
        ]


        agregator.add_sheet(key_sheet, value_sheet,formula_templates=formula_templates)

        agregator.style_all(
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
        agregator.style_row(
            sheet_name=key_sheet,
            row_index=1,
            start_col_letter="A",
            style={
                "font": Font(name="Calibri", size=12, bold=True, color="FFFFFF"),
                "fill": PatternFill(fill_type="solid", fgColor="4F81BD"),
                "alignment": Alignment(horizontal="center", vertical="center"),
            },
        )

        agregator.set_column_format(key_sheet, "B", "MM-DD-YYYY")
        agregator.autofit_column_widths(key_sheet, offset_dim=2)

    agregator.save(CFG.data_dir + "/output1.xlsx")

if __name__ == '__main__':
    main()