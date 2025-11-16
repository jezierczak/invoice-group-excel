
from invoice_group_excel.config import CFG
from invoice_group_excel.invoice_aggregator import InvoiceAggregator



def main() -> None:

    agregator = InvoiceAggregator(CFG.input_data_file,['Tabela1'])
    # print(agregator.get_sheet_input_data_with_conditions(
    #     'Tabela1',
    #     lambda x: x['Task']=='wyburzenia' and x['TaskV2']=='smieci'))

    agregator.add_sheets_grouped_by_key('Tabela1',['Task'],'Type')
    # print(agregator.get_sheet_names())
    agregator.add_sheet_with_contitions('Tabela1','illia_', lambda row: row["Type"] == 'illia','Task')
    agregator.add_sheet_with_contitions('Tabela1', 'salary_', lambda row: row["Type"] == 'salary')
    agregator.add_sheet_with_contitions('Tabela1', 'fuel_', lambda row: row["Type"] == 'fuel')
    agregator.add_sheet_with_contitions('Tabela1', 'material_consumables', lambda row: row["Type"] == 'material' or row["Type"] == 'consumables',"Task")
    agregator.update_input_data()
    # print(agregator.get_input_data().keys())
    agregator.add_sheet_with_contitions('posadzka', 'posadzka_material_consumables', lambda row: row["Type"] == 'material' or row["Type"] == 'consumables','Type')

    agregator.save_and_style(CFG.output_data_file)

if __name__ == '__main__':
    main()