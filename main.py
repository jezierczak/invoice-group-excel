
from invoice_group_excel.config import CFG
from invoice_group_excel.invoice_aggregator import InvoiceAggregator



def main() -> None:

    agregator = InvoiceAggregator(CFG.input_data_file,['Tabela1'])
    agregator.add_sheets_grouped_by_key(key_to_group = ['Task'],sort_by='Type')

    agregator.add_sheet_with_conditions('Tabela1', output_sheet_name='illia_', condition_fn=lambda row: row["Type"] == 'illia', sort_by='Task')
    agregator.add_sheet_with_conditions('Tabela1', output_sheet_name='salary_',condition_fn= lambda row: row["Type"] == 'salary')
    agregator.add_sheet_with_conditions('Tabela1', output_sheet_name='fuel_', condition_fn=lambda row: row["Type"] == 'fuel')
    agregator.add_sheet_with_conditions('Tabela1', output_sheet_name='material_consumables',
                                        condition_fn=lambda row: row["Type"] == 'material' or row["Type"] == 'consumables', sort_by=["Net"])
    agregator.update_input_data()

    agregator.add_sheet_with_conditions('posadzka', output_sheet_name='posadzka_material_consumables', condition_fn=lambda row: row["Type"] == 'material' or row["Type"] == 'consumables', sort_by=['Type','TaskV2'])

    agregator.save_and_style(CFG.output_data_file)

if __name__ == '__main__':
    main()