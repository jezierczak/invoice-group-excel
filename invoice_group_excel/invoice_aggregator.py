from typing import Any, Callable

from invoice_group_excel.kmexcel.excel.manager import ExcelManager


class InvoiceAggregator(ExcelManager):
    def __init__(self, input_data_file: str, input_data_sheets: list[str]) -> None:
        super().__init__(input_data_file)
        # self.input_data_file = input_data_file
        self.input_data_sheets = input_data_sheets
        # self._excel_manager = ExcelManager(input_data_file)
        self._input_data: dict[str, list[dict[str, Any]]] = {}
        for data_sheet in input_data_sheets:
            self._input_data[data_sheet] = self.read_sheet(data_sheet)


        # input_data: dict[str, list[dict[str, Any]]] = excel_manager.read_sheet('Tabela1')
    def get_input_data(self) -> dict[str,list[dict[str, Any]]]:
        return self._input_data

    def get_sheet_input_data(self,sheet_name: str) -> list[dict[str, Any]]:
        return self._input_data.get(sheet_name)

    def get_sheet_input_data_with_conditions(self,sheet_name: str, condition_fn: Callable[[dict[str,Any]],bool]) -> list[dict[str, Any]]:
        return [output for output in self.get_sheet_input_data(sheet_name) if condition_fn(output) ]
            #for key,value in output.items() if condition_fn(key,value)]

    def group_sheets_by_key(self,sheet_name: str, key_to_group: str) -> dict[str,list[dict[str, Any]]]:
        # list_of_keys = self.get_sheet_input_data(sheet_name)[0].keys()
        output = {}
        # for key in list_of_keys:
        set_of_values: set[str] = set()
        input_data = self.get_sheet_input_data(sheet_name)

        for row in input_data:
            set_of_values.add(row[key_to_group])

        for value in set_of_values:
            output[value] = self.get_sheet_input_data_with_conditions(sheet_name, lambda v: v[key_to_group] == value)

        return output

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


    # def add_sheets_grouped_by_key(self,sheet_name: str, key_to_group: str) -> None:




