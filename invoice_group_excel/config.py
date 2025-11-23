from pathlib import Path


class Config:

    data_dir: str = 'data'
    input_data_file: str = data_dir + '/costs_list_data.xlsx'
    output_data_file: str = data_dir + '/output_costs_list_data.xlsx'

    date_columns =["B"]
    date_format = "MM-DD-YYYY"


    summary_col_description = "A","SUMMARY"

CFG: Config = Config()