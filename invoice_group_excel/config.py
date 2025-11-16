from pathlib import Path


class Config:

    data_dir: str = 'data'
    input_data_file: str = data_dir + '/costs_list_data.xlsx'
    output_data_file: str = data_dir + '/output_costs_list_data.xlsx'

    date_columns =["B"]
    date_format = "MM-DD-YYYY"

    summary_description = {"A":"SUMMARY","D":"=SUM(D2:D{last})","E":"=SUM(E2:E{last})","F":"=SUM(F2:F{last})","G":"=SUM(G2:G{last})"}


CFG: Config = Config()