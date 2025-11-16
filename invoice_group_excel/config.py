from pathlib import Path


class Config:

    data_dir: str = 'data'
    input_data_file: str = data_dir + '/costs_list_data.xlsx'


CFG: Config = Config()