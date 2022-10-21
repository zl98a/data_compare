"""
    @功能：用于将json数据转换为列表
"""
from src.utils import file_read_write_utils


def read_json_into_list(json_file_path, model, char_set):
    return file_read_write_utils.read_from_json(json_file_path, model, char_set)


if __name__ == '__main__':
    my_json_file_path = '../../resource/ads_year_province_gdp.json'
    read_json_into_list(my_json_file_path, "", "")