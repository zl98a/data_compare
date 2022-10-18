"""
    @功能：用于比对Excel中数据和json中存储数据的栏位值是否相同（全部转成String进行比较），并将具体差异输出到新的Excel中
"""
from src.constant import sys_constant
from src.utils import read_excel_into_list_utils, read_json_into_list_utils, lists_compare_utils_key_product_info, file_read_write_utils


def process():
    # 1、读取excel数据，转为字典格式
    print("=======================Excel数据和Json中数据比对处理开始=======================")
    print("\n=======================1、读取Excel文件中对应sheet页数据，转化为列表格式=======================")
    excel_data_list= read_excel_into_list_utils.read_excel_into_list(sys_constant.excel_file_path,
                                                                      sys_constant.my_sheet_name)

    # 2、读取json数据，转为字典格式
    print("\n=======================2、读取json报文数据，转化为列表格式=======================")
    json_list = read_json_into_list_utils.read_json_into_list(sys_constant.json_file_path, 'r', 'UTF-8')

    # 3、开始比较excel数据和json数据栏位取值是否相同
    print("\n=======================3、开始比较excel数据和json栏位取值是否相同=======================")
    diff_vals_list = lists_compare_utils.compare_all_different(excel_data_list, json_list)

    # 4、将具体差异写入到Excel中
    print("\n=======================4、将具体差异写入到Excel中=======================")
    file_read_write_utils.write_diff_into_excel(sys_constant.excel_data_diff_file_path, diff_vals_list)

    # 5、读取Excel中的具体差异
    print("\n=======================5、读取Excel中的具体差异=======================")
    file_read_write_utils.read_excel_xls(sys_constant.excel_data_diff_file_path)


if __name__ == '__main__':
    # 具体比对处理
    process()