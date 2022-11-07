"""
    @功能：文件读取、写入等相关操作
"""
import json
import os

import xlrd2
import xlwt
from xlutils.copy import copy


# 将json文件对象中的数据直接转换成 Python 列表
def read_from_json(json_file_path, model, char_set):
    i = 0
    if '' == json_file_path.strip():
        json_file_path = "../../resource/ads_product_key_info.json"
    if '' == char_set.strip():
        char_set = "UTF-8"
    if '' == model.strip():
        model = "r"

    with open(json_file_path, model, encoding=char_set) as f:
        json_data = json.load(f)
        for table in json_data['data']:
            i += 1
        # print("根据json报文读取到的列表为：\n" + str(json_data['data']))
        print(f"共{i}行")   # 3224行
        return json_data['data']


if __name__ == '__main__':
    read_from_json('../../resource/ads_enterprise_environment_punishment.json', 'r', 'UTF-8')


# 根据excel路径将数据读入到工作薄中（读入内存）
def read_from_excel(excel_file_path):
    return xlrd2.open_workbook(excel_file_path, 'rw')


def read_excel_xls(path):
    workbook = xlrd2.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    for i in range(0, worksheet.nrows):
        for j in range(0, worksheet.ncols):
            print(worksheet.cell_value(i, j), "\t", end="")  # 逐行逐列读取数据
        print()


# 将数据写入到Excel中
def write_diff_into_excel(excel_file_path, write_data_list):
    if os.path.exists(excel_file_path):
        write_excel_xls_append(excel_file_path, write_data_list)
    else:
        write_excel_xls(excel_file_path, 'sheet1', write_data_list)


def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)  # 在工作簿中新建一个表格

    # 创建我们需要的第一行的标头数据
    heads = ['类型', '差异']
    ls = 0
    # 将标头循环写入表中
    for head in heads:
        sheet.write(0, ls, head)
        ls += 1

    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i+1, j, value[i][j])  # 像表格中写入数据（对应的行和列）
            # sheet.write(i+2, j, ' ')  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd2.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            # 多加一个空行
            new_worksheet.write(i + rows_old + 1, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")
