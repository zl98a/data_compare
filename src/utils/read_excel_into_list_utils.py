"""
    @功能：用于读取Excel中数据，并将其转换为列表
"""
from src.utils import file_read_write_utils


# 根据excel路径及 sheet页名称将对应的sheet页内容读入到字典中
def read_excel_into_list(excel_file_path, sheet_name):
    # 1、根据excel路径将数据读入到工作薄中（读入内存）
    work_book = file_read_write_utils.read_from_excel(excel_file_path)
    # 2、将对应的sheet页内容读入到字典中
    sheet = work_book.sheet_by_name(sheet_name)
    return convert_sheet_context_into_list(sheet,sheet_name)


# 将对应的sheet页内容读入到字典中
def convert_sheet_context_into_list(sheet, sheet_name):
    fld_key_list = []
    fld_val_list = []
    new_val_list = []
    new_val_list2 = []
    for col_index in range(sheet.nrows):
        for row_index in range(sheet.ncols):
            # 将栏位名称存入key_list
            fld_key_list.append(sheet.cell_value(0, row_index))
            # 将第二行及之后栏位值存入value_list
            if col_index >= 1:
                fld_val_list.append(sheet.cell_value(col_index, row_index))
        if col_index >= 1:
            dict1 = dict(zip(fld_key_list, fld_val_list))
            new_val_list.append(dict1)

    # 产品关键信息表
    if sheet_name == '产品关键信息':
        for table in new_val_list:
            if table['产品或服务产能'] != '' or table['产品或服务产量'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 分国家年度GDP表
    if sheet_name == '分国家年度GDP':
        for table in new_val_list:
            if table['gdp'] != '' or table['gdp_person'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 分年份分省份GDP表
    if sheet_name == '分年份分省份GDP':
        for table in new_val_list:
            if table['gdp'] != '' or table['gdp_person'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 分国家年度人口数量表
    if sheet_name == '分国家年度人口数':
        for table in new_val_list:
            if table['总人口'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 分年份分省份人口数量表
    if sheet_name == '分年份分省份人口数量':
        for table in new_val_list:
            if table['总人口'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 中国百家上市公司双碳领导力排行榜表
    if sheet_name == '中国百家上市公司双碳领导力排行榜':
        for table in new_val_list:
            if table['公司简称'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 分行业企业双碳领导力排行榜表
    if sheet_name == '分行业企业双碳领导力排行榜':
        for table in new_val_list:
            if table['公司简称'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))
    return new_val_list2


# 表内数据空值转换
def table_none_format(new_val_list):
    i = 0
    for table in new_val_list:
        i += 1
        for value in table:
            if table[value] == '':
                table[value] = None
    return i


if __name__ == '__main__':
    my_excel_file_path = '../../resource/mysqlTest.xlsx'
    my_sheet_name = '分行业企业双碳领导力排行榜'
    final_dict = read_excel_into_list(my_excel_file_path, my_sheet_name)
