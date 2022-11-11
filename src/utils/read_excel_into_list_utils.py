"""
    @功能：用于读取Excel中数据，并将其转换为列表
"""
from decimal import Decimal

import pandas as pd

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

    # 企业数据————企业产值碳排放强度表
    if sheet_name == '企业产值碳排放强度':
        for table in new_val_list:
            if table['企业产值碳排放强度'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业产品碳排放强度表
    if sheet_name == '企业产品碳排放强度':
        for table in new_val_list:
            if table['企业产品碳排放强度'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业碳排放表
    if sheet_name == '企业碳排放':
        for table in new_val_list:
            if table['企业碳排放总量（t）'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————产品关键信息表
    if sheet_name == '产品关键信息':
        for table in new_val_list:
            if table['产品或服务产能'] != '' or table['产品或服务产量'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业分结构碳排放表
    if sheet_name == '企业分结构碳排放':
        for table in new_val_list:
            table['企业分结构碳排放量（t）'] = round(table['企业分结构碳排放量（t）'], 4)
            if table['排放结构'] != '' or table['企业分结构碳排放量（t）'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业分排放源碳排放表
    if sheet_name == '企业分排放源碳排放':
        for table in new_val_list:
            table['企业分排放源碳排放量（t）'] = round(table['企业分排放源碳排放量（t）'], 4)
            if table['排放源'] != '' or table['企业分排放源碳排放量（t）'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业细分排放源碳排放表
    if sheet_name == '企业细分排放源碳排放':
        for table in new_val_list:
            table['企业细分排放源碳排放量（t）'] = round(table['企业细分排放源碳排放量（t）'], 4)
            if table['细分排放源'] != '' or table['企业细分排放源碳排放量（t）'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业能源消费标准量表
    if sheet_name == '企业能源消费标准量':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业能源消费总量表
    if sheet_name == '企业能源消费总量':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业分品种能源消费实物量表
    if sheet_name == '企业分品种能源消费实物量':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业分品种能源消费标准量表
    if sheet_name == '企业分品种能源消费标准量':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业能效水平表
    if sheet_name == '企业能效水平':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业污染物排放表
    if sheet_name == '企业污染物排放':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业工业固体废物排放及处理利用表
    if sheet_name == '企业工业固体废物排放及处理利用':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业经济表
    if sheet_name == '企业经济':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业污染物物理排放强度表
    if sheet_name == '企业污染物物理排放强度':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # 企业数据————企业污染物经济排放强度表
    if sheet_name == '企业污染物经济排放强度':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界产品碳排放强度表
    if sheet_name == '分边界产品碳排放强度':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界产值碳排放强度表
    if sheet_name == '分边界产值碳排放强度':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界分结构碳排放表
    if sheet_name == '分边界分结构碳排放':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界分排放源碳排放表
    if sheet_name == '分边界分排放源碳排放':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界细分排放源碳排放表
    if sheet_name == '分边界细分排放源碳排放':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界碳排放表
    if sheet_name == '分边界碳排放':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界能源消费总量表
    if sheet_name == '分边界能源消费总量':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界分品种能源消费实物量表
    if sheet_name == '分边界分品种能源消费实物量':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界分品种能源消费标准量表
    if sheet_name == '分边界分品种能源消费标准量':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界能效水平表
    if sheet_name == '分边界能效水平':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界污染物排放表
    if sheet_name == '分边界污染物排放':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界工业固体废物排放及处理利用表
    if sheet_name == '分边界工业固体废物排放及处理利用':
        for table in new_val_list:
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

    # 强度榜&总量榜单表
    if sheet_name == '强度榜&总量榜单':
        for table in new_val_list:
            if table['公司简称'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 企业ESG评级表
    if sheet_name == 'ESG':
        for table in new_val_list:
            print(type(table['E_实际得分']))
            table['E_实际得分'] = round(table['E_实际得分'], 9)
            if table['ESG级别'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 企业TCFD评级表
    if sheet_name == '企业TCFD评级':
        for table in new_val_list:
            # print(type(table['G治理_权重']))
            table['G治理_权重'] = round(table['G治理_权重'], 2)
            table['S战略_权重'] = round(table['S战略_权重'], 2)
            table['R风险_权重'] = round(table['R风险_权重'], 2)
            table['M&T指标与目标部分_权重'] = round(table['M&T指标与目标部分_权重'], 2)
            table['总得分'] = round(table['总得分'], 2)
            table['G治理_得分率'] = round(table['G治理_得分率'], 9)
            table['S战略_得分率'] = round(table['S战略_得分率'], 9)
            table['R风险_得分率'] = round(table['R风险_得分率'], 9)
            table['M&T指标与目标部分_得分率'] = round(table['M&T指标与目标部分_得分率'], 9)
            table['G治理_得分率算数平均值'] = round(table['G治理_得分率算数平均值'], 9)
            table['S战略_得分率算数平均值'] = round(table['S战略_得分率算数平均值'], 9)
            table['R风险_得分率算数平均值'] = round(table['R风险_得分率算数平均值'], 9)
            table['M&T指标与目标部分_得分率算数平均值'] = round(table['M&T指标与目标部分_得分率算数平均值'], 9)
            if table['评级结果'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 环境违法（第一批）——罚款,责令改正
    if sheet_name == '罚款,责令改正':
        for table in new_val_list:
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
    my_excel_file_path = '../../resource/zmd_use_case.xlsx'
    my_sheet_name = 'ESG'
    final_dict = read_excel_into_list(my_excel_file_path, my_sheet_name)
