"""
    @功能：用于读取Excel中数据，并将其转换为列表
"""
from decimal import Decimal
from types import NoneType

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

    # (分边界)企业数据————分边界污染物物理排放强度表
    if sheet_name == '分边界污染物物理排放强度':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')

    # (分边界)企业数据————分边界污染物经济排放强度表
    if sheet_name == '分边界污染物经济排放强度':
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
            table['年度碳排放（万吨）'] = round(table['年度碳排放（万吨）'], 0)
            table['年度碳排放强度（吨/万元）'] = round(table['年度碳排放强度（吨/万元）'], 2)
            table['年度营业收入（亿元）'] = round(table['年度营业收入（亿元）'], 0)
            if table['公司简称'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 企业ESG评级表
    if sheet_name == '企业ESG评级':
        for table in new_val_list:
            environment_total_score = table['E_总分']
            environment_actual_score = table['E_实际得分']
            environment_score_rate = table['E-得分率']
            environment_average_score_rate = table['E_得分率算数平均值']
            environment_weight = table['E_权重']
            society_total_score = table['S_总分']
            society_actual_score = table['S_实际得分']
            society_score_rate = table['S-得分率']
            society_average_score_rate = table['S_得分率算数平均值']
            society_weight = table['S_权重']
            governance_total_score = table['G_总分']
            governance_actual_score = table['G_实际得分']
            governance_score_rate = ['G-得分率']
            governance_average_score_rate = table['G_得分率算数平均值']
            governance_weight = table['G_权重']
            esg_total_score = table['ESG总分']
            if type(environment_total_score) is float:
                table['G_得分率算数平均值'] = round(environment_total_score, 9)
            if type(environment_actual_score) is float:
                table['E_实际得分'] = round(environment_actual_score, 9)
            if type(environment_score_rate) is float:
                table['E-得分率'] = round(environment_score_rate, 9)
            if type(environment_average_score_rate) is float:
                table['E_得分率算数平均值'] = round(environment_average_score_rate, 9)
            if type(environment_weight) is float:
                table['E_权重'] = round(environment_weight, 9)
            if type(society_total_score) is float:
                table['S_总分'] = round(society_total_score, 9)
            if type(society_actual_score) is float:
                table['S_实际得分'] = round(society_actual_score, 9)
            if type(society_score_rate) is float:
                table['S-得分率'] = round(society_score_rate, 9)
            if type(society_average_score_rate) is float:
                table['S_得分率算数平均值'] = round(society_average_score_rate, 9)
            if type(society_weight) is float:
                table['S_权重'] = round(society_weight, 9)
            if type(governance_total_score) is float:
                table['G_总分'] = round(governance_total_score, 9)
            if type(governance_actual_score) is float:
                table['G_实际得分'] = round(governance_actual_score, 9)
            if type(governance_score_rate) is float:
                table['G-得分率'] = round(governance_score_rate, 9)
            if type(governance_average_score_rate) is float:
                table['G_得分率算数平均值'] = round(governance_average_score_rate, 9)
            if type(governance_weight) is float:
                table['G_权重'] = round(governance_weight, 9)
            if type(esg_total_score) is float:
                table['ESG总分'] = round(esg_total_score, 4)
            if table['ESG级别'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

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
            table['G治理_得分率算数平均值'] = round(table['G治理_得分率算数平均值'], 2)
            table['S战略_得分率算数平均值'] = round(table['S战略_得分率算数平均值'], 2)
            table['R风险_得分率算数平均值'] = round(table['R风险_得分率算数平均值'], 2)
            table['M&T指标与目标部分_得分率算数平均值'] = round(table['M&T指标与目标部分_得分率算数平均值'], 2)
            if table['评级结果'] != '':
                new_val_list2.append(table)
                i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 环境违法-罚款,责令关键字-抽检
    if sheet_name == '罚款,责令关键字':
        for table in new_val_list:
            table['罚款金额(万元)'] = round(table['罚款金额(万元)'], 4)
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 环境违法-罚款,责令关键字-抽检
    if sheet_name == '市生态环境局关键字':
        for table in new_val_list:
            table['罚款金额(万元)'] = round(table['罚款金额(万元)'], 4)
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))


    # 排污许可-废气&废水关键字-抽检
    if sheet_name == '废气&废水':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 环境信用-抽检（江苏）
    if sheet_name == '江苏':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 碳市场-机组及生产设施信息
    if sheet_name == '机组及生产设施信息':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 碳市场-委托检测关键词
    if sheet_name == '委托检测关键词':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 碳市场-排放量信息
    if sheet_name == '排放量信息':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 碳市场-基本信息
    if sheet_name == '基本信息':
        for table in new_val_list:
            # table['统一社会信用代码'] = table['统一社会信用代码'].split('\'')[1]
            # table[''] = table['统一社会信用代码'].split('\'')[1]
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 绿色荣誉评价-绿色工厂
    if sheet_name == '绿色工厂':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 绿色荣誉评价-绿色供应链管理示范企业
    if sheet_name == '绿色供应链管理示范企业':
        for table in new_val_list:
            new_val_list2.append(table)
            i = table_none_format(new_val_list2)
        print(f'共{i}行')
    # print("根据Excel表格读取到的列表为: \n" + str(new_val_list2))

    # 绿色荣誉评价-绿色设计产品
    if sheet_name == '绿色设计产品':
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
    my_excel_file_path = '../../resource/carbon_market_confirm_method.xlsx'
    my_sheet_name = '测试'
    final_dict = read_excel_into_list(my_excel_file_path, my_sheet_name)
