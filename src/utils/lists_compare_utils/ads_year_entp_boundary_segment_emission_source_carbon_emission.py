"""
    @功能：比较两个列表中的字典是否相同
    注意：1、由于比对的mysql表精度和Excel精度可能存在差异
        2、
"""
from src.utils.read_excel_into_list_utils import read_excel_into_list
from src.utils.read_json_into_list_utils import read_json_into_list


def compare_all_different(list1, list2):
    diff = []
    i = 0
    a = 0
    print("开始比较list1(表格读取) 和 list2（json读取） 的所有差异:")
    # 首先判断Excel 和 json 数据总数是否一致
    # print(count(list1))
    # print(count(list2))
    if count(list1) != count(list2):
        print("表格数据总数和json数据总数对应不上！")
    else:
        # 其次获取list1中的一条字典数据，再获取list2中对应的一条字典数据，进行两条字典数据的比对
        for dict1 in list1:
            year = dict1['年份']  # 年份
            enterprise_name = dict1['企业名称']  # 企业名称
            boundary = dict1['边界']  # 边界
            subdivide_emission_source = dict1['细分排放源']  # 细分排放源
            carbon_emission = dict1['分边界细分排放源碳排放量（t）']
            # 通过三个字段标志另一个列表中的唯一字典
            dict2 = get_dict_wih_same_key(new_list1=[year, enterprise_name, boundary, subdivide_emission_source, carbon_emission], list2=list2)
            # 接下来就是两条字典数据比对
            try:
                differ = set(dict1.items()) ^ set(dict2.items())
                a += 1
                if len(differ) != 0:
                    i += 1
                    print(f"【{i}】--\ndict1（表格读取）：\n{dict1}\ndict2（json读取）:\n{dict2}\n相同关键字的栏位取值有差异，差异是:{differ}")
                    for item in list(differ):
                        diff.append(item)
                else:
                    # i += 1
                    # print(f"第{i}行数据比对一致")
                    pass
            except AttributeError:
                pass
        print(a)
        return diff


def get_dict_wih_same_key(new_list1, list2):
    i = 0
    for dict2 in list2:
        if dict2['年份'] == new_list1[0] and dict2['企业名称'] == new_list1[1] and dict2['边界'] == new_list1[2] \
                and dict2['细分排放源'] == new_list1[3] and dict2['分边界细分排放源碳排放量（t）'] == new_list1[4]:
            return dict2


def compare_different_wih_same_key(dict1, dict2):
    # 关键字不同是否跳出开关，默认关闭
    break_switch = False
    print("==字典1为：" + str(dict1))
    print("==字典2为：" + str(dict2))
    print("1）比较两个字典dict1 和 dict2 的关键字是否相同")
    has_diff_key = dict1.keys() ^ dict2
    if len(has_diff_key) > 0:
        print(" " * 3 + "两个字典存在不一样的key")
        if len(dict1.keys() - dict2.keys()) > 0:
            print(" " * 3 + "在字典1中存在但在字典2中不存在的key为：" + str(dict1.keys() - dict2.keys()))

        if len(dict2.keys() - dict1.keys()) > 0:
            print(" " * 3 + "在字典2中存在但在字典1中不存在的key为：" + str(dict2.keys() - dict1.keys()))

        if break_switch:
            return has_diff_key

    print("2）比较相同关键字的两个字典dict1 和 dict2 的所有差异")
    diff = dict1.keys() & dict2
    diff_vals = [(k, dict1[k], dict2[k]) for k in diff if dict1[k] != dict2[k]]
    # print(diff_vals)
    if len(diff_vals) == 0:
        print("3）字典1和字典2相同关键字的栏位取值完全相同")
    else:
        print("4）典1和字典2相同关键字的栏位取值所有差异为：" + str(diff_vals))
    return diff_vals


def count(list3):
    i = 0
    for table in list3:
        i += 1
    return i


if __name__ == '__main__':
    my_excel_file_path = '../../../resource/enterprise_data_boundary.xlsx'
    my_sheet_name = '分边界细分排放源碳排放'
    list1 = read_excel_into_list(my_excel_file_path, my_sheet_name)
    my_json_file_path = '../../../resource/ads_year_entp_boundary_segment_emission_source_carbon_emission.json'
    list2 = read_json_into_list(my_json_file_path, "", "")
    compare_all_different(list1, list2)