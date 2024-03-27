"""
coding: utf-8
@Filename: pingcode.py
@DateTime: 2024/3/27 10:23
@Author: Sophia
"""
import xlwt
import logging
import os
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path
import datetime

def xmind_to_zentao_csv_file(xmind_file):
    """Convert XMind file to a pingcode xls file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to pingcode file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)
    pingcode_file = xmind_file[:-6] + '.xls'
    if os.path.exists(pingcode_file):
        os.remove(pingcode_file)

    style_header= xlwt.easyxf('pattern: pattern solid, fore_colour 0x16; font: bold on;alignment:HORZ CENTER;'
                             'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A')
    style = xlwt.easyxf('pattern: pattern solid, fore_colour White;'
                             'borders:left 1,right 1,top 1,bottom 1')#未居中无背景颜色0x34
    style_center = xlwt.easyxf('pattern: pattern solid, fore_colour White;alignment:HORZ CENTER;'
                             'borders:left 1,right 1,top 1,bottom 1')  # 无背景颜色居中
    style_num = xlwt.easyxf('pattern: pattern solid, fore_colour White;alignment:HORZ CENTER;'
                                   'borders:left 1,right 1,top 1,bottom 1,bottom_colour 0x3A', num_format_str='#,##0.00')  # 无背景颜色居中

    with open(pingcode_file, 'w', encoding='utf8') as f:
        f.close()
        workbook = xlwt.Workbook()
        testcase_worksheet = workbook.add_sheet('功能测试用例', cell_overwrite_ok='True')
        # for i in range(13):
        #     testcase_worksheet.col(i).width = (13 * 367)
        #     if i in [6, 7]:
        #         testcase_worksheet.col(i).width = (25 * 500)
        #     if i in [7]:
        #         testcase_worksheet.col(i).width = (20 * 400)
        #     if i in [4, 8, 9, 10, 11, 12]:
        #         testcase_worksheet.col(i).width = (13 * 200)
        fileheader = ['模块', '编号', '*标题', '维护人', '用例类型', '重要程度', '测试类型', '预估工时', '剩余工时',
                      '关联工作项', '前置条件', '步骤描述', '预期结果', '关注人', '备注']
        pingcode_testcase_rows = [['import testcase'], fileheader]
        for testcase in testcases:
            row = gen_a_testcase_row(testcase)
            pingcode_testcase_rows.append(row)
        row_index = 0
        for row in pingcode_testcase_rows:
            index = 0
            for col in row:
                testcase_worksheet.write(row_index, index, col, style)
                index += 1
            row_index += 1
        workbook.save(pingcode_file)
        logging.info('Convert XMind file(%s) to a pingcde xls file(%s) successfully!', xmind_file, pingcode_file)
    return pingcode_file

def gen_a_testcase_row(testcase_dict):
    case_names = str(testcase_dict['name']).split("#")
    sub_module = case_names[:-1]
    case_module = gen_case_module(testcase_dict['suite']) + ('/' + '/'.join(sub_module) if sub_module else '')
    case_title = case_names[-1].strip()
    case_precontion = testcase_dict['preconditions']
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])
    case_priority = gen_case_priority(testcase_dict['importance'])
    # case_label = testcase_dict['execution_type']
    # case_label = case_label.replace('，', ',')
    # row = [case_module, case_title, case_label, case_priority, case_precontion, case_step, case_expected_result]
    row = [case_module, '', case_title, '', '功能测试' , case_priority, '手动', '', '', '', case_precontion, case_step, case_expected_result]
    return row

def gen_case_module(module_name):
    if module_name:
        module_name = module_name.replace('（', '(')
        module_name = module_name.replace('）', ')')
    else:
        module_name = '/'
    return module_name

def gen_case_step_and_expected_result(steps):
    case_step = ''
    case_expected_result = ''

    for step_dict in steps:
        case_step += str(step_dict['step_number']) + '. ' + step_dict['actions'].replace('\n', '').strip() + '\n'
        case_expected_result += str(step_dict['step_number']) + '. ' + \
            step_dict['expectedresults'].replace('\n', '').strip() + '\n' \
            if step_dict.get('expectedresults', '') else ''

    return case_step, case_expected_result

def gen_case_priority(priority):
    mapping = {1: 'P1', 2: 'P2', 3: 'P3'}
    if priority in mapping.keys():
        return mapping[priority]
    else:
        return 'P2'


if __name__ == '__main__':
    xmind_file = '/Users/sophia/seal/walrus-testcases/walrus.xmind'
    pingcode_xls_file = xmind_to_zentao_csv_file(xmind_file)
    print('Conver the xmind file to a pingcode csv succssfully: %s', pingcode_xls_file)