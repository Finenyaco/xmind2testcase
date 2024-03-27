#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import logging
import os
import re
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path

"""
Convert XMind fie to Zentao testcase csv file 

Zentao official document about import CSV testcase file: https://www.zentao.net/book/zentaopmshelp/243.mhtml 
"""


def xmind_to_zentao_csv_file(xmind_file):
    """Convert XMind file to a zentao csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to zentao file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["所属模块", "用例名称", "标签", "用例等级", "前置条件", "步骤描述", "预期结果", "备注"]
    zentao_testcase_rows = [fileheader]
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        zentao_testcase_rows.append(row)

    zentao_file = xmind_file[:-6] + '.csv'
    if os.path.exists(zentao_file):
        os.remove(zentao_file)
        # logging.info('The zentao csv file already exists, return it directly: %s', zentao_file)
        # return zentao_file

    with open(zentao_file, 'w', encoding='utf8') as f:
        writer = csv.writer(f)
        writer.writerows(zentao_testcase_rows)
        logging.info('Convert XMind file(%s) to a zentao csv file(%s) successfully!', xmind_file, zentao_file)

    return zentao_file

def gen_a_testcase_row(testcase_dict):
    # case_module = '/' + gen_case_module(testcase_dict['suite'])
    case_names = str(testcase_dict['name']).split("#")
    sub_module = case_names[:-1]
    # sub_module = re.findall(r'[\w]+', sub_modules)
    case_module = '/' + gen_case_module(testcase_dict['suite']) + ('/' + '/'.join(sub_module) if sub_module else '')
    case_module = case_module.replace('>', '').strip()
    # if sub_module:
    #     case_module = '/' + gen_case_module(testcase_dict['suite'])+ '/' + '/'.join(sub_module)
    case_title = case_names[-1].replace('>','',1).strip()
    # case_title = testcase_dict['name']
    case_precontion = testcase_dict['preconditions']
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])
    # case_keyword = ''
    case_priority = gen_case_priority(testcase_dict['importance'])
    # case_type = gen_case_type(testcase_dict['execution_type'])
    case_type = testcase_dict['execution_type']
    case_type = case_type.replace('，', ',')
    # case_apply_phase = '迭代测试'
    row = [case_module, case_title, case_type, case_priority, case_precontion, case_step, case_expected_result]
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


def gen_case_type(case_type):
    mapping = {1: '手动', 2: '自动'}
    if case_type in mapping.keys():
        return mapping[case_type]
    else:
        return '手动'


if __name__ == '__main__':
    xmind_file = '../docs/zentao_testcase_template.xmind'
    zentao_csv_file = xmind_to_zentao_csv_file(xmind_file)
    print('Conver the xmind file to a zentao csv file succssfully: %s', zentao_csv_file)