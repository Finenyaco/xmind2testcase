#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import csv
import logging
import os
import re
import argparse
from xmind2testcase.utils import get_xmind_testcase_list, get_absolute_path

"""
Convert XMind fie to GPUStack testcase csv file 

"""

def xmind_to_gpustack_csv_file(xmind_file):
    """Convert XMind file to a csv file"""
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to csv file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ["所属模块", "用例名称", "标签", "用例等级", "前置条件", "步骤描述", "预期结果", "备注"]
    gpustack_testcase_rows = [fileheader]
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        gpustack_testcase_rows.append(row)

    csv_file = xmind_file[:-6] + '.csv'
    if os.path.exists(csv_file):
        os.remove(csv_file)
        # logging.info('The csv file already exists, return it directly: %s', csv_file)
        # return csv_file

    with open(csv_file, 'w', encoding='utf8') as f:
        writer = csv.writer(f)
        writer.writerows(gpustack_testcase_rows)
        logging.info('Convert XMind file(%s) to csv file(%s) successfully!', xmind_file, csv_file)

    return csv_file


def gen_a_testcase_row(testcase_dict):
    # case_module = '/' + gen_case_module(testcase_dict['suite'])
    case_names = str(testcase_dict['name']).split("#")
    sub_module = case_names[:-1]
    # sub_module = re.findall(r'[\w]+', sub_modules)
    case_module = '/' + gen_case_module(testcase_dict['suite']) + ('/' + '/'.join(sub_module) if sub_module else '')
    case_module = case_module.replace('>', '').strip()
    # if sub_module:
    #     case_module = '/' + gen_case_module(testcase_dict['suite'])+ '/' + '/'.join(sub_module)
    case_title = format_case_title(case_names[-1])
    # case_title = testcase_dict['name']
    case_precontion = testcase_dict['preconditions']
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['steps'])
    # case_keyword = ''
    case_priority = gen_case_priority(testcase_dict['importance'])
    # case_type = gen_case_type(testcase_dict['execution_type'])
    case_type = testcase_dict['execution_type']
    case_type = case_type.replace('，', ',')
    # case_apply_phase = '迭代测试'
    summary = testcase_dict['summary']
    row = [case_module, case_title, case_type, case_priority, case_precontion, case_step, case_expected_result, summary]
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
        # 第一行输入是不用添加换行符，其他行添加时需要先换行再输入
        if case_step == '':
            case_step += str(step_dict['step_number']) + '. ' + step_dict['actions'].replace('\n', '').strip()
        else:
            case_step += '\n' + str(step_dict['step_number']) + '. ' + step_dict['actions'].replace('\n', '').strip()
        if case_expected_result == '':
            case_expected_result += str(step_dict['step_number']) + '. ' + \
                step_dict['expectedresults'].replace('\n', '').strip() \
                if step_dict.get('expectedresults', '') else ''
        else:
            case_expected_result += '\n' + str(step_dict['step_number']) + '. ' + \
                                    step_dict['expectedresults'].replace('\n', '').strip() \
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

def format_case_title(case_title):
    case_title = case_title.strip()
    if case_title.startswith('>'):
        case_title = case_title[1:]
    else:
        case_title = case_title
    return case_title


def main():
    parser = argparse.ArgumentParser(description='Convert XMind to GPUStack CSV')
    parser.add_argument('xmind_file', nargs='?', default='', help='Path to XMind file')
    args = parser.parse_args()

    if not args.xmind_file:
        # 如果未提供参数，提示用户输入
        args.xmind_file = input("请输入XMind文件路径: ").strip()

    gpustack_csv_file = xmind_to_gpustack_csv_file(args.xmind_file)
    print(f'转换成功: {gpustack_csv_file}')


if __name__ == '__main__':
    main()
    # xmind_file = '/Users/sophia/seal/gpustack-test/测试用例模板规则.xmind'
    # xmind_file = '/Users/sophia/seal/gpustack-test/GPUStack-testcase.xmind'
    # gpustack_csv_file = xmind_to_gpustack_csv_file(xmind_file)
    # print('Conver the xmind file to csv file succssfully: %s', gpustack_csv_file)