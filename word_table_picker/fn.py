#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
@File   : docx_table_spider.py
@Time   : 2021/05/11 16:31
@Version: 1.0
@Author : silon-j
@Desc   : 批量抓取docx文件中表格文件内容
"""

import time
import os
import logging
from docx import Document
from openpyxl import Workbook
from win32com import client as wc


def check_dir(dir_path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)


def data_clean(data):
    # clean redundant data for exporting xlsx
    ripe_data = []
    ripe_data.append(data[1][1])
    ripe_data.append(data[1][3])
    ripe_data.append(data[1][5])
    ripe_data.append(data[2][1])
    ripe_data.append(data[2][3])
    ripe_data.append(data[2][5])
    ripe_data.append(data[3][1])
    ripe_data.append(data[3][3])
    ripe_data.append(data[4][1])
    ripe_data.append(''.join(data[6][0].split('\n')[1:]))
    return ripe_data


def gen_output_xlsx(data):
    check_dir('./output')
    # create a xlsx file
    wb = Workbook()
    ws = wb.worksheets[0]
    # set header
    ws.append(['姓名', '教龄', '职称', '联系电话', '微信号', '授课科目', '学校名称', '参赛类型', '作品名称', '视频作品介绍'])
    for file_data in data:
        ws.append(file_data)
    wb.save('./output/{time}.xlsx'.format(time=time.strftime("%m-%d__%H-%M", time.localtime())))


def save_doc_to_docx(dir_path):  # doc转docx
    '''
    :param rawpath: 传入和传出文件夹的路径
    :return: None
    '''
    word = wc.Dispatch("Word.Application")
    # 不能用相对路径，老老实实用绝对路径
    # 需要处理的文件所在文件夹目录
    files = os.listdir(input_dir)
    for i in range(0, len(files)):
        file_name = files[i]
        # 找出文件中以.doc结尾并且不以~$开头的文件（~$是为了排除临时文件的）
        if file_name.endswith('.doc') and not file_name.startswith('~$'):
            filepath = os.path.join(os.path.abspath(input_dir), files[i])
            print('正在转换：{}'.format(filepath))
            try:
                # 打开文件
                doc = word.Documents.Open(filepath)
                # # 将文件名与后缀分割
                rename = os.path.splitext(file_name)
                # 将文件另存为.docx
                doc.SaveAs(os.path.join(os.path.abspath(input_dir), '{}.docx'.format(rename[0])), 12)  # 12表示docx格式
                doc.Close()
            except Exception as e:
                print('转换doc至docx失败，请检查该文件：{}'.format(files[i]))
                print("error:{}".format(e))
    word.Quit()


if __name__ == '__main__':
    # get all doc files
    input_dir = './input_doc'
    check_dir(input_dir)

    # trans all .doc files to .docx
    # print(os.path.abspath(input_dir))
    save_doc_to_docx(input_dir)
    files = os.listdir(input_dir)

    # operate
    data_all = []
    for i in range(0, len(files)):
        filename = files[i]
        if filename.endswith('docx'):
            path = os.path.join(input_dir, files[i])
            # print('正在提取：{}'.format(path))
            try:
                doc = Document(path)
                tables = doc.tables
                details = []
                for i, row in enumerate(tables[0].rows[:]):
                    row_content = []
                    # 读一行中的所有单元格
                    for cell in row.cells[:]:
                        c = cell.text
                        row_content.append(c)
                    details.append(row_content)
                ripe_file_data = data_clean(details)
                data_all.append(ripe_file_data)
            except Exception as e:
                logging.warning('docx提取失败，请检查该文件：{}'.format(filename))
        else:
            print('非docx文件未处理：{}'.format(filename))

    # write down in xlsx
    gen_output_xlsx(data_all)
    print('运行结束')
    input('输入任何值退出')
