#!/usr/bin/python 3.10
# -*- coding: utf-8 -*- 
#
# @Time    : 2022-10-28 14:28
# @Author  : 发发
# @QQ      : 1315337973
# @File    : CUIT_LongQuan_activity_certificate_automation_main.py
# @Software: PyCharm

import os
import shutil
from time import sleep

from pandas import read_excel
from tqdm import tqdm
from docx import Document
from docxcompose.composer import Composer
from docx2pdf import convert

from CUIT_LongQuan_activity_certificate_automation_tools import one_page_and_save_file

print('欢迎使用 发发 的 CUIT_LongQuan_activity_certificate_automation V1.0 （成都信息工程大学龙泉校区活动证明自动化 V1.0）')
print('源码随后开源 与 版本迭代更新：https://github.com/lovely-fafa')
print('程序员不对结果准确性承担责任！！！')
print('接下来会让你输入一些东西，输入一个按一下回车（憨憨应该知道什么是回车吧...）')
sleep(2)
print('*' * 100)

while True:
    print('现在我需要一个 Excel 文件，你只需要保证 里面的第一个工作簿 中有 "姓名"、"学号"、"学院" 这 3 列（顺序可以不一样哦）')
    while True:
        road = input(r'请输入 Excel 的绝对路径（例如 C:\Users\lenovo\Downloads\191413348_2_游园会活动证明_14_14.xlsx）：').strip()
        if not os.path.exists(road):
            if '.xlsx' not in road:
                print('憨憨咩，文件名都不加后缀名？')
            else:
                print('憨憨咩，文件都不存在...')
        else:
            break

    data = read_excel(road)
    flag = True
    for i in ["姓名", "学号", "学院"]:
        if i not in data.columns:
            flag = False

    if not flag:
        print('憨憨去改 Excel 文件')
        sleep(2)
    else:
        break
data['学号'] = data['学号'].astype(str)

print('*' * 100)
activity_info_dict = {}
for prompt, info_name in zip(['活动时间（例如 2022年10月22日18:00至10月23日21:30 请注意英文冒号和中间的 "至" 字）：',
                              '举办组织/协会（例如 共青团成都信息工程大学委员会学生社团管理部）：',
                              '活动名称（例如 全国大学生数学建模竞赛 注意后面不要有 "活动" 两个字！)：',
                              '落款时间（例如 2022年2月31日）：'],
                             ['activity_date', 'organization_name', 'activity_name', 'inscribe_date']):
    while True:
        info_input = input(prompt).strip()
        if len(info_input) == 0:
            print('憨憨不要装怪 [○･｀Д´･ ○]')
        else:
            activity_info_dict[info_name] = info_input
            break

print('*' * 100)
print('嘿嘿，我开始处理啦...')
print('第一步：创建依赖文件夹...')
if not os.path.exists(r'C:/output_fafa/'):
    os.makedirs(r'C:/output_fafa/')

if not os.path.exists(f'C:/output_fafa/{activity_info_dict["activity_name"]}'):
    os.makedirs(f'C:/output_fafa/{activity_info_dict["activity_name"]}')

road_all = f'C:/output_fafa/{activity_info_dict["activity_name"]}/{activity_info_dict["activity_name"]}_all'
road_docx = f'C:/output_fafa/{activity_info_dict["activity_name"]}/{activity_info_dict["activity_name"]}_docx'
if not os.path.exists(road_all):
    os.makedirs(road_all)
if not os.path.exists(road_docx):
    os.makedirs(road_docx)
print('创建依赖文件夹成功！')

print('*' * 100)
print('第二步：填入 docx 文件...')
all_data_list = []
for i, j, k in zip(data['姓名'], data['学院'], data['学号']):
    all_data_list.append([i.strip(), j.strip(), k.strip()])

max_file = len(all_data_list) // 13

for i in tqdm(range(max_file + 1)):
    one_page_and_save_file(one_page_table_data=all_data_list[i * 13: (i + 1) * 13],
                           file_name=f'{road_docx}/{i}_{activity_info_dict["activity_name"]}.docx',
                           **activity_info_dict)
print('填入 docx 文件成功！')

print('*' * 100)
print('第三步：合并 word 文件...')

original_docx_path = road_docx
new_docx_path = f'{road_all}/{activity_info_dict["activity_name"]}.docx'

all_file_path = []
for file_name in os.listdir(original_docx_path):
    all_file_path.append(f'{original_docx_path}/{file_name}')

if len(all_file_path) > 1:
    first_document = Document(all_file_path[0])
    first_document.add_page_break()
    middle_new_docx = Composer(first_document)

    for index, word in tqdm(enumerate(all_file_path[1:])):
        word_document = Document(word)
        if index != len(all_file_path) - 2:
            word_document.add_page_break()
        middle_new_docx.append(word_document)

    middle_new_docx.save(new_docx_path)

else:
    shutil.copyfile(all_file_path[0], new_docx_path)
print('合并 word 文件成功！')

print('*' * 100)
print('第四步：转为 PDF 文件...')
convert(new_docx_path, new_docx_path.rsplit('.')[0] + '.pdf')
print('转为 PDF 文件成功！')

print('*' * 100)
print(f'小可爱快去 "C:/output_fafa/{activity_info_dict["activity_name"]}" 这一个文件夹看看叭 (*╹▽╹*)')
input("please input any key to exit!")
