# coding: utf-8
# Author: Alice

import sys
from datetime import datetime

def add_week_indicator():
    from week_indicator import indicateWeeks
    process = indicateWeeks()
    process.set_attribute()
    process.main_process()

def readExcel():
    from excel_reader import excelReader
    process = excelReader()
    process.main()

def main(i):
    welcome_text = '''
欢迎使用该工具
0.退出
1.生成周数指示器
2.读取生成ICAL文件
...待续...
'''
    if i== 1:
        print(welcome_text)

    func = input('请输入选择: ')
    print()
    if func == '0':
        print('Thanks!\n')
        sys.exit()
    elif func == '1':
        add_week_indicator()
        print('\nFinished...')
    elif func == '2':
        readExcel()
        print('\nFinished...')
    else:
        print('Wrong')


# 初始化程序会出现欢迎语，执行完某功能后能再次选择
if __name__ == '__main__':
    i = 1
    while True:
        main(i)
        # print('', end='\n')
        # i += 1

