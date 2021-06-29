# coding: utf-8
# Author: Alice

import sys
# import json
import xlrd
# import os
from random import randint
from uuid import uuid4 as uid
from datetime import datetime, timedelta



class excelReader:
    def __init__(self):
        # 指定信息在 xls 表格内的列数，第一列是第 0 列
        self.config = dict()
        self.config['time'] = 0
        self.config['monday'] = 1
        self.config['tuesday'] = 2
        self.config['wednesday'] = 3
        self.config['thursday'] = 4
        self.config['friday'] = 5
        self.config['saturday'] = 6
        self.config['sunday'] = 7

        self.config['col'] = {0: 'A', 1: 'B', 2: 'C', 3: 'D', 4: 'E', 5: 'F', 6: 'G', 7: 'H'}
        self.config['row'] = {0: '1', 1: '2', 2: '3', 3: '4', 4: '5', 5: '6', 6: '7', 7: '8', 8: '9'}

        self.first_week_start_str = ''
        # self.weeks = 0
        self.created = '19700101T000000Z'

        self.color = '#70DB4A' #green
        self.name = '我的小小课程表'

        # read excel file
        try:
            self.data = xlrd.open_workbook('export.xlsx')
        except FileNotFoundError:
            print('File do not exist!')
            sys.exit()
        else:
            print('File open success\n')
        self.table = self.data.sheets()[0]
        # basic information
        self.numOfRow = self.table.nrows # get number of rows
        self.numOfCol = self.table.ncols # get number of columns

    def confirm_excel_random(self):
        print('Random check:')
        print(f'TOTAL row: {self.numOfRow}   col: {self.numOfCol}')
        colr = randint(1, 7)
        rowr = randint(3, 8)
        print(f'''{self.config['col'][colr]}{self.config['row'][rowr]}:''')
        if self.table.cell(rowr, colr).value == '':
            print('Obviuosly it\'s empty')
        else:
            print(self.table.cell(rowr, colr).value)
            # place = self.table.cell(rowr, colr).value.find('[')
            # count1 = 1
            # while place >= 0:
            #     print(f'''the {count1}th [ is at {place}''')
            #     place = self.table.cell(rowr, colr).value.find('[', place + 1)
            #     count1 += 1
        option = input('\n回车继续，输入其他内容退出\n')
        if option:
            return 1
        else:
            return 0

    # 获取三个变量
    def set_attribute(self):
        self.first_week_start_str = input('请输入第一周第一天的日期(YYYYMMDD):\n')
        # self.weeks = input('请输入总周数: ')
        self.created = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    
    def excelTOical(self):
        # name\n[teacher]\n[time][place]

        # 将开始日期转换为标准时间元组
        self.first_week_start = datetime.strptime(self.first_week_start_str, '%Y%m%d')
        #获取时间变化量
        self.delta_1 = timedelta(days = 1)
        self.delta_7 = timedelta(days = 7)

        #create file
        #head
        try:
            self.name = self.table.cell(0, 0).value
            with open(f'timetable.ics', 'w', encoding='UTF-8') as f:
                ical_begin_base = f'''
BEGIN:VCALENDAR
VERSION:2.0
CALSCALE:GREGORIAN
X-WR-CALNAME:{self.name}
X-APPLE-CALENDAR-COLOR:{self.color}
X-WR-TIMEZONE:Asia/Shanghai
BEGIN:VTIMEZONE
TZID:Asia/Shanghai
X-LIC-LOCATION:Asia/Shanghai
BEGIN:STANDARD
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
TZNAME:CST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE
'''
                f.write(ical_begin_base)
                f.close()
        except:
            print('file head write failure!')
            sys.exit()
        else:
            print('file head write success.')

        #body
        # name\n[teacher]\n[time][place]
        daytimestart = {3: '083000', 4: '103000', 5: '140000', 6: '160000', 7: '184500', 8: '204500'}
        daytimeend   = {3: '101500', 4: '121500', 5: '154500', 6: '174500', 7: '203000', 8: '223000'}

        for icol in range(1, 8):
            for irow in range(3, 9):
                celldata = self.table.cell(irow, icol).value

                #整合在一起的话，最后收尾须抽象成复数，且在循环中更新须在抽象的使用完毕之后
                #根据自然顺序
                left3 = -1
                right3 = -2 #最后有一个换行符
                left1 = celldata.find('[', left3 + 1)
                while left1 >= 0:
                    left1 = celldata.find('[', left3 + 1)
                    cname = celldata[right3 + 2: left1 - 1] # name\n 不包含换行符

                    right1 = celldata.find(']', right3 + 2)
                    tname = celldata[left1 + 1: right1] #如果是实验，则tname为第几节

                    left2 = celldata.find('[', left1 + 1)
                    right2 = celldata.find(']', right1 + 1)
                    timestr = celldata[left2 + 1: right2 - 1] + ',' #替换“周”字为“,”

                    left3 = celldata.find('[', left2 + 1)
                    right3 = celldata.find(']', right2 + 1)
                    place = celldata[left3 + 1: right3]

                    # timestr: 3-5,7-19周
                    #由逗号分隔时间刻或时间段
                    #由于“周”字最后一个逗号后面特殊处理（改为修改“周”字为“,”）

                    lastcomma = -1
                    comma = timestr.find(',', lastcomma + 1)
                    while comma >= 0:
                        smalltimestr = timestr[lastcomma + 1: comma] # 2 or 3-5
                        lastcomma = comma
                        #由-判断是否为时间段
                        dash = smalltimestr.find('-')
                        if dash == -1:
                            weeknum = int(smalltimestr)
                            curr_week = self.first_week_start + self.delta_7 * (weeknum - 1) + self.delta_1 * (icol - 1)
                            curr_week_str = curr_week.strftime('%Y%m%d')

                            try:
                                with open(f'timetable.ics', 'a', encoding='UTF-8') as f:
                                    ical_body_base = f'''
BEGIN:VEVENT
CREATED:{self.created}
DTSTAMP:{self.created}
SUMMARY:{cname}
DESCRIPTION:{tname}
LOCATION:{place}
TZID:Asia/Shanghai
SEQUENCE:0
UID:{uid()}
DTSTART;TZID=Asia/Shanghai:{curr_week_str}T{daytimestart[irow]}
DTEND;TZID=Asia/Shanghai:{curr_week_str}T{daytimeend[irow]}
X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:Class starts in fifteen minutes :-).
TRIGGER:-P0DT0H15M0S
X-WR-ALARMUID:{uid()}
UID:{uid()}
END:VALARM
END:VEVENT
'''
                                    f.write(ical_body_base)
                                    f.close()
                            except:
                                print('file body write failure!')
                                sys.exit()
                            else:
                                print('file body write success.')

                        else:
                            weekstart = int(smalltimestr[0: dash])

                            interval_2 = smalltimestr.find('单')
                            interval_another = smalltimestr.find('双')
                            if interval_2 < interval_another:
                                interval_2 = interval_another
                            if interval_2 == -1:
                                weekend = int(smalltimestr[dash + 1:])
                            else:
                                weekend = int(smalltimestr[dash + 1: interval_2])

                            curr_week = self.first_week_start + self.delta_7 * (weekstart - 1) + self.delta_1 * (icol - 1)
                            curr_week_str = curr_week.strftime('%Y%m%d')

                            curr_week_end = self.first_week_start + self.delta_7 * (weekend - 1) + self.delta_1 * (icol - 1)
                            curr_week_end_str = curr_week_end.strftime('%Y%m%d')

                            try:
                                with open(f'timetable.ics', 'a', encoding='UTF-8') as f:
                                    ical_body_base = f'''
BEGIN:VEVENT
CREATED:{self.created}
DTSTAMP:{self.created}
SUMMARY:{cname}
DESCRIPTION:{tname}
LOCATION:{place}
TZID:Asia/Shanghai
SEQUENCE:0
UID:{uid()}
RULE:FREQ=WEEKLY;UNTIL={curr_week_end_str}T{daytimeend[irow]}Z;INTERVAL=1
DTSTART;TZID=Asia/Shanghai:{curr_week_str}T{daytimestart[irow]}
DTEND;TZID=Asia/Shanghai:{curr_week_str}T{daytimeend[irow]}
X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC
BEGIN:VALARM
ACTION:DISPLAY
DESCRIPTION:Class starts in fifteen minutes :-).
TRIGGER:-P0DT0H15M0S
X-WR-ALARMUID:{uid()}
UID:{uid()}
END:VALARM
END:VEVENT
'''
                                    f.write(ical_body_base)
                                    f.close()
                            except:
                                print('file body write failure!')
                                sys.exit()
                            else:
                                print('file body write success.')

                        
                        comma = timestr.find(',', lastcomma + 1)

                    #最后一个时间段, 想了想把“周”字改为“,”，并到上面去
                    left1 = celldata.find('[', left3 + 1)





        #tail
        try:
            with open(f'timetable.ics', 'a', encoding='UTF-8') as f:
                ical_end_base = f'''
END:VCALENDAR
'''
                f.write(ical_end_base)
                f.close()
        except:
            print('file tail write failure!')
            sys.exit()
        else:
            print('file tail write success.')

        #success information
        print(f'indicateWeeks.ics created success')

    def main(self):
        if self.confirm_excel_random():
            sys.exit()
        self.set_attribute()
        self.excelTOical()




if __name__ == '__main__':
    my = excelReader()
    my.main()
