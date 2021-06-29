# coding: utf-8
# Author: Alice

import sys
from uuid import uuid4 as uid
from datetime import datetime, timedelta

# 该ICS文件需要的几个参数@param
# 创建时间 created
# 第一周第一天的日期YYYYMMDD first_week_start_str
# 总周数 weeks


class indicateWeeks:
    def __init__(self):
        self.first_week_start_str = ''
        self.weeks = 0
        self.created = '19700101T000000Z'

        self.color = '#A8A8A8' #浅灰色
        self.name = '这是第几周?'

    # 获取三个变量
    def set_attribute(self):
        self.first_week_start_str = input('请输入第一周第一天的日期(YYYYMMDD): ')
        self.weeks = input('请输入总周数: ')
        self.created = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
        self.name = f'''这是第几周? from {self.first_week_start_str} by {self.weeks}'''

    def main_process(self):
        # 将开始日期转换为标准时间元组
        first_week_start = datetime.strptime(self.first_week_start_str, '%Y%m%d')
        #获取时间变化量
        delta_7 = timedelta(days = 7)
        # delta_6 = timedelta(days = 6)

        #create file
        #head
        try:
            with open(f'indicateWeeks.ics', 'w', encoding='UTF-8') as f:
                ical_begin_base = f'''
BEGIN:VCALENDAR
VERSION:2.0
X-WR-CALNAME:{self.name}
X-APPLE-CALENDAR-COLOR:{self.color}
'''
                f.write(ical_begin_base)
                f.close()
        except:
            print('file head write failure!')
            sys.exit()
        else:
            print('file head write success.')

        #body
        for i in range(int(self.weeks)):
            curr_week = i + 1
            curr_week_start = first_week_start + i * delta_7
            curr_week_end = curr_week_start + delta_7

            start_date = curr_week_start.strftime('%Y%m%d')
            end_date = curr_week_end.strftime('%Y%m%d')

            ical_base = f'''
BEGIN:VEVENT
CREATED:{self.created}
DTSTAMP:{self.created}
TZID:Asia/Shanghai
SEQUENCE:0
SUMMARY:第 {curr_week} 周
DTSTART;VALUE=DATE:{start_date}
DTEND;VALUE=DATE:{end_date}
UID:{uid()}
END:VEVENT
'''
            try:
                with open(f'indicateWeeks.ics', 'a', encoding='UTF-8') as f:
                    f.write(ical_base)
                    f.close()
            except:
                print('file body write failure')
            else:
                print(f'the {i+1}th week file body write success')

        #tail
        try:
            with open(f'indicateWeeks.ics', 'a', encoding='UTF-8') as f:
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

        
if __name__ == '__main__':
    my = indicateWeeks()
    my.set_attribute()
    my.main_process()