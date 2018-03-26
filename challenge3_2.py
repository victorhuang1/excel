# -*- coding: utf-8 -*-

from openxyl import load_workbook
import datetime
from openxyl import Workbook

wb = load_workbook('courses.xlsx')
sheet1 = wb['students']
sheet2 = wb['time']

def sheet3():
    sheet3 = wb.create_sheet(title='combine')
    sheet3.append(['创建时间','课程名称','学习人数','学习时间'])
    for stu in sheet1.values:
        if stu[2] != '学习人数':
            for time in sheet2.values:
                if time[1] == stu[1]:
                    sheet3.append(list(stu) + [time[2]])
    wb.save('courses.xlsx')

if __name__ == '__main__':
    combine()
