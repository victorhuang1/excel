# -*- coding: utf-8 -*-

from openxyl import load_workbook
import datetime
from openxyl import Workbook

wb = load_workbook('courses.xlsx')
sheet1 = wb['students']
sheet2 = wb['time']

def sheet3():
    sheet3 = wb.create_sheet(title='combine')
    sheet3.append(['����ʱ��','�γ�����','ѧϰ����','ѧϰʱ��'])
    for stu in sheet1.values:
        if stu[2] != 'ѧϰ����':
            for time in sheet2.values:
                if time[1] == stu[1]:
                    sheet3.append(list(stu) + [time[2]])
    wb.save('courses.xlsx')

if __name__ == '__main__':
    combine()
