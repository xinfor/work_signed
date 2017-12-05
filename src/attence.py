# -*- coding: utf-8 -*-
"""
Created on Mon Oct 23 09:27:11 2017

@author: XiaoY
"""

import xlrd
import xlwt

# 引入工具模块中的工具类
from tools import AttenceMetaData, SheetAttenceParse


punch_in_out = xlrd.open_workbook('../考勤报表.xls')    # 读取Excel
attence_meta_data = AttenceMetaData(punch_in_out.sheets()[2])   # 考勤基本信息，主要是生成的Excel头部不变的部分

day_count = len(attence_meta_data.date_all) # 记录的天数
moon = attence_meta_data.date_start.split('-')[1] # 本月，用来生成本月报表命名

staff_attence_list = [] # 用来存储所有的考勤
for sheet in punch_in_out.sheets()[2:]: # 由于SheetAttenceParse仅实现了单页Sheet的考勤解析，所以需要循环进行获取全部
    sheet_attence_parse = SheetAttenceParse(sheet, day_count)
    staff_attence_list.extend(sheet_attence_parse._parse_staff_attence())
"""
!!!未实现!!!
为保证数据的准确度，在此处应进行日期校验，实现思路为：
将程序获取的日期与签到时间结合，与原文件进行比对。
不实现了，请进行手动检查。
"""

attence_wb = xlwt.Workbook(encoding='utf-8')    # 新建Excel
attence_sheet = attence_wb.add_sheet('Sheet1')  # 创建Sheet

# 标题
align_center = xlwt.Alignment()
align_center.horz = xlwt.Alignment.HORZ_CENTER
align_center.vert = xlwt.Alignment.VERT_CENTER

font_title = xlwt.Font()
font_title.name = '宋体'
font_title.bold = True

style_title = xlwt.XFStyle()
style_title.alignment = align_center
style_title.font = font_title

attence_sheet.write_merge(0, 1, 0, day_count-1, attence_meta_data.name, style_title)

# 考勤日期范围
align_left = xlwt.Alignment()
align_left.horz = xlwt.Alignment.HORZ_LEFT
align_left.vert = xlwt.Alignment.VERT_CENTER

font_date_range = xlwt.Font()
font_date_range.name = '宋体'

style_date_range = xlwt.XFStyle()
style_date_range.alignment = align_left
style_date_range.font = font_date_range

attence_sheet.write_merge(2, 2, 0, 5, '考勤日期：'+attence_meta_data.date_start+'~'+attence_meta_data.date_end, style_date_range)

# 创建日期
align_right = xlwt.Alignment()
align_right.horz = xlwt.Alignment.HORZ_RIGHT
align_right.vert = xlwt.Alignment.VERT_CENTER

style_date_create = xlwt.XFStyle()
style_date_create.alignment = align_right
style_date_create.font = font_date_range

attence_sheet.write_merge(2, 2, day_count-7, day_count-1, '创建日期：'+attence_meta_data.date_create, style_date_create)

# 日期
day_style = xlwt.easyxf('font: name 宋体, bold on; align: wrap on, vert centre, horiz center; pattern: pattern_fore_colour 5, pattern SOLID_PATTERN; borders: top thin, left thin, right thin;')
week_style = xlwt.easyxf('font: name 宋体, bold on; align: wrap on, vert centre, horiz center; pattern: pattern_fore_colour 5, pattern SOLID_PATTERN; borders: bottom thin, left thin, right thin;')
week_list = []
for date_day in attence_meta_data.date_all:
    day, week = date_day.split(' ')
    week_list.append(week)
    
    attence_sheet.write(3, int(day)-1, int(day), day_style)
    attence_sheet.write(4, int(day)-1, week, week_style)

# 考勤
staff_info_style = xlwt.easyxf('font: name 宋体, bold on; pattern: pattern_fore_colour 7, pattern SOLID_PATTERN; borders: top thin, bottom thin, left thin, right thin;')
forget_signed_style = xlwt.easyxf('font: name 宋体, bold on; align: wrap on, vert centre, horiz center; pattern: pattern_fore_colour 52, pattern SOLID_PATTERN; borders: top thin, bottom thin, left thin, right thin;')
unusual_signed_style = xlwt.easyxf('font: name 宋体, colour_index 2, bold on; align: wrap on, vert centre, horiz center; borders: top thin, bottom thin, left thin, right thin;')
normal_signed_style = xlwt.easyxf('font: name 宋体; align: wrap on, vert centre, horiz center; borders: top thin, bottom thin, left thin, right thin;')
for index, attence in enumerate(staff_attence_list):
    row_index = 3 * index
    
    attence_sheet.write_merge(5+row_index, 5+row_index, 0, 7, '工号：'+str(attence['工号']), staff_info_style)
    #attence_sheet.write(5+row_index, 1, attence['工号'])
    attence_sheet.write_merge(5+row_index, 5+row_index, 8, 15, '姓名：'+attence['姓名'], staff_info_style)
    attence_sheet.write_merge(5+row_index, 5+row_index, 16, day_count-1, '部门：'+attence['部门'], staff_info_style)
    
    for signed_index, signed_time in enumerate(attence['记录']):
        if signed_time[0] == '': # 未签到
            if week_list[signed_index] != '六' and week_list[signed_index] != '日': # 若是周末，则不进行处理
                attence_sheet.write(5+row_index+1, signed_index, '未打卡', forget_signed_style)
        else: # 已签到
            if signed_time[0] <= '09:10': # 正常签到
                attence_sheet.write(5+row_index+1, signed_index, signed_time[0], normal_signed_style)
            else: # 迟到
                attence_sheet.write(5+row_index+1, signed_index, signed_time[0], unusual_signed_style)
        if signed_time[1] == '': # 未签退
            if week_list[signed_index] != '六' and week_list[signed_index] != '日': # 若是周末，则不进行处理
                attence_sheet.write(5+row_index+2, signed_index, '未打卡', forget_signed_style)
        else: # 已签退
            if signed_time[1] >= '18:00': # 正常签退
                attence_sheet.write(5+row_index+2, signed_index, signed_time[1], normal_signed_style)
            else: # 早退
                attence_sheet.write(5+row_index+2, signed_index, signed_time[1], unusual_signed_style)

attence_wb.save('../'+moon+'月考勤.xls')
