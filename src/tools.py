#!/usr/bin/env python3
# encoding: utf-8

__auchor__ = 'XiaoY'

from xlrd import xldate_as_tuple

"""
考勤工具模块
"""

class AttenceMetaData(object):
    """
    获取日期等信息
    """
    def __init__(self, sheet2):
        self.sheet = sheet2
        self.date_start, self.date_end = self._get_date_start_end()
        self.date_all = self._get_date_all()
        self.date_create = self._get_date_create()
        self.name = '考勤记录表'
        if self._check_date() == -1:
            return None
    
    def _get_date_start_end(self):
        """
        获取考勤开始记录的时间
        """
        return self.sheet.cell_value(1, 33).split('：')[1].split('～')
    
    def _get_date_create(self):
        """
        获取考勤创建时间
        """
        return self.sheet.cell_value(2, 33).split('：')[1].split(' ')[0]

    def _get_date_all(self):
        """
        获取考勤记录的所有日期
        """
        return self.sheet.col_values(0)[12:]

    def _check_date(self):
        """
        检查日期是否有错误
        """
        if self.date_all[-1].split(' ')[0] != self.date_end.split('-')[-1]:
            print('日期校验错误!\n请检查考勤日期或报告XiaoY。')
            return -1

class SheetAttenceParse(object):
    """
    解析每张sheet的考勤记录
    """
    
    staff_row_count = 14    # 每个员工所占的单元格列数
    staff_count = 3         # 每张sheet表的员工数
    
    def __init__(self, sheet, day_count):
        if sheet.name == '排班记录表' or sheet.name == '考勤汇总表':
            print('不需要解析前两张sheet表。')
            return None
        self.sheet = sheet
        self.day_count = day_count
    
    def get_sheet_attence(self):
        """
        获取sheet表的所有考勤
        """
        pass
    
    def _parse_staff_attence(self):
        """
        解析考勤信息
        """
        
        staff_attence_list = []
        staff_name_list = []
        
        offset = self.staff_row_count + 1
        for staff_index in range(self.staff_count):
            staff_attence = {}
            
            id = self.sheet.cell_value(4, 9+staff_index*offset)
            name = self.sheet.cell_value(3, 9+staff_index*offset)
            if id == '' or name == '':  # 若没有工号或名字，则进行跳过
                continue
            if name in staff_name_list: # 本意实现重复员工的剔除，但由于处理单页Sheet，没有获取全部员工，所以没有达到效果
                continue
            staff_name_list.append(name)
            staff_id = int(id)  # 工号
            staff_name = name   # 名字
            staff_department  = self.sheet.cell_value(3, 1+staff_index*offset)  # 部门

            staff_time_signed = []
            for day_index in range(self.day_count):
                # 解析签到的时间，思路为：
                # 获取每个人每天的全部签到存进times列表，若列表长度为0，则整天未签到，为1，则未签到一次，为2，则全部签到
                # 以字典形式保存，0对应签到，1对应签退，值为''时表示未签到
                times_tmp = []
                for work_mode in range(1, 13):
                    times_tmp.append(self.sheet.cell_value(12+day_index, work_mode+staff_index*offset))
                times = []
                for time in times_tmp:
                    if time == '':
                        continue
                    times.append(time)
                #print([xldate_as_tuple(x, 0) for x in times if isinstance(x, float)])
                if len(times) == 0:
                    staff_time_signed.append({0:'', 1:''})
                elif len(times) == 1:
                    #print(times[0])
                    if times[0] > 0.5:
                        staff_time_signed.append({0:'', 1:'{0:0>2}:{1:0>2}'.format(str(xldate_as_tuple(times[0], 0)[3]), str(xldate_as_tuple(times[0], 0)[4]))})
                    else:
                        staff_time_signed.append({0:'{0:0>2}:{1:0>2}'.format(str(xldate_as_tuple(times[0], 0)[3]), str(xldate_as_tuple(times[0], 0)[4])), 1:''})
                else:
                    staff_time_signed.append({0:'{0:0>2}:{1:0>2}'.format(str(xldate_as_tuple(min(times), 0)[3]), str(xldate_as_tuple(min(times), 0)[4])), 1:'{0:0>2}:{1:0>2}'.format(str(xldate_as_tuple(max(times), 0)[3]), str(xldate_as_tuple(max(times), 0)[4]))})
            staff_attence.update({'工号': staff_id, '姓名': staff_name, '部门': staff_department, '记录': staff_time_signed})
            staff_attence_list.append(staff_attence)
        return staff_attence_list
