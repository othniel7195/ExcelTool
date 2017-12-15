#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from openpyxl import load_workbook
import os
import sys
import time
import re
import json

__metaclass__ = type


class SheetData:
    """
    :type wb: :class:`openpyxl.workbook.Workbook`
    :type ws: `openpyxl.worksheet.Worksheet`
    """

    def __init__(self, workbook_path):
        """
        
        :type workbook_path: string
        """
        self.wb = load_workbook(workbook_path)

    def sheetForWorkbook(self, sheet_name):
        ws = self.wb[sheet_name.decode("utf-8")]
        return ws

    def iterate_worksheet(self, ws):

        """
        获取worksheet 所有数据
        :rtype: list
        """

        data = []

        for row_iterate in ws.iter_rows(None, ws.min_row, ws.max_row, ws.min_column, ws.max_column):
            every_row = []
            for cel in row_iterate:

                every_row.append((cel.coordinate, cel.value))

            data.append(every_row)

        return data

    def search_sheet_col(self, data, search_str):
        """
        获取查询的title行数据
        :type data: list
        :type search_str: str
        :rtype: tuple (row, col, 值)
        """
        for idx, every_row_data in enumerate(data):
            for num, value in enumerate(every_row_data):
                if value[1] and value[1] == search_str.decode("utf-8"):
                    return idx, num, value

    def get_every_row_of_search(self, search_tuple, data):
        """
        获取每一行的查询用的key
        :type search_tuple: tuple(1, 4, ["E2", "起保日期"])
        :type data: list
        :rtype: list
        """

        all_of_result = []
        for search_num in range(1, len(data) - 1):
            idx = search_tuple[0] + search_num
            all_of_result.append((idx, data[idx][search_tuple[1]]))

        return all_of_result

    def get_all_dir(self):
        """
        获取当前路径下的所有文件夹名字和创建时间
        :rtype:list 
        """
        folders = []
        for folder in os.listdir(sys.path[0]):
            if os.path.isdir(folder):
                file_info = []
                file_info.append(folder)
                file_info.append(self.get_dir_create_time(folder))
                folders.append(file_info)

        return folders

    def get_all_file(self, path):
        """
        获取文件夹下所有xlsx文件
        :type path: str
        :rtype:  list
        """
        files = []
        dir_path = "{0}/{1}".format(sys.path[0], path)
        # print "文件夹路径:{0}".format(dir_path)
        for file in os.listdir(dir_path):
            if file.endswith("xlsx"):
                files.append(file)

        return files

    def get_dir_create_time(self, path):
        time_struct = time.localtime(os.path.getctime(path))
        return time.strftime('%Y-%m-%d', time_struct)

    def get_date_string(self, dateStr):
        """
        匹配字符串中的时间字符串
        :type dateStr: str
        :rtype: str
        """
        return re.findall(r"\d{2,4}[-|.]\d+[-|.]\d+", dateStr)

    def compare_date_equal(self, date1, date2):

        """
        比较时间是否相等
        :type date1: str 在保日期
        :type date2: str 文件夹日期
        :rtype: bool
        """
        if len(date1) == 0:
            return False
        if len(date2) == 0:
            return False

        d1 = date1[0].replace(".", "-")
        d2 = date2[0].replace(".", "-")
        datel1 = d1.split("-")
        datel2 = d2.split("-")

        #假设 数组count = 2 意味着没有年份,如果是没有月份或者没有日期 就是文件夹命名有问题
        num1 = len(datel1)
        num2 = len(datel2)
        date_compare_result = False
        if num1 == 2:
            datel1.insert(0, time.localtime()[0])
        if num2 == 2:
            datel2.insert(0, time.localtime()[0])


        # print "比较年份"
        y1 = datel1[0]
        y2 = datel2[0]
        year_compare_result = False
        if len(y1) == len(y2) and int(y1) == int(y2):
            year_compare_result = True
        elif len(y1) != len(y2):
            y1 = y1[-2:]
            y2 = y2[-2:]
            # print "y1:{0} y2:{0}".format(y1, y2)
            if int(y1) == int(y2):
                year_compare_result = True

        # print "比较月份"
        m1 = datel1[1]
        m2 = datel2[1]
        month_compare_result = False
        if int(m1) == int(m2):
            month_compare_result = True

        # print "比较日期"
        d1 = datel1[2]
        d2 = datel2[2]
        day_compare_result = False
        #在保日期 = 是文件夹日期 +1
        d2 = int(d2) + 1
        if int(d1) == d2:
            day_compare_result = True

        if year_compare_result and month_compare_result and day_compare_result:
            date_compare_result = True

        return date_compare_result

    def get_dir_path_from_protect(self, iarg):
        if iarg[1][1]:
            # 获取文件夹名字
            for date_string in self.get_all_dir():
                date_result = self.compare_date_equal(self.get_date_string(iarg[1][1]),
                                                      self.get_date_string(date_string[0]))

                if date_result:
                    return date_string
                else:
                    date_result = self.compare_date_equal(self.get_date_string(iarg[1][1]),
                                                          self.get_date_string(date_string[1]))
                    if date_result:
                        return date_string

        return None

if __name__ == "__main__":
    basicExcel = SheetData("基础数据表单.xlsx")
    ws1 = basicExcel.sheetForWorkbook("基础数据表单")
    ws1 = basicExcel.iterate_worksheet(ws1)

    # 原始数据
    print u"原始查询数据:"
    for i in ws1:
        print json.dumps(i, encoding="UTF-8", ensure_ascii=False)

    # 查询起保日期
    print u"起保日:"
    t1 = basicExcel.search_sheet_col(ws1, "起保日期")
    print json.dumps(t1, encoding="UTF-8", ensure_ascii=False)

    # print u"每一行的起保数据:"
    # 获取第每一行行需要查询的数据
    ar1 = basicExcel.get_every_row_of_search(t1, ws1)
    # print json.dumps(ar1, encoding="UTF-8", ensure_ascii=False)

    # 通过日期查询文件夹
    # print u"日期:"
    for dd in ar1:
        print dd
        ds = basicExcel.get_dir_path_from_protect(dd)
        if ds:
            print "找到了文件夹======================="
            # 遍历文件夹下的文件
            modifyFiles = basicExcel.get_all_file(ds[0]) # 找到文件夹下所有文件

            for modifyFile in modifyFiles:

                path = "{0}/{1}".format(ds[0], modifyFile)
                modifyExcel = SheetData(path)
                add_sheet = modifyExcel.sheetForWorkbook("增加被保险人")
                add_sheet = modifyExcel.iterate_worksheet(add_sheet)

                print u"增加被保险人:"
                for add in add_sheet:
                    print json.dumps(add, encoding="UTF-8", ensure_ascii=False)

                del_sheet = modifyExcel.sheetForWorkbook("减少被保险人")
                del_sheet = modifyExcel.iterate_worksheet(del_sheet)

                print u"减少被保险人:"
                for dell in del_sheet:
                    print json.dumps(dell, encoding="UTF-8", ensure_ascii=False)

                person_data_change_sheet = modifyExcel.sheetForWorkbook("信息变更")
                person_data_change_sheet = modifyExcel.iterate_worksheet(person_data_change_sheet)

                print u"信息变更:"
                for pdchange in person_data_change_sheet:
                    print json.dumps(pdchange, encoding="UTF-8", ensure_ascii=False)


                insure_data_change_sheet = modifyExcel.sheetForWorkbook("保障变更")
                insure_data_change_sheet = modifyExcel.iterate_worksheet(insure_data_change_sheet)

                print u"保障变更:"
                for idchange in insure_data_change_sheet:
                    print json.dumps(idchange, encoding="UTF-8", ensure_ascii=False)


            print "=================================="





































