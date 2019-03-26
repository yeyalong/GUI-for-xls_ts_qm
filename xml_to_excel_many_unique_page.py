#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xml.dom.minidom as xmldom  #通过minidom解析xml文件
import os
import xlwt
import sys
import xlrd
from xlutils.copy import copy

class GenerateExcel():
    def __init__(self):
        self.list_source_value = []
        self.j = 0
        self.row_first_ = 0   #第一行
        self.row_second_ = 1  #第二行
        self.row_third_ = 2   #第三行
        self.col_first_ = 0   #第一列
        self.col_second_ = 1  #第二列
        self.sheet_index = -1


    def XmlToExcelManyUnique(self, excel_path,  sheet_name, *xml_paths):
        global sheet
        # 存在excel，更新
        if os.path.exists(excel_path):
            book_open = xlrd.open_workbook(excel_path)
            sheets = book_open.sheet_names()
            book = copy(book_open)
            # 若有这个表
            for index, element in enumerate(sheets):
                if element == sheet_name:
                    self.sheet_index = index
                    sheet = book.get_sheet(self.sheet_index)
            # 若没有这个表
            if self.sheet_index == -1:
                sheet = book.add_sheet(sheet_name, cell_overwrite_ok=True)
        # 创建一个Workbook对象，这就相当于创建了一个Excel文件
        else:
            book = xlwt.Workbook(encoding='utf-8', style_compression=0)
            sheet = book.add_sheet(sheet_name, cell_overwrite_ok=True)
        # 遍历可变参数，读索引读值
        for index, element in enumerate(xml_paths):
            xmlfilepath = os.path.abspath(element)
            domobj = xmldom.parse(xmlfilepath)  # 得到文档对象
            elementobj = domobj.documentElement  # 得到元素对象
            elementobj_source = elementobj.getElementsByTagName("source")  # 获得source子标签,区分相同标签名
            elementobj_translation = elementobj.getElementsByTagName("translation")  # 获得translation子标签,区分相同标签名

            sheet.write(xml_to_excel_many_unique.row_second_, xml_to_excel_many_unique.col_first_, "source")
            sheet.write(xml_to_excel_many_unique.row_second_, xml_to_excel_many_unique.col_second_, "type")
            sheet.write(xml_to_excel_many_unique.row_first_, index + 2, elementobj.getAttribute("language"))
            sheet.write(xml_to_excel_many_unique.row_second_, index + 2, element)

            for i in range(len(elementobj_source)):
                if elementobj_source[i].firstChild.data not in xml_to_excel_many_unique.list_source_value:  # 筛选出不重复的source的value
                    xml_to_excel_many_unique.list_source_value.append(elementobj_source[i].firstChild.data)
                    for xml_to_excel_many_unique.j in range (len(xml_to_excel_many_unique.list_source_value)):
                        if index == 0:  # 从第三行开始，第一列写入source的value
                            sheet.write(xml_to_excel_many_unique.j + xml_to_excel_many_unique.row_third_,
                                        xml_to_excel_many_unique.col_first_, xml_to_excel_many_unique.list_source_value[xml_to_excel_many_unique.j])
                    if index == 0:  # 从第三行开始，第二列写入translation的type
                        sheet.write(xml_to_excel_many_unique.j + xml_to_excel_many_unique.row_third_,
                                    xml_to_excel_many_unique.col_second_, elementobj_translation[i].getAttribute("type"))
                    # 从第三行开始，从第三列开始的后面每列依次写入translation的value
                    if elementobj_translation[i].hasChildNodes():  # translation的value不为空
                        sheet.write(xml_to_excel_many_unique.j + xml_to_excel_many_unique.row_third_, index + 2,
                                    elementobj_translation[i].firstChild.data)
                    else:  # translation的value为空
                        sheet.write(xml_to_excel_many_unique.j + xml_to_excel_many_unique.row_third_, index + 2, "")  # 写入translation的value
            xml_to_excel_many_unique.list_source_value.clear()

        # if os.path.exists(excel_path):
        #     os.remove(excel_path)
        book.save(excel_path)

    def FindAllTs(self, dirname):
        for maindir, subdir, file_name_list in os.walk(dirname):
            for filename in file_name_list:
                apath = os.path.join(maindir, filename)
                if 'cutter' in filename:
                    result_cutter_all.append(apath)
                    result_cutter.append(filename)
                if 'parts' in filename:
                    result_parts_all.append(apath)
                    result_parts.append(filename)
                if 'widget' in filename:
                    result_widget_all.append(apath)
                    result_widget.append(filename)

if __name__ == '__main__':
    xml_to_excel_many_unique = GenerateExcel()
    # xml_to_excel_many_unique.XmlToExcelManyUnique(sys.argv[1], *sys.argv[2:])
    result_cutter_all = []
    result_parts_all = []
    result_widget_all = []
    result_cutter = []
    result_parts = []
    result_widget = []
    xml_to_excel_many_unique.FindAllTs('./source')
    xls_name = "f7000语言.xls"
    sheet_name1 = "cutter"
    xml_to_excel_many_unique.XmlToExcelManyUnique(xls_name, sheet_name1, *result_cutter_all)
    sheet_name2 = "parts"
    xml_to_excel_many_unique.XmlToExcelManyUnique(xls_name, sheet_name2, *result_parts_all)
    sheet_name3 = "widget"
    xml_to_excel_many_unique.XmlToExcelManyUnique(xls_name, sheet_name3, *result_widget_all)