#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import xml.dom.minidom as xmldom  #通过minidom解析xml文件
import os
import shutil
import xlwt

class ExcelToXml():
    def __init__(self):
        self.col_first_ = 0   #第一列
        self.col_second_ = 1  #第二列
        self.list_translate_value_ = []
        self.language_col_number_ = 0
        self.row_first_ = 0  # 第一行
        self.list_source_value = []
        self.list_translate_type = []
        self.list_translate_value = []
        self.language_value = 0

    def ReadExcel(self, excel_path, sheet_name, language):
        book = xlrd.open_workbook(excel_path)
        table = book.sheet_by_name(sheet_name)# 通过sheet名字获得sheet对象
        nrows = table.nrows    # 获取行总数
        ncols = table.ncols    # 获取列总数
        self.language_value = language

        #遍历excel
        for nrow in range(0, nrows):  #遍历每一行
            col_first = table.cell(nrow, self.col_first_).value  #取第一列的值
            col_second = table.cell(nrow, self.col_second_).value  #取第二列的值
            self.list_source_value.append(col_first)
            self.list_translate_type.append(col_second)

            for i in range(0, ncols):
                row_first = table.cell(self.row_first_, i).value  # 取第一行的值
                if row_first == language:
                    self.language_col_number_ = i

            # 取language列，放list中
            for nrow in range(0, nrows):  # 遍历每一行
                supplier_col_language = table.cell(nrow, self.language_col_number_).value  # 取language列的值
                self.list_translate_value_.append(supplier_col_language)

    def WriteXml(self, xml_path):
        xmlfilepath = os.path.abspath(xml_path)
        domobj = xmldom.parse(xmlfilepath)  #得到文档对象
        elementobj = domobj.documentElement  #得到元素对象
        elementobj_source = elementobj.getElementsByTagName("source")  #获得source子标签,区分相同标签名
        elementobj_translation = elementobj.getElementsByTagName("translation")  #获得translation子标签,区分相同标签名

        # 遍历第一列，根据第一列的值到excel中寻找
        for i in range(len(elementobj_source)):
            for j in range(len(self.list_source_value)):
                if self.list_source_value[j] == elementobj_source[i].firstChild.data:
                    # 把excel中list_translate_value更新到xml中translate的value
                    if elementobj_translation[i].hasChildNodes() == False:
                        translation_value = domobj.createTextNode(self.list_translate_value_[j])
                        elementobj_translation[i].appendChild(translation_value)
                        elementobj_translation[i].childNodes[0].nodeValue = ''
                        elementobj_translation[i].childNodes[0].nodeValue = self.list_translate_value_[j]  # 刚开始
                    else:
                        translation_value = domobj.createTextNode(self.list_translate_value_[j])
                        elementobj_translation[i].appendChild(translation_value)
                        elementobj_translation[i].childNodes[0].nodeValue = ''
                        elementobj_translation[i].nodeValue = self.list_translate_value_[j]  #之后
                    # 把excel中list_translate_type更新到xml中translate的type
                    elementobj_translation[i].setAttribute("type", self.list_translate_type[j])
                    if self.list_translate_type[j] == "":
                        elementobj_translation[i].removeAttribute("type")
                    break
        with open(xmlfilepath, 'w', encoding='utf-8') as xml_write_xml:
            domobj.writexml(xml_write_xml, encoding='utf-8')
        self.list_source_value.clear()
        self.list_translate_type.clear()
        self.list_translate_value_.clear()

if __name__ == '__main__':
    excel_to_xml = ExcelToXml()
    xls_name = "f7000语言.xls"
    book = xlrd.open_workbook(xls_name)
    sheets = book.sheet_names()
    for index, element in enumerate(sheets):
        table = book.sheet_by_name(element)  # 通过sheet名字获得sheet对象
        ncols = table.ncols  # 获取列总数
        for i in range(2, ncols):
            language_name = table.cell(0, i).value  # 取第一行的值
            source_file = './source/' + language_name + '/' + element + "_" + language_name[0:2] + ".ts"
            source_dir = './source/' + language_name + '/'
            default_file = './default/' + element + "_en.ts"

            if os.path.exists(source_file) == False:
                os.makedirs(source_dir)
                shutil.copy(default_file, source_file)
                domobj_source = xmldom.parse(source_file)  # 得到文档对象
                elementobj_source = domobj_source.documentElement  # 得到元素对象
                elementobj_source.setAttribute("language", language_name)
                with open(source_file, 'w', encoding='utf-8') as source_file_write:
                    domobj_source.writexml(source_file_write, encoding='utf-8')

            translate_file = './translate/' + language_name + '/' + element + "_" + language_name[0:2] + ".ts"
            translate_dir = './translate/' + language_name + '/'
            if os.path.exists(translate_dir) == False:
                os.makedirs(translate_dir)
            shutil.copy(source_file, translate_file)
            excel_to_xml.ReadExcel(xls_name, element, language_name)
            excel_to_xml.WriteXml(translate_file)

