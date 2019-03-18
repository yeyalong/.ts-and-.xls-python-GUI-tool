#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import xml.dom.minidom as xmldom  #通过minidom解析xml文件
import os
import sys

class ExcelToXml():
    def __init__(self):
        self.col_first_ = 0   #第一列
        self.col_second_ = 1  #第二列
        self.list_translate_value_ = []
        self.language_col_number_ = 0
        self.row_second_ = 1  # 第二行
        self.list_source_value = []
        self.list_translate_type = []
        self.list_translate_value = []

    def ReadExcel(self, excel_path, language):
        book = xlrd.open_workbook(excel_path)
        table  = book.sheet_by_index(0)# 通过sheet索引获得sheet对象
        nrows = table.nrows    # 获取行总数
        ncols = table.ncols    # 获取列总数

        #遍历excel
        for nrow in range(0, nrows):  #遍历每一行
            col_first = table.cell(nrow, self.col_first_).value  #取第一列的值
            col_second = table.cell(nrow, self.col_second_).value  #取第二列的值
            self.list_source_value.append(col_first)
            self.list_translate_type.append(col_second)

            for i in range(0, ncols):
                row_second = table.cell(self.row_second_, i).value  # 取第二行的值
                if row_second == language:
                    self.language_col_number = i
            # 取language列，放list中
            for nrow in range(0, nrows):  # 遍历每一行
                supplier_col_language = table.cell(nrow, self.language_col_number).value  # 取language列的值
                self.list_translate_value_.append(supplier_col_language)

    def WriteXml(self,xml_path):
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
                    translation_value = domobj.createTextNode(self.list_translate_value_[j])
                    elementobj_translation[i].appendChild(translation_value)
                    elementobj_translation[i].childNodes[0].nodeValue = ""
                    elementobj_translation[i].nodeValue = self.list_translate_value_[j]
                    # 把excel中list_translate_type更新到xml中translate的type
                    elementobj_translation[i].setAttribute("type", self.list_translate_type[j])
                    if self.list_translate_type[j] == "":
                        elementobj_translation[i].removeAttribute("type")
                    break
        with open(xmlfilepath, 'w', encoding='utf-8') as fail_write_xml:
            domobj.writexml(fail_write_xml, encoding='utf-8')

if __name__ == '__main__':
    excel_to_xml = ExcelToXml()
    # 第一个参数是输入的excel文件，第二个参数是指定的语言
    excel_to_xml.ReadExcel(sys.argv[1], sys.argv[2])
    # 写入到的xml文件
    excel_to_xml.WriteXml(sys.argv[3])
    # excel_to_xml.ReadExcel("xml_to_excel_unique.xls", "zh_CN")
    # excel_to_xml.WriteXml("cutter_en.ts", "")