#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xml.dom.minidom as xmldom  #通过minidom解析xml文件
import os
import xlwt
import sys

class GenerateExcel():
    def __init__(self):
        self.list_source_value = []
        self.j = 0
        self.row_first_ = 0   #第一行
        self.row_second_ = 1  #第二行
        self.row_third_ = 2   #第三行
        self.col_first_ = 0   #第一列
        self.col_second_ = 1  #第二列

    def XmlToExcelManyUnique(self, excel_path, xml_paths):
        # 创建一个Workbook对象，这就相当于创建了一个Excel文件
        book = xlwt.Workbook(encoding='utf-8', style_compression=0)
        sheet = book.add_sheet('test', cell_overwrite_ok=True)
        # 遍历可变参数，读索引读值
        for index, element in enumerate(xml_paths):
            xmlfilepath = element
            domobj = xmldom.parse(xmlfilepath)  # 得到文档对象
            elementobj = domobj.documentElement  # 得到元素对象
            elementobj_source = elementobj.getElementsByTagName("source")  # 获得source子标签,区分相同标签名
            elementobj_translation = elementobj.getElementsByTagName("translation")  # 获得translation子标签,区分相同标签名

            sheet.write(self.row_second_, self.col_first_, "source")
            sheet.write(self.row_second_, self.col_second_, "type")
            sheet.write(self.row_second_, index + 2, elementobj.getAttribute("language"))
            sheet.write(self.row_first_, index + 2, element)

            for i in range(len(elementobj_source)):
                if elementobj_source[i].firstChild.data not in self.list_source_value:  # 筛选出不重复的source的value
                    self.list_source_value.append(elementobj_source[i].firstChild.data)
                    for self.j in range (len(self.list_source_value)):
                        if index == 0:  # 从第三行开始，第一列写入source的value
                            sheet.write(self.j + self.row_third_,
                                        self.col_first_, self.list_source_value[self.j])
                    if index == 0:  # 从第三行开始，第二列写入translation的type
                        sheet.write(self.j + self.row_third_,
                                    self.col_second_, elementobj_translation[i].getAttribute("type"))
                    # 从第三行开始，从第三列开始的后面每列依次写入translation的value
                    if elementobj_translation[i].hasChildNodes():  # translation的value不为空
                        sheet.write(self.j + self.row_third_, index + 2,
                                    elementobj_translation[i].firstChild.data)
                    else:  # translation的value为空
                        sheet.write(self.j + self.row_third_, index + 2, "")  # 写入translation的value
            self.list_source_value.clear()

        if os.path.exists(excel_path):
            os.remove(excel_path)
        book.save(excel_path)

if __name__ == '__main__':
    xml_to_excel_many_unique = GenerateExcel()
    xml_to_excel_many_unique.XmlToExcelManyUnique(sys.argv[1], sys.argv[2].split('-'))
    # 第一个参数是输出的excel文件，名字自己起。后面参数都是输入的xml文件
    # xml_to_excel_many_unique.XmlToExcelManyUnique("xml_to_excel_many_unique.xls", "cutter_zh.ts", "cutter_zh.ts")