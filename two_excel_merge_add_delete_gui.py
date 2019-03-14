#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import xlrd
import xlwt
from xlutils.copy import copy
import sys

class TwoExcelMerge():
    def __init__(self):
        self.col_first_ = 0   #第一列
        self.row_second_ = 1  # 第二行
        self.language_col_number = 0
        self.list_source_value_supplier = []
        self.list_translate_value_supplier = []
        self.list_source_value_user = []
        self.list_add_index = []
        self.list_delete_index = []
        self.list_first_col_after_delete_value = []
        self.nrows_supplier = 0
        self.nrows_user = 0
        self.ncols_user = 0
        self.table_supplier = xlrd.sheet.Sheet
        self.table_user = xlrd.sheet.Sheet
        self.select_language = ""
        self.new_language_number = 0

    def ReadSupplier(self, supplier_excel_path, language):
        book_supplier = xlrd.open_workbook(supplier_excel_path)
        self.table_supplier = book_supplier.sheet_by_index(0)  # 通过sheet索引获得sheet对象
        self.nrows_supplier = self.table_supplier.nrows  # 获取行总数
        ncols_supplier = self.table_supplier.ncols  # 获取列总数
        for i in range(0, ncols_supplier): #遍历列，寻找language列
            row_second = self.table_supplier.cell(self.row_second_, i).value  # 取第二行的值
            if row_second == language:
                self.language_col_number = i
        #取第一和language列，放list中
        for nrow in range(0, self.nrows_supplier):  # 遍历每一行
            supplier_col_first = self.table_supplier.cell(nrow, self.col_first_).value  # 取第一列的值
            supplier_col_language = self.table_supplier.cell(nrow, self.language_col_number).value  # 取language列的值
            self.list_source_value_supplier.append(supplier_col_first)
            self.list_translate_value_supplier.append(supplier_col_language)
        self.select_language = language

    def ReadUser(self, user_excel_path):
        book_user = xlrd.open_workbook(user_excel_path)
        self.table_user = book_user.sheet_by_index(0)  # 通过sheet索引获得sheet对象
        self.nrows_user = self.table_user.nrows  # 获取行总数
        self.ncols_user = self.table_user.ncols  # 获取列总数
        # 取第一列，放list中
        for nrow in range(0, self.nrows_user):  # 遍历每一行
            user_col_first = self.table_user.cell(nrow, self.col_first_).value  # 取第一列的值
            self.list_source_value_user.append(user_col_first)

    def WriteExcel(self, new_excel_path):
        # 创建一个Workbook对象，这就相当于创建了一个Excel文件
        book_new_excel = xlwt.Workbook(encoding='utf-8', style_compression=0)
        sheet_new_excel = book_new_excel.add_sheet('test', cell_overwrite_ok=True)

        #取在user中第一列删除的index
        for n in range(0, self.nrows_user):
            user_col_first = self.table_user.cell(n, 0).value
            if user_col_first not in self.list_source_value_supplier:
                self.list_delete_index.append(n)

        #生成删除后的第一列
        for a in range(len(self.list_delete_index)):  # 遍历index，a:需要删除的那行
            for i in range(0, self.nrows_user):
                if len(self.list_first_col_after_delete_value) <= (len(self.list_source_value_user) -len(self.list_delete_index)):
                    self.list_first_col_after_delete_value.append(self.table_user.cell(i, 0).value)
            self.list_first_col_after_delete_value.pop(self.list_delete_index[a])
        if self.list_delete_index ==[]:
            self.list_first_col_after_delete_value = self.list_source_value_user

        # 生成删除后的中间列
        list_after_delete_center = []
        for ncol in range(1, self.ncols_user):  # 遍历中间的几列
            for a in range(len(self.list_delete_index)):  # 遍历index，a:需要删除的那行
                for i in range(0, self.nrows_user):  # 遍历user的行
                    if len(list_after_delete_center) <= (len(self.list_source_value_user) - len(self.list_delete_index)):
                        list_after_delete_center.append(self.table_user.cell(i, ncol).value)
                list_after_delete_center.pop(self.list_delete_index[a])
            if self.list_delete_index ==[]:
                for i in range(0, self.nrows_user):  # 遍历user的行
                    list_after_delete_center.append(self.table_user.cell(i, ncol).value)
            for n in range(len(list_after_delete_center)):
                sheet_new_excel.write(n, ncol, list_after_delete_center[n])  #依次把中间的每列写入到表中
            list_after_delete_center.clear()

        # 写第一列，并取第一列相对删除后的增加的index
        for nrow in range(0, self.nrows_supplier):
            supplier_col_first = self.table_supplier.cell(nrow, 0).value
            sheet_new_excel.write(nrow, 0, supplier_col_first)
            if supplier_col_first not in self.list_first_col_after_delete_value:
                self.list_add_index.append(nrow)

        #写入
        if os.path.exists(new_excel_path):
            os.remove(new_excel_path)
        book_new_excel.save(new_excel_path)

        # 修改中间几列
        book_new_excel_old = xlrd.open_workbook(new_excel_path)
        book_new_excel_new = copy(book_new_excel_old)
        sheet_new_excel_old = book_new_excel_old.sheet_by_index(0)
        sheet_new_excel_new = book_new_excel_new.get_sheet(0)
        sheet_new_excel_old_ncols = sheet_new_excel_old.ncols  # 获取列总数

        # 在刚才中间几列删除后的基础上在增加添加的行
        list_center_add = []
        for ncol in range(1, self.ncols_user):  # 遍历中间几列
            for m in range(len(self.list_add_index)):  # 遍历index，m:需要插入的那行
                for i in range(0, len(self.list_first_col_after_delete_value)):  # 遍历删除后的list
                    if len(list_center_add) <= len(self.list_first_col_after_delete_value):
                        list_center_add.append(sheet_new_excel_old.cell(i, ncol).value)   ####
                list_center_add.insert(self.list_add_index[m],"")
            if self.list_add_index ==[]:
                for i in range(0, len(self.list_first_col_after_delete_value)):  # 遍历删除后的list
                    list_center_add.append(sheet_new_excel_old.cell(i, ncol).value)  ####
            for n in range(len(list_center_add)):
                sheet_new_excel_new.write(n, ncol, list_center_add[n])   # 把中间列写入表中
            list_center_add.clear()

        # 将language列更新原有language列
        for j in range(0, sheet_new_excel_old_ncols):  # 遍历列，寻找language列
            sheet_new_row_second = sheet_new_excel_old.cell(self.row_second_, j).value  # 取第二行的值
            if sheet_new_row_second == self.select_language:
                self.new_language_number = j
        for i in range(len(self.list_translate_value_supplier)):
            if self.new_language_number != 0:
                sheet_new_excel_new.write(i, self.new_language_number, self.list_translate_value_supplier[i])
            else:
                sheet_new_excel_new.write(i, sheet_new_excel_old_ncols, self.list_translate_value_supplier[i])

        book_new_excel_new.save(new_excel_path)

if __name__ == '__main__':
    two_excel_merge = TwoExcelMerge()
    two_excel_merge.ReadSupplier(sys.argv[1], sys.argv[2])
    two_excel_merge.ReadUser(sys.argv[3])
    two_excel_merge.WriteExcel(sys.argv[4])