from tkinter import *
import tkinter.filedialog
import tkinter.messagebox
from tkinter import ttk
import os
import xlrd
import xml_to_excel_many_unique_gui
import excel_to_xml_gui
import two_excel_merge_add_delete_gui

class Interface():
    def __init__(self):
        self.first_filenames = ""
        self.first_write_text =""
        self.label_fir_write_after_text = ""
        self.col_second_ = 1  # 第二列
        self.row_second_ = 1  # 第二行
        self.language_col_number_ = 0
        self.second_xls_filename = ""
        self.second_list_language_value_ = []
        self.label_sec_select_xls_after_text = ""
        self.second_language_chosen = ""
        self.label_sec_select_lan_after_text = ""
        self.third_fir_xls_filename = ""
        self.label_thi_select_fir_xls_after_text = ""
        self.third_sec_xls_filename = ""
        self.label_thi_select_sec_xls_after_text = ""
        self.third_write_text = ""
        self.label_fir_write_after_text = ""
        self.third_list_language_value_ = []
        self.third_language_chosen = ""
        self.label_thi_select_lan_after_text = ""
        self.label_thi_write_after_text = ""

    def design_gui(self):
        root = Tk()
        root.title("界面")
        root.geometry("690x300")
        my_notebook = ttk.Notebook(root)
        my_notebook.place(relx=0.022, rely=0.062, relwidth=0.956, relheight=0.876)
        root.columnconfigure(0, weight=1)
        page1 = ttk.Frame(my_notebook)
        page2 = ttk.Frame(my_notebook)
        page3 = ttk.Frame(my_notebook)

        def first(self):
            def first_select():
                self.first_filenames = tkinter.filedialog.askopenfilenames(filetypes=[("ts文件", "*.ts")])
                if len(self.first_filenames) != 0:
                    first_string_filename =""
                    for i in range(0, len(self.first_filenames)):
                        first_string_filename += str(self.first_filenames[i])+"\n"
                    label_fir_select.config(text="您选择的ts文件是：" + "\n" + first_string_filename)
                else:
                    label_fir_select.config(text="您没有选择任何文件！")

            def first_write():
                self.first_write_text = entry_fir_write.get()
                if (len(self.first_write_text) > 4) and (self.first_write_text[-4:] == '.xls'):
                    label_fir_write_after.config(text="您设置生成的xls文件名是：" + "\n" + self.first_write_text)
                    if os.path.exists(self.first_write_text):
                        label_fir_write_after.config(text="您设置生成的xls文件目录下已存在，" + "\n" + "继续执行将覆盖！")
                else:
                    label_fir_write_after.config(text="您设置生成的xls文件名格式错误！")
                self.label_fir_write_after_text = label_fir_write_after.cget("text")

            def first_run():
                if (len(self.first_filenames) != 0) and (self.label_fir_write_after_text !='您设置生成的xls文件名格式错误！') and (
                        self.label_fir_write_after_text !=''):
                    ge_xls = xml_to_excel_many_unique_gui.GenerateExcel()
                    ge_xls.XmlToExcelManyUnique(self.first_write_text, self.first_filenames)
                    if os.path.exists(self.first_write_text):
                        tkinter.messagebox.showinfo('提示', '成功')
                else:
                    tkinter.messagebox.showerror('错误', '失败')
            # 第一页"选择ts文件"按钮
            butten_fir_select = Button(page1, text="选择ts文件", command=first_select)
            butten_fir_select.place(x=10, y=10, height=30, width=200)
            label_fir_select = Label(page1, text="", wraplength=200, justify='left')
            label_fir_select.place(x=10, y=40)
            # 第一页右边按钮
            entry_fir_write = Entry(page1)
            entry_fir_write.place(x=445, y=10, height=30, width=150)
            label_fir_write = Label(page1, text="生成xls文件名：")
            label_fir_write.place(x=345, y=10, height=30, width=100)
            label_fir_write_after = Label(page1, text="")
            label_fir_write_after.place(x=445, y=40)
            butten_fir_write = Button(page1, text="确认", command=first_write)
            butten_fir_write.place(x=595, y=10, height=30, width=50)
            # 第一页"运行"按钮
            butten_fir_run = Button(page1, text="运行", command=first_run)
            butten_fir_run.place(x=280, y=180, height=30, width=100)
        first(self)
        my_notebook.add(page1, text='ts文件生成xls文件')

        def second(self):
            def second_select_xls():
                self.second_xls_filename = tkinter.filedialog.askopenfilename(filetypes=[("xls文件", "*.xls")])
                if len(self.second_xls_filename) != 0:
                    label_sec_select_xls.config(text="您选择的xls文件是：" + "\n" + self.second_xls_filename)
                else:
                    label_sec_select_xls.config(text="您没有选择任何文件！")
                self.label_sec_select_xls_after_text = label_sec_select_xls.cget("text")

            def second_select_lan(*args):
                self.second_language_chosen = second_languageChosen.get()
                if (self.label_sec_select_xls_after_text == "") or (self.label_sec_select_xls_after_text == "您没有选择任何文件！"):
                    label_sec_select_lan.config(text="请先选择xls文件！")
                else:
                    second_book_xls = xlrd.open_workbook(self.second_xls_filename)
                    second_table_xls = second_book_xls.sheet_by_index(0)  # 通过sheet索引获得sheet对象
                    second_ncols_xls = second_table_xls.ncols  # 获取列总数
                    for i in range(0, second_ncols_xls):
                        second_row_second = second_table_xls.cell(self.row_second_, i).value  # 取第二行的值
                        self.second_list_language_value_.append(second_row_second)
                    if self.second_language_chosen in self.second_list_language_value_:
                        label_sec_select_lan.config(text="您选择的翻译语言是：" + "\n" + self.second_language_chosen)
                    else:
                        label_sec_select_lan.config(text="该xls文件中没有此语言！")
                self.label_sec_select_lan_after_text = label_sec_select_lan.cget("text")

            def second_select_ts():
                self.second_ts_filename = tkinter.filedialog.askopenfilename(filetypes=[("ts文件", "*.ts")])
                if len(self.second_ts_filename) != 0:
                    label_sec_select_ts.config(text="您选择的ts文件是：" + "\n" + self.second_ts_filename)
                else:
                    label_sec_select_ts.config(text="您没有选择任何ts文件！")

            def second_run():
                if (len(self.second_xls_filename) != 0) and (len(self.second_language_chosen) != 0) and (
                        self.label_sec_select_lan_after_text !='该xls文件中没有此语言！') and(len(self.second_ts_filename) != 0):
                    ge_ts = excel_to_xml_gui.ExcelToXml()
                    ge_ts.ReadExcel(self.second_xls_filename, self.second_language_chosen)
                    ge_ts.WriteXml(self.second_ts_filename)
                    tkinter.messagebox.showinfo('提示', '成功')
                else:
                    tkinter.messagebox.showerror('错误', '失败')
            # 第二页"选择xls文件"按钮
            butten_sec_select_xls = Button(page2, text="选择xls文件", command=second_select_xls)
            butten_sec_select_xls.place(x=10, y=10, height=30, width=180)
            label_sec_select_xls = Label(page2, text="", wraplength=180, justify='left')
            label_sec_select_xls.place(x=10, y=40, width=180)
            # 第二页"选择翻译语言："按钮
            label_sec_lan = Label(page2, text="选择翻译语言：")
            label_sec_lan.place(x=235, y=10, height=30, width=100)
            second_language = StringVar()
            second_languageChosen = ttk.Combobox(page2, width=5, textvariable=second_language, state='readonly')
            second_languageChosen['values'] = ('zh_CH', 'en_US', 'en_EN', 'jan_JAN', 'rus_RUS')  # 设置下拉列表的值
            second_languageChosen.place(x=335, y=10, height=30, width=80)
            second_languageChosen.bind("<<ComboboxSelected>>", second_select_lan)
            label_sec_select_lan = Label(page2, text="")
            label_sec_select_lan.place(x=265, y=40)
            # 第二页"选择ts文件"按钮
            butten_sec_select_ts = Button(page2, text="选择ts文件", command=second_select_ts)
            butten_sec_select_ts.place(x=466, y=10, height=30, width=180)
            label_sec_select_ts = Label(page2, text="", wraplength=180, justify='left')
            label_sec_select_ts.place(x=466, y=40)
            # 第二页"运行"按钮
            butten_sec_run = Button(page2, text="运行", command=second_run)
            butten_sec_run.place(x=280, y=180, height=30, width=100)
        second(self)
        my_notebook.add(page2, text='xls文件更新ts文件内容')

        def third(self):
            def third_select_fir_xls():
                self.third_fir_xls_filename = tkinter.filedialog.askopenfilename(filetypes=[("xls文件", "*.xls")])
                if len(self.third_fir_xls_filename) != 0:
                    label_thi_select_fir_xls.config(text="您选择的xls文件是：" + "\n" + self.third_fir_xls_filename)
                else:
                    label_thi_select_fir_xls.config(text="您没有选择任何文件！")
                if self.third_fir_xls_filename == self.third_sec_xls_filename:
                    label_thi_select_fir_xls.config(text="您选择的原有xls文件与新增xls文件相同，请重新选择！")
                self.label_thi_select_fir_xls_after_text = label_thi_select_fir_xls.cget("text")

            def third_select_sec_xls():
                self.third_sec_xls_filename = tkinter.filedialog.askopenfilename(filetypes=[("xls文件", "*.xls")])
                if len(self.third_sec_xls_filename) != 0:
                    label_thi_select_sec_xls.config(text="您选择的xls文件是：" + "\n" + self.third_sec_xls_filename)
                else:
                    label_thi_select_sec_xls.config(text="您没有选择任何文件！")
                if self.third_fir_xls_filename == self.third_sec_xls_filename:
                    label_thi_select_sec_xls.config(text="您选择的新增xls文件与原有xls文件相同，请重新选择！")
                self.label_thi_select_sec_xls_after_text = label_thi_select_sec_xls.cget("text")

            def third_select_lan(*args):
                self.third_language_chosen = third_languageChosen.get()
                if (self.label_thi_select_sec_xls_after_text == "") or (self.label_thi_select_sec_xls_after_text == "您没有选择任何文件！"):
                    label_thi_select_lan.config(text="请先选择新增xls文件！")
                else:
                    third_book_xls = xlrd.open_workbook(self.third_sec_xls_filename)
                    third_table_xls = third_book_xls.sheet_by_index(0)  # 通过sheet索引获得sheet对象
                    third_ncols_xls = third_table_xls.ncols  # 获取列总数
                    for i in range(0, third_ncols_xls):
                        third_row_second = third_table_xls.cell(self.row_second_, i).value  # 取第二行的值
                        self.third_list_language_value_.append(third_row_second)
                    if self.third_language_chosen in self.third_list_language_value_:
                        label_thi_select_lan.config(text="您选择的新增语言是：" + "\n" + self.third_language_chosen)
                    else:
                        label_thi_select_lan.config(text="该xls文件中没有此语言！")
                self.label_thi_select_lan_after_text = label_thi_select_lan.cget("text")

            def third_write():
                self.third_write_text = entry_thi_write.get()
                if (len(self.third_write_text) > 4) and (self.third_write_text[-4:] == '.xls'):
                    label_thi_write_after.config(text="您设置生成的xls文件名是：" + "\n" + self.third_write_text)
                    if os.path.exists(self.third_write_text):
                        label_thi_write_after.config(text="您设置生成的xls文件目录下已存在，" + "\n" + "继续执行将覆盖！")
                else:
                    label_thi_write_after.config(text="您设置生成的xls文件名格式错误！")
                self.label_thi_write_after_text = label_thi_write_after.cget("text")

            def third_run():
                if (len(self.third_fir_xls_filename) != 0) and (len(self.third_sec_xls_filename) != 0)and (
                        len(self.third_language_chosen) != 0) and (
                        self.label_thi_select_lan_after_text != '该xls文件中没有此语言！') and (
                        self.label_thi_write_after_text !='您设置生成的xls文件名格式错误！') and (
                        self.label_thi_write_after_text !=''):
                    merge_xls = two_excel_merge_add_delete_gui.TwoExcelMerge()
                    merge_xls.ReadSupplier(self.third_sec_xls_filename, self.third_language_chosen)
                    merge_xls.ReadUser(self.third_fir_xls_filename)
                    merge_xls.WriteExcel(self.third_write_text)
                    tkinter.messagebox.showinfo('提示', '成功')
                else:
                    tkinter.messagebox.showerror('错误', '失败')

            # 第三页"选择原有xls文件"按钮
            butten_thi_select_fir_xls = Button(page3, text="选择原有xls文件", command=third_select_fir_xls)
            butten_thi_select_fir_xls.place(x=10, y=10, height=30, width=180)
            label_thi_select_fir_xls = Label(page3, text="", wraplength=180, justify='left')
            label_thi_select_fir_xls.place(x=10, y=40)
            # 第三页"选择新增xls文件"按钮
            butten_thi_select_sec_xls = Button(page3, text="选择新增xls文件", command=third_select_sec_xls)
            butten_thi_select_sec_xls.place(x=280, y=10, height=30, width=180)
            label_thi_select_sec_xls = Label(page3, text="", wraplength=180, justify='left')
            label_thi_select_sec_xls.place(x=280, y=40)
            # 第二页"选择翻译语言："按钮
            label_thi_lan = Label(page3, text="选择新增语言：")
            label_thi_lan.place(x=465, y=10, height=30, width=100)
            third_language = StringVar()
            third_languageChosen = ttk.Combobox(page3, width=5, textvariable=third_language, state='readonly')
            third_languageChosen['values'] = ('zh_CH', 'en_US', 'en_EN', 'jan_JAN', 'rus_RUS')  # 设置下拉列表的值
            third_languageChosen.place(x=565, y=10, height=30, width=80)
            third_languageChosen.bind("<<ComboboxSelected>>", third_select_lan)
            label_thi_select_lan = Label(page3, text="")
            label_thi_select_lan.place(x=500, y=40)
            # 第三页生成文件按钮
            entry_thi_write = Entry(page3)
            entry_thi_write.place(x=265, y=120, height=30, width=150)
            label_thi_write = Label(page3, text="生成xls文件名：")
            label_thi_write.place(x=165, y=120, height=30, width=100)
            label_thi_write_after = Label(page3, text="")
            label_thi_write_after.place(x=265, y=150)
            butten_thi_write = Button(page3, text="确认", command=third_write)
            butten_thi_write.place(x=415, y=120, height=30, width=50)
            # 第三页"运行"按钮
            butten_thi_run = Button(page3, text="运行", command=third_run)
            butten_thi_run.place(x=280, y=200, height=30, width=100)

        third(self)
        my_notebook.add(page3, text='在一个xls文件中增加另一个xls文件中的一种语言')

        root.mainloop()

if __name__ == '__main__':
    show_interface = Interface()
    show_interface.design_gui()