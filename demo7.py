from tkinter import *
import tkinter.filedialog
import tkinter.messagebox
from tkinter import ttk
import os

class Interface():
    def __init__(self):
        self.first_filenames = ""
        self.write_text =""

    def set(self):
        root = Tk()
        root.title("界面")
        root.geometry("690x300")
        my_notebook = ttk.Notebook(root)
        my_notebook.place(relx=0.062, rely=0.071, relwidth=0.887, relheight=0.876)
        root.columnconfigure(0, weight=1)
        page1 = ttk.Frame(my_notebook)
        page2 = ttk.Frame(my_notebook)

        def first(self):
            def first_select():
                self.first_filenames = tkinter.filedialog.askopenfilenames(filetypes=[("ts文件", "*.ts"), ("文本文件", "*.txt")])
                if len(self.first_filenames) != 0:
                    string_filename =""
                    for i in range(0, len(self.first_filenames)):
                        string_filename += str(self.first_filenames[i])+"\n"
                    label_fir_select.config(text="您选择的文件是：" + "\n" + string_filename)
                else:
                    label_fir_select.config(text="您没有选择任何文件！")

            def first_write():
                self.write_text = entry_fir_write.get()
                label_fir_write_after.config(text="您设置生成的xls文件名是：" + "\n" + self.write_text)

            def first_run():
                os.system("python xml_to_excel_many_gui.py {} {}".format(self.write_text, '-'.join(list(self.first_filenames))))
                if os.path.exists(self.write_text):
                    tkinter.messagebox.showinfo('提示', '成功')
                else:
                    tkinter.messagebox.showerror('错误', '失败')

            butten_fir_select = Button(page1, text="选择ts文件", command=first_select)
            butten_fir_select.place(x=10, y=10, height=30, width=200)
            label_fir_select = Label(page1, text="")
            label_fir_select.place(x=10, y=40)

            entry_fir_write = Entry(page1)
            entry_fir_write.place(x=400, y=10, height=30, width=150)
            label_fir_write = Label(page1, text="生成xls文件名：")
            label_fir_write.place(x=300, y=10, height=30, width=100)
            label_fir_write_after = Label(page1, text="")
            label_fir_write_after.place(x=400, y=40)
            butten_fir_write = Button(page1, text="确定", command=first_write)
            butten_fir_write.place(x=550, y=10, height=30, width=50)

            butten_fir_run = Button(page1, text="运行", command=first_run)
            butten_fir_run.place(x=250, y=180, height=30, width=100)
        first(self)
        my_notebook.add(page1, text='One')

        def second(self):
            print("你好")
        second(self)
        my_notebook.add(page2, text='Two')

        root.mainloop()

if __name__ == '__main__':
    show_interface = Interface()
    show_interface.set()