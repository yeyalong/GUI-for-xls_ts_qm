from tkinter import *
import tkinter.messagebox
from tkinter import ttk
import os

class Interface():
    def design_gui(self):
        root = Tk()
        root.title("界面")
        root.geometry("590x130")
        my_notebook = ttk.Notebook(root)
        my_notebook.place(relx=0.022, rely=0.062, relwidth=0.956, relheight=0.876)
        root.columnconfigure(0, weight=1)

        def first():
            a = os.system("python xml_to_excel_many_unique_page.py")
            if a == 0:
                tkinter.messagebox.showinfo('提示', '生成xls文件成功')
            else:
                tkinter.messagebox.showerror('错误', '生成xls文件失败')

        def second():
            b = os.system("python excel_to_xml_page.py")
            if b == 0:
                tkinter.messagebox.showinfo('提示', '更新ts文件成功')
            else:
                tkinter.messagebox.showerror('错误', '更新ts文件失败')

        def third():
            c = os.system("python ts_qm.py")
            if c == 0:
                tkinter.messagebox.showinfo('提示', '生成qm文件成功')
            else:
                tkinter.messagebox.showerror('错误', '生成qm文件失败')
        # "生成xls文件"按钮
        butten_first = Button(root, text="生成xls文件", command=first)
        butten_first.place(x=30, y=50, height=30, width=130)
        # "更新ts文件"按钮
        butten_second = Button(root, text="更新ts文件", command=second)
        butten_second.place(x=230, y=50, height=30, width=130)
        # "生成qm文件"按钮
        butten_third = Button(root, text="生成qm文件", command=third)
        butten_third.place(x=430, y=50, height=30, width=130)

        root.mainloop()

if __name__ == '__main__':
    show_interface = Interface()
    show_interface.design_gui()