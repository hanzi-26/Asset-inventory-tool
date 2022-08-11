
import random
import time
import tkinter as tk
import tkinter.font as tkFont
import tkinter.ttk
from tkinter.filedialog import (askopenfilename)
import tkinter.messagebox as messagebox
import pandas as pd
from windnd import windnd
import PySimpleGUI as sg
import xlrd
import xlwt


class App:
    filename = None
    file_message = None
    progress_bar = sg.ProgressBar

    def __init__(self, root):
        # 设置窗口标题
        root.title("资产盘点转换工具")

        # 设置窗口大小
        width = 600
        height = 380
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        # root.resizable(width=False, height=False)

        # 添加文本内,设置字体的前景色和背景色，和字体类型、大小
        text = tk.Label(root, text="1.选择或直接拖入要转换的Excel文件\n2.点击转换按钮", fg="black", font=('Times', 20, 'bold'))
        # 将文本内容放置在主窗口内
        text.pack()

        # 选择文件按钮
        button_select_file = tk.Button(root)
        button_select_file["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times', size=10)
        button_select_file["font"] = ft
        button_select_file["fg"] = "#000000"
        button_select_file["justify"] = "center"
        button_select_file["text"] = "选择文件"
        button_select_file.place(x=20, y=100, width=70, height=25)
        button_select_file["command"] = self.choose_file

        # 文件拖拽
        windnd.hook_dropfiles(root, func=self.dragged_files)

        # 选择文件结果
        file_message = tk.Label(root)
        file_message.pack(side="left", fill="both")
        ft = tkFont.Font(family='Times', size=10)
        file_message["font"] = ft
        file_message["bg"] = "#e0e0e0"
        file_message["fg"] = "#777777"
        file_message["text"] = "文件未导入"
        file_message["justify"] = "center"
        file_message.place(x=130, y=100, bordermode="outside", width=430, height=120)
        self.file_message = file_message

        # 转换按钮
        button_convert = tk.Button(root, command=self.show)
        button_convert["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times', size=10)
        button_convert["font"] = ft
        button_convert["fg"] = "#000000"
        button_convert["justify"] = "center"
        button_convert["text"] = "转换"
        button_convert.place(x=20, y=180, width=70, height=25)
        button_convert["command"] = self.convert

        # 清空按钮
        button_clear = tk.Button(root)
        button_clear["bg"] = '#f0f0f0'
        ft = tkFont.Font(family='Times', size = 10)
        button_clear["font"] = ft
        button_clear["fg"] = "#000000"
        button_clear["justify"] = "center"
        button_clear["text"] = "清空"
        button_clear.place(x=20, y=260, width=70, height=25)
        button_clear["command"] = self.clear

        # 进度条
        self.progress_bar = tkinter.ttk.Progressbar(root)
        self.progress_bar.place(x=130, y=260, width=430, height=20)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 100

    def show(self):
        for i in range(100):
            time.sleep(random.random()*0.02 + 0.001)
            self.progress_bar['value'] += 1
            root.update()

    def clear(self):
        self.filename = None
        self.file_message["text"] = "文件未导入"
        self.progress_bar["value"] = 0
        self.update_ui()

    def dragged_files(self, files):
        for items in files:
            if items[-4:] in [b'.xls', b'.xlsx']:
                msg = '\n'.join(i.decode('gbk') for i in files)
                self.filename = msg
                self.update_ui()

    def choose_file(self):
        filename = askopenfilename(title="请选择输入的Excel文件", filetypes=[('Excel文件','*.xls')])
        self.filename = filename
        self.update_ui()

    def update_ui(self):
        self.file_message["text"] = self.filename

    def convert(self):
        if not self.filename:
            messagebox.showinfo("无法转换", "未选择文件")
            return
        if '\n' in self.filename:
            a = []
            b = self.filename
            a.append(b.split('\n'))
            for i in range(len(a[0])):
                if not self.process(a[0][i]):
                    return
        else:
            if not self.process(self.filename):
                return

        self.progress_bar["value"] = 100
        messagebox.showinfo("转换完成", "已覆盖原文件")
        return

    def process(self, file):
        try:
            df = pd.read_excel(file, sheet_name=0)  # sheet索引号从0开始

            values = df.values[1:]

            new_data = {"卡片编码": [], "epc编码": [], "资产名称": [], "规格型号": [], "使用部门": [], "原值": [], "使用日期": []}
            for value in values:
                new_data["卡片编码"].append(value[2])
                new_data["epc编码"].append(value[3].ljust(24, 'F'))
                new_data["资产名称"].append(value[5])
                new_data["规格型号"].append(value[6])
                new_data["使用部门"].append(value[11])
                new_data["原值"].append(value[7])
                new_data["使用日期"].append(value[14])

            writer = pd.ExcelWriter(file)
            new_data_pd = pd.DataFrame(new_data)
            new_data_pd.to_excel(writer, index=False, sheet_name="打印机格式")
            worksheet = writer.sheets["打印机格式"]
            worksheet.col(0).width = 8000
            worksheet.col(1).width = 8000
            worksheet.col(2).width = 8000
            worksheet.col(3).width = 8000
            worksheet.col(4).width = 8000
            worksheet.col(5).width = 8000
            worksheet.col(6).width = 8000

            writer.save()
            return True
        except Exception as ex:
            print(ex)
            messagebox.showerror("无法转换", "内部处理错误")
            return False


if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap("a.ico")
    app = App(root)
    root.mainloop()
