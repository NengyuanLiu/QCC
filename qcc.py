# !/user/bin/env Python3
# -*- coding:utf-8 -*-

import tkinter as tk
import tkinter.messagebox
from tkinter import filedialog, dialog
import os


class QCC(object):
    def __init__(self):
        self.window = tk.Tk()
        self.window.title('测试')
        self.window.geometry('500x500')

        self.excel_file_path = ''
        self.out_put_file_path = ''

        # 是否处理成功标志位
        self.is_process_sucess = False

        # 输入的excel文件路径
        self.tk_excel_file_path = tk.StringVar()

        # layout
        # 按钮1：选择excel文件，设置显示变量，显示excel文件路径
        self.label1 = tk.Label(self.window, text="输入Excel文件")
        self.label1.grid(row=0, column=0)
        self.en1 = tk.Entry(self.window, textvariable=self.tk_excel_file_path, width=30)
        self.en1.grid(row=0, column=1)
        self.bt1 = tk.Button(self.window, text='选择...', width=5, command=self.select_excel_file)
        self.bt1.grid(row=0, column=2)

        # 坐标输入框
        self.label2 = tk.Label(self.window, text="坐标与焊点规格")
        self.label2.grid(row=1, column=0)
        self.text2 = tk.Text(self.window, width=40)
        self.text2.grid(row=1, column=1)

        # 开始处理
        self.bt2 = tk.Button(self.window, text='开始处理', width=30, command=self.process_file)
        self.bt2.grid(row=2, columnspan=2)

        # 导出文件
        self.bt3 = tk.Button(self.window, text='导出文件', width=30, command=self.select_output_dir)
        self.bt3.grid(row=3, columnspan=2)

    # 选择excel文件
    def select_excel_file(self):
        self.excel_file_path = filedialog.askopenfilename(title=u'选择文件')
        self.tk_excel_file_path.set(self.excel_file_path)

    # 输出文件目录
    def select_output_dir(self):
        # 基本校验
        if self.is_process_sucess is False:
            tkinter.messagebox.showwarning(title="warning", message="未处理成功，请重新处理")
            return
        self.out_put_file_path = filedialog.asksaveasfilename(title=u'选择保存文件')

    # 处理函数
    def process_file(self):
        # 处理前获取最新的输入excel文件路径，支持手动输入路径
        self.excel_file_path = self.tk_excel_file_path.get()
        # 获取坐标信息
        text_context = self.text2.get(1.0, "end")
        print(text_context)

        # 基本校验
        if os.path.exists(self.excel_file_path) is False:  # or os.path.isdir(self.excel_file_path):# mac下该函数无法准确判断
            self.is_process_sucess = False
            tkinter.messagebox.showwarning(title="warning", message="请重新选择输入excel文件")
            return

        print('打开文件：', self.excel_file_path)
        print('输出文件：', self.out_put_file_path)
        self.is_process_sucess = True


if __name__ == '__main__':
    '''
    显示逻辑
    '''
    app = QCC()
    app.window.mainloop()  # 显示
