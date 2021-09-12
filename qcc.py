# !/user/bin/env Python3
# -*- coding:utf-8 -*-

import tkinter as tk
import tkinter.messagebox
from tkinter import filedialog, dialog
import os


class QCC(object):
    def __init__(self):
        # 坐标进制映射关系
        self.alpha2number = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7, 'H': 8, 'I': 9, 'J': 10, 'K': 11,
                             'L': 12, 'M': 13, 'N': 14, 'O': 15, 'P': 16, 'Q': 17, 'R': 18, 'S': 19, 'T': 20, 'U': 21,
                             'V': 22, 'W': 23, 'X': 24, 'Y': 25, 'Z': 26}
        self.number2alpha = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K',
                             12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U',
                             22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z'}

        # GUI初始化
        self.window = tk.Tk()
        self.window.title('QCC')
        self.window.geometry('500x500')

        self.excel_file_path = ''
        self.out_put_file_path = ''
        self.default_interval = 0
        self.input_raw_interval = ''

        # 是否处理成功标志位
        self.is_process_sucess = False

        # 输入的excel文件路径
        self.tk_excel_file_path = tk.StringVar()

        # layout
        self.label1 = tk.Label(self.window, text="默认间距")
        self.label1.grid(row=0, column=0)
        self.en1 = tk.Entry(self.window, width=30)
        self.en1.grid(row=0, column=1)

        # 按钮1：选择excel文件，设置显示变量，显示excel文件路径
        self.label2 = tk.Label(self.window, text="输入Excel文件")
        self.label2.grid(row=1, column=0)
        self.en2 = tk.Entry(self.window, textvariable=self.tk_excel_file_path, width=30)
        self.en2.grid(row=1, column=1)
        self.bt2 = tk.Button(self.window, text='选择...', width=5, command=self.select_excel_file)
        self.bt2.grid(row=1, column=2)

        # 坐标输入框
        self.label3 = tk.Label(self.window, text="坐标与焊点规格")
        self.label3.grid(row=2, column=0)
        self.text3 = tk.Text(self.window, width=40)
        self.text3.grid(row=2, column=1)

        # 开始处理
        self.bt3 = tk.Button(self.window, text='开始处理', width=30, command=self.process_file)
        self.bt3.grid(row=3, columnspan=2)

        # 导出文件
        self.bt4 = tk.Button(self.window, text='导出文件', width=30, command=self.select_output_dir)
        self.bt4.grid(row=4, columnspan=2)

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

    def get_user_input(self):
        # 处理前获取最新的输入excel文件路径，这样可以支持手动输入路径
        self.excel_file_path = self.tk_excel_file_path.get()
        # 打印默认间距信息
        self.default_interval = self.en1.get()
        # 获取坐标信息
        self.input_raw_interval = self.text3.get(1.0, "end")

    def alpha_to_cartesian(self, alpha_str) -> int:
        # eg. 'AA' -->  27
        scale = len(self.alpha2number)
        value = 0
        for c in alpha_str:
            value = self.alpha2number[c] + value * scale
        return value

    def cartesian_to_alpha(self, num) -> str:
        # eg. 27 --> 'AA'
        alpha = ''
        scale = len(self.number2alpha)
        temp = num
        while temp >= scale:
            alpha = alpha + self.number2alpha[temp // scale]
            temp = temp % scale
        if temp != 0:
            alpha += self.number2alpha[temp]
        return alpha

    def resolve_excel_file(self):
        # 基本校验
        if os.path.exists(self.excel_file_path) is False:  # or os.path.isdir(self.excel_file_path):# mac下该函数无法准确判断
            self.is_process_sucess = False
            tkinter.messagebox.showwarning(title="warning", message="输入文件不存在，请重新选择输入excel文件")
            return

    def process_file(self):
        # 处理主函数，做一下几件事：
        # 1. 获取用户输入值
        self.get_user_input()
        # 2. 解析excel文件
        # 3. 创建管脚-网络映射矩阵，建立管脚-坐标映射矩阵
        # 4. 解析用户输入间距，循环更新矩阵
        print('打开文件：', self.excel_file_path)
        print('输出文件：', self.out_put_file_path)
        self.is_process_sucess = True


if __name__ == '__main__':
    '''
    显示逻辑
    '''
    app = QCC()
    print(app.cartesian_to_alpha(29))
    app.window.mainloop()  # 显示
