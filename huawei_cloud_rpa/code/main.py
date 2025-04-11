import tkinter as tk
# from tkinter.filedialog import askdirectory
from tkinter import filedialog
import threading
import pandas as pd
# import DI
from tkinter import messagebox
# import table_handling
from time import sleep
import os
# import utils

class App(object):
    def __init__(self, root):
        self.filePath = {}

        root.title("华为云RPA")
        root.geometry('500x500')

        self.two_four_data = tk.StringVar()
        label01 = tk.Label(root, text="24年数据：")
        label01.grid(row=0, column=0)
        entry01 = tk.Entry(root, textvariable=self.two_four_data, width=40)
        entry01.grid(row=0, column=1)
        btn01 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.two_four_data))
        btn01.grid(row=0, column=2)

        self.two_five_data = tk.StringVar()
        label02 = tk.Label(root, text="25年业绩表：")
        label02.grid(row=1, column=0)
        entry02 = tk.Entry(root, textvariable=self.two_five_data, width=40)
        entry02.grid(row=1, column=1)
        btn02 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.two_five_data))
        btn02.grid(row=1, column=2)

        self.product_details = tk.StringVar()
        label03 = tk.Label(root, text="2025产品明细：")
        label03.grid(row=2, column=0)
        entry03 = tk.Entry(root, textvariable=self.product_details, width=40)
        entry03.grid(row=2, column=1)
        btn03 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.product_details))
        btn03.grid(row=2, column=2)

        self.customer_cor = tk.StringVar()
        label04 = tk.Label(root, text="客户对应关系表：")
        label04.grid(row=3, column=0)
        entry04 = tk.Entry(root, textvariable=self.customer_cor, width=40)
        entry04.grid(row=3, column=1)
        btn04 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.customer_cor))
        btn04.grid(row=3, column=2)

        self.data_requirements = tk.StringVar()
        label05 = tk.Label(root, text="数据需求表：")
        label05.grid(row=4, column=0)
        entry05 = tk.Entry(root, textvariable=self.data_requirements, width=40)
        entry05.grid(row=4, column=1)
        btn05 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.data_requirements))
        btn05.grid(row=4, column=2)

        btn005 = tk.Button(root, text="匹配", command=self.start)
        btn005.grid(row=7, column=2)

        # 结果打印框
        self.text = tk.Text(selectbackground="red", insertbackground="blue", spacing2=10, bd=0)
        self.text.grid(row=9, column=0, columnspan=10)

    def start(self):
        self.T = threading.Thread(target=self.ming_xi_handling)
        self.T.setDaemon(True)
        self.T.start()

    def selectPath(self, path):
        path_ = filedialog.askopenfilename()
        path.set(path_)

    def ming_xi_handling(self):
        try:
            self.text.insert(tk.END, "开始...\r\n")
            two_four_path = self.two_four_data.get()
            two_five_path = self.two_five_data.get()
            product_details_path = self.product_details.get()
            customer_cor_path = self.customer_cor.get()
            data_requirements_path = self.data_requirements.get()
            if not two_four_path or not two_five_path or not product_details_path or not customer_cor_path or not data_requirements_path:
                self.text.insert(tk.END, "文件路径不能为空！\r\n")
                return
            self.text.insert(tk.END, "检查客户对应关系表是否完整...\r\n")
            # 检查客户对应关系表是否完整
            flag = self.check_customer_cor(customer_cor_path)
            if not flag:
                return

            self.text.insert(tk.END, "数据解析中...\r\n")
            # 数据入库


            self.text.insert(tk.END, "25年业绩表增加BI到BO列...\r\n")
            # todo 25年业绩表增加BI到BO列

            self.text.insert(tk.END, "数据第一部分...\r\n")
            self.text.insert(tk.END, "数据第二部分...\r\n")
            self.text.insert(tk.END, "数据第三部分...\r\n")
            self.text.insert(tk.END, "数据第四部分...\r\n")
            self.text.insert(tk.END, "数据第五部分...\r\n")
            self.text.insert(tk.END, "数据第六部分...\r\n")
            self.text.insert(tk.END, "数据第七部分...\r\n")
            self.text.insert(tk.END, "数据第八部分...\r\n")
            self.text.insert(tk.END, "数据第九部分...\r\n")

            self.text.insert(tk.END, "结果表下载中...\r\n")

            self.text.insert(tk.END, "处理完成！\r\n")
        except BaseException as e:
            self.text.insert(tk.END, "发生错误！\r\n")
            self.text.insert(tk.END, e)

    # 验证客户对应关系表是否完整
    def check_customer_cor(self, customer_cor_path):
        # 读取客户对应关系表，判断“客户名称”不为空的行，是否都有“销售员”和“区域”
        try:
            # 读取Excel文件
            df = pd.read_excel(customer_cor_path)

            # 筛选有客户名称但缺失销售员或区域的行
            missing_data = df[
                (df['客户名称'].notnull()) &
                (df[['销售员', '区域']].isnull().any(axis=1))
                ]

            if not missing_data.empty:
                error_info = "\n".join([
                    f"第{row.Index + 2}行 {row['客户名称']}缺少: "
                    f"{'销售员' if pd.isnull(row['销售员']) else ''} "
                    f"{'区域' if pd.isnull(row['区域']) else ''}".strip()
                    for row in missing_data.itertuples()
                ])
                self.text.insert(tk.END, "客户对应关系表不完整！\r\n")
                self.text.insert(tk.END, error_info + "\r\n")
                return False
            return True
        except Exception as e:
            self.text.insert(tk.END, f"验证客户表时发生错误: {str(e)}\r\n")
            return False


if __name__ == '__main__':
    root = tk.Tk()

    App(root)
    root.mainloop()
