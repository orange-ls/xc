import threading
import tkinter as tk
from tkinter import filedialog
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import sys
import re
import oa_expense
import crm_download
import oa_general

business_scope_dict = {
    '4001': 'MH400003',
    'QH01': 'MHQH0003',
    'MU01': 'MHMU0002',
}
file_download_path = r"D:\Google\test"

class App(object):
    def __init__(self, root):
        self.filePath = {}

        root.title("费用报销")
        root.geometry('500x500')

        self.order_num_table = tk.StringVar()
        label02 = tk.Label(root, text="工单号表：")
        label02.grid(row=1, column=0)
        entry02 = tk.Entry(root, textvariable=self.order_num_table, width=40)
        entry02.grid(row=1, column=1)
        btn02 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.order_num_table))
        btn02.grid(row=1, column=2)

        self.asp_table = tk.StringVar()
        label03 = tk.Label(root, text="25**ASP表：")
        label03.grid(row=2, column=0)
        entry03 = tk.Entry(root, textvariable=self.asp_table, width=40)
        entry03.grid(row=2, column=1)
        btn03 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.asp_table))
        btn03.grid(row=2, column=2)

        self.config_file = tk.StringVar()
        # todo 调整为客户指定路径
        self.config_file.set(r"C:\Users\user\Desktop\费用报销rpa配置表.xlsx")
        label03 = tk.Label(root, text="配置文件表：")
        label03.grid(row=3, column=0)
        entry03 = tk.Entry(root, textvariable=self.config_file, width=40)
        entry03.grid(row=3, column=1)
        btn03 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.config_file))
        btn03.grid(row=3, column=2)

        btn005 = tk.Button(root, text="执行", command=self.start)
        btn005.grid(row=8, column=2)

        # 结果打印框
        self.text = tk.Text(selectbackground="red", insertbackground="blue", spacing2=10, bd=0)
        self.text.grid(row=9, column=0, columnspan=10)

        # todo 增加发票名称模板、工单号表名称模板、ASP表名称模板
        self.text.insert(tk.END, "ASP表表名规则：数字年月+月+ASP，例：2501月ASP.xlsx \r\n")


    def start(self):
        self.T = threading.Thread(target=self.data_handling)
        self.T.setDaemon(True)
        self.T.start()

    def selectPath(self, path):
        path_ = filedialog.askopenfilename()
        path.set(path_)

    def data_handling(self):
        try:
            self.text.insert(tk.END, "开始...\r\n")
            order_num_path = self.order_num_table.get()
            asp_table_path = self.asp_table.get()
            config_file_path = self.config_file.get()

            # if not os.path.exists(order_num_path) or not os.path.exists(asp_table_path) or not os.path.exists(config_file_path):
            #     self.text.insert(tk.END, "文件路径错误！\r\n")
            #     return
            # 读取配置文件为字典
            config_df = pd.read_excel(config_file_path, sheet_name='流程配置')[['名称', '值']]
            config_dict = {row['名称']: row['值'] for i, row in config_df.iterrows()}
            service_cor_df = pd.read_excel(config_file_path, sheet_name='服务商对应表')
            service_cor_list = service_cor_df.to_dict(orient='records')
            service_cor_dict = {row['服务商名称']: row for row in service_cor_list}

            # 将asp表的名字加入到配置文件字典中
            asp_table_name = os.path.basename(asp_table_path).replace('.xlsx', '')
            config_dict['asp表名称'] = asp_table_name

            # 获取文件的月份
            month = re.search(r'(?:^|.*?)(\d+月)', os.path.basename(asp_table_name)).group(1)
            month = month[2:].replace('0', '')
            config_dict['月份'] = month

            # 读取工单号表
            all_sheets = pd.ExcelFile(order_num_path).sheet_names
            sheet_names = [s for s in all_sheets if '例外服务' in s]
            driver = self.create_browser(config_dict['谷歌浏览器下载路径'])
            # 登录CRM系统，跳转到工单搜索界面
            crm_download.login_crm(driver, config_dict)
            # todo config_dict中增加税前金额 config_dict['ASP名称-MU01'] = 金额，如果asp表不能抓取出金额，则使用这个值

            for sheet_name in sheet_names:
                order_num_table = pd.read_excel(order_num_path, sheet_name=sheet_name)[['工单号/项目交付单号', 'ASP名称', 'ASP金额']]
                # 删除工单号为空的行
                order_num_table = order_num_table.dropna(subset=['工单号/项目交付单号'])
                order_num_list = order_num_table['工单号/项目交付单号'].tolist()
                asp_name = order_num_table['ASP名称'].tolist()[0]
                asp_amount = sum(order_num_table['ASP金额'].tolist())
                file_name = f"{month}_{asp_name}_{sheet_name}_{asp_amount}"      # 压缩包名：2月_北京神州光大科技有限公司_金额总和

                # 搜索工单，下载文件
                # 压缩包文件保存位置 使用配置文件管理
                crm_download.crm_download_file(driver, order_num_list, config_dict['谷歌浏览器下载路径'], config_dict['CRM文件保存路径'], file_name)
            driver.quit()

            # 读取asp表
            asp_table = pd.read_excel(asp_table_path, sheet_name='Sheet1')[['项目编号', '技服预提金额', '外包供应商名称', '业务范围', '项目总收入']].sort_values(by='外包供应商名称')

            # 按'外包供应商名称', '业务范围'为键，其他字段为值的字典
            asp_dict = {}
            for i, row in asp_table.iterrows():
                key = (row['外包供应商名称'], row['业务范围'])
                # 转换业务范围 MU01 -> MHMU0002
                value = {
                    '外包供应商名称': row['外包供应商名称'],
                    '业务范围': business_scope_dict.get(row['业务范围'], row['业务范围']),
                    '项目编号': row['项目编号'],
                    '技服预提金额': row['技服预提金额'],
                    '项目总收入': row['项目总收入'],
                }
                if key not in asp_dict:
                    asp_dict[key] = []
                asp_dict[key].append(value)

            # todo 增加税前金额
            for key, value in asp_dict.items():
                sum_amount = sum([v['技服预提金额'] for v in value])
                for v in value:
                    v['税前金额'] = sum_amount

            driver = self.create_browser(config_dict['谷歌浏览器下载路径'])
            oa_expense.login_oa(driver, config_dict)
            for key, value in asp_dict.items():
                oa_expense.create_expense_reimbursement(driver, value, config_dict, service_cor_dict.get(value[0]['外包供应商名称']))

            driver.quit()
            self.text.insert(tk.END, "执行完毕！\r\n")

        except Exception as e:
            self.text.insert(tk.END, "发生错误！\r\n")
            self.text.insert(tk.END, e)

    def get_chromedriver_path(self):
        # 开发环境路径
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        driver_path = os.path.join(base_path, "chromedriver_v138.exe")
        return driver_path

    def create_browser(self, download_path):
        # 创建设置浏览器对象
        q1 = Options()
        q1.add_argument('--no-sandbox')
        q1.add_argument('--start-maximized')
        q1.add_experimental_option('detach', True)

        # 设置浏览器下载路径
        download_path = os.path.normpath(download_path)
        os.makedirs(download_path, exist_ok=True)
        q1.add_experimental_option('prefs', {'download.default_directory': download_path})

        driver_path = self.get_chromedriver_path()
        driver = webdriver.Chrome(service=Service(driver_path), options=q1)
        # 隐性等待30秒
        driver.implicitly_wait(30)
        return driver


if __name__ == '__main__':
    root = tk.Tk()

    App(root)
    root.mainloop()