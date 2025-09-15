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
from datetime import datetime
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
    'QG01': 'MHQG0001',
}

class App(object):
    def __init__(self, root):
        self.filePath = {}

        root.title("费用报销")
        root.geometry('565x500')

        self.asp_table = tk.StringVar()
        label03 = tk.Label(root, text="25**ASP表：")
        label03.grid(row=2, column=0)
        entry03 = tk.Entry(root, textvariable=self.asp_table, width=55)
        entry03.grid(row=2, column=1)
        btn03 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.asp_table))
        btn03.grid(row=2, column=2)

        self.config_file = tk.StringVar()
        # 自动获取配置文件路径
        excel_path = self.get_excel_path('费用报销rpa配置表.xlsx')
        self.config_file.set(excel_path)
        label03 = tk.Label(root, text="配置文件表：")
        label03.grid(row=3, column=0)
        entry03 = tk.Entry(root, textvariable=self.config_file, width=55)
        entry03.grid(row=3, column=1)
        btn03 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.config_file))
        btn03.grid(row=3, column=2)

        btn005 = tk.Button(root, text="执行", command=self.start)
        btn005.grid(row=8, column=2)

        # 结果打印框
        self.text = tk.Text(selectbackground="red", insertbackground="blue", spacing2=10, bd=0)
        self.text.grid(row=9, column=0, columnspan=10)

        self.text.insert(tk.END, "--------------------------------------------------------------------------------\r\n")
        self.text.insert(tk.END, "1、ASP表名规则：数字年月+月+ASP，例：2501月ASP.xlsx  读取名为“Sheet1“的Sheet\r\n\n")
        self.text.insert(tk.END, "2、派单记录表名规则：数字月+月份ASP上门派单记录-公司名-服务种类，例：1月份ASP上门派单记录-北京神州光大科技有限公司-例外服务-MU01.xlsx \r\n\n")
        self.text.insert(tk.END, "3、路径和表格中涉及到公司名，请使用全称，如北京神州光大科技有限公司\r\n\n")
        self.text.insert(tk.END, "4、发票命名规则：月份-公司名-税前金额，例：1月-北京神州光大科技有限公司-1000.pdf\r\n")
        self.text.insert(tk.END, "--------------------------------------------------------------------------------\r\n\n")


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
            asp_table_path = self.asp_table.get()
            config_file_path = self.config_file.get()
            if not os.path.exists(asp_table_path) or not os.path.exists(config_file_path):
                self.text.insert(tk.END, "文件路径错误！\r\n")
                return
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
            date = re.search(r'(?:^|.*?)(\d+月)', os.path.basename(asp_table_name)).group(1)
            month = date[2:].replace('0', '')
            config_dict['月份'] = month   # 2月
            year = f"20{date[:2]}"   # 2025
            config_dict['年份'] = year
            config_dict['数字月份'] = date[2:].replace('月', '')

            # 读取asp表
            asp_table = pd.read_excel(asp_table_path, sheet_name='Sheet1')[['项目编号', '技服预提金额', '外包供应商名称', '业务范围', '项目总收入']].sort_values(by='外包供应商名称')

            # 按'外包供应商名称', '业务范围'为键，其他字段为值的字典
            asp_dict = {}
            for i, row in asp_table.iterrows():
                key = (row['外包供应商名称'], row['业务范围'])
                # 转换业务范围 MU01 -> MHMU0002
                value = {
                    '外包供应商名称': row['外包供应商名称'],
                    '业务范围': business_scope_dict.get(str(row['业务范围']), row['业务范围']),
                    '项目编号': row['项目编号'],
                    '技服预提金额': row['技服预提金额'],
                    '项目总收入': row['项目总收入'],
                }
                if key not in asp_dict:
                    asp_dict[key] = []
                asp_dict[key].append(value)

            # 拼接出订单表的路径
            asp_suppliers = list({key[0] for key in asp_dict.keys()})
            order_num_paths = [os.path.join(config_dict['派单记录表路径'], f"{supplier}\{year}\{month}份ASP上门派单记录-{supplier}.xlsx") for supplier in asp_suppliers]
            # 判断文件是否存在，不存在就报错提示
            for path in order_num_paths:
                if not os.path.exists(path):
                    self.text.insert(tk.END, f"{path} 文件不存在！\r\n")
                    return
            config_dict['工单号表'] = order_num_paths

            # 判断通用报销文件是否存在  按照ASP表中的公司名，找到对应公司总表中的sheet名，再在对应的路径中寻找是否有包含sheet名的文件
            general_path_list = []
            list_server = ['标准服务', '高级服务']
            for supplier in asp_suppliers:
                for server in list_server:
                    path = os.path.join(config_dict['派单记录表路径'], f"{supplier}\{year}\{month}份ASP上门派单记录-{supplier}-{server}.xlsx")
                    if os.path.exists(path):
                        general_path_list.append(path)

            self.text.insert(tk.END, "\n------开始下载现场服务报告------\r\n")
            is_crm_login = None
            for order_num_path in order_num_paths:
                # 读取工单号表
                all_sheets = pd.ExcelFile(order_num_path).sheet_names
                sheet_names = [s for s in all_sheets if 'Sheet1' not in s]

                for sheet_name in sheet_names:
                    order_num_table = pd.read_excel(order_num_path, sheet_name=sheet_name, converters={'工单号/项目交付单号': lambda x: f"{int(x)}" if "." in str(x) else str(x)})[['工单号/项目交付单号', 'ASP名称', 'ASP金额']]
                    # 删除工单号为空的行
                    order_num_table = order_num_table.dropna(subset=['工单号/项目交付单号'])
                    order_num_list = order_num_table['工单号/项目交付单号'].tolist()
                    asp_name = order_num_table['ASP名称'].tolist()[0]
                    asp_amount = sum(order_num_table['ASP金额'].tolist())

                    # config_dict中增加税前金额 config_dict['ASP名称-例外服务-MU01'] = 金额，如果asp表不能抓取出金额，则使用这个值
                    config_dict[f'{asp_name}-{sheet_name}'] = asp_amount

                    file_name = f"{asp_name}\{config_dict['年份']}\现场服务单\{config_dict['数字月份']}\{month}-{asp_name}-{sheet_name}-现场服务报告"      # 压缩包名：2月-北京神州光大科技有限公司-例外服务-MU01-现场服务报告
                    zip_name = os.path.join(config_dict['验收文件保存路径'], f"{file_name}.zip")    # D:\工作\01-合作管理\01-派单记录\河北华恒信通信技术有限公司\2025\现场服务单\02\2月-北京神州光大科技有限公司-例外服务-MU01-现场服务报告

                    # 先判断压缩包保存文件夹是否存在，不存在 就创建；存在 再判断压缩包是否存在，存在就跳过
                    if not os.path.exists(os.path.dirname(zip_name)):
                        os.makedirs(os.path.dirname(zip_name))
                    else:
                        if os.path.exists(zip_name):
                            # 判断压缩包是否存在，存在就跳过
                            self.text.insert(tk.END, f"{zip_name} 文件已存在！\r\n")
                            continue

                    if is_crm_login:
                        pass
                    else:
                        # 登录CRM系统，跳转到工单搜索界面
                        driver = self.create_browser(config_dict['谷歌浏览器下载路径'])
                        driver = crm_download.login_crm(driver, config_dict)
                        is_crm_login = driver.current_window_handle

                    # 搜索工单，下载文件。 压缩包文件保存位置 使用配置文件管理
                    crm_download.crm_download_file(driver, order_num_list, config_dict['谷歌浏览器下载路径'], config_dict['验收文件保存路径'], file_name)
            if is_crm_login:
                driver.quit()
                is_crm_login = None

            self.text.insert(tk.END, "\n\n------开始技服外包报销------\r\n")
            driver = self.create_browser(config_dict['谷歌浏览器下载路径'])
            # 登录OA系统，跳转到报销系统界面
            oa_expense.login_oa(driver, config_dict)
            # 进入技服外包报销
            for key, value in asp_dict.items():
                service_cor_data = service_cor_dict.get(value[0]['外包供应商名称'])
                if not service_cor_data:
                    self.text.insert(tk.END, f"配置表中没有找到{value[0]['外包供应商名称']}的配置信息！\r\n")
                    continue

                path = os.path.join(config_dict['派单记录表路径'], f"{key[0]}\{year}\{month}份ASP上门派单记录-{key[0]}-例外服务-{key[1]}.xlsx")
                if os.path.exists(path):
                    oa_expense.create_expense_reimbursement(driver, key[1], value, config_dict, service_cor_data)
                else:
                    self.text.insert(tk.END, f"\n跳过技服外包报销，{path} 文件不存在！\r\n")
            # 开始通用报销处理
            self.text.insert(tk.END, "\n\n ------开始通用报销处理------\r\n")
            for general_path in general_path_list:
                asp_name = os.path.basename(general_path).split('-')[1]
                service_cor_data = service_cor_dict.get(asp_name)
                if not service_cor_data:
                    self.text.insert(tk.END, f"配置表中没有找到{asp_name}的配置信息！\r\n")
                    continue
                oa_general.create_general_reimbursement(driver, config_dict, service_cor_data, general_path)

            driver.quit()
            self.text.insert(tk.END, "执行完毕！\r\n")
        except Exception as e:
            self.text.insert(tk.END, "\n发生错误！\r\n")
            self.text.insert(tk.END, e)
            self.text.insert(tk.END, '\n')

    def get_chromedriver_path(self):
        # 开发环境路径
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        driver_path = os.path.join(base_path, "chromedriver_v138.exe")
        return driver_path

    def get_excel_path(self, name):
        # 获取exe文件所在目录
        if getattr(sys, 'frozen', False):
            # 如果程序是打包后的exe
            exe_dir = os.path.dirname(sys.executable)
        else:
            # 如果程序是直接运行的Python脚本
            exe_dir = os.path.dirname(os.path.abspath(__file__))
        # 拼接Excel文件路径
        excel_path = os.path.join(exe_dir, name)
        return excel_path

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