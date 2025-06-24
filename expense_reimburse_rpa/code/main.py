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


class App(object):
    def __init__(self, root):
        self.filePath = {}

        root.title("费用报销")
        root.geometry('500x500')

        self.two_five_data = tk.StringVar()
        label02 = tk.Label(root, text="工单号表：")
        label02.grid(row=1, column=0)
        entry02 = tk.Entry(root, textvariable=self.two_five_data, width=40)
        entry02.grid(row=1, column=1)
        btn02 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.two_five_data))
        btn02.grid(row=1, column=2)

        self.asp_table = tk.StringVar()
        label03 = tk.Label(root, text="25**ASP表：")
        label03.grid(row=2, column=0)
        entry03 = tk.Entry(root, textvariable=self.asp_table, width=40)
        entry03.grid(row=2, column=1)
        btn03 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.asp_table))
        btn03.grid(row=2, column=2)


        btn005 = tk.Button(root, text="执行", command=self.start)
        btn005.grid(row=8, column=2)

        # 结果打印框
        self.text = tk.Text(selectbackground="red", insertbackground="blue", spacing2=10, bd=0)
        self.text.grid(row=9, column=0, columnspan=10)


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
            two_five_path = self.two_five_data.get()
            asp_table_path = self.asp_table.get()

            # 读取asp表
            asp_table = pd.read_excel(asp_table_path, sheet_name='Sheet1')[['项目编号', '技服预提金额', '外包供应商名称', '业务范围', '项目总收入']].sort_values(by='外包供应商名称')

            # 按'外包供应商名称', '业务范围'为键，其他字段为值的字典
            asp_dict = {}
            # 遍历每一行，按两列组合为键，存储所有重复项的值
            for i, row in asp_table.iterrows():
                key = (row['外包供应商名称'], row['业务范围'])
                value = {
                    '外包供应商名称': row['外包供应商名称'],
                    '业务范围': row['业务范围'],
                    '项目编号': row['项目编号'],
                    '技服预提金额': row['技服预提金额'],
                    '项目总收入': row['项目总收入'],
                }
                if key not in asp_dict:
                    asp_dict[key] = []
                asp_dict[key].append(value)

            driver = self.login_oa()
            for key, value in asp_dict.items():
                self.create_expense_reimbursement(driver, value)

            print("执行完毕！")


        except Exception as e:
            self.text.insert(tk.END, "发生错误！\r\n")
            self.text.insert(tk.END, e)

    def get_chromedriver_path(self):
        # 开发环境路径
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        driver_path = os.path.join(base_path, "chromedriver_v137.exe")
        # 打包后路径检测
        if not os.path.exists(driver_path):
            # 尝试上级目录
            driver_path = os.path.join(base_path, "../driver/chromedriver_v137.exe")
        return driver_path

    def create_browser(self):
        # 创建设置浏览器对象
        q1 = Options()
        q1.add_argument('--no-sandbox')
        q1.add_argument('--start-maximized')
        q1.add_experimental_option('detach', True)

        # driver = webdriver.Chrome(service=Service('chromedriver_v137.exe'), options=q1)
        driver = webdriver.Chrome(service=Service(r'D:\Workspace\xc\expense_reimburse_rpa\chromedriver_v137.exe'), options=q1)  # todo 需要替换成自己的chromedriver路径
        # driver_path = self.get_chromedriver_path()
        # driver = webdriver.Chrome(service=Service(driver_path), options=q1)
        # 隐性等待30秒
        driver.implicitly_wait(30)
        return driver

    def login_oa(self):
        # 登录OA，跳转到财务报销系统
        driver = self.create_browser()
        driver.get('https://newportal.digitalchina.com')

        # 登录
        driver.find_element(By.XPATH, '//*[@id="usernameInput"]').send_keys('lishuaiae')
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[3]/td/input[1]').click()
        driver.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[3]/td/input[2]').send_keys('Mm2002902L.')
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/div[3]/table/tbody/tr[5]/td/img').click()

        try:
            # 跳转到财务报销系统
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[4]/div[2]/div/div[1]/div/div[1]/ul/li[5]/div/span/span'))).click()
            # 点击“报销和借款”按钮
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[4]/div[3]/div[1]/div/div/div/div[1]/div/div/table/tbody/tr[3]/td[1]/div/div/div[3]/div[1]/div/div/div/div/div[3]/div[2]/div/div/div[1]/div/div[2]/div[1]').click()
        except:
            time.sleep(1)
            driver.refresh()
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[4]/div[2]/div/div[1]/div/div[1]/ul/li[5]/div/span/span'))).click()
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[4]/div[3]/div[1]/div/div/div/div[1]/div/div/table/tbody/tr[3]/td[1]/div/div/div[3]/div[1]/div/div/div/div/div[3]/div[2]/div/div/div[1]/div/div[2]/div[1]').click()
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])
        return driver


    def create_expense_reimbursement(self, driver, datas):
        '''
        跳转到技服外包报销
        :param driver: 浏览器对象
        :param datas: [{}, {}...]
        :return:
        '''
        reimburse_handels = driver.current_window_handle    # 财务报销系统 标签页的句柄
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[1]/ul/li[8]/div/div/div').click()
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[1]/ul/li[8]/ul/li[11]/div/div').click()

        time.sleep(3)
        driver.switch_to.window(driver.window_handles[-1])
        # todo 填写基本信息
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[8]/td[5]/div/div/span[1]/div/div/div/div[1]').send_keys('231090301')

        # 填写 项目明细
        for i in range(3):  # 尝试填写3次
            i = 4
            for data in datas:
                # 点击加号
                try:
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[2]/td[1]/div/div/i[1]').click()
                except:
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[1]/table/tbody/tr[2]/td[1]/div/div/i[1]').click()
                time.sleep(0.5)

                try:
                    driver.find_element(By.XPATH, f'/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[3]/div/div/input')
                except:
                    try:
                        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[2]/td[1]/div/div/i[1]').click()
                    except:
                        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[1]/table/tbody/tr[2]/td[1]/div/div/i[1]').click()
                    time.sleep(0.5)

                # 填写必填项
                driver.find_element(By.XPATH, f'/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[3]/div/div/input').send_keys(data['项目编号'])
                driver.find_element(By.XPATH,  f'/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[10]/div/div/input').send_keys(data['技服预提金额'])
                driver.find_element(By.XPATH, f'/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[{i}]/td[13]/div/div/input').send_keys(data['项目总收入'])
                time.sleep(1)
                i += 1
            # 验证是否填写完整
            if i != len(datas)+4:
                # 如果不完整，删除全部行，重新填写
                try:
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[3]/td[1]/span/label/span/input').click()
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[2]/table/tbody/tr[2]/td[1]/div/div/i[2]').click()
                except:
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[1]/table/tbody/tr[3]/td[1]/span/label/span/input').click()
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[45]/td[4]/div/div/div/div/div[1]/table/tbody/tr[2]/td[1]/div/div/i[2]').click()
                # 点击弹窗确定按钮
                time.sleep(1)
                driver.find_element(By.XPATH, '/html/body/div[8]/div/div[2]/div/div[1]/div/div/div[2]/button[1]').click()
                continue
            # todo 如果有弹窗提示没有填写完整，则需要重新填写某一行
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[44]/td[6]/div/input').click()


            break


        # todo 点击保存
        # 关闭新建标签页，切换到财务报销系统标签页
        driver.close()
        driver.switch_to.window(reimburse_handels)





if __name__ == '__main__':
    root = tk.Tk()

    App(root)
    root.mainloop()