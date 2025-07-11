from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os


def search_basic_infor(driver, but_address, tr_address):
    '''
    基本信息部分：搜索并选择
    :param driver:
    :param but_address: 搜索按钮
    :param tr_address: 搜索结果
    '''
    for i in range(10):
        time.sleep(1)
        # 判断是否正在加载
        rows = driver.execute_script('return document.getElementsByClassName("ant-spin-dot ant-spin-dot-spin");')
        if len(rows) > 0:
            time.sleep(2)
            continue
        time.sleep(1)
        # 如果搜索结果不是一个，重新点击搜索按钮
        rows = driver.find_elements(By.XPATH, tr_address)
        if len(rows) != 1:
            driver.find_element(By.XPATH, but_address).click()
            continue
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, tr_address))).click()
        break

def create_general_reimbursement(driver, datas, config_dict, service_cor_dict):
    '''
    跳转到技服外包报销，创建技服费用报销单
    :param driver: 浏览器对象
    :param datas: [{}, {}...]
    :param config_dict: 配置文件
    :param service_cor_dict: 服务商信息字典
    '''
    reimburse_handels = driver.current_window_handle  # 财务报销系统 标签页的句柄
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[1]/ul/li[8]/div/div/div'))).click()
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[1]/ul/li[8]/ul/li[10]/div/div').click()
    time.sleep(3)
    driver.switch_to.window(driver.window_handles[-1])

    # 进入创建报销单页面
    # 检查页面是否加载完成
    for i in range(3):
        try:
            time.sleep(3)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[5]/td[5]/div/div/span[1]/div/div/div/div[2]/button')))
            break
        except:
            driver.refresh()
            continue

    time.sleep(1)
    # 填写基本信息
    # 申请人
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[5]/td[5]/div/div/span[1]/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys('00072593')  # 陈月青
    but_address = "//*[@class='wea-search-tab']/div/div/button"
    tr_address = "//*[@class='wea-crm-list']/ul/li"
    driver.find_element(By.XPATH, but_address).click()
    search_basic_infor(driver, but_address, tr_address)
    time.sleep(1)

    # # 费用成本中心
    # driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[6]/td[5]/div/div/span[1]/div/div/div/div[2]/button').click()
    # driver.find_element(By.XPATH, "//*[@class='wea-tab-outer-SearchAd']/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input").send_keys(datas[0].get('业务范围'))
    # but_address = "//*[@class='wea-search-tab']/span/button"
    # tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
    # driver.find_element(By.XPATH, but_address).click()
    # search_basic_infor(driver, but_address, tr_address)
    # time.sleep(1)

    # 平台
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[7]/td[5]/div/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys('北京')
    but_address = "//*[@class='wea-search-tab']/div/div/button"
    tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
    driver.find_element(By.XPATH, but_address).click()
    search_basic_infor(driver, but_address, tr_address)
    time.sleep(1)

    # 费用是由
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[8]/td[5]/div/div/span[1]/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, "//*[@class='wea-tab-outer-SearchAd']/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input").send_keys('231090301')
    but_address = "//*[@class='wea-search-tab']/span/button"
    tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
    driver.find_element(By.XPATH, but_address).click()
    search_basic_infor(driver, but_address, tr_address)
    time.sleep(1)

    # 结算方式
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[9]/td[5]/div/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys('电汇')
    but_address = "//*[@class='wea-search-tab']/div/div/button"
    tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
    driver.find_element(By.XPATH, but_address).click()
    search_basic_infor(driver, but_address, tr_address)
    time.sleep(1)

    # 汇入省
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[12]/td[8]/div/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys(service_cor_dict['省份'])
    but_address = "//*[@class='wea-search-tab']/div/div/button"
    tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
    driver.find_element(By.XPATH, but_address).click()
    search_basic_infor(driver, but_address, tr_address)
    time.sleep(1)


