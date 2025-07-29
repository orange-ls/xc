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


def go_reimbursement(driver):
    # 进入技服外包报销界面
    for i in range(3):
        try:
            reimburse_handels = driver.current_window_handle  # 财务报销系统 标签页的句柄
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[text()='报销申请']")))
            # 通过js快速定位元素
            service_claim_js = """return document.evaluate("//div[text()='其他通用报销']", document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null).snapshotLength;"""
            element_count = driver.execute_script(service_claim_js)
            if element_count == 0:
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//div[text()='报销申请']"))).click()
            driver.find_element(By.XPATH, "//div[text()='其他通用报销']").click()
            time.sleep(3)
            driver.switch_to.window(driver.window_handles[-1])
            break
        except:
            driver.refresh()
            time.sleep(3)
            if i == 2:
                raise Exception("打开报销系统失败！")
            continue
    return driver, reimburse_handels


def create_general_reimbursement(driver, config_dict, service_cor_dict, general_path):
    '''
    跳转到技服外包报销，创建技服费用报销单
    :param driver: 浏览器对象
    :param datas: [{}, {}...]
    :param config_dict: 配置文件
    :param service_cor_dict: 服务商信息字典
    :param general_path: 通用报销 工单号表路径
    '''
    type = '标准服务' if '标准服务' in general_path else '高级服务'
    service_provider_name = service_cor_dict['服务商名称']
    config_amount_name = f"{service_provider_name}-{type}"  # 配置文件 对应未税金额的键
    if config_amount_name not in config_dict:
        raise FileNotFoundError(f"文件路径不存在:{config_amount_name}")
    # 进入技服外包报销界面
    driver, reimburse_handels = go_reimbursement(driver)
    for index in range(3):
        try:
            for i in range(3):
                # 进入创建报销单页面
                # 检查页面是否加载完成
                try:
                    time.sleep(3)
                    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="field353734span"]/div[2]/button')))
                    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[6]')))
                    break
                except:
                    driver.refresh()
                    continue

            time.sleep(1)
            # 填写基本信息
            # 申请人
            driver.find_element(By.XPATH, '//*[@id="field353734span"]/div[2]/button').click()
            driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys('00072593')  # 陈月青
            but_address = "//*[@class='wea-search-tab']/div/div/button"
            tr_address = "//*[@class='wea-crm-list']/ul/li"
            driver.find_element(By.XPATH, but_address).click()
            search_basic_infor(driver, but_address, tr_address)
            time.sleep(1)

            # 费用成本中心
            driver.find_element(By.XPATH, '//*[@id="field353738span"]/div[2]/button').click()
            driver.find_element(By.XPATH, "//*[@class='wea-tab-outer-SearchAd']/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input").send_keys('MH480002')
            but_address = "//*[@class='wea-search-tab']/span/button"
            tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
            driver.find_element(By.XPATH, but_address).click()
            search_basic_infor(driver, but_address, tr_address)
            time.sleep(1)

            # 平台
            driver.find_element(By.XPATH, '//*[@id="field353772span"]/div[2]/button').click()
            driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys('北京')
            but_address = "//*[@class='wea-search-tab']/div/div/button"
            tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
            driver.find_element(By.XPATH, but_address).click()
            search_basic_infor(driver, but_address, tr_address)
            time.sleep(1)

            # 费用是由
            num = '407100101' if '标准服务' in general_path else '407100201'
            driver.find_element(By.XPATH, '//*[@id="field353788span"]/div[2]/button').click()
            driver.find_element(By.XPATH, "//*[@class='wea-tab-outer-SearchAd']/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input").send_keys(num)
            but_address = "//*[@class='wea-search-tab']/span/button"
            tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
            driver.find_element(By.XPATH, but_address).click()
            search_basic_infor(driver, but_address, tr_address)
            time.sleep(1)

            # 结算方式
            driver.find_element(By.XPATH, '//*[@id="field353773span"]/div[2]/button').click()
            driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys('电汇')
            but_address = "//*[@class='wea-search-tab']/div/div/button"
            tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
            driver.find_element(By.XPATH, but_address).click()
            search_basic_infor(driver, but_address, tr_address)
            time.sleep(1)

            # 汇入省
            driver.find_element(By.XPATH, '//*[@id="field353785span"]/div[2]/button').click()
            driver.find_element(By.XPATH, "//*[@class='wea-search-tab']/div/div/span[1]/input").send_keys(service_cor_dict['省份'])
            but_address = "//*[@class='wea-search-tab']/div/div/button"
            tr_address = "//*[@class='ant-table-body']/table/tbody/tr"
            driver.find_element(By.XPATH, but_address).click()
            search_basic_infor(driver, but_address, tr_address)
            time.sleep(1)

            # 收款单位
            driver.find_element(By.XPATH, '//*[@id="field353782"]').send_keys(service_cor_dict['服务商名称'])
            # 判断是否有confirm确认框弹出
            driver.find_element(By.XPATH, '//*[@id="field353783"]').click()
            try:
                alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
                alert.dismiss()
            except:
                print("no alert")
            # 收款单位开户行
            # driver.find_element(By.XPATH, '//*[@id="field353783"]').send_keys(service_cor_dict['收款单位开户行'])
            # 收款单位开户行 获取收款单位客户行的值
            for i in range(6):
                customer_bank_account = driver.find_element(By.XPATH, '//*[@id="field353783"]').get_attribute('value')
                if customer_bank_account == service_cor_dict['收款单位开户行']:
                    break
                else:
                    driver.find_element(By.XPATH, '//*[@id="field353783"]').clear()
                    driver.find_element(By.XPATH, '//*[@id="field353783"]').send_keys(service_cor_dict['收款单位开户行'])
            # 收款账号
            driver.find_element(By.XPATH, '//*[@id="field353784"]').send_keys(service_cor_dict['收款账号'])
            # 是否冲借款
            driver.find_element(By.XPATH, '//*[@id="weaSelect_5"]/div/label[1]/span[1]/input').click()
            # 用途说明/备注
            driver.find_element(By.XPATH, '//*[@id="field353777"]').send_keys(f"FY{config_dict['asp表名称']}委托统计费用结算")
            # 汇入市
            driver.find_element(By.XPATH, '//*[@id="field353786"]').send_keys(service_cor_dict['城市'])

            # # 上传发票
            # invoice_name = f"{config_dict['月份']}-{service_provider_name}-{config_dict[config_amount_name]}.pdf"
            # invoice_path = os.path.join(config_dict['发票保存路径'], invoice_name)
            # if not os.path.exists(invoice_path):
            #     raise FileNotFoundError(f"文件路径不存在:{invoice_path}")
            # # 点击"选择发票"按钮
            # driver.find_element(By.XPATH, '//*[@id="oTable0"]/tbody/tr[2]/td[1]/div/div/button').click()
            # iframe = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//*[@class='ec-iframe']")))
            # driver.switch_to.frame(iframe)
            # # 点击"发票录入"按钮
            # WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//*[@class='el-button-group']/button[1]"))).click()
            # # 上传发票文件
            # driver.find_element(By.XPATH, "//*[@class='c-csifr-tip-content']/input").send_keys(invoice_path)
            # # 点击"开始识别"按钮
            # driver.find_element(By.XPATH, "//*[@class='c-ccsi-footer']/button[2]").click()
            # time.sleep(3)
            # WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//*[@class='c-ccsi-footer']/button[4]")))
            # # 点击"确认"按钮
            # driver.find_element(By.XPATH, "//*[@class='c-ccsi-footer']/button[2]").click()
            # # 等待"发票录入"按钮出现
            # WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//*[@class='el-button-group']/button[1]")))
            # time.sleep(3)
            # # 点击"确认"按钮
            # driver.find_element(By.XPATH, "//*[@class='c-ccsi-footer']/button[2]").click()
            # # 等待"确认"按钮消失
            # # WebDriverWait(driver, 60).until(EC.invisibility_of_element_located((By.XPATH, "//*[@class='c-ccsi-footer']/button[2]")))
            # time.sleep(3)
            # # 切换回主文档
            # driver.switch_to.default_content()
            # time.sleep(1)
            # WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="oTable0"]/tbody/tr[2]/td[1]/div/div/button')))

            # 附件
            # 派单记录表
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[48]/td[5]/div/div/span/div/div[2]/span[1]/span/div/input').send_keys(general_path)
            # 完税证明
            file_path = os.path.join(config_dict['纳税文件'], service_provider_name)
            file_name = sorted(os.listdir(file_path), reverse=True)
            if not file_name:
                raise FileNotFoundError(f"文件路径不存在:{file_path}")
            file_name = os.path.join(file_path, file_name[0])
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[48]/td[5]/div/div/span/div/div[2]/span[1]/span/div/input').send_keys(file_name)
            # ASP框架服务合同
            file_name = []
            for name in os.listdir(config_dict['ASP框架服务合同']):
                if f'ASP框架协议-{service_provider_name}' in name and '脱敏版' not in name and '.pdf' in name:
                    file_name.append(name)
            if not file_name:
                raise FileNotFoundError(f"文件路径不存在:{config_dict['ASP框架服务合同']}\\ASP框架协议-{service_provider_name}")
            file_name = sorted(file_name, reverse=True)[0]
            file_path = os.path.join(config_dict['ASP框架服务合同'], file_name)
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[48]/td[5]/div/div/span/div/div[2]/span[1]/span/div/input').send_keys(file_path)
            # 验收文件
            file_name = f"{service_provider_name}\{config_dict['年份']}\现场服务单\{config_dict['数字月份']}\{config_dict['月份']}-{config_amount_name}-现场服务报告.zip"
            # file_name = f"{config_dict['月份']}-{config_amount_name}-{config_dict[config_amount_name]}.zip"
            file_name = os.path.join(config_dict['验收文件保存路径'], file_name)
            if not os.path.exists(file_name):
                raise FileNotFoundError(f"文件路径不存在:{file_name}")
            driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[48]/td[5]/div/div/span/div/div[2]/span[1]/span/div/input').send_keys(file_name)

            # # 审批信息
            # # 搜索部门预审
            # driver.find_element(By.XPATH, "//*[@id='field353752span']/div[2]/button").click()
            # WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//*[@class='wea-hr-muti-input-left']/div[1]/div[1]/div/span/input"))).send_keys(config_dict['部门预审'])
            # but_address = "//*[@class='wea-hr-muti-input-left']/div[1]/div[1]/div/button"
            # tr_address = "//*[@class='wea-hr-muti-input-left']/div[3]/div/div[1]/div/ul/li"
            # driver.find_element(By.XPATH, but_address).click()
            # search_basic_infor(driver, but_address, tr_address)
            # # 确认部门预审
            # time.sleep(1)
            # driver.find_element(By.XPATH, "//*[@class='wea-transfer-opration']/div/button[3]").click()
            # time.sleep(1)
            # driver.find_element(By.XPATH, '/html/body/div[14]/div/div[2]/div/div[1]/div[3]/button[1]').click()
            #
            # # 部门一级审批
            # driver.find_element(By.XPATH, '//*[@id="field353753_sel"]/div/div/div/div/span').click()
            # time.sleep(0.5)
            # element = driver.find_element(By.XPATH, '/html/body/div[15]/div/div/div/ul')
            # element.find_element(By.XPATH, f".//li[contains(.,'{config_dict['部门一级审批']}')]").click()
            #
            # # 部门终审
            # driver.find_element(By.XPATH, '//*[@id="field353757_sel"]/div/div/div/div/span').click()
            # time.sleep(0.5)
            # element = driver.find_element(By.XPATH, '/html/body/div[16]/div/div/div/ul')
            # element.find_element(By.XPATH, f".//li[contains(.,'{config_dict['部门终审']}')]").click()
            #
            # # 业务单元一级加签
            # driver.find_element(By.XPATH, '//*[@id="field353758_sel"]/div/div/div/div/span').click()
            # time.sleep(0.5)
            # element = driver.find_element(By.XPATH, '/html/body/div[17]/div/div/div/ul')
            # element.find_element(By.XPATH, f".//li[contains(.,'{config_dict['业务单元一级加签']}')]").click()

            # 等待上传附件完成
            WebDriverWait(driver, 600).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[48]/td[5]/div/div/span/div/div[1]/div[4]')))

            # 点击保存
            driver.find_element(By.XPATH, '//*[@class="wea-new-top-req-wapper "]/div[1]/div/div[3]/div/div[2]/div/span[2]/button').click()
            time.sleep(3)
            # 关闭新建标签页，切换到财务报销系统标签页
            driver.close()
            driver.switch_to.window(reimburse_handels)
            break
        except Exception as e:
            if '文件路径不存在' in str(e):
                raise Exception(e)
            if index == 2:
                raise Exception(f"技服外包报销失败：{e}")
            driver.refresh()
            continue
