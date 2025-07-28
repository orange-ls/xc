from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import zipfile
import os


def login_crm(driver, config_dict):
    # 登录crm，跳转到工单搜索界面
    driver.get('https://www.fxiaoke.com/proj/page/login')

    # 登录
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div/div[1]/ul/li[2]'))).click()
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/div/div[1]/input').send_keys(config_dict['CRM账号'])
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/div/div[2]/input').send_keys(config_dict['CRM密码'])
    # 勾选同意
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div/div[2]/span[1]').click()
    # 点击登录
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/div/div[6]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div[2]/div/div[3]/ul/li[3]').click()

    time.sleep(2)

    # 关闭弹窗
    for i in range(3):
        try:
            rows = driver.execute_script('return document.getElementsByClassName("el-dialog__body");')
            if len(rows) > 0:
                time.sleep(1)
                # WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="new-ui-guide-dialog-header"/span'))).click()
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="el-dialog__body"]/div[1]/span]'))).click()
            rows = driver.execute_script('return document.getElementsByClassName("el-message-box");')
            if len(rows) > 0:
                time.sleep(1)
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="el-message-box"]/div[1]/button'))).click()
        except:
            time.sleep(1)
            driver.refresh()
            continue

    # 点击服务通、工单管理、服务报告
    for i in range(3):
        try:
            # WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app-portal"]/header/div/div[1]/div[2]/div[1]/div/ul/li[2]'))).click()
            # driver.find_element(By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[2]/div[1]/div/div/div/div[3]/div/div/div[1]/ul[1]/li[6]').click()
            # WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sub-tpl"]/div/div[2]/div[1]/div/div/div/div[3]/div/div/div[1]/ul[1]/li[6]/div/div/ul/li[4]'))).click()
            # WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sub-tpl"]/div/div[2]/div[2]/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div[5]/div/div/div[3]/form/div/input')))

            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app-portal"]/header/div/div[1]/div/div[1]/div/ul/li[2]'))).click()
            driver.find_element(By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[1]/div/div[1]/div/div[2]/div[3]/div/div/div[1]/ul[1]/li[6]').click()
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[1]/div/div[1]/div/div[2]/div[3]/div/div/div[1]/ul[1]/li[6]/div/div/ul/li[4]'))).click()
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div/div[5]/div/div/div[3]/form/div/input')))
        except:
            time.sleep(1)
            driver.refresh()
            continue
        break
    return driver


def crm_download_file(driver, order_num_list, source_dir, output_dir, file_name):
    # 处理文件路径
    output_dir = os.path.normpath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # 清除下载路径中的文件
    clean_source_file(source_dir)

    # 开始按工单号搜索
    # 将搜索字段修改为“工单”
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div/div[5]/div/div/div[1]/div/div/div[3]'))).click()
    element = driver.find_element(By.CSS_SELECTOR, '.crm-w-select.crm-widget.crm-w-panel.bl')
    element.find_element(By.XPATH, 'div/ul/li[4]').click()


    order_num_exist = []    # 保存已经下载文件的工单
    for order_num in order_num_list:
        for i in range(5):
            # 输入工单号，点击搜索
            driver.find_element(By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div/div[5]/div/div/div[3]/form/div/input').clear()
            driver.find_element(By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div/div[5]/div/div/div[3]/form/div/input').send_keys(order_num)
            driver.find_element(By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div[2]/div/div[5]/div/div/span').click()

            # 等待搜索结果
            for i in range(10):
                is_exist = driver.find_element(By.CSS_SELECTOR, '.dt-loading.b-g-hide.lg').is_displayed()
                if not is_exist:
                    break
                time.sleep(1)
            # 获取搜索结果
            rows = driver.find_elements(By.XPATH, '//*[@id="sub-tpl"]/div[3]/div[2]/div[2]/div/div/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div[5]/div[3]/div[2]/table/tbody/tr')
            row_text = rows[0].text.strip()
            if not row_text:
                # print(f'跳过工单号：{order_num}')
                # continue
                break
            for row in rows:
                name = row.find_element(By.XPATH, 'td[2]').text
                if '02_' in name or '05_' in name or '设备健康检查' in name:
                    row.find_element(By.XPATH, 'td[4]/div/div/div/div/a').click()
                    order_num_exist.append(order_num)
                    break
            break

    time.sleep(3)
    # 判断目录source_dir中是否存在目标工单号开头的文件
    for i in range(10):
        filenames = os.listdir(source_dir)  # 单次目录读取
        missing_order = next((order for order in order_num_exist if not any(f.startswith(order) for f in filenames)), None)
        if not missing_order:
            break
        time.sleep(5)

    # 下载完成后，进入保存文件的文件夹，将文件打包成压缩包
    create_zip(source_dir, output_dir, file_name)
    print('done')


def create_zip(source_dir, output_dir, file_name):
    """
    智能创建ZIP压缩包
    :param source_dir: 需要压缩的源目录路径
    :param output_dir: 压缩文件输出目录路径
    :param file_name: 压缩文件名
    """
    # 生成压缩文件名
    zip_name = os.path.join(output_dir, f"{file_name}.zip")
    total_size = 0

    for i in range(3):
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = os.path.join(root, file)

                    # 计算相对路径
                    arc_path = os.path.relpath(file_path, start=source_dir)
                    # 添加文件到压缩包
                    zipf.write(file_path, arc_path)

                    # 更新总大小
                    total_size += os.path.getsize(file_path)
        if total_size <= 1:
            continue
        break

    total_size = round(total_size / (1024 * 1024), 2)
    if total_size >= 50:
        print(f"超过50MB，压缩文件大小：{total_size}MB")

    clean_source_file(source_dir)


def clean_source_file(source_dir):
    # 删除源目录中的文件
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            file_path = os.path.join(root, file)
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"删除失败：{file_path} - {str(e)}")