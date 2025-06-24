# Selenium环境安装
'''
    # 1:浏览器安装(谷歌Chrome、火狐Firefox、微软Edge、苹苹果Safari等)
    # Chrome官方下载地址:https://www.google.cn/chrome/
    # 2:浏览器驱动安装(ChromeDriver、GeckoDriver、msedggedriver等
    # 官方最新驱动下载地址:
    # https://googlechromelabs.github.io/chrome-for-testing/
    # 注意:驱动版本号要和浏览器版本号对应符合(至少大版本对应),否则失效。

    # 关闭自动更新:(防止更新导致驱动失效)
    # 开始内搜索services.msc找到 Google更新组件全部禁用
    # 安装三方库:selenium
'''

# 设置浏览器
'''
    # 禁用沙盒模式:add_argument('--no-sandbox')
    # 保持浏览器打开状态:add_experimental_option('detach',Truue)
    # 创建并启动浏览器:webdriver.Chrome()
    
    # 导包:from selenium import webdriver
    # 导包:from selenium.webdriver.chrome.options import Options
    # 导包:from selenium.webdriver.chrome.service import Service
'''

#定位一个元素
#定位多个元素
'''
    浏览器查找多个元素:document.getElementById('元素值')
    元素定位导包:from selenium.webdriver.common.by import By
'''

import os
import sys
from selenium import webdriver  # 用于操作浏览器
from selenium.webdriver.chrome.options import Options   # 用于设置谷歌浏览器参数
from selenium.webdriver.chrome.service import Service   # 用于设置谷歌驱动路径
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import ActionChains


def get_chromedriver_path():
    # 开发环境路径
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    driver_path = os.path.join(base_path, "chromedriver_v137.exe")
    # 打包后路径检测
    if not os.path.exists(driver_path):
        # 尝试上级目录
        driver_path = os.path.join(base_path, "../driver/chromedriver_v137.exe")
    return driver_path

def create_browser():
    # 创建设置浏览器对象
    q1 = Options()

    q1.add_argument('--no-sandbox')     # 禁用沙盒模式(增加兼容性)
    q1.add_argument('--start-maximized')     # 最大化浏览器窗口
    # 保持浏览器打开状态(默认是代码执行完毕自动关闭)
    q1.add_experimental_option('detach', True)

    # 创建并启动浏览器
    # a1 = webdriver.Chrome(service=Service('chromedriver_v137.exe'), options=q1)
    driver_path = get_chromedriver_path()
    a1 = webdriver.Chrome(service=Service(driver_path), options=q1)
    # 元素定位隐性等待(多少秒内找到元素就立刻执行，超时就报错)
    a1.implicitly_wait(10)  # 设置类型的语句，设置一次就可以让所有的a1都使用这个等待时间
    return a1


# a1 = create_browser()
# driver = create_browser()

# 打开指定网址
# a1.get('https://baidu.com/')
# a1.get('https://www.bilibili.com/')

# time.sleep(2)

# 关闭当前标签页
# a1.close()

# 退出浏览器并释放驱动
# a1.quit()

# # 浏览器最大化
# a1.maximize_window()
# time.sleep(2)
#
# # 浏览器最小化
# a1.minimize_window()
# time.sleep(2)

# # 浏览器打开位置
# a1.set_window_position(0, 0)
# # 浏览器打开尺寸
# a1.set_window_size(600, 600)

# # 浏览器截图
# a1.get_screenshot_as_file('1.png')  # 参数是图片保存路径
# time.sleep(3)

# # 刷新当前网页
# a1.refresh()


# 定位一个元素 (找到的话返回结果,找不到的话报错)
# a2 = a1.find_element(By.ID, 'kw')   # 查找到百度的输入框
# print(a2)

# # 定位多个元素 (找到的话返回结果列表,找不到的话返回空列表)
# a2 = a1.find_elements(By.ID, 'kw')
# print(a2)
# '''
#     查找多个元素：
#     1、使用python代码，find_elements()方法查找多个元素，返回结果为列表。
#     2、开发者工具-控制台 document.getElementById()方法查找。
# '''

# # 元素的输入
# a2.send_keys('python')
#
# # # 元素清空
# # a2.clear()
#
# a2 = a1.find_element(By.ID, 'su')   # 查找到百度的搜索按钮
# # 元素点击
# a2.click()


# # 元素定位-ID
# # 1、通过ID定位元素,一般比较准确。
# # 2、并不是所有网页或者元素都有ID值
# a1.find_element(By.ID, 'kw').send_keys('python')
# a1.find_element(By.ID, 'su').click()

# 元素定位-NAME
# 1、通过NAME定位元素,一般比较准确。
# 2、并不是所有网页或者元素都有NAME值
# a1.find_element(By.NAME, 'wd').send_keys('python')

# 元素定位-CLASS_NAME
# 1、class值不能有空格,否则报错
# 2、class值重复的有很多,需要切片
# 3、class值有的网站是随机的
# a1.find_elements(By.CLASS_NAME, 'channel-icons_item')[1].click()

# 元素定位-LINK_TEXT
# 1、通过精准链接文本找到标签a的元素
# 2、有重复的文本,需要切片
# a1.find_element(By.LINK_TEXT, '音乐').click()

# 元素定位-PARTIAL_LINK_TEXT
# 1、通过模糊链接文本找到标签a的元素[模糊文本定位]
# 2、有重复的文本,需要切片
# a1.find_element(By.PARTIAL_LINK_TEXT, '音').click()

# 元素定位-CSSSELECTOR
# 1,#id=井号+id值通过id定位
# 2, class=点+class值通过class定位
# 3,不加修饰符=标签头通过标签头定位
# 4,通过任意类型定位:"[类型="精准值']"
# 5,通过任意类型定位:"[类型*='模糊值']"
# 6,通过任意类型定位:"[类型^='开头值']"
# 7,通过任意类型定位:"[类型$'结尾值']"
# 以上这些方法都属于理论定位法
# 8,更简单的定位方式:在谷歌控制台直接复制 SELECTOR
# a1.find_element(By.CSS_SELECTOR, '#kw').send_keys('python')

# 元素定位-XPATH
# 1,复制谷歌浏览器 Xpath (通过属性+路径定位,属性如果是随机的,可能定位不到)
# 2,复制谷歌浏览器 Xpath完整路径 (缺点是定位值比较长,优点是基本100%准确)
# a1.find_element(By.XPATH, '//*[@id="input"]"]').send_keys('python')
# a1.find_element(By.XPATH, '/html/body/ntp-app//div/div[2]/cr-searchbox//div/input"]').send_keys('python')

#获取全部标签页句柄
# a2 = a1.window_handles
# print(a2)
#
# #通过句柄切换标签页
# a1.switch_to.window(a2[-1])
#
# #获取当前标签页句柄
# a2 = a1.current_window_handle
# print(a1.current_window_handle)
# # 警告框(alert)元素交互（只有一个确定按钮的弹框）
# # 获取弹窗内的文本内容
# print(a1.switch_to.alert.text)
# # 点击弹窗确定按钮
# a1. switch_to.alert.accept()

# # 确认框(confirm)元素交互（有确定和取消两个按钮的弹框）
# # 点击弹窗确定按钮
# a1.switch_to.alert.accept()
# # 点击弹窗取消按钮
# a1.switch_to.alert.dismiss()

# 网页后退
# a1.back()
#
# # 网页前进
# a1.forward()

# 判断元素是否存在
# is_email_visible = a1.find_element(By.NAME, "email_input").is_displayed()
# print(is_email_visible)


def login_oa():
    # 登录OA，跳转到财务报销系统
    driver = create_browser()
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

driver = login_oa()

reimburse_handels = driver.current_window_handle    # 财务报销系统 标签页的句柄
driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[1]/ul/li[8]/div/div/div').click()
driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[1]/ul/li[8]/ul/li[11]/div/div').click()

time.sleep(3)
driver.switch_to.window(driver.window_handles[-1])

# todo 填写基本信息
# element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[5]/td[5]/div/div/span[1]/div/div/div')
# ActionChains(driver).move_to_element_with_offset(element, 10, 0).click().perform()
# element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[5]/td[5]/div/div/span[1]/div/div/div').click()
# driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[5]/td[5]/div/div/span[1]/div/div/div').send_keys('MHQH0003')
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[5]/td[5]/div/div/span[1]/div/div/div/div[2]/button')))
time.sleep(15)
# 填写基本信息
# 申请人
try:
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[5]/td[5]/div/div/span[1]/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/span[1]/input').send_keys('chenyqr')
    driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div[1]/div/div[1]/div/ul/li/div[1]'))).click()
except:
    driver.find_element(By.XPATH, '/html/body/div[7]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/span[1]/input').send_keys('chenyqr')
    driver.find_element(By.XPATH, '/html/body/div[7]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div[1]/div/div[1]/div/ul/li/div[1]'))).click()
time.sleep(1)
# 费用成本中心
try:
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[6]/td[5]/div/div/span[1]/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, '/html/body/div[5]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input').send_keys('MHQH0003')
    driver.find_element(By.XPATH, '/html/body/div[5]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/span/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/div[2]/div/div[1]/div[2]/div/div[3]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
except:
    driver.find_element(By.XPATH, '/html/body/div[8]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input').send_keys('MHQH0003')
    driver.find_element(By.XPATH, '/html/body/div[8]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/span/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div/div[2]/div/div[1]/div[2]/div/div[3]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
time.sleep(1)
# 平台
try:
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[7]/td[5]/div/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, '/html/body/div[6]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/span[1]/input').send_keys('北京')
    driver.find_element(By.XPATH, '/html/body/div[6]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
except:
    driver.find_element(By.XPATH, '/html/body/div[9]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/span[1]/input').send_keys('北京')
    driver.find_element(By.XPATH, '/html/body/div[9]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[9]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
time.sleep(1)
# 费用是由
try:
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[8]/td[5]/div/div/span[1]/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, '/html/body/div[7]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input').send_keys('231090301')
    driver.find_element(By.XPATH, '/html/body/div[7]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/span/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[7]/div/div[2]/div/div[1]/div[2]/div/div[3]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
except:
    driver.find_element(By.XPATH, '/html/body/div[10]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div/div[1]/form/div/div/div[1]/div[2]/div/div/div/span/input').send_keys('231090301')
    driver.find_element(By.XPATH, '/html/body/div[10]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/span/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/div/div[2]/div/div[1]/div[2]/div/div[3]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
time.sleep(1)
# 结算方式
try:
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[1]/div/div/div[2]/div[1]/div/table/tbody/tr[9]/td[5]/div/div/div/div/div[2]/button').click()
    driver.find_element(By.XPATH, '/html/body/div[8]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/span[1]/input').send_keys('电汇')
    driver.find_element(By.XPATH, '/html/body/div[8]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[8]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
except:
    driver.find_element(By.XPATH, '/html/body/div[11]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/span[1]/input').send_keys('电汇')
    driver.find_element(By.XPATH, '/html/body/div[11]/div/div[2]/div/div[1]/div[2]/div/div[1]/div[1]/div[2]/div/div/div/button').click()
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[11]/div/div[2]/div/div[1]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div/div/div/span/div[2]/table/tbody/tr'))).click()
time.sleep(1)

print("aaaa")
