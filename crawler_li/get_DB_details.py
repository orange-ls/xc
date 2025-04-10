'''
模拟登录，获取待办事项表格HTML内容，并解析出每个待办事项的详细信息。
'''
import requests
from bs4 import BeautifulSoup
import re
import json


class WorkflowCrawler:
    def __init__(self, base_url):
        self.base_url = base_url
        self.session = requests.Session()
        # 配置浏览器级请求头
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Accept-Encoding": "gzip, deflate"
        })

    def login(self, username, password):
        """处理完整登录流程"""
        try:
            # Step 1: 获取登录页面参数
            login_page = self.session.get(
                f"{self.base_url}/Login.aspx",
                verify=False,
                timeout=10
            )
            login_page.raise_for_status()

            # 提取ASP.NET隐藏字段
            soup = BeautifulSoup(login_page.text, 'html.parser')
            login_data = {
                '__VIEWSTATE': soup.find('input', {'id': '__VIEWSTATE'})['value'],
                '__EVENTVALIDATION': soup.find('input', {'id': '__EVENTVALIDATION'})['value'],
                '__VIEWSTATEGENERATOR': soup.find('input', {'id': '__VIEWSTATEGENERATOR'})['value'],
                'txtName': username,
                'txtPsw': password,
                'imgLogin.x': 23,
                'imgLogin.y': 18
            }

            # Step 2: 提交登录请求（禁止自动重定向）
            login_response = self.session.post(
                f"{self.base_url}/Login.aspx",
                data=login_data,
                allow_redirects=False,
                verify=False,
                timeout=15
            )

            # Step 3: 处理302重定向链
            if login_response.status_code == 302:
                return self.follow_redirect_chain(login_response.headers['Location'])
            return False

        except Exception as e:
            print(f"登录过程异常: {str(e)}")
            return False

    def follow_redirect_chain(self, initial_location):
        """处理多级重定向"""
        # 从login界面重定向到MainPage，然后再跳转到MainPageInfos,最后到达PendingWorkItem.aspx
        try:
            current_url = f"{self.base_url}{initial_location}"

            # 最多处理5次重定向
            for _ in range(5):
                response = self.session.get(
                    current_url,
                    verify=False,
                    timeout=10,
                    allow_redirects=False
                )

                # 处理特殊中间页面
                if "MainPage.aspx" in current_url:
                    response = self.session.get(
                        f"{self.base_url}/Pages/MainPageInfos.aspx",
                        verify=False,
                        timeout=10,
                        allow_redirects=False
                    )

                # 正常重定向处理,将重定向到的URL加入待处理队列
                if response.status_code in [301, 302, 303, 307, 308]:
                    new_location = response.headers.get('Location')
                    if not new_location.startswith('http'):
                        new_location = f"{self.base_url}{new_location}"
                    current_url = new_location
                    print(f"重定向到: {current_url}")
                else:
                    break

            response.raise_for_status()

            # 最终验证是否到达目标页面
            if "PendingWorkItem.aspx" in response.url:
                print("成功到达目标页面")
                html = response.text
                return html
            print(f"最终停留在: {response.url}")
            return False

        except Exception as e:
            print(f"重定向处理异常: {str(e)}")
            return False

    def parse_work_table(self, html):
        """解析待办事项表格"""
        try:
            soup = BeautifulSoup(html, 'lxml')

            # 精确匹配表格（根据提供的HTML结构）
            headers_table = soup.find('table', {'class': 'rgMasterTable'})
            body_table = soup.find('table', {'id': 'MainContent_RadGrid1_ctl00'})

            # 提取表头
            headers = [th.get_text(strip=True) for th in headers_table.find_all('th', class_='rgHeader')]

            # 确保表头数量正确
            if len(headers) != 5:
                print(f"表头数量异常，预期5列，实际{len(headers)}列")
                return None

            # 提取数据行
            data = []
            for row in body_table.select('tbody tr[class^="rg"]'):  # 匹配rgRow/rgAltRow
                cells = []
                for idx, td in enumerate(row.find_all('td')):
                    # 特殊处理第一列的链接信息
                    if idx == 0:
                        link = td.find('a')
                        if link:
                            # 提取链接文本和URL
                            cell_data = {
                                'title': link.get('title', ''),
                                'href': link.get('href', '').replace('&amp;', '&'),
                                # 'text': link.get_text(strip=True)
                            }
                        else:
                            cell_data = {'title': '', 'href': '', 'text': ''}
                    else:
                        cell_data = td.get_text(strip=True)

                    cells.append(cell_data)

                # 验证数据列数
                if len(cells) == len(headers):
                    data.append(cells)
                else:
                    print(f"数据列数不匹配，跳过该行。预期{len(headers)}列，实际{len(cells)}列")

            return {
                "headers": headers,
                "data": data,
                # "table_info": {
                #     "total_rows": len(data),
                #     "first_row_date": data[0][2] if data else None,
                #     "last_row_date": data[-1][2] if data else None
                # }
            }

        except Exception as e:
            print(f"表格解析异常: {str(e)}")
            return None

    def parse_todo_list(self, todo_url):
        """进入每一条待办事宜的详情页，解析详情页的待办事项列表"""
        try:
            # # 拼接待办事项详情页URL
            todo_url = f"http://10.1.126.177:8992/db-contract-application/db-api/contract/detail/{todo_url}"
            # 从Cookies中获取odbtoken
            odbtoken = self.session.cookies.get("XBSystemUserLoginName", "")
            # 配置请求头
            headers = {
                "odbtoken": odbtoken,
                "Accept": "application/json, text/plain, */*",
                "Referer": f"{self.base_url}/Pages/PendingWorkItem.aspx"
            }

            response = self.session.get(
                todo_url,
                headers=headers,
                verify=False,
                timeout=10
            )
            if todo_url != response.url:
                return {}

            # 解析json数据
            html_json = response.text
            html_dict = json.loads(html_json)
            if html_dict.get('message') == 'SUCCESS':
                data = html_dict.get('data', {})
                return data
        except Exception as e:
            print(f"详情解析异常: {str(e)}")
            return None


if __name__ == "__main__":
    # 配置参数
    BASE_URL = "http://10.1.126.177:88"
    USERNAME = "xinchao"
    PASSWORD = "1"

    crawler = WorkflowCrawler(BASE_URL)

    html = crawler.login(USERNAME, PASSWORD)

    if html:
        print("\n正在抓取待办事项...")
        result = crawler.parse_work_table(html)

        if result:
            print("\n抓取“待办事宜”成功！")
            for row in result['data']:
                recordid = None
                details = {}
                record_list = row[0].get('href', '').split('&')
                for record in record_list:
                    if 'recordid' in record:
                        recordid = record.split('=')[1]
                        break
                if recordid:
                    # 获取详情页数据
                    details = crawler.parse_todo_list(recordid)
                row.append(details)
            print(result)
        else:
            print("数据抓取失败")
    else:
        print("登录失败，程序终止")

    # 验证最终状态
    print("\n当前Cookies:", crawler.session.cookies.get_dict())