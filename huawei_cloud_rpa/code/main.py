import tkinter as tk
# from tkinter.filedialog import askdirectory
from tkinter import filedialog
import threading
import pandas as pd
from sqlalchemy import create_engine, text
# import DI
from tkinter import messagebox
# import table_handling
# from time import sleep
import os
# import utils

class App(object):
    def __init__(self, root):
        self.filePath = {}

        root.title("华为云RPA")
        root.geometry('500x500')

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

        self.two_four_data = tk.StringVar()
        label01 = tk.Label(root, text="24年数据：")
        label01.grid(row=5, column=0)
        entry01 = tk.Entry(root, textvariable=self.two_four_data, width=40)
        entry01.grid(row=5, column=1)
        btn01 = tk.Button(root, text="选择", command=lambda: self.selectPath(self.two_four_data))
        btn01.grid(row=5, column=2)

        btn007 = tk.Button(root, text="导入24年数据", command=self.import_24_data)
        btn007.grid(row=7, column=0)

        btn005 = tk.Button(root, text="匹配", command=self.start)
        btn005.grid(row=7, column=2)

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
            connect_db = self.connect_db()
            if not connect_db:
                return
            self.common_excel_to_db(connect_db, product_details_path, customer_cor_path)




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

    # 导入24年数据
    def import_24_data(self):

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

    # 配置数据库连接
    def connect_db(self):
        try:
            # 数据库配置
            DB_HOST = 'localhost'
            DB_PORT = 3306
            DB_USER = 'root'
            DB_PASS = '1234'
            DB_NAME = 'test_sync'

            # 创建数据库连接
            engine = create_engine(f'mysql+pymysql://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}')
            return engine
        except Exception as e:
            self.text.insert(tk.END, f"数据库连接失败: {str(e)}\r\n")
            return False

    # 客户关系表和2025年产品明细表入库
    def common_excel_to_db(self, engine, product_details_path, customer_cor_path):
        # 先把客户关系表和2025年产品明细表入库
        try:
            # 读取Excel文件
            df_customer_cor = pd.read_excel(customer_cor_path).rename(columns={'客户名称': 'customer_name', '销售员': 'salesperson', '区域': 'region'})

            # 筛选出2025产品明细 需要入库的字段
            cloud_services = pd.read_excel(product_details_path, sheet_name='云服务名称')[['云服务编码', '服务产品部']].rename(columns={'云服务编码': 'cloud_services_code', '服务产品部': 'service_department'})
            details_flow = pd.read_excel(product_details_path, sheet_name='流量产品清单')[['产品类型编码', '产品类型']].rename(columns={'产品类型编码': 'product_code', '产品类型': 'product_type'})
            details_special = pd.read_excel(product_details_path, sheet_name='2025年产品专项', header=1)[['L4层产品编码', '名称']].rename(columns={'L4层产品编码': 'product_code', '名称': 'product_name'})
            details_collaborate = pd.read_excel(product_details_path, sheet_name='企业协同')[['云服务编码', '云服务名称']].rename(columns={'云服务编码': 'cloud_services_code', '云服务名称': 'cloud_services_name'})

            # # 读取24年数据
            # df_two_four = pd.read_excel(two_four_path)[['业绩ID', '业绩金额(¥)', '业绩形成时间', '二级经销商名称', '客户名称', '产品类型编码', '客户标签', '销售纵队', '服务产品部', '是否流量型产品', '专线产品', '企业协同', '销售员', '区域', '季度']]
            # df_two_four = df_two_four.rename(columns={'业绩ID': 'performance_id', '业绩金额(¥)': 'sales_amount', '业绩形成时间': 'performance_date', '二级经销商名称': 'secondary_dealer', '客户名称': 'customer_name', '产品类型编码': 'product_code', '客户标签': 'customer_tag', '销售纵队': 'sales_team', '服务产品部': 'service_department', '是否流量型产品': 'is_traffic_product', '专线产品': 'leased_line_product', '企业协同': 'enterprise_coop', '销售员': 'salesperson', '区域': 'region', '季度': 'quarter'})
            # 写入数据库
            with engine.begin() as conn:  # 自动提交/回滚
                # 阶段1：清空旧数据
                conn.execute(text("DELETE FROM customer_correspondence"))
                conn.execute(text("DELETE FROM two_five_details_cloud_services"))
                conn.execute(text("DELETE FROM two_five_details_flow"))
                conn.execute(text("DELETE FROM two_five_details_special"))
                conn.execute(text("DELETE FROM two_five_details_collaborate"))

                # 阶段2：插入新数据.'append'表示在现有表中追加数据（保留原有数据）、False表示不写入索引、'multi'表示使用多行组合的批量插入语法、批量插入1000条
                df_customer_cor.to_sql(name='customer_correspondence', con=conn, if_exists='append', index=False, method='multi', chunksize=1000)
                cloud_services.to_sql(name='two_five_details_cloud_services', con=conn, if_exists='append', index=False, method='multi', chunksize=1000)
                details_flow.to_sql(name='two_five_details_flow', con=conn, if_exists='append', index=False, method='multi', chunksize=1000)
                details_special.to_sql(name='two_five_details_special', con=conn, if_exists='append', index=False, method='multi', chunksize=1000)
                details_collaborate.to_sql(name='two_five_details_collaborate', con=conn, if_exists='append', index=False, method='multi', chunksize=1000)

            # 写入24年数据
            # for _, row in df_two_four.iterrows():
            #     # 使用INSERT ... ON DUPLICATE KEY UPDATE语法
            #     stmt = text("""
            #             INSERT INTO hw_two_four_data
            #             (performance_id, sales_amount, performance_date, secondary_dealer, customer_name, product_code, customer_tag, sales_team, service_department, is_traffic_product, leased_line_product, enterprise_coop, salesperson, region, quarter)
            #             VALUES (:performance_id, :sales_amount, :performance_date, :secondary_dealer, :customer_name, :product_code, :customer_tag, :sales_team, :service_department, :is_traffic_product, :leased_line_product, :enterprise_coop, :salesperson, :region, :quarter)
            #             ON DUPLICATE KEY UPDATE
            #             performance_id = VALUES(performance_id)
            #         """)
            #     conn.execute(stmt, row.to_dict())


        except Exception as e:
            self.text.insert(tk.END, f"写入表时发生错误: {str(e)}\r\n")
            return False


if __name__ == '__main__':
    root = tk.Tk()

    App(root)
    root.mainloop()
