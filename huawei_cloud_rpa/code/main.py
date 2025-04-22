import gc
import tkinter as tk

from tkinter import filedialog
import threading

import numpy as np
import pandas as pd
from sqlalchemy import create_engine, text
import result_table_processing
import os
from python_calamine import CalamineWorkbook

# 字段映射配置
column_mapping = {
    '业绩ID': 'performance_id',
    '业绩金额(¥)': 'sales_amount',
    '业绩形成时间': 'performance_date',
    '特殊返点类型': 'special_rebate_type',
    '二级经销商名称': 'secondary_dealer',
    '客户名称': 'customer_name',
    '产品类型编码': 'product_code',
    '客户标签': 'customer_tag',
    '销售纵队': 'sales_team',
    '服务产品部': 'service_department',
    '是否流量型产品': 'is_traffic_product',
    '专线产品': 'leased_line_product',
    '企业协同': 'enterprise_coop',
    '销售员': 'salesperson',
    '区域': 'region',
    '季度': 'quarter'
}
selected_columns = list(column_mapping.keys())

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

        btn007 = tk.Button(root, text="导入24年数据", command=self.import_24_data_to_db)
        btn007.grid(row=7, column=0)

        btn005 = tk.Button(root, text="匹配", command=self.start)
        btn005.grid(row=7, column=2)

        # 结果打印框
        self.text = tk.Text(selectbackground="red", insertbackground="blue", spacing2=10, bd=0)
        self.text.grid(row=9, column=0, columnspan=10)

        self.text.insert(tk.END, "24年数据导入一次即可，不用重复导入！\r\n")
        self.text.insert(tk.END, "25年数据增量导入，全量导入会导致执行时间过长\r\n")
        self.text.insert(tk.END, "2025产品明细、客户对应关系表、数据需求表 必须选择\r\n\r\n")

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
            product_details_path = self.product_details.get()
            customer_cor_path = self.customer_cor.get()
            data_requirements_path = self.data_requirements.get()
            if not product_details_path or not customer_cor_path or not data_requirements_path:
                self.text.insert(tk.END, "文件路径不能为空！\r\n")
                return
            self.text.insert(tk.END, "检查客户对应关系表是否完整...\r\n")
            # 检查客户对应关系表是否完整
            flag = self.check_customer_cor(customer_cor_path)
            if not flag:
                return

            self.text.insert(tk.END, "数据解析中...\r\n")
            engine = self.connect_db()
            if not engine:
                return
            # 基础表 数据入库
            self.common_excel_to_db(engine, product_details_path, customer_cor_path)

            self.text.insert(tk.END, "25年业绩表增加BI到BO列...\r\n")
            # 25年业绩表增加BI到BO列
            if two_five_path:
                self.add_bi_to_bo(engine, two_five_path)

            self.text.insert(tk.END, "数据第一部分...\r\n")
            result_one = result_table_processing.result_table_one(engine)
            self.text.insert(tk.END, "数据第二部分...\r\n")
            result_two = result_table_processing.result_table_two(engine)
            self.text.insert(tk.END, "数据第三部分...\r\n")
            result_three = result_table_processing.result_table_three(engine)
            self.text.insert(tk.END, "数据第四部分...\r\n")
            result_four = result_table_processing.result_table_four(engine)
            self.text.insert(tk.END, "数据第五部分...\r\n")
            result_five = result_table_processing.result_table_five(engine)
            self.text.insert(tk.END, "数据第六部分...\r\n")
            result_six = result_table_processing.result_table_six(engine)
            self.text.insert(tk.END, "数据第七部分...\r\n")
            result_seven = result_table_processing.result_table_seven(engine)
            self.text.insert(tk.END, "数据第八部分...\r\n")
            result_eight = result_table_processing.result_table_eight(engine)
            self.text.insert(tk.END, "数据第九部分...\r\n")
            result_nine = result_table_processing.result_table_nine(engine)
            self.text.insert(tk.END, "结果表下载中...\r\n")

            self.text.insert(tk.END, "处理完成！\r\n")
            engine.dispose()
        except BaseException as e:
            self.text.insert(tk.END, "发生错误！\r\n")
            self.text.insert(tk.END, e)

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

        except Exception as e:
            self.text.insert(tk.END, f"写入表时发生错误: {str(e)}\r\n")
            return False

    # 数据入库方法
    def big_data_to_db(self, table, engine, df):
        # 数据库写入优化（使用批量UPSERT）
        chunk_size = 5000  # 根据测试调整最佳值
        columns = list(column_mapping.keys())
        db_columns = [column_mapping[col] for col in columns]

        # 准备SQL模板
        insert_sql = text(f"""
                               INSERT INTO {table} ({','.join(db_columns)})
                               VALUES ({','.join([':%s' % col for col in db_columns])})
                               ON DUPLICATE KEY UPDATE
                                   {','.join([f"{db_col}=VALUES({db_col})"
                                              for db_col in db_columns if db_col != 'performance_id'])}
                           """)

        with engine.begin() as conn:
            # 分块处理数据
            for i in range(0, len(df), chunk_size):
                chunk = df.iloc[i:i + chunk_size]
                # 转换为字典列表（内存优化）
                data = chunk.rename(columns=column_mapping) \
                    .replace({np.nan: None}) \
                    .to_dict('records')
                try:
                    # 批量执行
                    conn.execute(insert_sql, data)
                except Exception as e:
                    self.text.insert(tk.END, f"批量写入失败: {str(e)}\r\n")
                    raise e

    # 24年数据入库
    def import_24_data_to_db(self):
        self.text.insert(tk.END, f"开始读取24年数据...\r\n")
        two_four_path = self.two_four_data.get()
        if not two_four_path:
            self.text.insert(tk.END, "文件路径不能为空！\r\n")
            return
        try:
            engine = self.connect_db()
            if not engine:
                raise ConnectionError("数据库连接失败")

            with CalamineWorkbook.from_path(two_four_path) as wb:
                # 获取第一个工作表
                sheet = wb.get_sheet_by_index(0)
                # 读取所有数据（包含标题行）
                rows = sheet.to_python()
                selected_indices = [rows[0].index(col) for col in selected_columns]
                # 转换为DataFrame
                df = pd.DataFrame([[row[i] for i in selected_indices] for row in rows[1:]], columns=selected_columns)
                self.text.insert(tk.END, f"开始导入24年数据！\r\n")
                self.big_data_to_db('hw_two_four_data', engine, df)
                self.text.insert(tk.END, "数据导入成功！\r\n")
                del df
                # 获取第二个工作表
                sheet = wb.get_sheet_by_name('SMBcore')
                # 读取所有数据（包含标题行）
                rows = sheet.to_python()
                selected_indices = [rows[0].index(col) for col in selected_columns]
                # 转换为DataFrame
                df = pd.DataFrame([[row[i] for i in selected_indices] for row in rows[1:]], columns=selected_columns)
                self.text.insert(tk.END, f"开始导入24年SMBcore数据！\r\n")
                self.big_data_to_db('hw_two_four_data_smbcore', engine, df)
                self.text.insert(tk.END, "数据导入成功！\r\n")
                # del df
                # # 获取第三个工作表
                # sheet = wb.get_sheet_by_name('NA')
                # # 读取所有数据（包含标题行）
                # rows = sheet.to_python()
                # selected_indices = [rows[0].index(col) for col in selected_columns]
                # # 转换为DataFrame
                # df = pd.DataFrame([[row[i] for i in selected_indices] for row in rows[1:]], columns=selected_columns)
                # self.text.insert(tk.END, f"开始导入24年NA数据！\r\n")
                # self.big_data_to_db('hw_two_four_data_na', engine, df)
                # self.text.insert(tk.END, "数据导入成功！\r\n")
                del df, rows
                gc.collect()
        except Exception as e:
            self.text.insert(tk.END, f"24年数据读取失败: {str(e)}\r\n")
        finally:
            engine.dispose()

    # 25年业绩表增加BI到BO列，含导出25年数据
    def add_bi_to_bo(self, engine, two_five_path):
        # 1. Excel读取优化 使用calamine引擎读取数据
        try:
            wb = CalamineWorkbook.from_path(two_five_path)
            # 获取第一个工作表
            sheet = wb.get_sheet_by_index(0)
            # 读取所有数据（包含标题行）
            rows = sheet.to_python()
            selected_indices = [rows[0].index(col) for col in selected_columns]
            # 转换为DataFrame
            two_five_df = pd.DataFrame([[row[i] for i in selected_indices] for row in rows[1:]], columns=selected_columns)
            wb.close()
        except Exception as e:
            raise ValueError(f"25年数据读取失败: {str(e)}")

        # 2. 数据库查询
        with engine.connect() as conn:
            # 批量获取所有字典数据
            # query = """
            #         SELECT cloud_services_code, service_department
            #         FROM two_five_details_cloud_services;
            #         SELECT product_code, product_type
            #         FROM two_five_details_flow;
            #         SELECT product_code, product_name
            #         FROM two_five_details_special;
            #         SELECT cloud_services_code, cloud_services_name
            #         FROM two_five_details_collaborate;
            #         SELECT customer_name, salesperson, region
            #         FROM customer_correspondence;
            #     """
            # # 使用pandas多查询读取（比原生驱动快3-5倍）
            # dfs = pd.read_sql_query(query, conn, chunksize=None)
            sql_list = ['SELECT cloud_services_code, service_department FROM two_five_details_cloud_services;',
                        'SELECT product_code, product_type FROM two_five_details_flow',
                        'SELECT product_code, product_name FROM two_five_details_special',
                        'SELECT cloud_services_code, cloud_services_name FROM two_five_details_collaborate',
                        'SELECT customer_name, salesperson, region FROM customer_correspondence']
            dfs = [
                pd.read_sql_query(sql, conn, chunksize=None)
                for sql in sql_list
            ]

            # 解析查询结果
            cloud_services_map = dfs[0].set_index('cloud_services_code')['service_department']
            flow_products = set(dfs[1]['product_code'])
            special_products_map = dfs[2].set_index('product_code')['product_name']
            collaborate_map = dfs[3].set_index('cloud_services_code')['cloud_services_name']
            customer_relations = dfs[4][['customer_name', 'salesperson', 'region']]

            two_five_df = two_five_df.merge(customer_relations, left_on='客户名称', right_on='customer_name',how='left')

        # 3. 数据加工
        try:
            # 服务产品部映射
            two_five_df['服务产品部'] = two_five_df['产品类型编码'].map(cloud_services_map).fillna('')
            # 流量产品标记
            two_five_df['是否流量型产品'] = np.where(two_five_df['产品类型编码'].isin(flow_products), '是', '否')
            # 专线产品映射
            two_five_df['专线产品'] = two_five_df['产品类型编码'].map(special_products_map).fillna('')
            # 企业协同映射
            two_five_df['企业协同'] = two_five_df['产品类型编码'].map(collaborate_map).fillna('')
            # 客户信息映射
            two_five_df['销售员'] = two_five_df['salesperson'].fillna('').astype('category')
            two_five_df['区域'] = two_five_df['region'].fillna('').astype('category')
            # 清理临时列
            two_five_df.drop(['customer_name', 'salesperson', 'region'], axis=1, inplace=True)
            # 季度
            months = pd.to_datetime(two_five_df['业绩形成时间']).dt.month
            two_five_df['季度'] = 'Q' + ((months - 1) // 3 + 1).astype(str)
        except KeyError as e:
            raise ValueError(f"数据加工异常，缺少关键字段: {str(e)}")
        # 4. 数据入库
        try:
            self.text.insert(tk.END, f"25年数据入库中...\r\n")
            self.big_data_to_db('hw_two_five_data', engine, two_five_df)
        except Exception as e:
            self.text.insert(tk.END, f"25年数据入库失败: {str(e)}\r\n")
        # 5. 导出25年的数据
        self.text.insert(tk.END, "正在导出25年数据...\r\n")

    # 生成结果表
    # def generate_result_table(self, one, ywo, three, four, five, six, seven, eight, nine):




if __name__ == '__main__':
    root = tk.Tk()

    App(root)
    root.mainloop()
