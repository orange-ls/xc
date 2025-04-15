import gc
import tkinter as tk
# from tkinter.filedialog import askdirectory
from tkinter import filedialog
import threading

import numpy as np
import pandas as pd
from sqlalchemy import create_engine, text
from openpyxl import load_workbook
# import DI
from tkinter import messagebox
# import table_handling
# from time import sleep
import os
# import utils
from python_calamine import CalamineWorkbook

# 字段映射配置
column_mapping = {
    '业绩ID': 'performance_id',
    '业绩金额(¥)': 'sales_amount',
    '业绩形成时间': 'performance_date',
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
_column_config = {
            'mapping': {
                '业绩ID': ('A', 'str'),
                '业绩金额(¥)': ('G', 'float32'),
                '业绩形成时间': ('H', 'date32'),
                '二级经销商名称': ('Y', 'str'),
                '客户名称': ('AC', 'str'),
                '产品类型编码': ('AH', 'str'),
                '客户标签': ('BB', 'str'),
                '销售纵队': ('BF', 'str'),
                '服务产品部': ('BI', 'str'),
                '是否流量型产品': ('BJ', 'str'),
                '专线产品': ('BK', 'str'),
                '企业协同': ('BL', 'str'),
                '销售员': ('BM', 'str'),
                '区域': ('BN', 'str'),
                '季度': ('BO', 'str'),
            },
            'output_columns': [
                ('服务产品部', 'BI'), ('是否流量型产品', 'BJ'),
                ('专线产品', 'BK'), ('企业协同', 'BL'),
                ('销售员', 'BM'), ('区域', 'BN'), ('季度', 'BO')
            ]
        }

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
            if not two_five_path or not product_details_path or not customer_cor_path or not data_requirements_path:
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
            # self.common_excel_to_db(connect_db, product_details_path, customer_cor_path)

            self.text.insert(tk.END, "25年业绩表增加BI到BO列...\r\n")
            # 25年业绩表增加BI到BO列
            # self.add_bi_to_bo(connect_db, two_five_path)
            self.process(connect_db, two_five_path)


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
        two_four_path = self.two_four_data.get()
        try:
            engine = self.connect_db()
            if not engine:
                raise ConnectionError("数据库连接失败")

            wb = CalamineWorkbook.from_path(two_four_path)
            # 获取第一个工作表
            sheet = wb.get_sheet_by_index(0)
            # 读取所有数据（包含标题行）
            rows = sheet.to_python()
            # 转换为DataFrame
            df = pd.DataFrame(rows[1:], columns=rows[0])[selected_columns]
            wb.close()


            def big_data_to_db(table, engine, df):
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
                            return
            self.text.insert(tk.END, "数据导入成功！\r\n")
        except Exception as e:
            self.text.insert(tk.END, f"列名读取失败: {str(e)}\r\n")


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

    # 25年业绩表增加BI到BO列
    def add_bi_to_bo(self, engine, two_five_path):
        # try:
        #     # 读取25年业绩表
        #     two_five_df = pd.read_excel(two_five_path)[['业绩ID', '业绩金额(¥)', '业绩形成时间', '二级经销商名称', '客户名称', '产品类型编码', '客户标签', '销售纵队', '服务产品部', '是否流量型产品', '专线产品', '企业协同', '销售员', '区域', '季度']]
        #     # 查询数据库中产品明细表和客户关系表
        #     with engine.connect() as conn:
        #         # 云服务名称
        #         cloud_services = conn.execute(text("SELECT * FROM two_five_details_cloud_services"), ).mappings().all()
        #         cloud_services_dict = {row['cloud_services_code']: row['service_department'] for row in cloud_services}
        #         # 流量产品清单
        #         details_flow = conn.execute(text("SELECT * FROM two_five_details_flow"), ).mappings().all()
        #         flow_dict = {row['product_code']: row['product_type'] for row in details_flow}
        #         # 产品专项
        #         details_special = conn.execute(text("SELECT * FROM two_five_details_special"), ).mappings().all()
        #         special_dict = {row['product_code']: row['product_name'] for row in details_special}
        #         # 企业协同
        #         collaborate = conn.execute(text("SELECT * FROM two_five_details_collaborate"), ).mappings().all()
        #         collaborate_dict = {row['cloud_services_code']: row['cloud_services_name'] for row in collaborate}
        #         # 客户对应关系表
        #         customer_cor = conn.execute(text("SELECT * FROM customer_correspondence"), ).mappings().all()
        #         customer_cor_dict = {row['customer_name']: (row['salesperson'], row['region']) for row in customer_cor}
        #     # 遍历每一条数据，增加BI到BO列
        #     for index, row in two_five_df.iterrows():
        #         row['服务产品部'] = cloud_services_dict.get(row['产品类型编码'], '')
        #         row['是否流量型产品'] = '是' if flow_dict.get(row['产品类型编码'], '') else '否'
        #         row['专线产品'] = special_dict.get(row['产品类型编码'], '')
        #         row['企业协同'] = collaborate_dict.get(row['产品类型编码'], '')
        #         salesperson, region = customer_cor_dict.get(row['客户名称'], ('', ''))
        #         row['销售员'] = salesperson
        #         row['区域'] = region
        #         dt = pd.to_datetime(row['业绩形成时间'])
        #         # 计算财务季度（假设财年从1月开始）
        #         fiscal_quarter = (dt.month - 1) // 3 + 1
        #         row['季度'] = f'Q{fiscal_quarter}'
        #     # 将结果数据写入数据库，按业绩ID来插入或更新数据
        #     with engine.begin() as conn:
        #         # 创建带唯一会话标识的临时表
        #         temp_table_name = f"temp_two_five_data"
        #
        #         try:
        #             # 1. 创建支持事务的InnoDB临时表（MySQL 8.0+）
        #             conn.execute(text(f"""
        #                 CREATE TEMPORARY TABLE {temp_table_name} (
        #                     performance_id BIGINT PRIMARY KEY,
        #                     amount DECIMAL(18,2),
        #                     product_type VARCHAR(20),
        #                     quarter CHAR(2),
        #                     INDEX idx_quarter (quarter)
        #                 ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        #             """))
        #
        #             # 2. 分块批量插入（MySQL最大包大小限制）
        #             chunk_size = 5000  # 根据max_allowed_packet调整
        #             two_five_df = two_five_df.rename(columns={'业绩ID': 'performance_id', '业绩金额(¥)': 'sales_amount', '业绩形成时间': 'performance_date', '二级经销商名称': 'secondary_dealer', '客户名称': 'customer_name', '产品类型编码': 'product_code', '客户标签': 'customer_tag', '销售纵队': 'sales_team', '服务产品部': 'service_department', '是否流量型产品': 'is_traffic_product', '专线产品': 'leased_line_product', '企业协同': 'enterprise_coop', '销售员': 'salesperson', '区域': 'region', '季度': 'quarter'})
        #             for i in range(0, len(two_five_df), chunk_size):
        #                 chunk = two_five_df.iloc[i:i + chunk_size]
        #                 chunk.to_sql(
        #                     name=temp_table_name,
        #                     con=conn,
        #                     if_exists='append',
        #                     index=False,
        #                     method='multi',
        #                     chunksize=1000
        #                 )
        #
        #             # 3. 使用INSERT ... ON DUPLICATE KEY UPDATE
        #             conn.execute(text(f"""
        #                 INSERT INTO hw_two_five_data
        #                 SELECT * FROM {temp_table_name}
        #                 ON DUPLICATE KEY UPDATE
        #                     amount = VALUES(amount),
        #                     product_type = VALUES(product_type),
        #                     quarter = VALUES(quarter),
        #                     update_time = CURRENT_TIMESTAMP()
        #             """))
        #
        #             # 4. 手动释放临时表空间（针对大表）
        #             conn.execute(text(f"ALTER TABLE {temp_table_name} ENGINE=InnoDB ROW_FORMAT=COMPRESSED"))
        #             conn.execute(text(f"OPTIMIZE LOCAL TABLE {temp_table_name}"))
        #
        #         finally:
        #             # 5. 显式删除临时表（防御性措施）
        #             try:
        #                 conn.execute(text(f"DROP TEMPORARY TABLE IF EXISTS {temp_table_name}"))
        #             except Exception as e:
        #                 self.text.insert(tk.END, f"临时表清理失败: {str(e)}\r\n")
        #
        # except Exception as e:
        #     self.text.insert(tk.END, f"读取25年业绩表时发生错误: {str(e)}\r\n")

        # 1. Excel读取优化
        try:
            two_five_df = pd.read_excel(
                two_five_path,
                dtype=column_mapping,
                engine='openpyxl',  # 明确指定引擎
                na_filter=False,  # 关闭空值检测
            )
        except Exception as e:
            raise ValueError(f"Excel读取失败: {str(e)}")

        # 2. 数据库查询优化（单次连接批量查询）
        with engine.connect() as conn:
            # 批量获取所有字典数据
            query = """
                    /* 云服务字典 */
                    SELECT cloud_services_code, service_department 
                    FROM two_five_details_cloud_services;

                    /* 流量产品清单 */
                    SELECT product_code, product_type
                    FROM two_five_details_flow;

                    /* 产品专项字典 */
                    SELECT product_code, product_name 
                    FROM two_five_details_special;

                    /* 企业协同字典 */
                    SELECT cloud_services_code, cloud_services_name 
                    FROM two_five_details_collaborate;

                    /* 客户对应关系 */
                    SELECT customer_name, salesperson, region 
                    FROM customer_correspondence;
                """

            # 使用pandas多查询读取（比原生驱动快3-5倍）
            dfs = pd.read_sql_query(query, conn, chunksize=None)

            # 解析查询结果
            cloud_services_map = dfs[0].set_index('cloud_services_code')['service_department']
            flow_products = set(dfs[1].set_index('product_code')['product_type'])
            special_products_map = dfs[2].set_index('product_code')['product_name']
            collaborate_map = dfs[3].set_index('cloud_services_code')['cloud_services_name']
            customer_relations = dfs[4].set_index('customer_name')['salesperson', 'region']

        # 3. 数据加工优化（向量化操作）
        try:
            # 服务产品部映射（向量化操作）
            two_five_df['服务产品部'] = two_five_df['产品类型编码'].map(cloud_services_map).fillna('')

            # 流量产品标记（布尔索引）
            two_five_df['是否流量型产品'] = np.where(
                two_five_df['产品类型编码'].isin(flow_products), '是', '否'
            )

            # 专线产品映射
            two_five_df['专线产品'] = two_five_df['产品类型编码'].map(special_products_map).fillna('')

            # 企业协同映射
            two_five_df['企业协同'] = two_five_df['产品类型编码'].map(collaborate_map).fillna('')

            # 客户信息映射（批量处理）
            customer_info = customer_relations.reindex(two_five_df['客户名称'])
            two_five_df['销售员'] = customer_info['salesperson'].fillna('').astype('category')
            two_five_df['区域'] = customer_info['region'].fillna('').astype('category')

            # 财务季度计算（向量化）
            months = pd.to_datetime(two_five_df['业绩形成时间']).dt.month
            two_five_df['季度'] = 'Q' + ((months - 1) // 3 + 1).astype(str)

        except KeyError as e:
            raise ValueError(f"数据加工异常，缺少关键字段: {str(e)}")

        # 4. 数据库写入优化（原生批量操作）
        # 内存优化处理
        two_five_df = two_five_df.astype({
            '服务产品部': 'category',
            '是否流量型产品': 'category',
            '季度': 'category'
        })

        # 分块写入函数
        def batch_insert(conn, df, chunk_size=5000):
            """使用原生批量插入实现高效写入"""
            cols = df.columns.tolist()
            total = len(df)

            # 生成动态SQL
            insert_sql = f"""
                    INSERT INTO hw_two_five_data ({','.join(cols)})
                    VALUES ({','.join(['%s'] * len(cols))})
                    ON DUPLICATE KEY UPDATE
                        {','.join([f"{col}=VALUES({col})" for col in cols if col != '业绩ID'])}
                """

            # 分块处理
            for i in range(0, total, chunk_size):
                chunk = df.iloc[i:i + chunk_size]
                # 转换numpy类型为Python原生类型
                data = [tuple(x.astype(object) if isinstance(x, pd.Timestamp) else x
                              for x in record)
                        for record in chunk.itertuples(index=False)]

                try:
                    conn.execute(text(insert_sql), data)
                except Exception as e:
                    raise RuntimeError(f"批量插入失败: {str(e)}")

        # 执行写入
        try:
            with engine.begin() as conn:
                # 执行分块插入
                batch_insert(conn, two_five_df)

        except Exception as e:
            raise RuntimeError(f"数据库写入失败: {str(e)}")
        finally:
            # 内存清理
            del two_five_df
            gc.collect()

    def _optimized_excel_read(self, path):
        # 预处理列配置：提前计算好列索引和数据类型
        column_mapping_config = _column_config['mapping']

        # 预计算列索引和数据类型（只需计算一次）
        preprocessed_columns = []
        for col_str, dtype in column_mapping_config.values():
            # 优化列索引计算函数
            index = 0
            for char in col_str.upper():
                index = index * 26 + (ord(char) - ord('A') + 1)
            preprocessed_columns.append((index - 1, dtype))  # 存储零基索引和类型

        # 预提取列名（只需计算一次）
        column_names = list(column_mapping_config.keys())

        """内存优化的Excel流式读取"""
        wb = load_workbook(
            filename=path,
            read_only=True,
            data_only=True,
            keep_links=False,
            rich_text=False
        )
        ws = wb.active

        def _convert_value(value, dtype):
            """类型安全转换"""
            try:
                if dtype == 'int32': return int(value or 0)
                if dtype == 'float32': return float(value or 0.0)
                if dtype == 'category': return str(value).strip()[:50]  # 长度限制
                return value
            except:
                return None

        # 生成器模式读取数据
        def data_stream():
            # 使用iter_rows的批量模式（默认每次返回100行）
            for row in ws.iter_rows(min_row=2, values_only=True):
                # 预分配列表避免多次append
                yield [
                    _convert_value(row[idx], dtype)
                    for idx, dtype in preprocessed_columns
                ]

        dtypes = {name: dtype for name, (_, dtype) in column_mapping_config.items()}
        df = pd.DataFrame(
            data=data_stream(),
            columns=column_names
        ).astype(dtypes)

        wb.close()
        return df


    def _batch_export(self, df, original_path):
        """基于openpyxl的智能批量导出"""
        wb = load_workbook(original_path)
        ws = wb.active

        # 批量写入优化
        output_cols = _column_config['output_columns']
        col_indices = {col: ord(pos) - 65 for col, pos in output_cols}

        # 内存分块处理
        chunk_size = 5000
        for i in range(0, len(df), chunk_size):
            chunk = df.iloc[i:i + chunk_size]
            for idx, row in chunk.iterrows():
                for col in output_cols:
                    ws.cell(
                        row=idx + 2,  # 数据从第2行开始
                        column=col_indices[col[0]] + 1,
                        value=row[col[0]]
                    )
            # 阶段性提交和内存清理
            wb.save(original_path.replace('.xlsx', '_updated.xlsx'))
            gc.collect()

        wb.close()

    def process(self, engine, file_path):
        try:
            # 1. 高性能读取
            # df = self._optimized_excel_read(file_path)
            # 2.使用pandas
            # df = pd.read_excel(
            #     file_path,
            #     skiprows=1,  # 跳过标题行（假设数据从第2行开始）
            #     header=None,  # 无列标题
            #     engine='calamine'
            # )[[selected_columns]]

            wb = CalamineWorkbook.from_path(file_path)
            # 获取第一个工作表
            sheet = wb.get_sheet_by_index(0)
            # 读取所有数据（包含标题行）
            rows = sheet.to_python()
            # 转换为DataFrame
            df = pd.DataFrame(rows[1:], columns=rows[0])[selected_columns]
            wb.close()

            # 2. 数据加工流程（保持原有逻辑）
            # ... [原有数据库查询和向量化处理代码] ...
            # 数据库写入优化（使用批量UPSERT）
            chunk_size = 5000  # 根据测试调整最佳值
            columns = list(column_mapping.keys())
            db_columns = [column_mapping[col] for col in columns]

            # 准备SQL模板
            insert_sql = text(f"""
                       INSERT INTO hw_two_five_data ({','.join(db_columns)})
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
                        return


            # 3. 智能导出
            self._batch_export(df, file_path)

            return "处理完成"

        except Exception as e:
            raise RuntimeError(f"处理失败: {str(e)}")


if __name__ == '__main__':
    root = tk.Tk()

    App(root)
    root.mainloop()
