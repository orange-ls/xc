'''
    第1个结果表是：各“区域”的“整体业绩”、“NA业绩”、“SMB业绩”、“SMBcore业绩”、“同期增长率”
    第2个结果表是：各“区域”的“全量业绩”等
    第3个结果表是：各“销售员”的“全量业绩”等
    第4个结果表是：各“区域”各“月份”的“SMBcore业绩”和“NA业绩”
    第5个结果表是：25年SMBcore业绩
    第6个结果表是：各“二级经销商和客户”的SMBcore业绩
    第7个结果表是：各“专线产品”季度业绩
    第8个结果表是：新增渠道
    第9个结果表是：新增客户
'''
from sqlalchemy import text


def result_table_one(engine, max_date):
    '''
    第1个结果表
    :param engine: 数据库连接
    :return: 结果表数据
    '''
    # 构建查询sql,用于“整体业绩”、“NA业绩”、“SMB业绩”、“SMBcore业绩”
    def select_sql(where_sql):
        select_sql = f'''
            -- 分地区统计渠道/直客金额并生成二维报表，强制显示"其他"行
            WITH all_regions AS (
                SELECT '北京' AS region UNION ALL
                SELECT '广州' UNION ALL SELECT '深圳' UNION ALL 
                SELECT '上海' UNION ALL SELECT '南京' UNION ALL 
                SELECT '成都' UNION ALL SELECT '其他'
            ),
            customer_types AS (
                SELECT '渠道' AS customer_type UNION ALL
                SELECT '直客'
            ),
            region_data AS (
                SELECT 
                    COALESCE(ar.region, 
                        CASE 
                            WHEN d.region NOT IN ('北京','广州','深圳','上海','南京','成都') 
                            THEN '其他' 
                        END
                    ) AS grouped_region,
                    ct.customer_type,
                    COALESCE(ROUND(SUM(d.sales_amount)/10000, 1), 0) AS total_sales
                FROM all_regions ar
                CROSS JOIN customer_types ct
                LEFT JOIN hw_two_five_data d 
                    ON d.region = ar.region
                    AND ct.customer_type = CASE 
                        WHEN d.secondary_dealer IS NULL OR d.secondary_dealer='' THEN '直客' 
                        ELSE '渠道' 
                    END
                    {where_sql}
                GROUP BY grouped_region, ct.customer_type
            ),
            pivot_data AS (
                SELECT 
                    grouped_region,
                    SUM(CASE WHEN customer_type = '渠道' THEN total_sales END) AS 渠道,
                    SUM(CASE WHEN customer_type = '直客' THEN total_sales END) AS 直客,
                    SUM(total_sales) AS 合计
                FROM region_data
                
                GROUP BY
                    grouped_region
            ),
            total_row AS (
                SELECT 
                    '总计' AS grouped_region,
                    SUM(渠道) AS 渠道,
                    SUM(直客) AS 直客,
                    SUM(合计) AS 合计
                FROM pivot_data
                WHERE grouped_region != '总计'  -- 避免重复累加
            )
            
            -- 合并数据并排序
            SELECT * FROM (
                SELECT * FROM pivot_data
                UNION ALL
                SELECT * FROM total_row
            ) AS final_data
            WHERE grouped_region IS NOT NULL  -- 过滤空值
            ORDER BY
                FIELD(grouped_region,'北京','广州','深圳','上海','南京','成都','其他','总计');
        '''
        return select_sql

    conn = engine.connect()
    select_params = {
        '整体业绩': "",
        'NA业绩': "AND d.sales_team = '华为云NA'",
        'SMB业绩': "AND d.sales_team in ('中长尾','电网销')",
        'SMBcore业绩': "AND d.sales_team in ('中长尾','电网销') AND d.is_traffic_product IN ('否', '')"
    }

    # 循环执行查询 “整体业绩”、“NA业绩”、“SMB业绩”、“SMBcore业绩” 并返回结果
    result_data = {}
    for k, v in select_params.items():
        result = [dict(row) for row in conn.execute(text(select_sql(v))).mappings().fetchall()]
        result = {re['grouped_region']: re for re in result}
        result_data[k] = result

    # 构建sql，查询“同期增长率”
    growth_rate_sql = f'''
        -- 分表统计24年与25年数据，计算增长率
        WITH 
        -- 1. 定义所有地区
        all_regions AS (
            SELECT '北京' AS region UNION ALL
            SELECT '广州' UNION ALL SELECT '深圳' UNION ALL
            SELECT '上海' UNION ALL SELECT '南京' UNION ALL
            SELECT '成都' UNION ALL SELECT '其他'
        ),
        
        -- 2. 定义业绩类型及对应条件
        performance_types AS (
            SELECT 
                '整体业绩' AS ptype,
                '' AS condition_clause
            UNION ALL
            SELECT 
                'NA业绩',
                "AND sales_team = '华为云NA'"
            UNION ALL
            SELECT 
                'SMB业绩',
                "AND sales_team IN ('中长尾','电网销')"
            UNION ALL
            SELECT 
                'SMBcore业绩',
                "AND sales_team IN ('中长尾','电网销') AND is_traffic_product IN ('否', '')"
        ),
        
        -- 3. 计算24年各维度数据
        data_2024 AS (
            SELECT 
                CASE 
                    WHEN region IN ('北京','广州','深圳','上海','南京','成都') 
                    THEN region 
                    ELSE '其他' 
                END AS grouped_region,
                pt.ptype,
                COALESCE(SUM(
                    CASE
                        WHEN pt.ptype = '整体业绩' THEN d.sales_amount
                        WHEN pt.ptype = 'NA业绩' AND d.sales_team = '华为云NA' THEN d.sales_amount
                        WHEN pt.ptype = 'SMB业绩' AND d.sales_team IN ('中长尾','电网销') THEN d.sales_amount
                        WHEN pt.ptype = 'SMBcore业绩' AND d.sales_team IN ('中长尾','电网销') AND d.is_traffic_product IN ('否', '') THEN d.sales_amount
                    END
                ), 0) AS amount_2024
            FROM hw_two_four_data d
            CROSS JOIN performance_types pt
                WHERE d.performance_date BETWEEN DATE(CONCAT(YEAR(CURDATE()) - 1, '-01-01')) AND '{max_date}'
            GROUP BY 
                CASE 
                    WHEN d.region IN ('北京','广州','深圳','上海','南京','成都') 
                    THEN d.region 
                    ELSE '其他' 
                END, 
                pt.ptype
        ),
        
        -- 4. 计算25年各维度数据（结构相同）
        data_2025 AS (
            SELECT 
                CASE 
                    WHEN region IN ('北京','广州','深圳','上海','南京','成都') 
                    THEN region 
                    ELSE '其他' 
                END AS grouped_region,
                pt.ptype,
                COALESCE(SUM(
                    CASE
                        WHEN pt.ptype = '整体业绩' THEN d.sales_amount
                        WHEN pt.ptype = 'NA业绩' AND d.sales_team = '华为云NA' THEN d.sales_amount
                        WHEN pt.ptype = 'SMB业绩' AND d.sales_team IN ('中长尾','电网销') THEN d.sales_amount
                        WHEN pt.ptype = 'SMBcore业绩' AND d.sales_team IN ('中长尾','电网销') AND d.is_traffic_product IN ('否', '') THEN d.sales_amount
                    END
                ), 0) AS amount_2025
            FROM hw_two_five_data d
            CROSS JOIN performance_types pt
            GROUP BY 
                CASE 
                    WHEN d.region IN ('北京','广州','深圳','上海','南京','成都') 
                    THEN d.region 
                    ELSE '其他' 
                END, 
                pt.ptype
        ),
        
        -- 5. 合并两年数据并计算增长率
        combined_data AS (
            SELECT 
                ar.region,
                pt.ptype,
                COALESCE(d24.amount_2024, 0) AS amount_2024,
                COALESCE(d25.amount_2025, 0) AS amount_2025,
                CASE 
                    WHEN COALESCE(d24.amount_2024, 0) = 0 THEN NULL  -- 处理除零
                    ELSE ROUND((d25.amount_2025 - d24.amount_2024) / d24.amount_2024 * 100, 0)
                END AS growth_rate
            FROM all_regions ar
            CROSS JOIN performance_types pt
            LEFT JOIN data_2024 d24 
                ON ar.region = d24.grouped_region AND pt.ptype = d24.ptype
            LEFT JOIN data_2025 d25 
                ON ar.region = d25.grouped_region AND pt.ptype = d25.ptype
        ),
        
        -- 6. 行列转换生成报表
        pivot_table AS (
            SELECT 
                region,
                MAX(CASE WHEN ptype = '整体业绩' THEN growth_rate END) AS all_sales,
                MAX(CASE WHEN ptype = 'NA业绩' THEN growth_rate END) AS na_sales,
                MAX(CASE WHEN ptype = 'SMB业绩' THEN growth_rate END) AS smb_sales,
                MAX(CASE WHEN ptype = 'SMBcore业绩' THEN growth_rate END) AS smbcore_sales
            FROM combined_data
            GROUP BY region
        ),
        
        -- 7. 生成总计行
        total_row AS (
            SELECT 
                '总计' AS region,
                ROUND(
                    (SUM(CASE WHEN ptype = '整体业绩' THEN amount_2025 END) - 
                     SUM(CASE WHEN ptype = '整体业绩' THEN amount_2024 END)) / 
                    NULLIF(SUM(CASE WHEN ptype = '整体业绩' THEN amount_2024 END), 0) * 100, 0
                ) AS all_sales,
                ROUND(
                    (SUM(CASE WHEN ptype = 'NA业绩' THEN amount_2025 END) - 
                     SUM(CASE WHEN ptype = 'NA业绩' THEN amount_2024 END)) / 
                    NULLIF(SUM(CASE WHEN ptype = 'NA业绩' THEN amount_2024 END), 0) * 100, 0
                ) AS na_sales,
                ROUND(
                    (SUM(CASE WHEN ptype = 'SMB业绩' THEN amount_2025 END) - 
                     SUM(CASE WHEN ptype = 'SMB业绩' THEN amount_2024 END)) / 
                    NULLIF(SUM(CASE WHEN ptype = 'SMB业绩' THEN amount_2024 END), 0) * 100, 0
                ) AS smb_sales,
                ROUND(
                    (SUM(CASE WHEN ptype = 'SMBcore业绩' THEN amount_2025 END) - 
                     SUM(CASE WHEN ptype = 'SMBcore业绩' THEN amount_2024 END)) / 
                    NULLIF(SUM(CASE WHEN ptype = 'SMBcore业绩' THEN amount_2024 END), 0) * 100, 0
                ) AS smbcore_sales
            FROM combined_data
        )
        
        -- 8. 最终结果
        SELECT 
            region AS grouped_region,
		    CONCAT(IFNULL(all_sales, '0'), '%') AS `整体业绩`,
            CONCAT(IFNULL(na_sales, '0'), '%') AS `NA业绩`,
            CONCAT(IFNULL(smb_sales, '0'), '%') AS `SMB业绩`,
            CONCAT(IFNULL(smbcore_sales, '0'), '%') AS `SMBcore业绩`
        FROM (
            SELECT * FROM pivot_table
            UNION ALL
            SELECT * FROM total_row
        ) AS final
        ORDER BY FIELD(region, '北京','广州','深圳','上海','南京','成都','其他','总计');
    '''
    result = [dict(row) for row in conn.execute(text(growth_rate_sql)).mappings().fetchall()]
    result = {re['grouped_region']: re for re in result}
    result_data['同期增长率'] = result

    # 构建sql，查询“24年同期数据”
    data_24_sql = f'''
        WITH 
        -- 1. 定义所有地区
        all_regions AS (
            SELECT '北京' AS region UNION ALL
            SELECT '广州' UNION ALL SELECT '深圳' UNION ALL
            SELECT '上海' UNION ALL SELECT '南京' UNION ALL
            SELECT '成都' UNION ALL SELECT '其他'
        ),
        
        -- 2. 定义业绩类型及对应条件
        performance_types AS (
            SELECT '整体业绩' AS ptype UNION ALL
            SELECT 'NA业绩' UNION ALL
            SELECT 'SMB业绩' UNION ALL
            SELECT 'SMBcore业绩'
        ),
        
        -- 3. 计算24年各维度数据
        data_2024 AS (
            SELECT 
                ar.region AS grouped_region,
                pt.ptype,
                COALESCE(ROUND(
                        SUM(
                    CASE
                        WHEN pt.ptype = '整体业绩' THEN d.sales_amount
                        WHEN pt.ptype = 'NA业绩' AND d.sales_team = '华为云NA' THEN d.sales_amount
                        WHEN pt.ptype = 'SMB业绩' AND d.sales_team IN ('中长尾','电网销') THEN d.sales_amount
                        WHEN pt.ptype = 'SMBcore业绩' AND d.sales_team IN ('中长尾','电网销') AND COALESCE(d.is_traffic_product, '') IN ('否', '') THEN d.sales_amount
                    END
                )/10000,1), 0) AS amount
            FROM all_regions ar
            CROSS JOIN performance_types pt
            LEFT JOIN hw_two_four_data d 
                ON ar.region = CASE 
                    WHEN d.region IN ('北京','广州','深圳','上海','南京','成都') 
                    THEN d.region 
                    ELSE '其他' 
                END
                AND d.performance_date BETWEEN DATE(CONCAT(YEAR(CURDATE()) - 1, '-01-01')) AND '{max_date}'
            GROUP BY ar.region, pt.ptype
        ),
        
        -- 4. 转换为宽表格式
        wide_data AS (
            SELECT 
                grouped_region,
                SUM(CASE WHEN ptype = '整体业绩' THEN amount ELSE 0 END) AS 整体业绩,
                SUM(CASE WHEN ptype = 'NA业绩' THEN amount ELSE 0 END) AS NA业绩,
                SUM(CASE WHEN ptype = 'SMB业绩' THEN amount ELSE 0 END) AS SMB业绩,
                SUM(CASE WHEN ptype = 'SMBcore业绩' THEN amount ELSE 0 END) AS SMBcore业绩
            FROM data_2024
            GROUP BY grouped_region
        ),
        
        -- 5. 添加总计行
        final_data AS (
            SELECT * FROM wide_data
            UNION ALL
            SELECT 
                '总计' AS grouped_region,
                SUM(整体业绩),
                SUM(NA业绩),
                SUM(SMB业绩),
                SUM(SMBcore业绩)
            FROM wide_data
        )
        
        -- 6. 按指定顺序排序
        SELECT 
            grouped_region,
            整体业绩,
            NA业绩,
            SMB业绩,
            SMBcore业绩
        FROM final_data
        ORDER BY 
            CASE grouped_region
                WHEN '北京' THEN 1
                WHEN '广州' THEN 2
                WHEN '深圳' THEN 3
                WHEN '上海' THEN 4
                WHEN '南京' THEN 5
                WHEN '成都' THEN 6
                WHEN '其他' THEN 7
                WHEN '总计' THEN 8
            END;
    '''
    result_24 = [dict(row) for row in conn.execute(text(data_24_sql)).mappings().fetchall()]
    result_24 = {re['grouped_region']: re for re in result_24}
    result_data['24年同期数据'] = result_24

    conn.close()
    return result_data


def result_table_two(engine):
    sql = '''
        SELECT 
            IFNULL(classified_region, '汇总') AS 区域,
            ROUND(SUM(national_num)/10000, 1) AS 全量业绩,
            ROUND(SUM(national_num_h1)/10000, 1) AS 全量H2进度,
            ROUND(SUM(national_year_num)/10000, 1) AS 全量全年进度,
            ROUND(SUM(smb_sales_h1)/10000, 1) AS SMBH2进度,
            ROUND(SUM(smb_sales_year)/10000, 1) AS SMB全年进度
        FROM (
            SELECT 
                CASE 
                    WHEN region IN ('北京','广州','深圳','上海','南京') THEN region
                    ELSE '其他' 
                END AS classified_region,
                sales_amount AS national_num,
                CASE WHEN performance_date >= '2025-07-01' THEN sales_amount ELSE 0 END AS national_num_h1,
                sales_amount AS national_year_num,
                CASE WHEN sales_team IN ('中长尾', '电网销') THEN sales_amount ELSE 0 END AS smb_sales,
                CASE WHEN performance_date >= '2025-07-01' AND sales_team IN ('中长尾', '电网销') 
                    THEN sales_amount ELSE 0 END AS smb_sales_h1,
                CASE WHEN sales_team IN ('中长尾', '电网销') THEN sales_amount ELSE 0 END AS smb_sales_year
            FROM hw_two_five_data
        ) AS sub
        GROUP BY classified_region WITH ROLLUP
        ORDER BY 
            CASE classified_region
                WHEN '北京' THEN 1
                WHEN '广州' THEN 2
                WHEN '深圳' THEN 3
                WHEN '上海' THEN 4
                WHEN '南京' THEN 5
                WHEN '其他' THEN 6
                ELSE 7
            END;
    '''

    result = [dict(row) for row in engine.connect().execute(text(sql)).mappings().fetchall()]
    result = {re['区域']: re for re in result}
    return result


def result_table_three(engine):
    sql = '''
        SELECT 
            IFNULL(classified, '汇总') AS 销售,
            ROUND(SUM(national_num)/10000, 1) AS 全量业绩,
            ROUND(SUM(national_num_h1)/10000, 1) AS 全量H2进度,
            ROUND(SUM(national_year_num)/10000, 1) AS 全量全年进度,
            ROUND(SUM(smb_sales_h1)/10000, 1) AS SMBH2进度,
            ROUND(SUM(smb_sales_year)/10000, 1) AS SMB全年进度
        FROM (
            SELECT 
                salesperson AS classified,
                sales_amount AS national_num,
                CASE WHEN performance_date >= '2025-07-01' THEN sales_amount ELSE 0 END AS national_num_h1,
                sales_amount AS national_year_num,
                CASE WHEN sales_team IN ('中长尾', '电网销') THEN sales_amount ELSE 0 END AS smb_sales,
                CASE WHEN performance_date >= '2025-07-01' AND sales_team IN ('中长尾', '电网销') 
                    THEN sales_amount ELSE 0 END AS smb_sales_h1,
                CASE WHEN sales_team IN ('中长尾', '电网销') THEN sales_amount ELSE 0 END AS smb_sales_year
            FROM hw_two_five_data
        ) AS sub
        GROUP BY classified WITH ROLLUP
    '''
    result = [dict(row) for row in engine.connect().execute(text(sql)).mappings().fetchall()]
    result = {re['销售']: re for re in result}
    return result


def result_table_four(engine, max_date):
    def select_sql(where_sql, max_date):
        sql = f'''
            WITH regions AS (
                SELECT '北京' AS 区域
                UNION ALL SELECT '广州'
                UNION ALL SELECT '深圳'
                UNION ALL SELECT '上海'
                UNION ALL SELECT '南京'
                UNION ALL SELECT '长春'
                UNION ALL SELECT '其他'
            ),
            filtered_data AS (
                  SELECT 
                    CASE WHEN region IN ('北京','广州','深圳','上海','南京','长春') THEN region ELSE '其他' END AS 区域,
                    MONTH(performance_date) AS month,
                    ROUND(COALESCE(SUM(sales_amount),0)/10000, 1) AS sales_amount
                FROM hw_two_five_data
                WHERE 
                    1=1
                    {where_sql}
                GROUP BY 区域, month
            ),
            last_year_data AS (
                -- 新增2024年同期数据部分
                SELECT 
                    CASE WHEN region IN ('北京','广州','深圳','上海','南京','长春') THEN region ELSE '其他' END AS 区域,
                    MONTH(performance_date) AS month,
                    ROUND(COALESCE(SUM(sales_amount),0)/10000, 1) AS sales_amount
                FROM hw_two_four_data
                WHERE 
                    performance_date BETWEEN DATE(CONCAT(YEAR(CURDATE()) - 1, '-01-01')) AND '{max_date}'
                    {where_sql}
                GROUP BY 区域, month
            )
            SELECT 
                r.区域,
                COALESCE(SUM(CASE WHEN d.month = 1 THEN d.sales_amount ELSE 0 END), 0) AS `1月`,
                COALESCE(SUM(CASE WHEN d.month = 2 THEN d.sales_amount ELSE 0 END), 0) AS `2月`,
                COALESCE(SUM(CASE WHEN d.month = 3 THEN d.sales_amount ELSE 0 END), 0) AS `3月`,
                COALESCE(SUM(CASE WHEN d.month = 4 THEN d.sales_amount ELSE 0 END), 0) AS `4月`,
                COALESCE(SUM(CASE WHEN d.month = 5 THEN d.sales_amount ELSE 0 END), 0) AS `5月`,
                COALESCE(SUM(CASE WHEN d.month = 6 THEN d.sales_amount ELSE 0 END), 0) AS `6月`,
                COALESCE(SUM(CASE WHEN d.month = 7 THEN d.sales_amount ELSE 0 END), 0) AS `7月`,
                COALESCE(SUM(CASE WHEN d.month = 8 THEN d.sales_amount ELSE 0 END), 0) AS `8月`,
                COALESCE(SUM(CASE WHEN d.month = 9 THEN d.sales_amount ELSE 0 END), 0) AS `9月`,
                COALESCE(SUM(CASE WHEN d.month = 10 THEN d.sales_amount ELSE 0 END), 0) AS `10月`,
                COALESCE(SUM(CASE WHEN d.month = 11 THEN d.sales_amount ELSE 0 END), 0) AS `11月`,
                COALESCE(SUM(CASE WHEN d.month = 12 THEN d.sales_amount ELSE 0 END), 0) AS `12月`,
                COALESCE(SUM(d.sales_amount), 0) AS 合计,
                COALESCE(SUM(l.sales_amount), 0) AS `24年同期`
            FROM regions r
            LEFT JOIN filtered_data d ON r.区域 = d.区域
            LEFT JOIN last_year_data l ON r.区域 = l.区域 AND d.month = l.month
            GROUP BY r.区域
            
            UNION ALL
            
            SELECT 
                '汇总' AS 区域,
                SUM(`1月`), SUM(`2月`), SUM(`3月`), SUM(`4月`),
                SUM(`5月`), SUM(`6月`), SUM(`7月`), SUM(`8月`),
                SUM(`9月`), SUM(`10月`), SUM(`11月`), SUM(`12月`),
                SUM(合计),SUM(`24年同期`)
            FROM (
                SELECT 
                    r.区域,
                    COALESCE(SUM(CASE WHEN d.month = 1 THEN d.sales_amount ELSE 0 END), 0) AS `1月`,
                    COALESCE(SUM(CASE WHEN d.month = 2 THEN d.sales_amount ELSE 0 END), 0) AS `2月`,
                    COALESCE(SUM(CASE WHEN d.month = 3 THEN d.sales_amount ELSE 0 END), 0) AS `3月`,
                    COALESCE(SUM(CASE WHEN d.month = 4 THEN d.sales_amount ELSE 0 END), 0) AS `4月`,
                    COALESCE(SUM(CASE WHEN d.month = 5 THEN d.sales_amount ELSE 0 END), 0) AS `5月`,
                    COALESCE(SUM(CASE WHEN d.month = 6 THEN d.sales_amount ELSE 0 END), 0) AS `6月`,
                    COALESCE(SUM(CASE WHEN d.month = 7 THEN d.sales_amount ELSE 0 END), 0) AS `7月`,
                    COALESCE(SUM(CASE WHEN d.month = 8 THEN d.sales_amount ELSE 0 END), 0) AS `8月`,
                    COALESCE(SUM(CASE WHEN d.month = 9 THEN d.sales_amount ELSE 0 END), 0) AS `9月`,
                    COALESCE(SUM(CASE WHEN d.month = 10 THEN d.sales_amount ELSE 0 END), 0) AS `10月`,
                    COALESCE(SUM(CASE WHEN d.month = 11 THEN d.sales_amount ELSE 0 END), 0) AS `11月`,
                    COALESCE(SUM(CASE WHEN d.month = 12 THEN d.sales_amount ELSE 0 END), 0) AS `12月`,
                    COALESCE(SUM(d.sales_amount), 0) AS 合计,
                    COALESCE(SUM(l.sales_amount), 0) AS `24年同期`
                FROM regions r
                LEFT JOIN filtered_data d ON r.区域 = d.区域
                LEFT JOIN last_year_data l ON r.区域 = l.区域 AND d.month = l.month
                GROUP BY r.区域
            ) AS sub
            ORDER BY 
                CASE 区域
                    WHEN '北京' THEN 1
                    WHEN '广州' THEN 2
                    WHEN '深圳' THEN 3
                    WHEN '上海' THEN 4
                    WHEN '南京' THEN 5
                    WHEN '长春' THEN 6
                    WHEN '其他' THEN 7
                    ELSE 8
                END;
        '''
        return sql

    select_params = {
        'SMBcore业绩': "AND sales_team in ('中长尾','电网销') AND is_traffic_product IN ('否', '')",
        'NA业绩': "AND sales_team = '华为云NA'"
    }
    conn = engine.connect()
    result_data = {}
    for k, v in select_params.items():
        result = [dict(row) for row in conn.execute(text(select_sql(v, max_date))).mappings().fetchall()]
        for re in result:
            re_sum = float(re['合计'])
            re_24 = float(re['24年同期'])
            re['增长率'] = '0'
            if re_sum and re_24:
                re['增长率'] = f'{int(round((re_sum - re_24)/re_24*100, 0))}%'
        result = {re['区域']: re for re in result}
        result_data[k] = result

    return result_data


def result_table_five(engine):
    sql = '''
        WITH base_data AS(
            SELECT
                CASE WHEN secondary_dealer != '' AND secondary_dealer IS NOT NULL THEN secondary_dealer
                ELSE '直客'
                END AS secondary_dealer_re,
                customer_name,
                MONTH(performance_date) AS month_re,
                ROUND(COALESCE(SUM(sales_amount),0)/10000, 1) AS sales_amount
            FROM hw_two_five_data
            WHERE
                sales_team IN ('中长尾', '电网销')
                AND is_traffic_product IN ('否','')
            GROUP BY secondary_dealer_re, customer_name, month_re
        )
        SELECT 
            secondary_dealer_re AS `渠道`,
            customer_name AS `客户`,
            COALESCE(SUM(CASE WHEN month_re = 1 THEN sales_amount ELSE 0 END), 0) AS `1月`,
            COALESCE(SUM(CASE WHEN month_re = 2 THEN sales_amount ELSE 0 END), 0) AS `2月`,
            COALESCE(SUM(CASE WHEN month_re = 3 THEN sales_amount ELSE 0 END), 0) AS `3月`,
            COALESCE(SUM(CASE WHEN month_re = 4 THEN sales_amount ELSE 0 END), 0) AS `4月`,
            COALESCE(SUM(CASE WHEN month_re = 5 THEN sales_amount ELSE 0 END), 0) AS `5月`,
            COALESCE(SUM(CASE WHEN month_re = 6 THEN sales_amount ELSE 0 END), 0) AS `6月`,
            COALESCE(SUM(CASE WHEN month_re = 7 THEN sales_amount ELSE 0 END), 0) AS `7月`,
            COALESCE(SUM(CASE WHEN month_re = 8 THEN sales_amount ELSE 0 END), 0) AS `8月`,
            COALESCE(SUM(CASE WHEN month_re = 9 THEN sales_amount ELSE 0 END), 0) AS `9月`,
            COALESCE(SUM(CASE WHEN month_re = 10 THEN sales_amount ELSE 0 END), 0) AS `10月`,
            COALESCE(SUM(CASE WHEN month_re = 11 THEN sales_amount ELSE 0 END), 0) AS `11月`,
            COALESCE(SUM(CASE WHEN month_re = 12 THEN sales_amount ELSE 0 END), 0) AS `12月`,
            COALESCE(SUM(sales_amount), 0) AS 合计
        FROM
            base_data
        GROUP BY secondary_dealer_re, customer_name
        ORDER BY 合计 DESC
    '''
    result = [dict(row) for row in engine.connect().execute(text(sql)).mappings().fetchall()]
    return result


def result_table_six(engine, max_date):
    sql = f'''
        WITH 
        base_data_25 AS(
            SELECT
                CASE WHEN secondary_dealer != '' AND secondary_dealer IS NOT NULL THEN secondary_dealer
                    ELSE customer_name
                END AS secondary_dealer_re,
                ROUND(COALESCE(SUM(sales_amount),0)/10000, 1) AS sales_amount
            FROM hw_two_five_data
            WHERE
                sales_team IN ('中长尾', '电网销')
                AND is_traffic_product IN ('否','')
            GROUP BY secondary_dealer_re
        ),
        base_data_24 AS(
            SELECT
                CASE WHEN secondary_dealer != '' AND secondary_dealer IS NOT NULL THEN secondary_dealer
                    ELSE customer_name
                END AS secondary_dealer_re,
                ROUND(COALESCE(SUM(sales_amount),0)/10000, 1) AS sales_amount
            FROM hw_two_four_data_smbcore
            WHERE
                performance_date BETWEEN DATE(CONCAT(YEAR(CURDATE()) - 1, '-01-01')) AND '{max_date}'
            GROUP BY secondary_dealer_re
        )
        
        SELECT
            bd25.secondary_dealer_re AS `SMBcore业绩`,
            bd25.sales_amount AS `25年截止目前业绩`,
            bd24.sales_amount AS `24年同期业绩`,
            CASE 
                WHEN bd24.sales_amount IS NULL THEN NULL
                ELSE CONCAT(ROUND((bd25.sales_amount - bd24.sales_amount) / bd24.sales_amount * 100, 0), '%')
            END AS `同期增长率`,
            bd25.sales_amount - IFNULL(bd24.sales_amount,0) AS `同比24年正负值`
        FROM
        base_data_25 bd25
        LEFT JOIN base_data_24 bd24 ON bd25.secondary_dealer_re = bd24.secondary_dealer_re
    '''
    result = [dict(row) for row in engine.connect().execute(text(sql)).mappings().fetchall()]
    result = {re['SMBcore业绩']: re for re in result}
    return result


def result_table_seven(engine, max_date):
    sql = f'''
        SELECT
            main.product,
            COALESCE(SUM(CASE WHEN current_year.quarter = 'Q1' THEN sales_amount_q ELSE 0 END), 0) AS 25Q1,
            COALESCE(SUM(CASE WHEN current_year.quarter = 'Q2' THEN sales_amount_q ELSE 0 END), 0) AS 25Q2,
            COALESCE(SUM(CASE WHEN current_year.quarter = 'Q3' THEN sales_amount_q ELSE 0 END), 0) AS 25Q3,
            COALESCE(SUM(CASE WHEN current_year.quarter = 'Q4' THEN sales_amount_q ELSE 0 END), 0) AS 25Q4,
            COALESCE(SUM(sales_amount_q), 0) AS `25年目前业绩`,
            COALESCE(last_year.same_performance_24, 0) AS `24年同期业绩`,
            CONCAT(ROUND((COALESCE(SUM(sales_amount_q), 0) - COALESCE(last_year.same_performance_24, 0)) 
                / NULLIF(COALESCE(last_year.same_performance_24, 0), 0) * 100, 0), '%') AS 同比增长
        FROM (
            SELECT DISTINCT leased_line_product AS product 
            FROM hw_two_five_data
            WHERE leased_line_product IN ('EI', 'PaaS', '安全', '媒体', '数据库', '网络')
        ) AS main
        LEFT JOIN (
            SELECT 
                leased_line_product,
                ROUND(SUM(sales_amount)/10000, 1) AS sales_amount_q,
                quarter
            FROM hw_two_five_data
            WHERE sales_team IN ('中长尾', '电网销')
            GROUP BY leased_line_product,quarter
        ) AS current_year ON main.product = current_year.leased_line_product
        LEFT JOIN (
            SELECT 
                leased_line_product,
                ROUND(SUM(sales_amount)/10000, 1) AS same_performance_24
            FROM hw_two_four_data
            WHERE sales_team IN ('中长尾', '电网销')
                        AND performance_date BETWEEN DATE(CONCAT(YEAR(CURDATE()) - 1, '-01-01')) AND '{max_date}'
            GROUP BY leased_line_product
        ) AS last_year ON main.product = last_year.leased_line_product
        GROUP BY product, last_year.same_performance_24
        
        UNION ALL
        
        SELECT 
            '企业协同' AS product,
            COALESCE(SUM(CASE WHEN quarter = 'Q1' THEN sales_amount_q ELSE 0 END), 0) AS 25Q1,
            COALESCE(SUM(CASE WHEN quarter = 'Q2' THEN sales_amount_q ELSE 0 END), 0) AS 25Q2,
            COALESCE(SUM(CASE WHEN quarter = 'Q3' THEN sales_amount_q ELSE 0 END), 0) AS 25Q3,
            COALESCE(SUM(CASE WHEN quarter = 'Q4' THEN sales_amount_q ELSE 0 END), 0) AS 25Q4,
            COALESCE(SUM(sales_amount_q), 0) AS `25年目前业绩`,
            COALESCE(same_performance_24, 0) AS `24年同期业绩`,
            CONCAT(ROUND(
                (COALESCE(SUM(sales_amount_q), 0) - COALESCE(same_performance_24, 0)) 
                / NULLIF(COALESCE(same_performance_24, 0), 0) * 100, 0
            ),'%') AS 同比增长
        FROM (
            SELECT 
                quarter,
                ROUND(SUM(sales_amount)/10000, 1) AS sales_amount_q
            FROM hw_two_five_data
            WHERE 
                enterprise_coop IS NOT NULL AND enterprise_coop <> ''
                AND special_rebate_type <> '特殊商务'
            GROUP BY quarter
        ) AS current_year_coop
        LEFT JOIN (
            SELECT 
                ROUND(SUM(sales_amount)/10000, 1) AS same_performance_24
            FROM hw_two_four_data
            WHERE 
                enterprise_coop IS NOT NULL AND enterprise_coop <> ''
                AND special_rebate_type <> '特殊商务'
                AND performance_date BETWEEN DATE(CONCAT(YEAR(CURDATE()) - 1, '-01-01')) 
                AND '{max_date}'
        ) AS last_year_coop ON 1=1
        GROUP BY same_performance_24
        
        ORDER BY FIELD(product, 'EI', 'PaaS', '安全', '媒体', '数据库', '网络', '企业协同');
    '''
    result = [dict(row) for row in engine.connect().execute(text(sql)).mappings().fetchall()]
    result = {re['product']: re for re in result}
    return result


def result_table_eight(engine):
    sql = '''
        SELECT
            secondary_dealer AS 新增渠道,
            ROUND(SUM(sales_amount)/10000, 1) AS 业绩金额,
            ROUND(SUM(IF(sales_team = '华为云NA', sales_amount, 0))/10000, 1) AS NA业绩,
            ROUND(SUM(IF(sales_team IN ('中长尾', '电网销'), sales_amount, 0))/10000, 1) AS SMB业绩,
            ROUND(SUM(IF(sales_team IN ('中长尾', '电网销') 
                       AND is_traffic_product IN ('否',''), sales_amount, 0))/10000, 1) AS SMBcore业绩,
            GROUP_CONCAT(DISTINCT salesperson) AS 销售员
        FROM hw_two_five_data
        WHERE secondary_dealer NOT IN (
            SELECT DISTINCT secondary_dealer 
            FROM hw_two_four_data
                WHERE secondary_dealer IS NOT NULL
        )
        GROUP BY secondary_dealer
        ORDER BY 业绩金额 DESC
    '''
    result = [dict(row) for row in engine.connect().execute(text(sql)).mappings().fetchall()]
    return result


def result_table_nine(engine):
    sql = '''
        WITH new_customers AS (
            SELECT DISTINCT customer_name
            FROM hw_two_five_data
            WHERE customer_name NOT IN (
                SELECT DISTINCT customer_name 
                FROM hw_two_four_data
                        WHERE customer_name IS NOT NULL
            )
        ),
        recent_info AS (
            SELECT 
                customer_name,
                secondary_dealer,
                customer_tag,
                salesperson
            FROM (
                SELECT 
                    customer_name,
                    secondary_dealer,
                    customer_tag,
                    salesperson,
                    ROW_NUMBER() OVER (
                        PARTITION BY customer_name 
                        ORDER BY performance_date DESC
                    ) AS rn
                FROM 
                    hw_two_five_data
                WHERE 
                    customer_name IN (SELECT customer_name FROM new_customers)
            ) t
            WHERE 
                rn = 1
        )
        SELECT
            five.customer_name AS `新增客户`,
            ri.secondary_dealer AS `渠道名称`,
            ROUND(SUM(five.sales_amount)/10000, 1) AS `业绩金额`,
            ROUND(SUM(CASE WHEN five.sales_team = '华为云NA' THEN five.sales_amount ELSE 0 END)/10000, 1) AS `NA业绩`,
            ROUND(SUM(CASE WHEN five.sales_team IN ('中长尾', '电网销') THEN five.sales_amount ELSE 0 END)/10000, 1) AS `SMB业绩`,
            ROUND(SUM(CASE 
                    WHEN five.sales_team IN ('中长尾', '电网销') 
                    AND five.is_traffic_product IN ('否','') 
                    THEN five.sales_amount ELSE 0 
                END)/10000, 1) AS `SMB-CORE`,
            ri.salesperson AS `销售员`,
            ri.customer_tag AS `客户标签`
        FROM 
            hw_two_five_data five
        INNER JOIN 
            new_customers nc ON five.customer_name = nc.customer_name
        LEFT JOIN 
            recent_info ri ON five.customer_name = ri.customer_name
        GROUP BY 
            five.customer_name, 
            ri.secondary_dealer, 
            ri.salesperson, 
            ri.customer_tag
        ORDER BY `业绩金额` DESC
    '''
    result = [dict(row) for row in engine.connect().execute(text(sql)).mappings().fetchall()]
    return result
