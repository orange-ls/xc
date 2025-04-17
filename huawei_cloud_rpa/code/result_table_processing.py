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

def result_table_one(engine):
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
                    COALESCE(SUM(d.sales_amount), 0) AS total_sales
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
    growth_rate_sql = '''
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
                WHERE d.performance_date BETWEEN DATE(CONCAT(YEAR(CURDATE()) - 1, '-01-01')) AND DATE_SUB(CURDATE(), INTERVAL 1 YEAR)
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
                    ELSE ROUND((d25.amount_2025 - d24.amount_2024) / d24.amount_2024 * 100, 2)
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
                    (SUM(amount_2025) - SUM(amount_2024)) / NULLIF(SUM(amount_2024), 0) * 100, 2
                ) AS all_sales,
                ROUND(
                    (SUM(CASE WHEN ptype = 'NA业绩' THEN amount_2025 END) - 
                     SUM(CASE WHEN ptype = 'NA业绩' THEN amount_2024 END)) / 
                    NULLIF(SUM(CASE WHEN ptype = 'NA业绩' THEN amount_2024 END), 0) * 100, 2
                ) AS na_sales,
                ROUND(
                    (SUM(CASE WHEN ptype = 'SMB业绩' THEN amount_2025 END) - 
                     SUM(CASE WHEN ptype = 'SMB业绩' THEN amount_2024 END)) / 
                    NULLIF(SUM(CASE WHEN ptype = 'SMB业绩' THEN amount_2024 END), 0) * 100, 2
                ) AS smb_sales,
                ROUND(
                    (SUM(CASE WHEN ptype = 'SMBcore业绩' THEN amount_2025 END) - 
                     SUM(CASE WHEN ptype = 'SMBcore业绩' THEN amount_2024 END)) / 
                    NULLIF(SUM(CASE WHEN ptype = 'SMBcore业绩' THEN amount_2024 END), 0) * 100, 2
                ) AS smbcore_sales
            FROM combined_data
        )
        
        -- 8. 最终结果
        SELECT 
            region,
                CONCAT(IFNULL(all_sales, 'N/A'), '%') AS all_sales,
            CONCAT(IFNULL(na_sales, 'N/A'), '%') AS na_sales,
            CONCAT(IFNULL(smb_sales, 'N/A'), '%') AS smb_sales,
            CONCAT(IFNULL(smbcore_sales, 'N/A'), '%') AS smbcore_sales
        FROM (
            SELECT * FROM pivot_table
            UNION ALL
            SELECT * FROM total_row
        ) AS final
        ORDER BY FIELD(region, '北京','广州','深圳','上海','南京','成都','其他','总计');
    '''
    result = [dict(row) for row in conn.execute(text(growth_rate_sql)).mappings().fetchall()]
    result = {re['region']: re for re in result}
    result_data['同期增长率'] = result

    conn.close()
    return result_data

# def result_table_two(engine):