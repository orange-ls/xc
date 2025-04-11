CREATE TABLE IF NOT EXISTS hw_two_five_data (
    performance_id VARCHAR(255) NOT NULL COMMENT '业绩ID',
    sales_amount DECIMAL(16,2) COMMENT '业绩金额(¥)',
    performance_date DATE COMMENT '业绩形成时间',
    secondary_dealer VARCHAR(255) COMMENT '二级经销商名称',
		customer_name VARCHAR(100) COMMENT '客户名称',
    product_code VARCHAR(50) COMMENT '产品类型编码',
    customer_tag VARCHAR(20) COMMENT '客户标签',
    sales_team VARCHAR(100) COMMENT '销售纵队',
    service_department VARCHAR(100) COMMENT '服务产品部',
    is_traffic_product VARCHAR(10) COMMENT '是否流量型产品（是/否）',
    leased_line_product VARCHAR(20) COMMENT '专线产品',
    enterprise_coop VARCHAR(100) COMMENT '企业协同',
    salesperson VARCHAR(50) COMMENT '销售员',
    region VARCHAR(50) COMMENT '区域',
    quarter VARCHAR(15) COMMENT '季度',
    PRIMARY KEY (performance_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='25年业绩表';


CREATE TABLE IF NOT EXISTS hw_two_four_data (
    performance_id VARCHAR(255) NOT NULL COMMENT '业绩ID',
    sales_amount DECIMAL(16,2) COMMENT '业绩金额(¥)',
    performance_date DATE COMMENT '业绩形成时间',
    secondary_dealer VARCHAR(255) COMMENT '二级经销商名称',
		customer_name VARCHAR(100) COMMENT '客户名称',
    product_code VARCHAR(50) COMMENT '产品类型编码',
    customer_tag VARCHAR(20) COMMENT '客户标签',
    sales_team VARCHAR(100) COMMENT '销售纵队',
    service_department VARCHAR(100) COMMENT '服务产品部',
    is_traffic_product VARCHAR(10) COMMENT '是否流量型产品（是/否）',
    leased_line_product VARCHAR(20) COMMENT '专线产品',
    enterprise_coop VARCHAR(100) COMMENT '企业协同',
    salesperson VARCHAR(50) COMMENT '销售员',
    region VARCHAR(50) COMMENT '区域',
    quarter VARCHAR(15) COMMENT '季度',
    PRIMARY KEY (performance_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='24年业绩表';


CREATE TABLE IF NOT EXISTS hw_two_four_data_smbcore (
    performance_id VARCHAR(255) NOT NULL COMMENT '业绩ID',
    sales_amount DECIMAL(16,2) COMMENT '业绩金额(¥)',
    performance_date DATE COMMENT '业绩形成时间',
    secondary_dealer VARCHAR(255) COMMENT '二级经销商名称',
		customer_name VARCHAR(100) COMMENT '客户名称',
    product_code VARCHAR(50) COMMENT '产品类型编码',
    customer_tag VARCHAR(20) COMMENT '客户标签',
    sales_team VARCHAR(100) COMMENT '销售纵队',
    service_department VARCHAR(100) COMMENT '服务产品部',
    is_traffic_product VARCHAR(10) COMMENT '是否流量型产品（是/否）',
    leased_line_product VARCHAR(20) COMMENT '专线产品',
    enterprise_coop VARCHAR(100) COMMENT '企业协同',
    salesperson VARCHAR(50) COMMENT '销售员',
    region VARCHAR(50) COMMENT '区域',
    quarter VARCHAR(15) COMMENT '季度',
    PRIMARY KEY (performance_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='24年业绩表SMBcore';


CREATE TABLE IF NOT EXISTS hw_two_four_data_na (
    performance_id VARCHAR(255) NOT NULL COMMENT '业绩ID',
    sales_amount DECIMAL(16,2) COMMENT '业绩金额(¥)',
    performance_date DATE COMMENT '业绩形成时间',
    secondary_dealer VARCHAR(255) COMMENT '二级经销商名称',
		customer_name VARCHAR(100) COMMENT '客户名称',
    product_code VARCHAR(50) COMMENT '产品类型编码',
    customer_tag VARCHAR(20) COMMENT '客户标签',
    sales_team VARCHAR(100) COMMENT '销售纵队',
    service_department VARCHAR(100) COMMENT '服务产品部',
    is_traffic_product VARCHAR(10) COMMENT '是否流量型产品（是/否）',
    leased_line_product VARCHAR(20) COMMENT '专线产品',
    enterprise_coop VARCHAR(100) COMMENT '企业协同',
    salesperson VARCHAR(50) COMMENT '销售员',
    region VARCHAR(50) COMMENT '区域',
    quarter VARCHAR(15) COMMENT '季度',
    PRIMARY KEY (performance_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='24年业绩表NA';


CREATE TABLE IF NOT EXISTS customer_correspondence (
    customer_name VARCHAR(100) COMMENT '客户名称',
    salesperson VARCHAR(50) COMMENT '销售员',
    region VARCHAR(50) COMMENT '区域'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='客户对应关系表';


CREATE TABLE IF NOT EXISTS two_five_details_cloud_services (
    cloud_services_code VARCHAR(50) COMMENT '云服务编码',
    service_department VARCHAR(50) COMMENT '服务产品部'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='25产品明细_云服务名称';


CREATE TABLE IF NOT EXISTS two_five_details_flow (
    product_code VARCHAR(50) COMMENT '产品类型编码',
    product_type VARCHAR(50) COMMENT '产品类型'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='25产品明细_流量产品清单';


CREATE TABLE IF NOT EXISTS two_five_details_special (
    product_code VARCHAR(50) COMMENT 'L4层产品编码',
    product_name VARCHAR(50) COMMENT '名称'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='25产品明细_产品专项';


CREATE TABLE IF NOT EXISTS two_five_details_collaborate (
    cloud_services_code VARCHAR(50) COMMENT '云服务编码',
    cloud_services_name VARCHAR(50) COMMENT '云服务名称'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
COMMENT='25产品明细_企业协同';
