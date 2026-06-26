-- ========================================
-- 报表系统数据库表结构
-- 创建日期：2026-06-03
-- 数据库名称：bth-report（与 .env 中 MYSQL_DATABASE 一致）
-- 本文件须保存为 UTF-8（无 BOM）
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;

-- 创建数据库（名称含 - 须用反引号）
CREATE DATABASE IF NOT EXISTS `bth-report` DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `bth-report`;

-- ========================================
-- 表 1：产品基础信息表
-- ========================================
DROP TABLE IF EXISTS product_info;
CREATE TABLE product_info (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    sku VARCHAR(100) NOT NULL COMMENT '产品编码（SKU）',
    product_id VARCHAR(100) COMMENT '商品ID',
    ean VARCHAR(50) COMMENT 'EAN码',
    category_l2 VARCHAR(100) COMMENT '二级分类',
    category_l3 VARCHAR(100) COMMENT '三级分类',
    supplier VARCHAR(100) COMMENT '供应商',
    operation_mode VARCHAR(50) COMMENT '运营模式',
    product_status VARCHAR(50) COMMENT '产品状态',
    purchase_price DECIMAL(10, 2) COMMENT '原始采购价（人民币）',
    declared_price DECIMAL(10, 2) COMMENT '申报价格',
    is_new_product_amz TINYINT DEFAULT 0 COMMENT 'AMZ新老品（1=新品，0=老品）',
    is_new_product_local TINYINT DEFAULT 0 COMMENT '本土平台新老品（1=新品，0=老品）',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    UNIQUE KEY uk_sku (sku),
    INDEX idx_product_id (product_id),
    INDEX idx_ean (ean),
    INDEX idx_category (category_l2, category_l3),
    INDEX idx_supplier (supplier)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='产品基础信息表';

-- ========================================
-- 表 2：站点映射表
-- ========================================
DROP TABLE IF EXISTS site_mapping;
CREATE TABLE site_mapping (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    account_name VARCHAR(100) NOT NULL COMMENT '账号英文名',
    site_code VARCHAR(20) NOT NULL COMMENT '站点代码（DE/FR/US/UK等）',
    platform VARCHAR(50) NOT NULL COMMENT '平台（AMAZON/OTTO/REAL/MANO等）',
    shop_name VARCHAR(100) COMMENT '店铺名称',
    shop_cn_name VARCHAR(100) COMMENT '店铺中文名',
    site_identifier VARCHAR(150) COMMENT '站点识别码（儿子-站点）',
    country_code VARCHAR(10) COMMENT '国家或地区代码',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    UNIQUE KEY uk_account_site_platform (account_name, site_code, platform),
    INDEX idx_platform (platform),
    INDEX idx_site_code (site_code),
    INDEX idx_site_identifier (site_identifier)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='站点账号映射表';

-- ========================================
-- 表 3：物流费用映射表
-- ========================================
DROP TABLE IF EXISTS logistics_mapping;
CREATE TABLE logistics_mapping (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    sku VARCHAR(100) NOT NULL COMMENT '产品编码',
    site_code VARCHAR(20) NOT NULL COMMENT '站点代码',
    logistics_type VARCHAR(50) NOT NULL COMMENT '物流类型（头程/尾程/MF/FBA/4PX等）',
    logistics_category VARCHAR(50) COMMENT '物流分类（MF-COMMF/MF-OHPAMF/非MF等）',
    fee_rmb DECIMAL(10, 2) COMMENT '费用（人民币）',
    fee_eur DECIMAL(10, 2) COMMENT '费用（欧元）',
    fee_usd DECIMAL(10, 2) COMMENT '费用（美元）',
    tariff_included DECIMAL(10, 2) COMMENT '关税（含税）',
    effective_date DATE COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    remark VARCHAR(200) COMMENT '备注',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    INDEX idx_sku_site (sku, site_code),
    INDEX idx_logistics_type (logistics_type),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='物流费用映射表';

-- ========================================
-- 表 4：销售负责人映射表
-- ========================================
DROP TABLE IF EXISTS sales_owner_mapping;
CREATE TABLE sales_owner_mapping (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    sku VARCHAR(100) COMMENT '产品编码',
    platform VARCHAR(50) COMMENT '平台',
    site_code VARCHAR(20) COMMENT '站点',
    mapping_type VARCHAR(50) NOT NULL COMMENT '映射类型（平台/站点/SKU）',
    sales_owner VARCHAR(100) NOT NULL COMMENT '销售负责人',
    sales_manager VARCHAR(100) COMMENT '销售经理',
    effective_date DATE COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    INDEX idx_sku (sku),
    INDEX idx_platform_site (platform, site_code),
    INDEX idx_sales_owner (sales_owner),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='销售负责人映射表';

-- ========================================
-- 表 5：VAT 和佣金费率表
-- ========================================
DROP TABLE IF EXISTS fee_rate_mapping;
CREATE TABLE fee_rate_mapping (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    platform VARCHAR(50) NOT NULL COMMENT '平台',
    site_code VARCHAR(20) NOT NULL COMMENT '站点',
    category VARCHAR(100) COMMENT '商品分类（如果按分类区分）',
    vat_rate DECIMAL(6, 4) COMMENT 'VAT税率（如 0.19 表示 19%）',
    commission_rate DECIMAL(6, 4) COMMENT '佣金率',
    withdrawal_fee_base DECIMAL(10, 2) COMMENT '提现费基础费用',
    withdrawal_fee_rate DECIMAL(6, 4) COMMENT '提现费比例',
    effective_date DATE NOT NULL COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    remark VARCHAR(200) COMMENT '备注',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    INDEX idx_platform_site (platform, site_code),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='VAT和佣金费率表';

-- ========================================
-- 表 6：汇率配置表
-- ========================================
DROP TABLE IF EXISTS exchange_rate;
CREATE TABLE exchange_rate (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    currency_from VARCHAR(10) NOT NULL COMMENT '源币种（RMB/USD/CAD/GBP等）',
    currency_to VARCHAR(10) NOT NULL COMMENT '目标币种（EUR）',
    rate DECIMAL(10, 6) NOT NULL COMMENT '汇率',
    rate_type VARCHAR(20) DEFAULT 'divide' COMMENT '汇率类型（divide=除法，multiply=乘法）',
    effective_date DATE NOT NULL COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    UNIQUE KEY uk_currency_date (currency_from, currency_to, effective_date),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='汇率配置表';

-- ========================================
-- 表 7：定价映射表
-- ========================================
DROP TABLE IF EXISTS price_mapping;
CREATE TABLE price_mapping (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    sku VARCHAR(100) NOT NULL COMMENT '产品编码',
    site_code VARCHAR(20) NOT NULL COMMENT '站点代码',
    platform VARCHAR(50) COMMENT '平台',
    price_eur DECIMAL(10, 2) COMMENT '产品单价（欧元）',
    price_usd DECIMAL(10, 2) COMMENT '产品单价（美元）',
    shipping_fee_eur DECIMAL(10, 2) COMMENT '运费回款（欧元）',
    effective_date DATE COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    INDEX idx_sku_site (sku, site_code),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='定价映射表';

-- ========================================
-- 表 8：仓租映射表
-- ========================================
DROP TABLE IF EXISTS warehouse_rent_mapping;
CREATE TABLE warehouse_rent_mapping (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    sku VARCHAR(100) NOT NULL COMMENT '产品编码',
    warehouse_type VARCHAR(50) NOT NULL COMMENT '仓库类型（HY/4PX/MF等）',
    site_code VARCHAR(20) COMMENT '站点代码',
    rent_eur DECIMAL(10, 2) COMMENT '仓租（欧元）',
    rent_rmb DECIMAL(10, 2) COMMENT '仓租（人民币）',
    effective_date DATE COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    INDEX idx_sku (sku),
    INDEX idx_warehouse_type (warehouse_type),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='仓租映射表';

-- ========================================
-- 表 9：测评费用映射表
-- ========================================
DROP TABLE IF EXISTS evaluation_fee_mapping;
CREATE TABLE evaluation_fee_mapping (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    refund_type VARCHAR(50) NOT NULL COMMENT '退款类型（空包退/测评退/测评退70%/佣金/好评返现）',
    calculation_rule VARCHAR(200) NOT NULL COMMENT '计算规则说明',
    base_fee DECIMAL(10, 2) COMMENT '基础费用',
    rate DECIMAL(6, 4) COMMENT '费率',
    currency VARCHAR(10) DEFAULT 'EUR' COMMENT '币种',
    effective_date DATE NOT NULL COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    INDEX idx_refund_type (refund_type),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='测评费用映射表';

-- ========================================
-- 表 10：二次上架规则表
-- ========================================
DROP TABLE IF EXISTS relisting_rules;
CREATE TABLE relisting_rules (
    id INT PRIMARY KEY AUTO_INCREMENT COMMENT '主键ID',
    platform VARCHAR(50) NOT NULL COMMENT '平台（OTTO/其他）',
    product_condition VARCHAR(50) NOT NULL COMMENT '商品状态（良品/不良品）',
    sku_suffix VARCHAR(20) COMMENT 'SKU后缀（如-NW）',
    apply_purchase_cost TINYINT DEFAULT 1 COMMENT '是否应用采购成本（1=是，0=否）',
    remark VARCHAR(200) COMMENT '备注说明',
    effective_date DATE NOT NULL COMMENT '生效日期',
    expired_date DATE COMMENT '失效日期（NULL 表示当前有效）',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
    INDEX idx_platform (platform),
    INDEX idx_effective_date (effective_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='二次上架采购成本规则表';

-- ========================================
-- 初始化数据
-- ========================================

-- 初始化汇率数据（示例）
INSERT INTO exchange_rate (currency_from, currency_to, rate, rate_type, effective_date) VALUES
('RMB', 'EUR', 7.3, 'divide', '2026-01-01'),
('USD', 'EUR', 0.858, 'multiply', '2026-01-01'),
('CAD', 'EUR', 0.6179, 'multiply', '2026-01-01'),
('GBP', 'EUR', 1.15, 'multiply', '2026-01-01'),
('CZK', 'EUR', 0.04133, 'multiply', '2026-01-01'),  -- 捷克克朗
('PLN', 'EUR', 0.237, 'multiply', '2026-01-01'),    -- 波兰兹罗提
('HUF', 'EUR', 0.002611, 'multiply', '2026-01-01'), -- 匈牙利福林
('SEK', 'EUR', 0.0934, 'multiply', '2026-01-01'),   -- 瑞典克朗
('RON', 'EUR', 0.196, 'multiply', '2026-01-01');    -- 罗马尼亚列伊

-- 初始化测评费用规则（根据 updateLog.md）
INSERT INTO evaluation_fee_mapping (refund_type, calculation_rule, base_fee, rate, effective_date) VALUES
('空包退订单金额', '平台佣金、VAT、提现费', 0.34, 0.05, '2026-01-01'),
('测评退订单金额', '订单金额*1.05+0.34', 0.34, 1.05, '2026-01-01'),
('测评退订单金额70%', '订单金额(*0.7)*1.05+0.34', 0.34, 1.05, '2026-01-01'),
('佣金', 'RMB金额兑换为欧元(￥/7.3)', NULL, NULL, '2026-01-01'),
('好评返现', 'RMB金额兑换为欧元(￥/7.3)', NULL, NULL, '2026-01-01');

-- 初始化二次上架规则（根据 updateLog.md）
INSERT INTO relisting_rules (platform, product_condition, sku_suffix, apply_purchase_cost, remark, effective_date) VALUES
('OTTO', '良品', NULL, 1, 'OTTO平台良品加采购成本', '2026-01-01'),
('OTTO', '不良品', NULL, 0, 'OTTO平台不良品采购成本为0', '2026-01-01'),
('OTTO', '任意', '-NW', 0, 'OTTO平台-NW尾缀采购成本为0', '2026-01-01'),
('其他平台', '任意', '-NW', 0, '其他平台-NW尾缀采购成本为0', '2026-01-01'),
('其他平台', '任意', '非-NW', 1, '其他平台非NW加采购成本', '2026-01-01');

-- ========================================
-- 视图：当前有效汇率
-- ========================================
CREATE OR REPLACE VIEW v_current_exchange_rate AS
SELECT 
    currency_from,
    currency_to,
    rate,
    rate_type,
    effective_date
FROM exchange_rate
WHERE (expired_date IS NULL OR expired_date > CURDATE())
  AND effective_date <= CURDATE()
ORDER BY effective_date DESC;

-- ========================================
-- 存储过程：查询映射（示例）
-- ========================================
DELIMITER //

CREATE PROCEDURE sp_get_product_mapping(
    IN p_sku VARCHAR(100),
    IN p_target_field VARCHAR(50)
)
BEGIN
    IF p_target_field = '商品ID' THEN
        SELECT product_id FROM product_info WHERE sku = p_sku;
    ELSEIF p_target_field = '二级分类' THEN
        SELECT category_l2 FROM product_info WHERE sku = p_sku;
    ELSEIF p_target_field = '采购价' THEN
        SELECT purchase_price FROM product_info WHERE sku = p_sku;
    ELSE
        SELECT NULL;
    END IF;
END //

DELIMITER ;

-- ========================================
-- 完成提示
-- ========================================
SELECT '✅ 数据库表结构创建完成！' AS message;
SELECT '📊 共创建 10 张表、1 个视图、1 个存储过程' AS summary;
SELECT '🔗 下一步：运行 migrate_excel_to_db.py 迁移数据' AS next_step;
