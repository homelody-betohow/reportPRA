-- 用途：产品 SKU 定价/成本主数据表
-- 数据来源：\\Betohow\数据报表\数据库\BTH全部SKU明细-v*.xlsx → sheet「基础数据维护」
-- 表头：第 1 行（一级表头）+ 第 2 行（二级表头），数据从第 3 行开始，共约 5203 条
--
-- 列映射说明（Excel列索引 → 字段名）：
--   [0]  仓库SKU（第1列，部分为空，部分与SKU相同）   → warehouse_sku
--   [1]  SKU                                         → product_sku  ← 业务唯一键
--   [2]  成本价                                       → cost_price_cny
--   [3]  重量（g）                                    → unit_weight_g
--   [4]  规格（欧规/通用/美规/英规/日本/澳规/常规/配件/国内配置）→ region_spec
--   [5]  内箱 长（cm）                                → inner_box_l_cm
--   [6]  内箱 宽（cm）                                → inner_box_w_cm
--   [7]  内箱 高（cm）                                → inner_box_h_cm
--   [8]  外箱 长（cm）                                → outer_box_l_cm
--   [9]  外箱 宽（cm）                                → outer_box_w_cm
--   [10] 外箱 高（cm）                                → outer_box_h_cm
--   [11] 外箱 每箱数量（个）                           → carton_qty
--   [12] 外箱 箱规毛重（g）                            → carton_gross_g
--   [13] 头程（RMB） EU/AU                            → first_leg_eu_au_cny
--   [14] 头程（RMB） US                               → first_leg_us_cny
--   [15] 头程（RMB） CA                               → first_leg_ca_cny
--   [16] 头程（RMB） JP                               → first_leg_jp_cny
--   [17] 头程（RMB） UK                               → first_leg_uk_cny
--   [18] 关税（含税） EU                              → duty_eu_cny
--   [19] 关税（含税） US                              → duty_us_cny
--   [20] 关税（含税） CA/AU                           → duty_ca_au_cny
--   [21] 关税（含税） JP                              → duty_jp_cny
--   [22] 关税（含税） UK                              → duty_uk_cny
--   [23] 关税（不含税） EU                            → duty_eu_notax_cny
--   [24] 关税（不含税） US                            → duty_us_notax_cny
--   [25] 关税（不含税） CA/AU                         → duty_ca_au_notax_cny
--   [26] 关税（不含税） JP                            → duty_jp_notax_cny
--   [27] 关税（不含税） UK                            → duty_uk_notax_cny
--   [28] 原始采购价                                   → purchase_price_orig_cny
--   [29] 供应商代码                                   → supplier_code
--   [30] 供应商（简称）                               → supplier_name
--   [31] 品类                                        → category
--   [32] 运营模式（代运营 / 自运营）                   → ops_mode
--   [33] 产品销售状态（正常销售/清仓/重点产品/新品/仅售后配件/正常售完下架）→ sales_status
--   [34] 名称                                        → product_name
--   [35] 默认采购价                                   → purchase_price_default_cny
--   [36] 开发负责人                                   → dev_owner
--   [37] 供应商（含代码全称，如 F003[江门市...]）      → supplier_full_name
--   [38] 产品开发时间                                 → product_dev_date
--   [39] 产品上架时间                                 → product_launch_date
--   [40] 代运营佣金点（常规）                          → commission_pct_regular
--   [41] 5月佣金点（DE）                              → commission_pct_may_de
--   [42] 5月佣金点（非DE）                            → commission_pct_may_non_de
--   [43] 5月采购价                                   → purchase_price_may_cny
--   [44] 6月佣金点（DE）                              → commission_pct_jun_de
--   [45] 6月佣金点（非DE）                            → commission_pct_jun_non_de
--   [46] 6月采购价                                   → purchase_price_jun_cny
--   [47] 9月佣金点（DE）                              → commission_pct_sep_de
--   [48] 9月佣金点（非DE）                            → commission_pct_sep_non_de
--   [49] 9月采购价                                   → purchase_price_sep_cny
--   [50] 单价修改 → 修改后单价                        → price_modified_cny
--   [51] 单价修改 → 预计修改时间                      → price_modified_date
--
-- 导入策略：
--   - line_hash 为业务行内容稳定哈希，UPSERT 时用于变更检测
--   - 业务幂等键为 product_sku（UNIQUE KEY）

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS product_sku_pricing (
  id         BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash  CHAR(64)        NULL COMMENT '行内容稳定哈希（SHA-256 hex），变更检测用',

  -- ===== 业务标识 =====
  warehouse_sku   VARCHAR(64)  NULL COMMENT '仓库SKU（第1列，部分为空或与product_sku相同）',
  product_sku     VARCHAR(64)  NOT NULL COMMENT '产品SKU（业务唯一键，Excel第2列）',

  -- ===== 基础物理参数 =====
  cost_price_cny  DECIMAL(18,6) NULL COMMENT '成本价（人民币）',
  unit_weight_g   DECIMAL(12,2) NULL COMMENT '单件重量（g）',
  region_spec     VARCHAR(16)   NULL COMMENT '规格/适用地区（欧规/美规/英规/日本/澳规/通用/常规/配件/国内配置）',

  -- ===== 内箱尺寸（cm）=====
  inner_box_l_cm  DECIMAL(10,2) NULL COMMENT '内箱长（cm）',
  inner_box_w_cm  DECIMAL(10,2) NULL COMMENT '内箱宽（cm）',
  inner_box_h_cm  DECIMAL(10,2) NULL COMMENT '内箱高（cm）',

  -- ===== 外箱参数 =====
  outer_box_l_cm  DECIMAL(10,2) NULL COMMENT '外箱长（cm）',
  outer_box_w_cm  DECIMAL(10,2) NULL COMMENT '外箱宽（cm）',
  outer_box_h_cm  DECIMAL(10,2) NULL COMMENT '外箱高（cm）',
  carton_qty      INT UNSIGNED  NULL COMMENT '箱规（每外箱产品数量，个）',
  carton_gross_g  DECIMAL(12,2) NULL COMMENT '箱规毛重（g）',

  -- ===== 头程运费（人民币/件，按市场）=====
  first_leg_eu_au_cny DECIMAL(18,6) NULL COMMENT '头程运费 EU/AU（RMB/件）',
  first_leg_us_cny    DECIMAL(18,6) NULL COMMENT '头程运费 US（RMB/件）',
  first_leg_ca_cny    DECIMAL(18,6) NULL COMMENT '头程运费 CA（RMB/件）',
  first_leg_jp_cny    DECIMAL(18,6) NULL COMMENT '头程运费 JP（RMB/件）',
  first_leg_uk_cny    DECIMAL(18,6) NULL COMMENT '头程运费 UK（RMB/件）',

  -- ===== 关税（含税，人民币/件，按市场）=====
  duty_eu_cny         DECIMAL(18,6) NULL COMMENT '关税含税 EU（RMB/件）',
  duty_us_cny         DECIMAL(18,6) NULL COMMENT '关税含税 US（RMB/件，通常为0）',
  duty_ca_au_cny      DECIMAL(18,6) NULL COMMENT '关税含税 CA/AU（RMB/件）',
  duty_jp_cny         DECIMAL(18,6) NULL COMMENT '关税含税 JP（RMB/件）',
  duty_uk_cny         DECIMAL(18,6) NULL COMMENT '关税含税 UK（RMB/件）',

  -- ===== 关税（不含税，人民币/件，按市场）=====
  duty_eu_notax_cny      DECIMAL(18,6) NULL COMMENT '关税不含税 EU（RMB/件）',
  duty_us_notax_cny      DECIMAL(18,6) NULL COMMENT '关税不含税 US（RMB/件）',
  duty_ca_au_notax_cny   DECIMAL(18,6) NULL COMMENT '关税不含税 CA/AU（RMB/件）',
  duty_jp_notax_cny      DECIMAL(18,6) NULL COMMENT '关税不含税 JP（RMB/件）',
  duty_uk_notax_cny      DECIMAL(18,6) NULL COMMENT '关税不含税 UK（RMB/件）',

  -- ===== 采购价格 =====
  purchase_price_orig_cny    DECIMAL(18,6) NULL COMMENT '原始采购价（人民币）',
  purchase_price_default_cny DECIMAL(18,6) NULL COMMENT '默认采购价（当前生效价，人民币）',
  purchase_price_may_cny     DECIMAL(18,6) NULL COMMENT '5月采购价（人民币）',
  purchase_price_jun_cny     DECIMAL(18,6) NULL COMMENT '6月采购价（人民币）',
  purchase_price_sep_cny     DECIMAL(18,6) NULL COMMENT '9月采购价（人民币）',

  -- ===== 代运营佣金点（0~1 小数，如 0.1 表示 10%）=====
  commission_pct_regular     DECIMAL(8,4)  NULL COMMENT '代运营佣金点-常规',
  commission_pct_may_de      DECIMAL(8,4)  NULL COMMENT '5月佣金点-DE站',
  commission_pct_may_non_de  DECIMAL(8,4)  NULL COMMENT '5月佣金点-非DE站',
  commission_pct_jun_de      DECIMAL(8,4)  NULL COMMENT '6月佣金点-DE站',
  commission_pct_jun_non_de  DECIMAL(8,4)  NULL COMMENT '6月佣金点-非DE站',
  commission_pct_sep_de      DECIMAL(8,4)  NULL COMMENT '9月佣金点-DE站',
  commission_pct_sep_non_de  DECIMAL(8,4)  NULL COMMENT '9月佣金点-非DE站',

  -- ===== 价格调整记录 =====
  price_modified_cny  DECIMAL(18,6) NULL COMMENT '修改后单价（人民币，有调价记录时填写）',
  price_modified_date DATE          NULL COMMENT '预计修改生效时间',

  -- ===== 供应商 =====
  supplier_code       VARCHAR(16)   NULL COMMENT '供应商代码（如 F003、F005）',
  supplier_name       VARCHAR(128)  NULL COMMENT '供应商简称（如 金铭、慕家）',
  supplier_full_name  VARCHAR(255)  NULL COMMENT '供应商全称（含代码，如 F003[江门市金铭卫浴有限公司]）',

  -- ===== 产品运营属性 =====
  category      VARCHAR(64)  NULL COMMENT '品类（如 面盆水龙头、淋浴、厨房水龙头、售后配件）',
  ops_mode      VARCHAR(16)  NULL COMMENT '运营模式（代运营 / 自运营）',
  sales_status  VARCHAR(32)  NULL COMMENT '产品销售状态（正常销售/清仓/重点产品/新品/仅售后配件/正常售完下架）',
  product_name  VARCHAR(128) NULL COMMENT '产品名称（中文简称）',
  dev_owner     VARCHAR(64)  NULL COMMENT '开发负责人姓名',

  -- ===== 时间节点 =====
  product_dev_date    DATE NULL COMMENT '产品开发时间',
  product_launch_date DATE NULL COMMENT '产品上架时间',

  -- ===== 来源与时间戳 =====
  source_type VARCHAR(24)  NULL DEFAULT 'Excel' COMMENT '来源类型：Excel/API',
  created_at  TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at  TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),
  UNIQUE KEY uk_psp_product_sku (product_sku),

  KEY idx_psp_warehouse_sku  (warehouse_sku),
  KEY idx_psp_supplier_code  (supplier_code),
  KEY idx_psp_category       (category),
  KEY idx_psp_ops_mode       (ops_mode),
  KEY idx_psp_sales_status   (sales_status),
  KEY idx_psp_region_spec    (region_spec)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='产品SKU定价与成本主数据表（来源：BTH全部SKU明细 → 基础数据维护）';
