-- 用途：存储 Excel「欧洲平台定价表.xlsx」的数据（按 sheet 分表 + 可扩展字段）
-- 字符集：utf8mb4（支持 € 等符号）
-- 建议：导入时每次生成一个 import_batch_id，便于回溯与重导

SET NAMES utf8mb4;

-- =========================
-- 1) 汇率表（sheet: 汇率表）
-- =========================
CREATE TABLE IF NOT EXISTS excel_eu_pricing_exchange_rate (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  import_batch_id VARCHAR(64) NOT NULL,
  source_row_no INT UNSIGNED NULL,

  platform_country VARCHAR(64) NULL COMMENT '平台国家，例如 Amazon-DE / Mano-FR',
  platform_rmb_rate DECIMAL(18,6) NULL COMMENT '平台对应人民币汇率（Excel 原字段）',
  currency_symbol VARCHAR(8) NULL COMMENT '币种符号，例如 € / £ / $',
  currency_name VARCHAR(64) NULL COMMENT '币种/含义（例如 欧元汇率/英镑汇率/欧元对英镑…）',
  rmb_rate DECIMAL(18,10) NULL COMMENT '对人民币汇率（Excel 原字段）',
  shipping_warehouse VARCHAR(64) NULL COMMENT '发货仓库（若后续有值）',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_exchange_rate_batch (import_batch_id),
  KEY idx_exchange_rate_platform_country (platform_country)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- =========================
-- 2) 基础表（sheet: 基础表）
-- 说明：该 sheet 表头是“多行表头”，建议导入前先把真正字段名整理出来。
-- 这里提供一个“可落地”的宽表，保证能把数据先存进去。
-- =========================
CREATE TABLE IF NOT EXISTS excel_eu_pricing_base (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  import_batch_id VARCHAR(64) NOT NULL,
  source_row_no INT UNSIGNED NULL,

  sku VARCHAR(64) NULL COMMENT '百途鸿SKU',
  category_lv3 VARCHAR(128) NULL COMMENT '三级类目',
  purchase_price_cny DECIMAL(18,6) NULL COMMENT '采购价（￥）',
  gross_weight_g INT NULL COMMENT '单个毛重（g）',

  inner_box_l_cm DECIMAL(18,4) NULL COMMENT '内盒规格 长(cm)',
  inner_box_w_cm DECIMAL(18,4) NULL COMMENT '内盒规格 宽(cm)',
  inner_box_h_cm DECIMAL(18,4) NULL COMMENT '内盒规格 高(cm)',

  outer_box_l_cm DECIMAL(18,4) NULL COMMENT '外箱规格 长(cm)',
  outer_box_w_cm DECIMAL(18,4) NULL COMMENT '外箱规格 宽(cm)',
  outer_box_h_cm DECIMAL(18,4) NULL COMMENT '外箱规格 高(cm)',
  carton_qty INT NULL COMMENT '每箱数量',
  carton_weight_g INT NULL COMMENT '每箱重量（g）',

  -- 尾程/派送费（Excel 内类似 HY-DE / 4PX-FR / MF-FR / FBA-DE 等列很多）
  -- 为避免未来列变动导致频繁改表，这里用 JSON 先承接。
  last_mile_rates_json JSON NULL COMMENT '尾程费率明细（列名->数值）',

  -- 销售仓租/调拨（同样列很多）
  warehouse_rent_json JSON NULL COMMENT '销售仓租（列名->数值）',
  transfer_fee_json JSON NULL COMMENT '调拨（列名->数值）',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_base_batch (import_batch_id),
  KEY idx_base_sku (sku)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- ==================================
-- 3) MANO 尾程/仓租费率（sheet: MANO尾程 仓租）
-- 建议按“尺寸段 + 重量上限 + 国家”做规范化存储
-- ==================================
CREATE TABLE IF NOT EXISTS excel_mano_tail_fee (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  import_batch_id VARCHAR(64) NOT NULL,
  source_row_no INT UNSIGNED NULL,

  size_group VARCHAR(64) NOT NULL COMMENT '尺寸段，例如 ≤50x40x30 / ≤100x60x60 / ≤150x100x80 / >150x100x80',
  max_weight_g INT NOT NULL COMMENT '重量上限(g)',
  country_code VARCHAR(8) NOT NULL COMMENT '国家：FR/ES/IT',
  fee_eur DECIMAL(18,6) NULL COMMENT '费用（欧元）',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uk_mano_tail_fee (import_batch_id, size_group, max_weight_g, country_code),
  KEY idx_mano_tail_fee_country (country_code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- ==================================
-- 4) 欧洲平台定价表（sheet: 欧洲平台定价表）
-- ==================================
CREATE TABLE IF NOT EXISTS excel_eu_platform_pricing (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  import_batch_id VARCHAR(64) NOT NULL,
  source_row_no INT UNSIGNED NULL,

  seq_no INT NULL COMMENT '序号',
  product_id VARCHAR(64) NULL COMMENT '商品ID',
  sku VARCHAR(64) NULL COMMENT 'SKU',
  name VARCHAR(255) NULL COMMENT '名称',

  first_leg_method VARCHAR(32) NULL COMMENT '头程方式',
  shipping_warehouse VARCHAR(64) NULL COMMENT '发货仓库',
  is_transfer VARCHAR(8) NULL COMMENT '是否调拨（是/否）',
  is_tax_included VARCHAR(8) NULL COMMENT '是否含税（是/否）',

  rma_ratio DECIMAL(10,6) NULL COMMENT 'RMA占比',
  ad_ratio DECIMAL(10,6) NULL COMMENT '广告占比',
  mgmt_cost DECIMAL(18,6) NULL COMMENT '管理成本',

  seckill_days INT NULL COMMENT '秒杀天数',
  review_qty INT NULL COMMENT '测评量',

  normal_price DECIMAL(18,6) NULL COMMENT '正常销价',
  normal_qty INT NULL COMMENT '正常量',
  promo_price DECIMAL(18,6) NULL COMMENT '促销价格',
  promo_qty INT NULL COMMENT '促销量',
  seckill_price DECIMAL(18,6) NULL COMMENT '秒杀价价格',
  seckill_qty INT NULL COMMENT '秒杀数量',

  total_qty INT NULL COMMENT '总销量',
  total_gross_margin_rate DECIMAL(18,10) NULL COMMENT '总毛利率',

  platform_fee DECIMAL(18,6) NULL COMMENT '平台费',
  sales_tax DECIMAL(18,6) NULL COMMENT '销售税',
  withdrawal_fee DECIMAL(18,6) NULL COMMENT '提现费',
  review_cost DECIMAL(18,6) NULL COMMENT '测评花费',
  seckill_cost DECIMAL(18,6) NULL COMMENT '秒杀花费',
  promo_cost DECIMAL(18,6) NULL COMMENT '促销费',

  purchase_price DECIMAL(18,6) NULL COMMENT '采购价',
  first_leg_tariff DECIMAL(18,6) NULL COMMENT '头程关税',
  last_mile_fee DECIMAL(18,6) NULL COMMENT '尾程派送费',
  warehouse_rent_fee DECIMAL(18,6) NULL COMMENT '销售仓租费',
  transfer_fee DECIMAL(18,6) NULL COMMENT '调拨费',

  avg_platform_fee DECIMAL(18,6) NULL COMMENT '平均平台费',
  avg_sales_tax DECIMAL(18,6) NULL COMMENT '平均销售税',
  avg_withdrawal_fee DECIMAL(18,6) NULL COMMENT '平均提现费',
  avg_purchase_price DECIMAL(18,6) NULL COMMENT '平均采购价',
  avg_first_leg_tariff DECIMAL(18,6) NULL COMMENT '平均头程关税',
  avg_last_mile_fee DECIMAL(18,6) NULL COMMENT '平均尾程派送费',
  avg_transfer_fee DECIMAL(18,6) NULL COMMENT '平均调拨费',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_eu_platform_pricing_batch (import_batch_id),
  KEY idx_eu_platform_pricing_sku (sku),
  KEY idx_eu_platform_pricing_product (product_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- =========================
-- 5) TEMU（sheet: TEMU）
-- =========================
CREATE TABLE IF NOT EXISTS excel_temu_pricing (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  import_batch_id VARCHAR(64) NOT NULL,
  source_row_no INT UNSIGNED NULL,

  seq_no INT NULL COMMENT '序号',
  product_id VARCHAR(64) NULL COMMENT '商品ID',
  sku_1pack VARCHAR(64) NULL COMMENT '1个装SKU',
  sku_2pack VARCHAR(64) NULL COMMENT '2个装SKU',
  qty INT NULL COMMENT '数量',

  shipping_warehouse VARCHAR(64) NULL COMMENT '发货仓库',
  first_leg_method VARCHAR(32) NULL COMMENT '头程方式',
  rma_ratio DECIMAL(10,6) NULL COMMENT 'RMA占比',

  normal_price_eur DECIMAL(18,6) NULL COMMENT '正常销价（欧元）',
  normal_price_cny DECIMAL(18,6) NULL COMMENT '正常售价（人民币）',
  total_gross_margin_rate DECIMAL(18,10) NULL COMMENT '总毛利率',

  purchase_price DECIMAL(18,6) NULL COMMENT '采购价',
  first_leg_tariff DECIMAL(18,6) NULL COMMENT '头程关税',
  last_mile_fee DECIMAL(18,6) NULL COMMENT '尾程派送费',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_temu_pricing_batch (import_batch_id),
  KEY idx_temu_pricing_sku1 (sku_1pack),
  KEY idx_temu_pricing_product (product_id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- ==================================================
-- 6) MANO-UK / RDC / Conforama（结构一致：SKU 维度定价&成本）
-- ==================================================
CREATE TABLE IF NOT EXISTS excel_platform_sku_pricing (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  import_batch_id VARCHAR(64) NOT NULL,
  source_row_no INT UNSIGNED NULL,

  platform_site VARCHAR(32) NOT NULL COMMENT '来源 sheet/站点：MANO-UK / RDC-FR / Conforama-FR',

  seq_no INT NULL COMMENT '序号',
  sku VARCHAR(64) NULL COMMENT 'SKU',
  name VARCHAR(255) NULL COMMENT '名称',
  spec VARCHAR(255) NULL COMMENT '规格',
  category VARCHAR(128) NULL COMMENT '品类',
  supplier VARCHAR(128) NULL COMMENT '供应商',

  is_pan_eu VARCHAR(8) NULL COMMENT '是否泛欧（是/否）',
  first_leg_method VARCHAR(32) NULL COMMENT '头程方式',
  shipping_warehouse VARCHAR(64) NULL COMMENT '发货仓库',
  is_transfer VARCHAR(8) NULL COMMENT '是否调拨（是/否）',

  rma_ratio DECIMAL(10,6) NULL COMMENT 'RMA占比',
  ad_ratio DECIMAL(10,6) NULL COMMENT '广告占比',
  offsite_ratio DECIMAL(10,6) NULL COMMENT '站外占比',
  review_ratio DECIMAL(10,6) NULL COMMENT '测评占比',

  seckill_cost DECIMAL(18,6) NULL COMMENT '秒杀花费',

  normal_price DECIMAL(18,6) NULL COMMENT '正常销价',
  normal_qty INT NULL COMMENT '正常量',
  review_qty INT NULL COMMENT '测评量',
  seckill_times INT NULL COMMENT '秒杀次数',
  promo_price DECIMAL(18,6) NULL COMMENT '促销价格',
  promo_qty INT NULL COMMENT '促销量',
  total_qty INT NULL COMMENT '总销量',

  total_gross_margin_rate DECIMAL(18,10) NULL COMMENT '总毛利率',
  platform_fee DECIMAL(18,6) NULL COMMENT '平台费',
  sales_tax DECIMAL(18,6) NULL COMMENT '销售税',
  withdrawal_fee DECIMAL(18,6) NULL COMMENT '提现费',
  review_cost DECIMAL(18,6) NULL COMMENT '测评花费',

  purchase_price DECIMAL(18,6) NULL COMMENT '采购价',
  operation_mode VARCHAR(64) NULL COMMENT '运营模式',
  managed_service_commission DECIMAL(18,6) NULL COMMENT '代运营佣金',
  first_leg_tariff DECIMAL(18,6) NULL COMMENT '头程关税',
  last_mile_fee DECIMAL(18,6) NULL COMMENT '尾程派送费',
  warehouse_rent_fee DECIMAL(18,6) NULL COMMENT '销售仓租费',
  transfer_fee DECIMAL(18,6) NULL COMMENT '调拨费',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  KEY idx_platform_sku_pricing_batch (import_batch_id),
  KEY idx_platform_sku_pricing_site (platform_site),
  KEY idx_platform_sku_pricing_sku (sku)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

