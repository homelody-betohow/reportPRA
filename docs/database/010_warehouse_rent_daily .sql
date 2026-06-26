CREATE TABLE warehouse_rent_daily (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT COMMENT '自增主键',

  provider VARCHAR(16) NOT NULL COMMENT '服务商/来源：HY 或 4PX',          -- 'HY' / '4PX'
  charge_date DATE NOT NULL COMMENT '计费日期（取日期部分）',
  warehouse_code VARCHAR(64) NULL COMMENT '仓库代码（HY，如 DEHY）',
  warehouse_name VARCHAR(128) NULL COMMENT '仓库名称（4PX，如 法国巴黎2仓）',
  currency CHAR(3) NOT NULL COMMENT '币种（如 EUR/RMB）',

  sku VARCHAR(128) NOT NULL COMMENT 'SKU（按日+SKU 维度汇总，建议 NOT NULL）',              -- 新增：SKU 维度（建议 NOT NULL）

  amount_total DECIMAL(18,6) NOT NULL COMMENT '金额合计（SUM(detail.amount)）',    -- SUM(detail.amount)
  discount_total DECIMAL(18,6) NULL COMMENT '优惠合计（SUM(detail.discount_amount)）',      -- SUM(detail.discount_amount)
  qty_total DECIMAL(18,6) NULL COMMENT '数量合计（SUM(detail.qty)）',           -- SUM(detail.qty)
  volume_total_m3 DECIMAL(18,6) NULL COMMENT '体积合计（SUM(detail.volume_m3)）（HY 常用）',     -- SUM(detail.volume_m3)
  line_count INT UNSIGNED NOT NULL COMMENT '明细行数（COUNT(*)）',       -- COUNT(*)

  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  UNIQUE KEY uk_daily_sku (provider, charge_date, warehouse_code, warehouse_name, currency, sku),
  INDEX idx_daily_sku_date (sku, charge_date),
  INDEX idx_daily_date (charge_date),
  INDEX idx_daily_wh (warehouse_code, warehouse_name, charge_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='仓租日汇总（按 日+仓+币种+SKU）';