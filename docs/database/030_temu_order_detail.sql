-- 用途：TEMU 后台导出「TEMU-订单详情.xlsx」按行落库（结构化字段 + line_hash 去重）
-- 说明：导入脚本 `python/v2/orders/import_temu_fee0.py` 步骤 1 按 Excel 行 UPSERT（`line_hash` 唯一）；
--       步骤 1 不落订单号/店铺/仓 SKU（上述字段先空串）；步骤 2 再按参考号匹配
--       `sales_order_shipped`（platform=semitemu）回写 order_no、warehouse_sku、product_sku、shop_*、platform_site。
-- 字符集：utf8mb4

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS temu_order_detail (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash CHAR(64) NOT NULL COMMENT '行内容哈希 SHA256，用于去重',

  -- Excel 来源信息
  excel_sheet VARCHAR(128) NOT NULL COMMENT '工作表名（店铺）',
  currency_group VARCHAR(16) NOT NULL COMMENT '分组：RMB/USD/ZINIAO',
  pay_currency VARCHAR(16) NULL COMMENT '付款币种 CNY/USD/EUR',

  -- 订单主体
  order_time DATETIME NOT NULL COMMENT '订单时间',
  order_no VARCHAR(128) NOT NULL COMMENT '订单号',
  ref_no VARCHAR(128) NULL COMMENT '参考号',

  -- SKU 信息
  sku_key VARCHAR(255) NULL COMMENT 'SKU ID',
  warehouse_sku VARCHAR(128) NOT NULL COMMENT '仓库SKU',
  product_sku VARCHAR(128) NOT NULL COMMENT '产品SKU',
  sku_quantity SMALLINT UNSIGNED NOT NULL COMMENT '购买数量',

  -- 金额字段（全部英文命名）
  unit_price_pay DECIMAL(18,6) NULL COMMENT '产品单价（付款币种）',
  sales_receipt DECIMAL(18,6) NULL COMMENT '销售回款',
  sales_reverse DECIMAL(18,6) NULL COMMENT '销售冲回',
  shipping_receipt DECIMAL(18,6) NULL COMMENT '运费回款',
  tax_income DECIMAL(18,6) NULL COMMENT '税金收入',
  shipping_tax_income DECIMAL(18,6) NULL COMMENT '运费税收入',
  deduction_estimate DECIMAL(18,6) NULL COMMENT '预估扣除金额',
  income_estimate DECIMAL(18,6) NULL COMMENT '预计收入',

  -- 店铺信息
  shop_name_en VARCHAR(128) NOT NULL COMMENT '店铺英文名',
  shop_alias VARCHAR(128) NOT NULL COMMENT '店铺别名',
  platform_site VARCHAR(64) NOT NULL COMMENT '站点',

  -- 状态
  row_kind VARCHAR(32) NOT NULL COMMENT 'valid正常 cancelled取消 invalid无效',
  imported_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '导入时间',

  PRIMARY KEY (id),
  UNIQUE KEY uk_line_hash (line_hash),
  KEY idx_order_time (order_time),
  KEY idx_order_no (order_no),
  KEY idx_shop_platform (shop_name_en, platform_site),
  KEY idx_sku (product_sku, warehouse_sku),
  KEY idx_ref_no (ref_no),
  KEY idx_imported (imported_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='TEMU-订单明细表';
