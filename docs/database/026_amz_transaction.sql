-- 用途：Amazon 交易明细 Excel 落库（python/excel/daily/amazon/transaction交易明细-*.xlsx）
-- 说明：「已发放订单」「已推迟订单」两份表头一致，可合并导入；用 line_hash 去重。
-- 字符集：utf8mb4

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS amz_transaction (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash CHAR(64) NOT NULL COMMENT '行内容稳定哈希（SHA-256 hex），用于去重与增量 UPSERT',

  source_kind          VARCHAR(32)  NULL COMMENT '来源分类：released=已发放订单 deferred=已推迟订单',

  period_label         VARCHAR(16)  NULL COMMENT '期间（Excel「期间」如 2026-05）',
  report_row_at        DATETIME     NULL COMMENT '报表原日期',
  shop_name            VARCHAR(128) NULL COMMENT '店铺名称',
  site_code            VARCHAR(16)  NULL COMMENT '站点',
  currency             VARCHAR(16)  NULL COMMENT '币种',
  payout_reconcile_status VARCHAR(64) NULL COMMENT '划款账单对账状态',
  payout_at            DATETIME     NULL COMMENT '划款时间',
  settlement_start_at  DATETIME     NULL COMMENT '结算时间-开始',
  settlement_end_at    DATETIME     NULL COMMENT '结算时间-结束',
  settlement_at        DATETIME     NULL COMMENT '结算时间',
  shipped_at           DATETIME     NULL COMMENT '发货时间',
  ship_warehouse        VARCHAR(255) NULL COMMENT '发货仓库',

  group_id             VARCHAR(128) NULL COMMENT 'Amazon group id（结算分组）',
  transaction_type     VARCHAR(64)  NULL COMMENT 'type：order/refund/serviceFee 等',
  amazon_order_id      VARCHAR(64)  NULL COMMENT 'order id',
  original_sales_order_no VARCHAR(64) NULL COMMENT '原销售订单号',
  merchant_order_id    VARCHAR(64)  NULL COMMENT 'merchantOrderId',
  fulfillment_channel  VARCHAR(32)  NULL COMMENT '配送方式 FBA/FBM 等',
  seller_sku           VARCHAR(255) NULL COMMENT 'seller sku',
  child_asin           VARCHAR(32)  NULL COMMENT '子ASIN',
  parent_asin          VARCHAR(32)  NULL COMMENT '父ASIN',
  warehouse_sku        VARCHAR(255) NULL COMMENT 'warehouse sku',
  line_description     VARCHAR(512) NULL COMMENT 'description',
  quantity             DECIMAL(18,4) NULL COMMENT '数量',
  marketplace          VARCHAR(64)  NULL COMMENT 'marketplace（如 Amazon.de）',

  product_sales               DECIMAL(18,6) NULL COMMENT 'product sales',
  product_sales_tax           DECIMAL(18,6) NULL COMMENT 'product sales tax',
  shipping_credits            DECIMAL(18,6) NULL COMMENT 'shipping credits',
  shipping_credits_tax        DECIMAL(18,6) NULL COMMENT 'shipping credits tax',
  gift_wrap_credits           DECIMAL(18,6) NULL COMMENT 'gift wrap credits',
  gift_wrap_credits_tax       DECIMAL(18,6) NULL COMMENT 'gift wrap credits tax',
  regulatory_fee              DECIMAL(18,6) NULL COMMENT 'regulatory fee',
  promotional_rebates         DECIMAL(18,6) NULL COMMENT 'promotional rebates',
  promotional_rebates_tax     DECIMAL(18,6) NULL COMMENT 'promotional rebates tax',
  marketplace_withheld_tax    DECIMAL(18,6) NULL COMMENT 'marketplace withheld tax',
  sales_tax_collected         DECIMAL(18,6) NULL COMMENT 'sales tax collected',
  low_value_goods             DECIMAL(18,6) NULL COMMENT 'low value goods',
  amazon_point_costs          DECIMAL(18,6) NULL COMMENT 'amazon point costs',
  selling_fees                DECIMAL(18,6) NULL COMMENT 'selling fees',
  fba_fees                    DECIMAL(18,6) NULL COMMENT 'fba fees',
  other_transaction_fees      DECIMAL(18,6) NULL COMMENT 'other transaction fees',
  other_amount                DECIMAL(18,6) NULL COMMENT 'other',
  total_amount                DECIMAL(18,6) NULL COMMENT 'total',

  purchase_cost               DECIMAL(18,6) NULL COMMENT '采购成本',
  purchase_shipping           DECIMAL(18,6) NULL COMMENT '采购运费',
  purchase_tax                DECIMAL(18,6) NULL COMMENT '采购税费',
  first_leg_shipping          DECIMAL(18,6) NULL COMMENT '头程运费',
  first_leg_tax               DECIMAL(18,6) NULL COMMENT '头程税费',
  fx_rate_cny                 DECIMAL(18,8) NULL COMMENT '转人民币汇率',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),
  UNIQUE KEY uk_amz_txn_line_hash (line_hash),
  KEY idx_amz_txn_order (amazon_order_id),
  KEY idx_amz_txn_shop_period (shop_name, period_label),
  KEY idx_amz_txn_settlement (settlement_at),
  KEY idx_amz_txn_group (group_id),
  KEY idx_amz_txn_type (transaction_type)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='Amazon-transaction交易明细';
