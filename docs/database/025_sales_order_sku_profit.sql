-- 用途：发货明细行级利润快照，与 sales_order_shipped 一一对应（唯一键 line_hash）
-- 来源：run_order_sku_profit.py / order_sku_profit_constants.py --init-sync 按行读取 shipped；
-- calc_node 与 shipped.profit_calc_node 呼应；订单退款不参与毛利（refund_* 恒 0，net=gross）
-- 字符集：utf8mb4
--
-- 若库中曾存在旧版「按 platform+order_no+warehouse_sku 聚合」的 sales_order_sku_profit，请先备份后
-- DROP TABLE 再执行本脚本，或见 docs/database/026_sales_order_sku_profit_migrate_note.sql

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS sales_order_sku_profit (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash            CHAR(64) NOT NULL COMMENT '与 sales_order_shipped.line_hash 一致',

  platform             VARCHAR(64)  NOT NULL DEFAULT '' COMMENT '平台',
  shop_name_en         VARCHAR(128) NULL     COMMENT '店铺英文名',
  platform_site        VARCHAR(64)  NULL     COMMENT '站点',
  order_type           VARCHAR(64)  NULL     COMMENT '订单类型',
  ref_no               VARCHAR(128) NOT NULL DEFAULT '' COMMENT '参考号',
  order_no             VARCHAR(128) NOT NULL COMMENT '订单号',
  product_sku          VARCHAR(128) NULL COMMENT '产品SKU',
  warehouse_sku        VARCHAR(128) NOT NULL COMMENT '仓库SKU',

  platform_sku         VARCHAR(255) NULL COMMENT '平台SKU',
  warehouse_name       VARCHAR(255) NULL COMMENT '仓库',
  shipping_method      VARCHAR(255) NULL COMMENT '运输方式',

  pay_currency         VARCHAR(16) NULL COMMENT '付款币种',
  base_currency        VARCHAR(16) NULL COMMENT '本位币种',

  pay_time             DATETIME NULL COMMENT '付款时间',
  ship_time            DATETIME NULL COMMENT '发货时间',

  shipped_qty          INT NOT NULL DEFAULT 0 COMMENT '本行仓库sku销量（同 shipped.warehouse_sku_qty）',

  order_total_pay          DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '订单总金额（付款币种）',
  order_goods_pay          DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '订单商品金额（付款币种）',
  platform_shipping_pay    DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '平台运费（付款币种）',
  payment_fee_pay          DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '支付手续费（付款币种）',
  platform_fee_pay         DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '平台手续费（付款币种）',
  fba_fee_pay              DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT 'FBA费用（付款币种）',
  platform_subsidy_pay     DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '平台补贴费（付款币种）',
  tax_pay                  DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '税费（付款币种）',
  other_fee_pay            DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '其他费用（付款币种）',
  purchase_cost_pay        DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '采购成本（付款币种）',
  purchase_shipping_pay    DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '采购运费（付款币种）',
  purchase_tax_pay         DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '采购税费（付款币种）',
  first_leg_shipping_pay   DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '头程运费（付款币种）',
  first_leg_tax_pay        DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '头程税费（付款币种）',
  packaging_fee_pay        DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '包材费用（付款币种）',
  delivery_shipping_pay    DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '派送运费（付款币种）',

  order_total_base         DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '订单总金额（本位币）',
  order_goods_base         DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '订单商品金额（本位币）',
  platform_shipping_base   DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '平台运费（本位币）',
  payment_fee_base         DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '支付手续费（本位币）',
  platform_fee_base        DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '平台手续费（本位币）',
  fba_fee_base             DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT 'FBA费用（本位币）',
  platform_subsidy_base    DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '平台补贴费（本位币）',
  tax_base                 DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '税费（本位币）',
  other_fee_base           DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '其他费用（本位币）',
  purchase_cost_base       DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '采购成本（本位币）',
  purchase_shipping_base   DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '采购运费（本位币）',
  purchase_tax_base        DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '采购税费（本位币）',
  first_leg_shipping_base  DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '头程运费（本位币）',
  first_leg_tax_base       DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '头程税费（本位币）',
  packaging_fee_base       DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '包材费用（本位币）',
  delivery_shipping_base   DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '派送运费（本位币）',
  total_fee_base           DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '总费用（本位币）',
  total_cost_base          DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '总成本（本位币）',
  gross_profit_base        DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '发货毛利（本位币）',
  gross_margin_rate        DECIMAL(10,6) NULL COMMENT '毛利率（小数），同明细行',

  refund_qty               INT NOT NULL DEFAULT 0 COMMENT '预留：退款数量；当前脚本不写 sales_order_refund，恒为 0',
  refund_amount_base       DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '预留：退款本位币额；当前不参与毛利，恒为 0',

  net_profit_base          DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '与发货毛利一致（本位币）：当前等于 gross_profit_base，不减退款',
  net_margin_rate          DECIMAL(10,6) NULL COMMENT '与 gross_margin_rate 一致（发货明细毛利率）',
  distribution_lev         TINYINT NULL COMMENT '分销等级：0=自营；1=分销',

  calc_node                VARCHAR(24) NULL COMMENT '与 sales_order_shipped.profit_calc_node 一致：本行利润写入/批次标记（最长24）',

  source_note              VARCHAR(255) NULL COMMENT '如 line_level:sales_order_shipped_no_refund',
  created_at               TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at               TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),
  UNIQUE KEY uk_sosp_line_hash (line_hash),
  KEY idx_sosp_order_no (order_no),
  KEY idx_sosp_pay_time (pay_time),
  KEY idx_sosp_warehouse_sku (warehouse_sku),
  KEY idx_sosp_calc_node (calc_node)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='销售订单SKU毛利表';
