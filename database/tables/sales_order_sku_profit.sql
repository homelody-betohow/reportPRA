-- ========================================
-- 表：sales_order_sku_profit
-- 数据库：rpa-report
-- 导出时间：2026-06-12 14:17:41
-- 来源：局域网数据库 172.18.188.18:3309
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS sales_order_sku_profit;
CREATE TABLE `sales_order_sku_profit` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `line_hash` char(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '与 sales_order_shipped.line_hash 一致',
  `platform` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '平台',
  `shop_name_en` varchar(128) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '店铺英文名',
  `platform_site` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '站点',
  `order_type` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '订单类型',
  `ref_no` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '参考号',
  `order_no` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '订单号',
  `product_sku` varchar(128) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '产品SKU',
  `warehouse_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '仓库SKU',
  `platform_sku` varchar(255) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '平台SKU',
  `warehouse_name` varchar(255) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '仓库',
  `warehouse_type` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '仓库类型',
  `shipping_method` varchar(255) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '运输方式',
  `pay_currency` varchar(16) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '付款币种',
  `base_currency` varchar(16) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '本位币种',
  `pay_time` datetime DEFAULT NULL COMMENT '付款时间',
  `ship_time` datetime DEFAULT NULL COMMENT '发货时间',
  `shipped_qty` int NOT NULL DEFAULT '0' COMMENT '本行仓库sku销量（同 shipped.warehouse_sku_qty）',
  `order_total_pay` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '订单总金额（付款币种）',
  `order_total_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '订单总金额（本位币）',
  `order_goods_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '订单商品金额（本位币）',
  `platform_shipping_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '平台运费（本位币）',
  `payment_fee_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '支付手续费（本位币）',
  `platform_fee_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '平台手续费（本位币）',
  `fba_fee_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT 'FBA费用（本位币）',
  `platform_subsidy_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '平台补贴费（本位币）',
  `vat_fee_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT 'VAT费用(本位币)',
  `tax_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '税费（本位币）',
  `other_fee_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '其他费用（本位币）',
  `purchase_cost_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '采购成本（本位币）',
  `purchase_shipping_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '采购运费（本位币）',
  `purchase_tax_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '采购税费（本位币）',
  `first_leg_shipping_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '头程运费（本位币）',
  `first_leg_tax_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '头程税费（本位币）',
  `packaging_fee_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '包材费用（本位币）',
  `delivery_shipping_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '派送运费（本位币）',
  `total_fee_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '总费用（本位币）',
  `total_cost_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '总成本（本位币）',
  `gross_profit_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '发货毛利（本位币）',
  `gross_margin_rate` decimal(10,6) DEFAULT NULL COMMENT '毛利率（小数），同明细行',
  `refund_qty` int NOT NULL DEFAULT '0' COMMENT '预留：退款数量；当前脚本不写 sales_order_refund，恒为 0',
  `refund_amount_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '预留：退款本位币额；当前不参与毛利，恒为 0',
  `net_profit_base` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '与发货毛利一致（本位币）：当前等于 gross_profit_base，不减退款',
  `net_margin_rate` decimal(10,6) DEFAULT NULL COMMENT '与 gross_margin_rate 一致（发货明细毛利率）',
  `distribution_lev` tinyint DEFAULT NULL COMMENT '分销等级：0=自营；1=分销',
  `report_hash` char(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '报告哈希（SHA-256 hex）',
  `calc_node` varchar(24) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '与 sales_order_shipped.profit_calc_node 一致：本行利润写入/批次标记（最长24）',
  `source_note` varchar(255) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '',
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_sosp_line_hash` (`line_hash`),
  KEY `idx_sosp_order_no` (`order_no`),
  KEY `idx_sosp_pay_time` (`pay_time`),
  KEY `idx_sosp_warehouse_sku` (`warehouse_sku`),
  KEY `idx_sosp_calc_node` (`calc_node`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='销售订单SKU毛利表';
