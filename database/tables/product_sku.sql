-- ========================================
-- 表：product_sku
-- 数据库：rpa-report
-- 导出时间：2026-06-12 14:17:41
-- 来源：局域网数据库 172.18.188.18:3309
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS product_sku;
CREATE TABLE `product_sku` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `line_hash` char(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '行内容稳定哈希（SHA-256 hex），变更检测用',
  `product_sku` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '产品SKU（如 E51033005）唯一键',
  `product_uid` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '商品ID/SPU 型号（如 25-CFLT-00001，多颜色变体共用）',
  `warehouse_ref` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '仓库识别码（如 JD-03-L-001）',
  `category_lv1` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '一级分类（如 厨房龙头、淋浴花洒套装）',
  `category_lv2` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '二级分类（如 厨房龙头、淋浴花洒套装）',
  `category_lv3` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '三级分类（如 传统恒温淋浴花洒套装）',
  `category_code` varchar(8) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '产品类别代码（命名规则代码，如 03/04/01/0/99/08）',
  `supplier_name` varchar(128) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '供应商名称',
  `product_unit` varchar(16) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '单位（如 套）',
  `product_color` varchar(32) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '产品颜色（如 铬色/雅黑色/拉丝色）',
  `amz_lifecycle` varchar(16) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'AMZ 新老品状态（新品/保留品/不保留老品）',
  `local_lifecycle` varchar(16) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '本土平台新老品状态（新品/保留品/不保留老品）',
  `accounting_class` varchar(32) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '核算分类（如 全新品/转厂新质新品/颜色变体新品/保留品）',
  `purchase_moq` int unsigned DEFAULT NULL COMMENT 'MOQ 最小起订量（件）',
  `purchase_lead_days` int unsigned DEFAULT NULL COMMENT '采购交期（天）',
  `carton_qty` int unsigned DEFAULT NULL COMMENT '箱规（每外箱产品数）',
  `cost_price_cny` decimal(18,6) DEFAULT NULL COMMENT '成本价（人民币）',
  `unit_weight_g` decimal(12,2) DEFAULT NULL COMMENT '单件重量（g）',
  `carton_gross_g` decimal(12,2) DEFAULT NULL COMMENT '箱规毛重（g）',
  `inner_box_l_cm` decimal(10,2) DEFAULT NULL COMMENT '内箱长（cm）',
  `inner_box_w_cm` decimal(10,2) DEFAULT NULL COMMENT '内箱宽（cm）',
  `inner_box_h_cm` decimal(10,2) DEFAULT NULL COMMENT '内箱高（cm）',
  `outer_box_l_cm` decimal(10,2) DEFAULT NULL COMMENT '外箱长（cm）',
  `outer_box_w_cm` decimal(10,2) DEFAULT NULL COMMENT '外箱宽（cm）',
  `outer_box_h_cm` decimal(10,2) DEFAULT NULL COMMENT '外箱高（cm）',
  `first_leg_eu_au_cny` decimal(18,6) DEFAULT NULL COMMENT '头程运费 EU/AU（RMB/件）',
  `first_leg_us_cny` decimal(18,6) DEFAULT NULL COMMENT '头程运费 US（RMB/件）',
  `first_leg_uk_cny` decimal(18,6) DEFAULT NULL COMMENT '头程运费 UK（RMB/件）',
  `duty_eu_cny` decimal(18,6) DEFAULT NULL COMMENT '关税 EU（RMB/件）',
  `duty_us_cny` decimal(18,6) DEFAULT NULL COMMENT '关税 US（RMB/件）',
  `duty_uk_cny` decimal(18,6) DEFAULT NULL COMMENT '关税 UK（RMB/件）',
  `source_type` varchar(24) COLLATE utf8mb4_unicode_ci DEFAULT 'Excel' COMMENT '来源类型：Excel/API',
  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_product_sku` (`product_sku`),
  KEY `idx_psm_product_uid` (`product_uid`),
  KEY `idx_psm_warehouse_ref` (`warehouse_ref`),
  KEY `idx_psm_category_lv2` (`category_lv2`),
  KEY `idx_psm_category_lv3` (`category_lv3`),
  KEY `idx_psm_supplier` (`supplier_name`),
  KEY `idx_psm_color` (`product_color`),
  KEY `idx_psm_amz_lifecycle` (`amz_lifecycle`)
) ENGINE=InnoDB AUTO_INCREMENT=1953 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='产品SKU数据表';
