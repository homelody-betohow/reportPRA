-- ========================================
-- 表：platform_shop
-- 数据库：rpa-report
-- 导出时间：2026-06-12 14:17:41
-- 来源：局域网数据库 172.18.188.18:3309
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS platform_shop;
CREATE TABLE `platform_shop` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `shop_hash` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺哈希（platform+站点+account）',
  `shop_name_en` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺英文名',
  `shop_name_cn` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '店铺中文名/别名（导入时取自订单 shop_alias）',
  `shop_alias` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '店铺别名（与 sales_order_shipped.shop_alias 一致，便于对账）',
  `store_account` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺账号',
  `store_secret` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺密钥',
  `platform` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '平台',
  `platform_site` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '站点',
  `market_region` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场区域',
  `market_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '市场编码',
  `currency` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '本位/结算币种',
  `fx_rate` decimal(18,8) NOT NULL DEFAULT '0.00000000' COMMENT '汇率（转本位币）',
  `commission_type` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '佣金类型',
  `commission_rate` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT '佣金率（小数，如 0.15 表示 15%）',
  `vat_rate_type` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT 'VAT 税率类型',
  `vat_rate` decimal(18,6) NOT NULL DEFAULT '0.000000' COMMENT 'VAT 税率（小数，如 0.19 表示 19%）',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '运营负责人（导入时取自订单 shop_owner）',
  `shop_status` tinyint unsigned NOT NULL DEFAULT '1' COMMENT '状态：1=启用 0=停用（停用则不参与利润计算）',
  `remark` varchar(512) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '备注',
  `use_rpa` tinyint(1) NOT NULL DEFAULT '0' COMMENT '使用RPA',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_psc_shop_hash` (`shop_hash`),
  KEY `idx_psc_platform_shop` (`platform`,`platform_site`,`shop_name_en`),
  KEY `idx_psc_platform` (`platform`),
  KEY `idx_psc_platform_site` (`platform_site`),
  KEY `idx_psc_store_account` (`store_account`),
  KEY `idx_psc_shop_name_en` (`shop_name_en`),
  KEY `idx_psc_shop_alias` (`shop_alias`),
  KEY `idx_psc_currency` (`currency`),
  KEY `idx_psc_ops_owner` (`ops_owner`),
  KEY `idx_psc_shop_status` (`shop_status`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='店铺信息表';
