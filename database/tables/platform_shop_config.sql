-- ========================================
-- 表：platform_shop_config
-- 数据库：rpa-report
-- 导出时间：2026-06-12 14:17:41
-- 来源：局域网数据库 172.18.188.18:3309
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS platform_shop_config;
CREATE TABLE `platform_shop_config` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `shop_hash` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺哈希',
  `shop_name_en` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺英文名',
  `shop_name_cn` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '店铺中文名',
  `platform` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '平台',
  `platform_site` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '站点',
  `currency` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '币种',
  `fx_rate` decimal(18,6) NOT NULL COMMENT '汇率',
  `commission_type` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '佣金类型',
  `commission_rate` decimal(18,6) NOT NULL COMMENT '佣金率',
  `vat_rate_type` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT 'VAT税率类型',
  `vat_rate` decimal(18,6) NOT NULL COMMENT 'VAT税率',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '运营负责人',
  `shop_status` tinyint(1) NOT NULL DEFAULT '1' COMMENT '店铺状态',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_psc_shop_hash` (`shop_hash`),
  KEY `idx_psc_platform_shop` (`platform`,`platform_site`,`shop_name_en`),
  KEY `idx_psc_platform` (`platform`),
  KEY `idx_psc_platform_site` (`platform_site`),
  KEY `idx_psc_shop_name_en` (`shop_name_en`),
  KEY `idx_psc_currency` (`currency`),
  KEY `idx_psc_ops_owner` (`ops_owner`)
) ENGINE=InnoDB AUTO_INCREMENT=9876 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='平台店铺配置表';
