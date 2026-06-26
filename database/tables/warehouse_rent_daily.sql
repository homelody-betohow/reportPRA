-- ========================================
-- 表：warehouse_rent_daily
-- 数据库：rpa-report
-- 导出时间：2026-06-12 14:17:41
-- 来源：局域网数据库 172.18.188.18:3309
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS warehouse_rent_daily;
CREATE TABLE `warehouse_rent_daily` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `provider` varchar(16) NOT NULL COMMENT '服务商/来源：HY 或 4PX',
  `charge_date` date NOT NULL COMMENT '计费日期（取日期部分）',
  `warehouse_code` varchar(64) DEFAULT NULL COMMENT '仓库代码（HY，如 DEHY）',
  `warehouse_name` varchar(128) DEFAULT NULL COMMENT '仓库名称（4PX，如 法国巴黎2仓）',
  `currency` char(3) NOT NULL COMMENT '币种（如 EUR/RMB）',
  `sku` varchar(128) NOT NULL COMMENT 'SKU（按日+SKU 维度汇总，建议 NOT NULL）',
  `amount_total` decimal(18,6) NOT NULL COMMENT '金额合计（SUM(detail.amount)）',
  `discount_total` decimal(18,6) DEFAULT NULL COMMENT '优惠合计（SUM(detail.discount_amount)）',
  `qty_total` decimal(18,6) DEFAULT NULL COMMENT '数量合计（SUM(detail.qty)）',
  `volume_total_m3` decimal(18,6) DEFAULT NULL COMMENT '体积合计（SUM(detail.volume_m3)）（HY 常用）',
  `line_count` int unsigned NOT NULL COMMENT '明细行数（COUNT(*)）',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_daily_sku` (`provider`,`charge_date`,`warehouse_code`,`warehouse_name`,`currency`,`sku`),
  KEY `idx_daily_sku_date` (`sku`,`charge_date`),
  KEY `idx_daily_date` (`charge_date`),
  KEY `idx_daily_wh` (`warehouse_code`,`warehouse_name`,`charge_date`)
) ENGINE=InnoDB AUTO_INCREMENT=53255 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='仓租日汇总（按 日+仓+币种+SKU）';
