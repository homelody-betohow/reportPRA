-- ========================================
-- 表：warehouse_rent_detail
-- 数据库：rpa-report
-- 导出时间：2026-06-12 14:17:41
-- 来源：局域网数据库 172.18.188.18:3309
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE rpa-report;

DROP TABLE IF EXISTS warehouse_rent_detail;
CREATE TABLE `warehouse_rent_detail` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `provider` varchar(16) NOT NULL COMMENT '服务商/来源：HY 或 4PX',
  `line_hash` char(64) NOT NULL COMMENT '明细行唯一指纹（sha256 hex，用于幂等去重）',
  `doc_no` varchar(128) DEFAULT NULL COMMENT '单据号/批次号：HY=Code；4PX=仓租单号（4PX不唯一，仅用于分组/追溯）',
  `charge_date` date NOT NULL COMMENT '计费日期（取日期部分）',
  `warehouse_code` varchar(64) DEFAULT NULL COMMENT '仓库代码（HY，如 DEHY）',
  `warehouse_name` varchar(128) DEFAULT NULL COMMENT '仓库名称（4PX，如 法国巴黎2仓）',
  `currency` char(3) NOT NULL COMMENT '币种（如 EUR/RMB）',
  `sku` varchar(128) DEFAULT NULL COMMENT 'SKU',
  `barcode` varchar(128) DEFAULT NULL COMMENT '条码/自定义编码（HY）',
  `product_name` varchar(255) DEFAULT NULL COMMENT '产品名称（HY）',
  `qty` decimal(18,6) DEFAULT NULL COMMENT '数量：HY=Quantity；4PX=SKU数量',
  `volume_m3` decimal(18,6) DEFAULT NULL COMMENT '体积（m³）（HY）',
  `weight_kg` decimal(18,6) DEFAULT NULL COMMENT '重量（kg）（HY）',
  `aging_days` int DEFAULT NULL COMMENT '库龄（天）（HY）',
  `rent_free_days` int DEFAULT NULL COMMENT '免租天数（HY）',
  `toll_days` int DEFAULT NULL COMMENT '收费天数（HY）',
  `receiving_no` varchar(128) DEFAULT NULL COMMENT '入库单号/收货单号（HY）',
  `putaway_at` datetime DEFAULT NULL COMMENT '上架时间（HY）',
  `aging_bucket` varchar(32) DEFAULT NULL COMMENT '库龄段（4PX，如 0-30天/30-60天...）',
  `service_category` varchar(32) DEFAULT NULL COMMENT '服务类别（4PX）',
  `service_product` varchar(64) DEFAULT NULL COMMENT '服务产品（4PX）',
  `fee_type` varchar(32) DEFAULT NULL COMMENT '计费类型（4PX）',
  `fee_name` varchar(64) DEFAULT NULL COMMENT '费用名称（4PX）',
  `amount` decimal(18,6) NOT NULL COMMENT '金额：HY=Product amount；4PX=应收金额（建议作为统一口径）',
  `billed_amount` decimal(18,6) DEFAULT NULL COMMENT '计费金额（4PX）',
  `discount_amount` decimal(18,6) DEFAULT NULL COMMENT '优惠金额（4PX）',
  `raw_row_json` json DEFAULT NULL COMMENT '原始行 JSON（用于追溯与排错，可包含 source_file/source_sheet 等）',
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_line` (`provider`,`line_hash`),
  KEY `idx_date_wh` (`charge_date`,`warehouse_code`,`warehouse_name`),
  KEY `idx_sku_date` (`sku`,`charge_date`),
  KEY `idx_docno` (`provider`,`doc_no`)
) ENGINE=InnoDB AUTO_INCREMENT=187595 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci COMMENT='仓租明细（HY/4PX 通用明细表）';
