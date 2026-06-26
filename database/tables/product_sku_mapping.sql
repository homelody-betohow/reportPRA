-- ========================================
-- 表：product_sku_mapping
-- 数据库：rpa-report
-- 说明：按「服务商（partner）」维度维护外部 SKU → 标准产品 SKU 的映射
--
-- 设计要点：
--   1. partner_type=platform  ：用 seller_sku   标识平台刊登 SKU
--   2. partner_type=warehouse ：用 warehouse_sku 标识仓储/WMS SKU
--   3. mapping_type=single    ：1 对 1，填写 product_sku；component_info 必须为 NULL
--   4. mapping_type=bundle    ：1 对多（组合品），填写 component_info；product_sku 固定空串
--   5. 未使用的 SKU 列统一填空串 ''（非 NULL），以便唯一索引与导入去重稳定
--   6. 站点信息由 shop_hash / market_* 承载；仓库名称由 partner_name 承载，不再单独存 partner_site_code / warehouse_code
--
-- component_info JSON 格式（bundle 时必填，数组按组件顺序排列）：
--   [{"product_sku":"E51033005","qty":1},{"product_sku":"E51033010","qty":2}]
--
-- line_hash 参与列（顺序固定，SHA-256 hex；改动需评估历史行迁移）：
--   partner_code, partner_type, partner_name, shop_hash,
--   seller_sku, warehouse_sku, mapping_type, product_sku
-- ========================================
SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `rpa-report`;

DROP TABLE IF EXISTS `product_sku_mapping`;
CREATE TABLE `product_sku_mapping` (
  `id` bigint unsigned NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  `line_hash` char(64) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '行内容稳定哈希（SHA-256 hex）；参与列见文件头注释',

  `partner_code` varchar(30) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '服务商/合作伙伴编码（如 AMZ/TEMU/HY/4PX/MANO）',
  `partner_type` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL COMMENT '映射维度：platform=平台侧 / warehouse=仓储侧',
  `partner_name` varchar(128) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '合作伙伴名称：platform=shop_alias / warehouse=warehouse_name',


  `market_region` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '市场区域（platform 侧从 platform_shop 同步）',
  `market_code` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '市场编码（platform 侧从 platform_shop 同步）',
  `seller_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '平台卖家 SKU（partner_type=platform 时必填）',
  `warehouse_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '仓储 SKU（partner_type=warehouse 时必填）',

  `mapping_type` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'single' COMMENT '映射形态：single=1对1 / bundle=组合品',
  `product_sku` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '标准产品 SKU（mapping_type=single 时必填；bundle 时固定空串）',
  `component_info` json DEFAULT NULL COMMENT '组合品组件明细（mapping_type=bundle 时必填；single 时必须为 NULL）',

  `dev_owner` varchar(128) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '开发负责人',
  `ops_owner` varchar(128) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '运营负责人',
  `shop_hash` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '店铺哈希（platform 侧；无则空串）',
  `source_type` varchar(24) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'Excel' COMMENT '来源类型：Excel/API/Manual',
  `is_active` tinyint(1) NOT NULL DEFAULT '1' COMMENT '是否启用：1=有效 / 0=停用（保留历史行）',

  `created_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (`id`),
  UNIQUE KEY `uk_psm_line_hash` (`line_hash`),
  UNIQUE KEY `uk_psm_partner_external` (
    `partner_code`,
    `partner_type`,
    `shop_hash`,
    `seller_sku`,
    `warehouse_sku`,
    `product_sku`
  ),
  KEY `idx_psm_platform_lookup` (`partner_code`, `shop_hash`, `seller_sku`),
  KEY `idx_psm_warehouse_lookup` (`partner_code`, `warehouse_sku`),
  KEY `idx_psm_product_sku` (`product_sku`),
  KEY `idx_psm_partner_type` (`partner_type`, `partner_code`),
  KEY `idx_psm_mapping_type` (`mapping_type`),
  KEY `idx_psm_market` (`market_region`, `market_code`),

  CONSTRAINT `chk_psm_partner_type` CHECK (`partner_type` IN ('platform', 'warehouse')),
  CONSTRAINT `chk_psm_mapping_type` CHECK (`mapping_type` IN ('single', 'bundle')),
  CONSTRAINT `chk_psm_sku_by_partner` CHECK (
    (`partner_type` = 'platform' AND `seller_sku` <> '' AND `warehouse_sku` = '')
    OR (`partner_type` = 'warehouse' AND `warehouse_sku` <> '' AND `seller_sku` = '')
  ),
  CONSTRAINT `chk_psm_mapping_payload` CHECK (
    (`mapping_type` = 'single' AND `product_sku` <> '' AND `component_info` IS NULL)
    OR (`mapping_type` = 'bundle' AND `product_sku` = '' AND `component_info` IS NOT NULL)
  )
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='产品SKU映射表（single→product_sku；bundle→component_info）';
