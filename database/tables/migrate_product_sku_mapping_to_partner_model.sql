-- ========================================
-- 迁移脚本：product_sku_mapping 从旧结构迁移到新的服务商维度模型
-- 执行前请务必备份数据！
-- 
-- 旧字段：platform, platform_site, shop_name_en, warehouse_name, platform_sku
-- 新字段：partner_code, partner_type, partner_site_code, shop_hash, 
--         seller_sku, warehouse_code, market_region, market_code, 
--         mapping_type, component_info, is_active
-- ========================================

USE `rpa-report`;

-- 步骤 1：备份旧表（可选，建议执行）
-- CREATE TABLE `product_sku_mapping_backup_20260618` LIKE `product_sku_mapping`;
-- INSERT INTO `product_sku_mapping_backup_20260618` SELECT * FROM `product_sku_mapping`;

-- 步骤 2：添加新字段
ALTER TABLE `product_sku_mapping`
  ADD COLUMN `partner_code` varchar(30) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '服务商/合作伙伴编码（如 AMZ/TEMU/HY/4PX/MANO）' AFTER `line_hash`,
  ADD COLUMN `partner_type` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'platform' COMMENT '映射维度：platform=平台侧 / warehouse=仓储侧' AFTER `partner_code`,
  ADD COLUMN `partner_name` varchar(128) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '合作伙伴名称（冗余展示，可选）' AFTER `partner_type`,
  ADD COLUMN `partner_site_code` varchar(64) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '站点/区域（平台站点或仓库区域；无则空串）' AFTER `partner_name`,
  
  ADD COLUMN `shop_hash` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '店铺哈希（platform 侧；无则空串）' AFTER `partner_site_code`,
  ADD COLUMN `market_region` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '市场区域（platform 侧从 platform_shop 同步）' AFTER `shop_hash`,
  ADD COLUMN `market_code` varchar(64) COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT '市场编码（platform 侧从 platform_shop 同步）' AFTER `market_region`,
  ADD COLUMN `seller_sku` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '平台卖家 SKU（partner_type=platform 时必填）' AFTER `market_code`,
  ADD COLUMN `warehouse_code` varchar(128) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '' COMMENT '仓库编码（warehouse 侧；无则空串）' AFTER `warehouse_sku`,
  
  ADD COLUMN `mapping_type` varchar(16) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT 'single' COMMENT '映射形态：single=1对1 / bundle=组合品' AFTER `warehouse_code`,
  ADD COLUMN `component_info` json DEFAULT NULL COMMENT '组合品组件明细（mapping_type=bundle 时必填；single 时必须为 NULL）' AFTER `product_sku`,
  
  ADD COLUMN `is_active` tinyint(1) NOT NULL DEFAULT '1' COMMENT '是否启用：1=有效 / 0=停用（保留历史行）' AFTER `source_type`;

-- 步骤 3：数据迁移（将旧字段数据同步到新字段）
-- 假设旧表有 platform, platform_site, shop_name_en, warehouse_name, platform_sku 字段

-- 3.1 更新 partner_code（从 platform）
UPDATE `product_sku_mapping` SET `partner_code` = IFNULL(`platform`, '');

-- 3.2 更新 partner_site_code（从 platform_site）
UPDATE `product_sku_mapping` SET `partner_site_code` = IFNULL(`platform_site`, '');

-- 3.3 更新 seller_sku（从 platform_sku）
UPDATE `product_sku_mapping` SET `seller_sku` = IFNULL(`platform_sku`, '');

-- 3.4 更新 warehouse_code（从 warehouse_name）
UPDATE `product_sku_mapping` SET `warehouse_code` = IFNULL(`warehouse_name`, '');

-- 3.5 计算 shop_hash（需要 platform + platform_site + shop_name_en）
-- 注意：这里的 shop_hash 计算逻辑需要与代码中的 stable_line_hash 保持一致
-- 由于 SQL 中无法直接计算 SHA-256，建议后续通过 Python 脚本批量更新
-- 暂时设置为空，等后续导入订单时自动填充
UPDATE `product_sku_mapping` SET `shop_hash` = '';

-- 3.6 设置 partner_type（所有旧数据默认为 platform 类型）
-- 如果需要区分，可以根据业务规则调整
UPDATE `product_sku_mapping` SET `partner_type` = 'platform';

-- 3.7 设置 mapping_type 为 single
UPDATE `product_sku_mapping` SET `mapping_type` = 'single';

-- 3.8 设置 is_active 为 1
UPDATE `product_sku_mapping` SET `is_active` = 1;

-- 步骤 4：删除旧字段（谨慎操作！确认数据迁移无误后再执行）
-- 建议先保留旧字段一段时间，确认新系统运行正常后再删除
-- ALTER TABLE `product_sku_mapping` DROP COLUMN `platform`;
-- ALTER TABLE `product_sku_mapping` DROP COLUMN `platform_site`;
-- ALTER TABLE `product_sku_mapping` DROP COLUMN `shop_name_en`;
-- ALTER TABLE `product_sku_mapping` DROP COLUMN `warehouse_name`;
-- ALTER TABLE `product_sku_mapping` DROP COLUMN `platform_sku`;

-- 步骤 5：添加新索引
ALTER TABLE `product_sku_mapping`
  ADD KEY `idx_psm_platform_lookup` (`partner_code`, `partner_site_code`, `shop_hash`, `seller_sku`),
  ADD KEY `idx_psm_warehouse_lookup` (`partner_code`, `partner_site_code`, `warehouse_code`, `warehouse_sku`),
  ADD KEY `idx_psm_partner_type` (`partner_type`, `partner_code`),
  ADD KEY `idx_psm_mapping_type` (`mapping_type`),
  ADD KEY `idx_psm_market` (`market_region`, `market_code`);

-- 步骤 6：删除可能冲突的旧索引（如果存在）
-- ALTER TABLE `product_sku_mapping` DROP INDEX `idx_psm_platform_sku` IF EXISTS;
-- ALTER TABLE `product_sku_mapping` DROP INDEX `idx_psm_shop` IF EXISTS;

-- 步骤 7：修改唯一键（需要先删除旧的，再添加新的）
-- 注意：这一步会失败如果有重复数据，请先清理重复数据
-- ALTER TABLE `product_sku_mapping` DROP INDEX `uk_psm_partner_product` IF EXISTS;
ALTER TABLE `product_sku_mapping`
  ADD UNIQUE KEY `uk_psm_partner_external` (
    `partner_code`,
    `partner_type`,
    `partner_site_code`,
    `shop_hash`,
    `warehouse_code`,
    `seller_sku`,
    `warehouse_sku`,
    `product_sku`
  );

-- 完成！
SELECT '迁移完成！请检查数据是否正确。' AS message;
