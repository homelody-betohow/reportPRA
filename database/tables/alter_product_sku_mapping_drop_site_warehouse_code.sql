-- 删除 product_sku_mapping 冗余字段：partner_site_code、warehouse_code
-- 站点信息由 shop_hash / market_* 承载；仓库名称由 partner_name 承载
--
-- 执行前请备份！若索引名与线上一致可直接执行。

USE `rpa-report`;

ALTER TABLE `product_sku_mapping`
  DROP INDEX `uk_psm_partner_external`,
  DROP INDEX `idx_psm_platform_lookup`,
  DROP INDEX `idx_psm_warehouse_lookup`;

ALTER TABLE `product_sku_mapping`
  DROP COLUMN `partner_site_code`,
  DROP COLUMN `warehouse_code`;

ALTER TABLE `product_sku_mapping`
  ADD UNIQUE KEY `uk_psm_partner_external` (
    `partner_code`,
    `partner_type`,
    `shop_hash`,
    `seller_sku`,
    `warehouse_sku`,
    `product_sku`
  ),
  ADD KEY `idx_psm_platform_lookup` (`partner_code`, `shop_hash`, `seller_sku`),
  ADD KEY `idx_psm_warehouse_lookup` (`partner_code`, `warehouse_sku`);
