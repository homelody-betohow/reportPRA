-- 给“已存在的老表”补字段用（不会 DROP 表）
--
-- 如果你线上库的 product_sku_mapping 还没有 product_sku 列，
-- 请在 Navicat / MySQL 客户端中对 rpa-report 库执行本文件里的 SQL。
--
-- 注意：不同 MySQL 版本对 IF NOT EXISTS 的支持不同；
-- 若执行失败，请删掉 IF NOT EXISTS 再执行一次。

ALTER TABLE `product_sku_mapping`
  ADD COLUMN IF NOT EXISTS `product_sku` VARCHAR(64) NOT NULL
  COMMENT '产品sku（默认= warehouse_sku）'
  AFTER `line_hash`;

CREATE INDEX IF NOT EXISTS `idx_psm_product_sku` ON `product_sku_mapping` (`product_sku`);

