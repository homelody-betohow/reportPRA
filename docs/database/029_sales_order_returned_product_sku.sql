-- 用途：旧库从「含 barcode」等结构迁到与 028 一致（product_sku，无 barcode）时参考。
-- 新建库只执行 docs/database/028_sales_order_returned.sql 即可。
--
-- 1) 增加 product_sku（若尚未存在）：
-- ALTER TABLE `sales_order_returned`
--   ADD COLUMN `product_sku` VARCHAR(128) NULL COMMENT '产品 SKU：warehouse_sku 去掉前缀 900008- 后的后缀'
--   AFTER `warehouse_sku`;
--
-- 2) 按规则回填（前缀「900008-」长度为 8，后缀从第 9 个字符起）：
-- UPDATE `sales_order_returned`
-- SET `product_sku` = SUBSTRING(`warehouse_sku`, 9)
-- WHERE `warehouse_sku` LIKE '900008-%' AND (`product_sku` IS NULL OR `product_sku` = '');
--
-- 3) 若存在 barcode 列且不再需要：
-- ALTER TABLE `sales_order_returned` DROP COLUMN `barcode`;

SET NAMES utf8mb4;
