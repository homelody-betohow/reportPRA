-- 为已有库增加 temu_order_item.line_hash（与 sales_order_shipped.line_hash 一致，用于快关联）
-- 新库请直接用 temu_order_item.sql 建表，无需执行本脚本。
-- 执行后建议运行：python scripts/dataImport/order_temu.py --backfill-line-hash

SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci;

ALTER TABLE `temu_order_item`
  ADD COLUMN `line_hash` CHAR(64) NULL
    COMMENT '与 sales_order_shipped.line_hash 一致，用于快关联'
    AFTER `id`;

ALTER TABLE `temu_order_item`
  ADD INDEX `idx_temu_line_hash` (`line_hash`);

-- 一次性回填（全表；可按需在 WHERE 中加 import_batch 条件缩小范围）
UPDATE `temu_order_item` AS t
INNER JOIN `sales_order_shipped` AS s
  ON TRIM(s.`ref_no`) COLLATE utf8mb4_unicode_ci = TRIM(t.`order_no`) COLLATE utf8mb4_unicode_ci
 AND TRIM(IFNULL(s.`platform_sku`, '')) COLLATE utf8mb4_unicode_ci
     = TRIM(IFNULL(t.`sku_id`, '')) COLLATE utf8mb4_unicode_ci
 AND TRIM(s.`platform`) COLLATE utf8mb4_unicode_ci = 'semitemu'
SET t.`line_hash` = s.`line_hash`
WHERE (t.`line_hash` IS NULL OR TRIM(t.`line_hash`) = '')
  AND s.`line_hash` IS NOT NULL
  AND TRIM(s.`line_hash`) <> '';
