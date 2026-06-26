-- 用途：已有库为 sales_order_sku_profit 增加 calc_node（与 sales_order_shipped.profit_calc_node 呼应）
-- 若建表已执行含 calc_node 的 025 全量脚本，则无需执行；若已存在该列会报错 Duplicate column，可忽略。

SET NAMES utf8mb4;

ALTER TABLE `sales_order_sku_profit`
  ADD COLUMN `calc_node` VARCHAR(24) NULL
    COMMENT '与 sales_order_shipped.profit_calc_node 一致：本行利润写入/批次标记（最长24）'
    AFTER `net_margin_rate`,
  ADD KEY `idx_sosp_calc_node` (`calc_node`);
