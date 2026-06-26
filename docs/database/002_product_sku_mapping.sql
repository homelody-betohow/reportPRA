-- 用途：产品 SKU 映射表（warehouse_sku ↔ platform_sku 在「店铺×站点×仓库」维度的关系）
-- 数据来源：订单 Excel 导入时 UPSERT（python/v2/orders/import_order_shipped.py）
-- 业务规则：
--   - product_sku 默认等于 warehouse_sku（订单流量中没有独立 product_sku 列）
--   - ops_owner 取自订单 Excel 的「平台sku负责人」列（即 sales_order_shipped.platform_sku_owner）
--   - line_hash 参与列：platform / platform_site / shop_name_en / warehouse_name /
--                      warehouse_sku / platform_sku（原平台 SKU 仅存 sales_order_shipped）
--     [改 line_hash 组成需评估历史行哈希迁移成本，谨慎修改]

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS product_sku_mapping (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash            CHAR(64) NOT NULL
    COMMENT '行内容稳定哈希（SHA-256 hex）；参与列见表头注释',

  product_sku          VARCHAR(64)  NOT NULL COMMENT '产品sku（默认= warehouse_sku）',
  platform             VARCHAR(64)  NULL COMMENT '平台',
  platform_site        VARCHAR(64)  NULL COMMENT '站点',
  shop_name_en         VARCHAR(128) NULL COMMENT '店铺英文名',
  warehouse_name       VARCHAR(255) NULL COMMENT '仓库',
  warehouse_sku        VARCHAR(128) NOT NULL COMMENT '仓库sku',
  platform_sku         VARCHAR(255) NULL COMMENT '平台sku',

  dev_owner            VARCHAR(128) NULL COMMENT '开发负责人',
  ops_owner            VARCHAR(128) NULL COMMENT '运营负责人（订单 Excel 的「平台sku负责人」列）',

  source_type          VARCHAR(24)  NULL DEFAULT 'Excel' COMMENT '来源类型：Excel/API',
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),
  UNIQUE KEY uk_mapping_hash (line_hash),
  KEY idx_psm_warehouse_sku (warehouse_sku),
  KEY idx_psm_platform_sku  (platform_sku(64)),
  KEY idx_psm_shop          (platform, platform_site, shop_name_en),
  KEY idx_psm_product_sku   (product_sku)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='产品SKU映射表';
