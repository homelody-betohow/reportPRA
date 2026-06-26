-- 用途：平台店铺主数据与计费配置（一家店 = platform + platform_site + shop_name_en）
-- 数据来源：
--   - 订单 Excel 导入时自动 INSERT IGNORE（python/v2/orders/import_order_shipped.py）
--   - 佣金/VAT/汇率等缺省字段可由财务在库内手工补全
-- 业务规则：
--   - shop_hash = SHA-256(platform, platform_site, store_account)，与 uk_psc_shop_hash 一致
--   - shop_name_cn 导入时默认取订单 Excel「店铺别名」（即 sales_order_shipped.shop_alias）
--   - shop_status=0：order_sku_profit 组装时跳过该 shop_name_en 的发货行（见 order_sku_profit_steps.py）
--   - step02 VAT 按 platform + platform_site + shop_name_en 关联本表 vat_rate
-- 字符集：utf8mb4

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS platform_shop (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  shop_hash VARCHAR(128) NOT NULL COMMENT '店铺哈希（platform+站点+account）',
  shop_name_en VARCHAR(128) NOT NULL COMMENT '店铺英文名',
  shop_name_cn VARCHAR(128) NOT NULL DEFAULT '' COMMENT '店铺中文名/别名（导入时取自订单 shop_alias）',
  shop_alias VARCHAR(128) NOT NULL DEFAULT '' COMMENT '店铺别名（与 sales_order_shipped.shop_alias 一致，便于对账）',
  store_account VARCHAR(128) NOT NULL COMMENT '店铺账号',
  store_secret VARCHAR(128) NOT NULL COMMENT '店铺密钥',

  platform VARCHAR(64) NOT NULL COMMENT '平台',
  platform_site VARCHAR(64) NOT NULL DEFAULT '' COMMENT '站点',
  market_region VARCHAR(64) NOT NULL DEFAULT '' COMMENT '市场区域',
  market_code VARCHAR(64) NOT NULL DEFAULT '' COMMENT '市场编码',

  currency VARCHAR(16) NOT NULL DEFAULT '' COMMENT '本位/结算币种',
  fx_rate DECIMAL(18,8) NOT NULL DEFAULT 0 COMMENT '汇率（转本位币）',

  commission_type VARCHAR(64) NOT NULL DEFAULT '' COMMENT '佣金类型',
  commission_rate DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT '佣金率（小数，如 0.15 表示 15%）',
  vat_rate_type VARCHAR(64) NOT NULL DEFAULT '' COMMENT 'VAT 税率类型',
  vat_rate DECIMAL(18,6) NOT NULL DEFAULT 0 COMMENT 'VAT 税率（小数，如 0.19 表示 19%）',

  ops_owner VARCHAR(128) NOT NULL DEFAULT '' COMMENT '运营负责人（导入时取自订单 shop_owner）',
  shop_status TINYINT UNSIGNED NOT NULL DEFAULT 1 COMMENT '状态：1=启用 0=停用（停用则不参与利润计算）',

  remark VARCHAR(512) NULL COMMENT '备注',
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),
  UNIQUE KEY uk_psc_shop_hash (shop_hash),
  KEY idx_psc_platform_shop (platform, platform_site, shop_name_en),
  KEY idx_psc_platform (platform),
  KEY idx_psc_platform_site (platform_site),
  KEY idx_psc_store_account (store_account),
  KEY idx_psc_shop_name_en (shop_name_en),
  KEY idx_psc_shop_alias (shop_alias),
  KEY idx_psc_currency (currency),
  KEY idx_psc_ops_owner (ops_owner),
  KEY idx_psc_shop_status (shop_status)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='平台店铺信息表';
