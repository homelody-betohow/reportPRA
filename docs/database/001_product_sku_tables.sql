-- 用途：产品 SKU 主数据表（python/excel/base/产品信息库.xlsx → sheet「产品信息表」）
-- 表头：第 5 行（一级表头）+ 第 6 行（二级表头），数据从第 7 行开始
-- 字符集：utf8mb4
--
-- 业务主键说明：
--   - 仓库识别码（warehouse_sku）在源数据中 100% 唯一（1983/1983），作为业务唯一键
--   - 产品编码（product_code）存在重复（约 29 条），仅做普通索引
--   - 商品ID（product_id，如 25-CFLT-00001）为 SPU/型号，多颜色变体共用，仅做普通索引
--
-- 导入策略：
--   - 每次导入生成 import_batch_id，便于回溯与回滚
--   - line_hash 为业务行内容稳定哈希，UPSERT 时用于变更检测
--   - 业务幂等键为 warehouse_sku（UNIQUE KEY）

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS product_sku (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash         CHAR(64)     NULL COMMENT '行内容稳定哈希（SHA-256 hex），变更检测用',

  -- ===== 业务标识（来自 Excel 第 6/12/16 列）=====
  product_sku       VARCHAR(64)  NOT NULL COMMENT '产品SKU（如 E51033005）唯一键',
  product_uid       VARCHAR(64)  NULL     COMMENT '商品ID/SPU 型号（如 25-CFLT-00001，多颜色变体共用）',
  warehouse_ref     VARCHAR(64)  NULL COMMENT '仓库识别码（如 JD-03-L-001）',

  -- ===== 分类信息 =====
  category_lv1      VARCHAR(64)  NULL COMMENT '一级分类（如 厨房龙头、淋浴花洒套装）',
  category_lv2      VARCHAR(64)  NULL COMMENT '二级分类（如 厨房龙头、淋浴花洒套装）',
  category_lv3      VARCHAR(64)  NULL COMMENT '三级分类（如 传统恒温淋浴花洒套装）',
  category_code     VARCHAR(8)   NULL COMMENT '产品类别代码（命名规则代码，如 03/04/01/0/99/08）',

  -- ===== 供应商 =====
  supplier_name     VARCHAR(128) NULL COMMENT '供应商名称',

  -- ===== 商品基础属性 =====
  product_unit      VARCHAR(16)  NULL COMMENT '单位（如 套）',
  product_color     VARCHAR(32)  NULL COMMENT '产品颜色（如 铬色/雅黑色/拉丝色）',

  -- ===== 商品生命周期 / 核算 =====
  amz_lifecycle     VARCHAR(16)  NULL COMMENT 'AMZ 新老品状态（新品/保留品/不保留老品）',
  local_lifecycle   VARCHAR(16)  NULL COMMENT '本土平台新老品状态（新品/保留品/不保留老品）',
  accounting_class  VARCHAR(32)  NULL COMMENT '核算分类（如 全新品/转厂新质新品/颜色变体新品/保留品）',

  -- ===== 采购参数 =====
  purchase_moq      INT UNSIGNED NULL COMMENT 'MOQ 最小起订量（件）',
  purchase_lead_days    INT UNSIGNED NULL COMMENT '采购交期（天）',
  carton_qty        INT UNSIGNED NULL COMMENT '箱规（每外箱产品数）',
  cost_price_cny    DECIMAL(18,6) NULL COMMENT '成本价（人民币）',

  -- ===== 重量 =====
  unit_weight_g     DECIMAL(12,2) NULL COMMENT '单件重量（g）',
  carton_gross_g    DECIMAL(12,2) NULL COMMENT '箱规毛重（g）',

  -- ===== 内箱尺寸（cm）=====
  inner_box_l_cm    DECIMAL(10,2) NULL COMMENT '内箱长（cm）',
  inner_box_w_cm    DECIMAL(10,2) NULL COMMENT '内箱宽（cm）',
  inner_box_h_cm    DECIMAL(10,2) NULL COMMENT '内箱高（cm）',

  -- ===== 外箱尺寸（cm）=====
  outer_box_l_cm    DECIMAL(10,2) NULL COMMENT '外箱长（cm）',
  outer_box_w_cm    DECIMAL(10,2) NULL COMMENT '外箱宽（cm）',
  outer_box_h_cm    DECIMAL(10,2) NULL COMMENT '外箱高（cm）',

  -- ===== 头程运费（人民币，按市场分摊到单件）=====
  first_leg_eu_au_cny DECIMAL(18,6) NULL COMMENT '头程运费 EU/AU（RMB/件）',
  first_leg_us_cny    DECIMAL(18,6) NULL COMMENT '头程运费 US（RMB/件）',
  first_leg_uk_cny    DECIMAL(18,6) NULL COMMENT '头程运费 UK（RMB/件）',

  -- ===== 关税（人民币，按市场分摊到单件）=====
  duty_eu_cny       DECIMAL(18,6) NULL COMMENT '关税 EU（RMB/件）',
  duty_us_cny       DECIMAL(18,6) NULL COMMENT '关税 US（RMB/件）',
  duty_uk_cny       DECIMAL(18,6) NULL COMMENT '关税 UK（RMB/件）',

  -- ===== 来源与时间戳 =====
  source_type       VARCHAR(24)  NULL DEFAULT 'Excel' COMMENT '来源类型：Excel/API',
  created_at        TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at        TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),
  -- 业务唯一键：产品SKU（源数据已验证 100% 唯一），UNIQUE KEY 自身即为索引，无需再建同列普通索引
  UNIQUE KEY uk_product_sku (product_sku),

  -- 高频查询索引
  KEY idx_psm_product_uid   (product_uid),
  KEY idx_psm_warehouse_ref (warehouse_ref),
  KEY idx_psm_category_lv2  (category_lv2),
  KEY idx_psm_category_lv3  (category_lv3),
  KEY idx_psm_supplier      (supplier_name),
  KEY idx_psm_color         (product_color),
  KEY idx_psm_amz_lifecycle (amz_lifecycle)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='产品SKU数据表';
