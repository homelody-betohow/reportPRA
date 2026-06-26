-- 用途：销售退件/退货入库明细落库（与发货 sales_order_shipped、退款 sales_order_refund 并列；
--       侧重「货物流回仓库」的事实：单号、SKU、数量、物流、入库与上架时间等）
-- 说明：python/excel/daily/鸿羽1仓-仓租明细*.xlsx 为仓租计费明细，字段应对齐 warehouse_rent_detail，
--       若业务退件导出 Excel 表头不同，导入脚本里做列映射即可，本表为通用退件行模型。
-- 字符集：utf8mb4

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS sales_order_returned (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash              CHAR(64) NOT NULL COMMENT '行内容稳定哈希（SHA-256 hex），用于去重与增量 UPSERT',

  platform               VARCHAR(64) NULL COMMENT '平台',
  shop_name_en           VARCHAR(128) NULL COMMENT '店铺英文名',
  shop_alias             VARCHAR(128) NULL COMMENT '店铺别名',
  platform_site          VARCHAR(64) NULL COMMENT '站点',
  warehouse_name         VARCHAR(255) NULL COMMENT '仓库名称',
  warehouse_code         VARCHAR(64) NULL COMMENT '仓库代码（与仓租明细 Warehouse Code 一致时可填）',

  return_doc_no          VARCHAR(128) NULL COMMENT '退件单号/WMS退件单号',
  receiving_no           VARCHAR(128) NULL COMMENT '入库单/收货单号（可与仓租明细 Receiving 对齐）',
  rma_no                 VARCHAR(128) NULL COMMENT 'RMA编号（若有，可与 sales_order_refund 业务关联）',

  orig_order_no          VARCHAR(128) NOT NULL COMMENT '原销售订单号（平台订单号）',
  orig_ref_no            VARCHAR(128) NOT NULL DEFAULT '' COMMENT '原订单参考号',
  orig_sales_order_no    VARCHAR(128) NULL COMMENT '原销售订单号（业务侧）',
  orig_tracking_no       VARCHAR(255) NULL COMMENT '原 outbound 跟踪号（可选）',

  return_type            VARCHAR(64) NULL COMMENT '退件类型：买家退货/拒收/截单退回/换货退回等',
  return_status          VARCHAR(64) NULL COMMENT '退件状态：待收货/已收货/已上架/质检不通过等',
  disposition            VARCHAR(64) NULL COMMENT '处置：重新上架/报废/待判/转库存等',

  platform_sku           VARCHAR(255) NULL COMMENT '平台 SKU',
  warehouse_sku          VARCHAR(128) NOT NULL COMMENT '仓库 SKU（如 900008-XXX）',
  product_sku            VARCHAR(128) NULL COMMENT '产品 SKU：导入时取 warehouse_sku 去掉前缀 900008- 后的后缀',
  product_name           VARCHAR(512) NULL COMMENT '产品名称',

  return_qty             DECIMAL(18,6) NOT NULL COMMENT '退件数量（支持小数场景时与源表一致）',
  received_qty           DECIMAL(18,6) NULL COMMENT '实收数量（签收/清点）',
  putaway_qty            DECIMAL(18,6) NULL COMMENT '已上架数量',

  return_tracking_no     VARCHAR(255) NULL COMMENT '退件物流跟踪号',
  carrier                VARCHAR(128) NULL COMMENT '承运商/物流商',
  shipping_method        VARCHAR(128) NULL COMMENT '运输方式',

  apply_at               DATETIME NULL COMMENT '退件申请时间',
  buyer_ship_at          DATETIME NULL COMMENT '买家寄出时间',
  received_at            DATETIME NULL COMMENT '仓库签收时间',
  putaway_at             DATETIME NULL COMMENT '上架完成时间（与仓租 Putaway Date 语义一致）',
  inspected_at           DATETIME NULL COMMENT '质检完成时间',

  length_cm              DECIMAL(18,6) NULL COMMENT '长 cm',
  width_cm               DECIMAL(18,6) NULL COMMENT '宽 cm',
  height_cm              DECIMAL(18,6) NULL COMMENT '高 cm',
  volume_m3              DECIMAL(18,6) NULL COMMENT '体积 m³',
  weight_kg              DECIMAL(18,6) NULL COMMENT '重量 kg',

  currency_code          VARCHAR(16) NULL COMMENT '币种（处理费/货值等）',
  handling_fee           DECIMAL(18,6) NULL COMMENT '退件处理费',
  declared_value         DECIMAL(18,6) NULL COMMENT '申报/核定货值（可选）',

  sales_owner            VARCHAR(128) NULL COMMENT '销售负责人',
  platform_sku_owner     VARCHAR(128) NULL COMMENT '平台 SKU 负责人',
  cs_remark              VARCHAR(2048) NULL COMMENT '客服备注',
  warehouse_remark       VARCHAR(2048) NULL COMMENT '仓库备注',

  profit_calc_node       VARCHAR(24) NULL COMMENT '利润/聚合批次标记（与 shipped/refund 同风格，可选）',
  source_type            VARCHAR(24) NULL COMMENT '来源类型：Excel/API',

  created_at             TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at             TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),
  UNIQUE KEY uk_return_line_hash (line_hash),
  KEY idx_ret_orig_order (orig_order_no),
  KEY idx_ret_received (received_at),
  KEY idx_ret_warehouse_sku (warehouse_sku),
  KEY idx_ret_product_sku (product_sku),
  KEY idx_ret_return_doc (return_doc_no),
  KEY idx_ret_receiving (receiving_no),
  KEY idx_ret_platform (platform)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='销售退件/退货入库明细（按行）';
