-- 用途：RMA 退款 Excel 导出落库（python/excel/daily/order/RMA-*.xlsx，表头在第 3 行）
-- 字符集：utf8mb4

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS sales_order_refund (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',
  line_hash            CHAR(64) NOT NULL COMMENT '行内容稳定哈希（SHA-256 hex），用于去重与增量 UPSERT',

  platform             VARCHAR(64) NULL COMMENT '平台',
  shop_alias           VARCHAR(128) NULL COMMENT '店铺别名',
  shop_name_en         VARCHAR(128) NULL COMMENT '店铺英文名',
  platform_site        VARCHAR(64) NULL COMMENT '站点',
  warehouse_name       VARCHAR(255) NULL COMMENT '仓库名称',
  order_dest_country   VARCHAR(64) NULL COMMENT '订单目的国家',

  rma_created_at       DATETIME NULL COMMENT 'RMA创建时间',
  rma_audit_at         DATETIME NULL COMMENT 'RMA审核时间',
  rma_refund_at        DATETIME NULL COMMENT 'RMA退款时间',
  orig_order_paid_at   DATETIME NULL COMMENT '原订单付款时间',

  refund_orig_order_no VARCHAR(128) NOT NULL COMMENT '退款原订单号',
  refund_orig_ref_no   VARCHAR(128) NOT NULL DEFAULT '' COMMENT '退款原订单参考号',
  refund_orig_track_no VARCHAR(255) NULL COMMENT '退款原订单跟踪号',
  paypal_refund_txn_no VARCHAR(255) NULL COMMENT 'PayPal退款交易号',
  refund_type          VARCHAR(64) NULL COMMENT '退款类型',
  shipping_method      VARCHAR(128) NULL COMMENT '运输方式',
  shipping_method_name VARCHAR(255) NULL COMMENT '运输方式名称',
  refund_status        VARCHAR(64) NULL COMMENT '退款状态',
  refund_method        VARCHAR(64) NULL COMMENT '退款方式',
  rma_product_sku      VARCHAR(128) NOT NULL COMMENT 'RMA产品',
  rma_product_qty      DECIMAL(18,6) NULL COMMENT 'RMA产品数量',
  currency_code        VARCHAR(16) NULL COMMENT '币种',
  product_name         VARCHAR(512) NULL COMMENT '产品名称',
  category_lv1         VARCHAR(128) NULL COMMENT '一级品类',
  category_lv2         VARCHAR(128) NULL COMMENT '二级品类',
  category_lv3         VARCHAR(128) NULL COMMENT '三级品类',
  product_style        VARCHAR(512) NULL COMMENT '产品款式',
  refund_amount        DECIMAL(18,6) NULL COMMENT '退款金额',
  refund_reason        VARCHAR(512) NULL COMMENT '退款原因',
  platform_refund_reason VARCHAR(512) NULL COMMENT '平台退款原因',
  created_by           VARCHAR(128) NULL COMMENT '创建人',
  refund_remark        VARCHAR(2048) NULL COMMENT '退款备注',
  finance_remark       VARCHAR(2048) NULL COMMENT '财务备注',
  default_buyer_acct   VARCHAR(128) NULL COMMENT '产品默认采购员账号',
  default_buyer_name   VARCHAR(128) NULL COMMENT '产品默认采购员',
  sales_owner_acct     VARCHAR(128) NULL COMMENT '销售负责人账号',
  sales_owner          VARCHAR(128) NULL COMMENT '销售负责人',
  dev_owner_acct       VARCHAR(128) NULL COMMENT '开发负责人账号',
  dev_owner            VARCHAR(128) NULL COMMENT '开发负责人',
  ops_owner            VARCHAR(128) NULL COMMENT '运营负责人',
  product_issue_type   VARCHAR(128) NULL COMMENT '产品问题类型',
  issue_category       VARCHAR(128) NULL COMMENT '问题分类',
  product_issue        VARCHAR(2048) NULL COMMENT '产品问题',
  profit_calc_node     VARCHAR(24) NULL COMMENT '利润计算节点',
  source_type          VARCHAR(24) NULL COMMENT '来源类型：Excel/API',

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),

  UNIQUE KEY uk_rma_line_hash (line_hash),
  KEY idx_rma_refund_orig (refund_orig_order_no),
  KEY idx_rma_refund_at (rma_refund_at),
  KEY idx_rma_platform (platform)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='RMA订单退款明细';
