CREATE TABLE warehouse_rent_detail (
  id BIGINT UNSIGNED PRIMARY KEY AUTO_INCREMENT COMMENT '自增主键',

  provider VARCHAR(16) NOT NULL COMMENT '服务商/来源：HY 或 4PX',          -- 'HY' / '4PX'
  line_hash CHAR(64) NOT NULL COMMENT '明细行唯一指纹（sha256 hex，用于幂等去重）',            -- sha256 hex（参与字段你后续在代码里定）
  doc_no VARCHAR(128) NULL COMMENT '单据号/批次号：HY=Code；4PX=仓租单号（4PX不唯一，仅用于分组/追溯）',               -- HY: Code（通常用于分组/批次号）；4PX: 仓租单号（不唯一）

  charge_date DATE NOT NULL COMMENT '计费日期（取日期部分）',
  warehouse_code VARCHAR(64) NULL COMMENT '仓库代码（HY，如 DEHY）',        -- HY: DEHY
  warehouse_name VARCHAR(128) NULL COMMENT '仓库名称（4PX，如 法国巴黎2仓）',       -- 4PX: 法国巴黎2仓
  currency CHAR(3) NOT NULL COMMENT '币种（如 EUR/RMB）',              -- EUR/RMB

  sku VARCHAR(128) NULL COMMENT 'SKU',
  barcode VARCHAR(128) NULL COMMENT '条码/自定义编码（HY）',              -- HY
  product_name VARCHAR(255) NULL COMMENT '产品名称（HY）',         -- HY

  qty DECIMAL(18,6) NULL COMMENT '数量：HY=Quantity；4PX=SKU数量',
  volume_m3 DECIMAL(18,6) NULL COMMENT '体积（m³）（HY）',
  weight_kg DECIMAL(18,6) NULL COMMENT '重量（kg）（HY）',

  aging_days INT NULL COMMENT '库龄（天）（HY）',                    -- HY
  rent_free_days INT NULL COMMENT '免租天数（HY）',                -- HY
  toll_days INT NULL COMMENT '收费天数（HY）',                     -- HY
  receiving_no VARCHAR(128) NULL COMMENT '入库单号/收货单号（HY）',         -- HY
  putaway_at DATETIME NULL COMMENT '上架时间（HY）',               -- HY

  aging_bucket VARCHAR(32) NULL COMMENT '库龄段（4PX，如 0-30天/30-60天...）',          -- 4PX: 库龄段
  service_category VARCHAR(32) NULL COMMENT '服务类别（4PX）',      -- 4PX
  service_product VARCHAR(64) NULL COMMENT '服务产品（4PX）',       -- 4PX
  fee_type VARCHAR(32) NULL COMMENT '计费类型（4PX）',              -- 4PX: 计费类型
  fee_name VARCHAR(64) NULL COMMENT '费用名称（4PX）',              -- 4PX: 费用名称

  amount DECIMAL(18,6) NOT NULL COMMENT '金额：HY=Product amount；4PX=应收金额（建议作为统一口径）',          -- HY: Product amount；4PX: 应收金额（建议）
  billed_amount DECIMAL(18,6) NULL COMMENT '计费金额（4PX）',       -- 4PX: 计费金额
  discount_amount DECIMAL(18,6) NULL COMMENT '优惠金额（4PX）',     -- 4PX: 优惠金额

  raw_row_json JSON NULL COMMENT '原始行 JSON（用于追溯与排错，可包含 source_file/source_sheet 等）',                 -- 原始行（也可在里面放 source_file/sheet/row 便于追溯）
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',

  UNIQUE KEY uk_line (provider, line_hash),
  INDEX idx_date_wh (charge_date, warehouse_code, warehouse_name),
  INDEX idx_sku_date (sku, charge_date),
  INDEX idx_docno (provider, doc_no)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='仓租明细（HY/4PX 通用明细表）';