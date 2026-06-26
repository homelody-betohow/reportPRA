-- 用途：销售月报表（明细表），由 Python 计算生成后写入
-- 数据样例：python/excel/base/4.1-4.30--月报.xlsx，按「月份 + 商品ID + SKU + 站点」维度记录销售/费用/毛利
-- 字符集：utf8mb4（兼容中文与符号）
--
-- 与日报（020_sales_daily_report.sql）的主要差异：
--   1. 月报 Excel 无「日期」列，月份隐含在文件名（如 4.1-4.30 → 2026-04），落库新增 report_month/period_start/period_end
--   2. 月报无「平台（报表）」「平台（报表识别码）」字段
--   3. 月报新增「仓租识别码」（= 平台 + 产品状态 拼接，属于冗余字段，已不落库）
--   4. 月报新增 3 个采购成本细分：订单采购成本 / 重发采购成本 / 二次上架采购成本
--
-- 字段对应（Excel 中文表头 -> 表字段）：
--   商品ID               -> product_uid
--   SKU                  -> product_sku
--   站点                 -> platform_site
--   平台                 -> platform
--   产品状态             -> product_status         （新品/保留品/不保留老品/分销）
--   二级分类             -> category_lv2
--   三级分类             -> category_lv3
--   销售经理             -> sales_manager
--   销售负责人           -> sales_owner
--   销量                 -> sales_qty
--   平台销售额           -> platform_sales_amount
--   退款数量             -> refund_qty
--   重发数量             -> reship_qty
--   退款额               -> refund_amount
--   销售额               -> sales_amount           （= 平台销售额 - 退款额）
--   测评费               -> review_fee
--   秒杀费               -> seckill_fee
--   广告费(AMZ)          -> ad_fee_amz
--   广告费(非AMZ)        -> ad_fee_non_amz
--   广告费合计           -> ad_fee_total
--   平台费(AMZ)          -> platform_fee_amz
--   平台费(非AMZ)        -> platform_fee_non_amz
--   平台费合计           -> platform_fee_total
--   销售税(AMZ)          -> sales_tax_amz
--   销售税(非AMZ)        -> sales_tax_non_amz
--   销售税合计           -> sales_tax_total
--   派送费               -> shipping_fee
--   海外仓仓租费         -> overseas_warehouse_rent
--   FBA仓租费            -> fba_warehouse_rent
--   仓租合计             -> warehouse_rent_total
--   提现费               -> withdrawal_fee
--   月租                 -> monthly_rent
--   赔偿金额             -> compensation_amount
--   其他分摊费用         -> other_allocated_fee
--   二次上架数量         -> relisting_qty
--   二次上架金额         -> relisting_amount
--   订单采购成本         -> order_purchase_cost     （订单实物总采购成本）
--   重发采购成本         -> reship_purchase_cost
--   二次上架采购成本     -> relisting_purchase_cost
--   采购成本             -> purchase_cost           （销售口径 COGS，按销量分摊后）
--   头程                 -> first_leg_fee
--   关税                 -> tariff
--   毛利                 -> gross_profit
--   毛利率               -> gross_margin_rate（DECIMAL，0.225 表示 22.5%）
--   运营模式             -> op_mode                 （自运营/代运营）
--   供应商               -> supplier
-- 说明：原表中的「平台商品ID识别码 / 站点商品ID识别码 / 仓租识别码」三列均为字段拼接，
--      属于冗余字段，已不落库；如需使用可在查询时用 CONCAT 拼接得到。

SET NAMES utf8mb4;

CREATE TABLE IF NOT EXISTS report_sales_monthly (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '自增主键',

  -- 月份维度（月报无日期列，由文件名解析得到）
  report_month         CHAR(7)  NOT NULL COMMENT '报表月份，格式 YYYY-MM，例如 2026-04',
  period_start         DATE     NOT NULL COMMENT '月份起始日（自然月第一天，例如 2026-04-01）',
  period_end           DATE     NOT NULL COMMENT '月份结束日（自然月最后一天，例如 2026-04-30）',

  -- 商品/站点/平台 维度
  product_uid          VARCHAR(64)  NOT NULL COMMENT '商品ID，例如 25-HWLY-00003',
  product_sku          VARCHAR(64)  NOT NULL COMMENT 'SKU，例如 E56011002',
  platform_site        VARCHAR(64)  NOT NULL COMMENT '站点，例如 AMAZON-BE-FB / MANO-DE-COMMF / TEMU-NF-A',
  platform             VARCHAR(32)  NULL COMMENT '平台，例如 AMAZON-EU / AMAZON-US / MANO-EU / TEMU / LM / OTTO / DLZ-EU / REAL / castorama / CD',

  product_status       VARCHAR(32)  NULL COMMENT '产品状态：新品/保留品/不保留老品/分销',
  category_lv2         VARCHAR(64)  NULL COMMENT '二级分类，例如 淋浴花洒套装 / 厨房龙头',
  category_lv3         VARCHAR(64)  NULL COMMENT '三级分类，例如 传统恒温淋浴花洒套装',

  -- 人员
  sales_manager        VARCHAR(64)  NULL COMMENT '销售经理（可能带平台后缀，如 陈晓佳AMZ）',
  sales_owner          VARCHAR(64)  NULL COMMENT '销售负责人（可能带平台后缀，如 刘思兰TEMU）',

  -- 销量与销售额
  sales_qty                 INT          NOT NULL DEFAULT 0   COMMENT '销量',
  platform_sales_amount     DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '平台销售额（平台口径含退款前金额）',
  refund_qty                INT          NOT NULL DEFAULT 0   COMMENT '退款数量',
  reship_qty                INT          NOT NULL DEFAULT 0   COMMENT '重发数量',
  refund_amount             DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '退款额',
  sales_amount              DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '销售额（= 平台销售额 - 退款额）',

  -- 推广/广告
  review_fee                DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '测评费',
  seckill_fee               DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '秒杀费',
  ad_fee_amz                DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '广告费(AMZ)',
  ad_fee_non_amz            DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '广告费(非AMZ)',
  ad_fee_total              DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '广告费合计（= AMZ + 非AMZ）',

  -- 平台费/税
  platform_fee_amz          DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '平台费(AMZ)',
  platform_fee_non_amz      DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '平台费(非AMZ)',
  platform_fee_total        DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '平台费合计',
  sales_tax_amz             DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '销售税(AMZ)',
  sales_tax_non_amz         DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '销售税(非AMZ)',
  sales_tax_total           DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '销售税合计',

  -- 物流/仓储
  shipping_fee              DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '派送费（尾程派送）',
  overseas_warehouse_rent   DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '海外仓仓租费',
  fba_warehouse_rent        DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT 'FBA 仓租费',
  warehouse_rent_total      DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '仓租合计（= 海外仓 + FBA）',

  -- 其他费用
  withdrawal_fee            DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '提现费',
  monthly_rent              DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '月租（FBA Monthly Storage 等月度费用）',
  compensation_amount       DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '赔偿金额（平台/物流赔付）',
  other_allocated_fee       DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '其他分摊费用',

  -- 二次上架
  relisting_qty             INT           NOT NULL DEFAULT 0  COMMENT '二次上架数量',
  relisting_amount          DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '二次上架金额',

  -- 采购成本（月报特有的细分）
  order_purchase_cost       DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '订单采购成本（订单实物总采购成本）',
  reship_purchase_cost      DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '重发采购成本',
  relisting_purchase_cost   DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '二次上架采购成本',
  purchase_cost             DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '采购成本（销售口径 COGS，按销量分摊后）',

  -- 物流成本
  first_leg_fee             DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '头程',
  tariff                    DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '关税',

  -- 毛利
  gross_profit              DECIMAL(18,6) NOT NULL DEFAULT 0  COMMENT '毛利',
  gross_margin_rate         DECIMAL(10,6) NULL  COMMENT '毛利率（小数形式，0.225 表示 22.5%）',

  -- 运营/供应商
  op_mode                   VARCHAR(16)   NULL  COMMENT '运营模式：自运营/代运营',
  supplier                  VARCHAR(64)   NULL  COMMENT '供应商，例如 启程/海镘/百途鸿',

  -- 时间戳
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '写入时间',
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',

  PRIMARY KEY (id),

  -- 业务自然键：同一月份内，同一商品+SKU+站点 应只有一行
  UNIQUE KEY uk_month_product_sku_site (report_month, product_uid, product_sku, platform_site),

  KEY idx_sales_monthly_month     (report_month),
  KEY idx_sales_monthly_period    (period_start, period_end),
  KEY idx_sales_monthly_sku_month (product_sku, report_month),
  KEY idx_sales_monthly_product   (product_uid),
  KEY idx_sales_monthly_site      (platform_site),
  KEY idx_sales_monthly_platform  (platform),
  KEY idx_sales_monthly_owner     (sales_owner),
  KEY idx_sales_monthly_category  (category_lv2, category_lv3)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='销售月报表';
