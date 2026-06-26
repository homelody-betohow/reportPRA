CREATE TABLE temu_order_item (

    id BIGINT PRIMARY KEY AUTO_INCREMENT,
    line_hash CHAR(64) NULL COMMENT '与 sales_order_shipped.line_hash 一致，用于快关联',
    -- =====================
    -- 订单信息
    -- =====================

    order_no VARCHAR(50) NOT NULL COMMENT '订单号',
    sub_order_no VARCHAR(50) COMMENT '子订单号',

    shop_account VARCHAR(100) COMMENT '店铺账号',
    shop_name_en VARCHAR(100) COMMENT '店铺英文名',
    shop_site_name VARCHAR(100) COMMENT '站点名称',
    country_region VARCHAR(50) COMMENT '国家地区',

    order_status VARCHAR(50) COMMENT '订单状态',
    shipping_method VARCHAR(100) COMMENT '发货方式',

    -- =====================
    -- 时间信息
    -- =====================

    created_time DATETIME COMMENT '订单创建时间',
    confirmed_time DATETIME COMMENT '订单确认时间',
    required_ship_deadline DATETIME COMMENT '要求最晚发货时间',

    shipped_time DATETIME COMMENT '实际发货时间',

    estimated_delivery_time DATETIME COMMENT '预计送达时间',

    delivered_time DATETIME COMMENT '实际签收时间',

    -- =====================
    -- 包裹信息
    -- =====================

    package_no VARCHAR(50) COMMENT '包裹号',

    tracking_no VARCHAR(100) COMMENT '运单号',

    logistics_company VARCHAR(50) COMMENT '物流公司',

    -- =====================
    -- 商品信息
    -- =====================

    sku_id varchar(100) COMMENT 'SKU ID',

    skc_id varchar(100) COMMENT 'SKC ID',

    spu_id varchar(100) COMMENT 'SPU ID',

    warehouse_sku varchar(100) COMMENT '仓库SKU',
    product_name VARCHAR(500) COMMENT '商品名称',

    variant_name VARCHAR(200) COMMENT '规格属性',

    quantity INT DEFAULT 1 COMMENT '购买数量',

    warehouse_name VARCHAR(100) COMMENT '库存扣减仓库',

    -- =====================
    -- 金额信息
    -- =====================

    declared_price DECIMAL(10,2) COMMENT '申报价格USD',
    
    order_payment DECIMAL(10,2) COMMENT '订单货款',
    second_payment DECIMAL(10,2) COMMENT '二次收款',

    sales_revenue DECIMAL(10,2) COMMENT '销售回款',
    sales_return DECIMAL(10,2) COMMENT '销售扣回',
    shipping_income DECIMAL(10,2) COMMENT '运费回款',
    shipping_deduction DECIMAL(10,2) COMMENT '运费扣回',
    expected_income DECIMAL(10,2) COMMENT '预计收入',

    currency VARCHAR(10) DEFAULT 'USD',

    -- =====================
    -- 收件人信息（可选）
    -- =====================

    receiver_name VARCHAR(100),

    receiver_phone VARCHAR(100),

    receiver_email VARCHAR(200),

    receiver_address TEXT,

    -- =====================
    -- 原始数据
    -- =====================

    raw_json JSON,

    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    UNIQUE KEY uk_order_sku(order_no, sku_id),

    INDEX idx_temu_line_hash(line_hash),

    INDEX idx_order_no(order_no),

    INDEX idx_sub_order_no(sub_order_no),

    INDEX idx_package_no(package_no),

    INDEX idx_tracking_no(tracking_no),

    INDEX idx_sku_id(sku_id),

    INDEX idx_created_time(created_time),

    INDEX idx_status(order_status)

) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;