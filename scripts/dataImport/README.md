# dataImport — 订单数据 Excel 导入

从 ERP 共享目录读取 Excel，写入局域网 MySQL，供后续报表计算使用。

**推荐入口**：使用 `run_batch.py` 一键按批次顺序导入；也可单独运行各子脚本。

---

## 目录说明

| 文件 | 作用 |
|------|------|
| `run_batch.py` | 生成/复用 `import_batch`，按固定顺序调用四个导入脚本 |
| `order_shipped.py` | 订单统计 Excel → `sales_order_shipped`（含 Temu 明细同步、SKU 映射） |
| `order_refund.py` | RMA 退款 Excel → `sales_order_refund` |
| `order_returned.py` | 鸿羽仓二次上架明细 → `sales_order_returned` |
| `order_temu.py` | 根据 `temu_order_item` 回填 `sales_order_shipped` 订单价格 |
| `amz_transaction.py` | transaction交易明细 Excel → `amz_transaction`（Amazon 交易明细） |
| `sku_mapping_import.py` | SKU 映射逻辑（由 `order_shipped.py` 调用，非独立入口） |
| `import_common.py` | 公共工具：单元格转换、`line_hash`、UPSERT |
| `batch_lock.py` | `run_batch.lock` 读写：当日批次复用 |
| `run_batch.lock` | 运行时锁文件（记录 `run_date`、`import_batch`） |

表结构 DDL 见 `report/database/tables/`。

---

## 执行顺序与依赖

```
run_batch.py
    │
    ├─ 1. order_shipped.py    订单发货（先写入发货表）
    ├─ 2. order_refund.py     RMA 退款
    ├─ 3. order_returned.py   二次上架退件（默认从发货表回填店铺/平台等）
    └─ 4. order_temu.py       更新 Temu 订单价格（依赖本批发货行 + temu_order_item）
```

- **必须先跑发货再跑退件**：`order_returned.py` 默认用 `orig_sales_order_no` 匹配 `sales_order_shipped.order_no`，补全 `platform`、`shop_name_en` 等空字段。
- **退款与发货无强依赖**，但流水线仍按「发货 → 退款 → 退件 → Temu 价格」顺序执行，便于用同一批次号核对。
- **Temu 价格更新依赖发货导入**：`order_shipped.py` 会将 `platform=semitemu` 且 `order_type=销售订单` 的行 UPSERT 到 `temu_order_item`；`order_temu.py` 再据此回填 `sales_order_shipped` 的价格字段。
- `order_shipped.py` 导入时还会：
  - 自动补全 `platform_shop`（按 `platform + platform_site + shop_name_en` 去重，不存在则 `INSERT IGNORE`）
  - 抽取 `order_offline_type=一票一件` 的订单行写入 `product_sku_mapping`（`INSERT IGNORE`）

---

## 快速开始

### 环境

1. 在 `report` 目录安装依赖：`pip install -r requirements.txt`
2. 配置数据库：`config/db_config.json`
3. 本机须能访问共享盘 `\\Betohow\...` 及局域网 MySQL

### 一键导入（推荐）

在 **`report` 目录**下执行：

```powershell
cd d:\py-project\report
python scripts\dataImport\run_batch.py
```

脚本会：

1. 解析批次号：显式 `--import-batch` 优先；否则当日 `run_batch.lock` 有效则复用，跨日自动清锁并生成新批次（格式 `YYYYMMDD_HHMMSS`）
2. 将同一批次号传给四个子脚本
3. 默认使用 `config/path_config.py` 中的日期与路径；某步失败则停止（可用 `--continue-on-error` 继续）
4. 执行结束后写入/更新 `run_batch.lock`

### 常用参数

```powershell
# 指定数据日期（覆盖默认 DATE_PATH）
python scripts\dataImport\run_batch.py --date 2026-06-09

# 手动指定批次号（便于对账、重跑；忽略锁文件）
python scripts\dataImport\run_batch.py --import-batch 20260616_120000

# 月报模式（路径子目录为 YYYY-MM）
python scripts\dataImport\run_batch.py --mode 每月

# 退件不从发货表回填
python scripts\dataImport\run_batch.py --no-shipped-enrich

# 某步失败后仍继续后续步骤
python scripts\dataImport\run_batch.py --continue-on-error

# 禁用终端彩色日志
python scripts\dataImport\run_batch.py --no-color
```

---

## 批次号（import_batch）

| 脚本 | 命令行参数 | 数据库字段 / 用途 |
|------|------------|-------------------|
| `run_batch.py` | `--import-batch` / `--batch` | 生成并下发给子脚本；写入 `run_batch.lock` |
| `order_shipped.py` | `--import-batch` / `--batch` | 写入 `sales_order_shipped.import_batch` |
| `order_refund.py` | `--import-batch` | 写入 `sales_order_refund.report_hash` |
| `order_returned.py` | `--import-batch` | 写入 `sales_order_returned.report_hash` |
| `order_temu.py` | `--import-batch` / `--batch` | 读取 `sales_order_shipped.import_batch` 过滤本批行（不写库） |

退款、退件表在库中使用 `report_hash` 列存储批次号；命令行统一使用 `--import-batch` 传参。

单独跑子脚本且未传 `--import-batch` 时：`order_shipped` 会尝试复用当日 `run_batch.lock` 或自动生成批次；`order_refund` / `order_returned` 不写入批次字段；`order_temu` 默认从锁文件读批次。

---

## 各脚本说明

### order_shipped.py — 订单发货

| 项 | 说明 |
|----|------|
| 源文件 | `订单统计-*.xlsx` / `*订单统计*.xlsx` |
| 目标表 | `sales_order_shipped` |
| 表头行 | Excel 第 5 行（`header=4`） |
| 有效行 | 同时有 `order_no`、`warehouse_sku` |
| 去重 | `line_hash` → 唯一键 `uk_order_line_hash` |
| 币种校验 | 读取 A3 单元格，须为 EUR，否则中止 |
| Temu 同步 | `semitemu` + `销售订单` → UPSERT `temu_order_item`（`file_name` 已有值则保留业务字段） |
| SKU 映射 | `一票一件` 行 → `product_sku_mapping`（platform + warehouse 两维度） |

```powershell
python scripts\dataImport\order_shipped.py
python scripts\dataImport\order_shipped.py --date 2026-05-26
python scripts\dataImport\order_shipped.py --file "\\Betohow\...\订单统计-5.1-5.26.xlsx"
python scripts\dataImport\order_shipped.py --import-batch 20260616_120000
```

### order_refund.py — RMA 退款

| 项 | 说明 |
|----|------|
| 源文件 | `RMA-*.xlsx` |
| 目标表 | `sales_order_refund` |
| 工作表 | 优先「RMA退款」，否则第一个 sheet |
| 表头行 | 第 3 行（`header=2`） |
| 有效行 | 同时有 `refund_orig_order_no`、`rma_product_sku` |
| 去重 | `line_hash` → 唯一键 `uk_rma_line_hash` |

```powershell
python scripts\dataImport\order_refund.py
python scripts\dataImport\order_refund.py --date 2026-06-09
python scripts\dataImport\order_refund.py --file "\\Betohow\...\RMA-6.1-6.9.xlsx"
```

### order_returned.py — 二次上架退件

| 项 | 说明 |
|----|------|
| 源文件 | `*二次上架明细-*.xls` / `*.xlsx` |
| 目标表 | `sales_order_returned` |
| 工作表 | 优先 `ReturnOrders`，否则第一个 sheet |
| 表头行 | 第 1 行 |
| 去重 | `line_hash` → 唯一键 `uk_return_line_hash` |
| 发货回填 | 默认开启；`--no-shipped-enrich` 关闭 |

```powershell
python scripts\dataImport\order_returned.py
python scripts\dataImport\order_returned.py --date 2026-06-09
python scripts\dataImport\order_returned.py --file "\\Betohow\...\鸿羽-二次上架明细-6.1-6.9.xls"
python scripts\dataImport\order_returned.py --no-shipped-enrich
```

### order_temu.py — Temu 订单价格回填

根据 `temu_order_item` 更新 `sales_order_shipped` 的价格字段（仅 `platform=semitemu` 且 `order_type=销售订单`）。

| 项 | 说明 |
|----|------|
| 数据源 | `temu_order_item`（由 `order_shipped.py` 同步或 RPA 导入） |
| 目标表 | `sales_order_shipped`（UPDATE 价格列） |
| 关联键 | 优先 `line_hash`；回退 `ref_no=order_no` + `platform_sku=sku_id` |
| 更新字段 | `pay_currency`、`unit_price_pay`、`order_goods_pay`、`order_total_pay`、`platform_shipping_pay` 及对应本位币 EUR 列 |
| 汇率 | `config/common.py`（如 `USD_to_EUR`、`RMB_di_EUR`） |
| 默认范围 | 本批 `import_batch` 对应的发货行 |

```powershell
# 流水线第 4 步（run_batch 自动调用，传入同一 import_batch）
python scripts\dataImport\order_temu.py

# 指定批次
python scripts\dataImport\order_temu.py --import-batch 20260624_115249

# 全表更新（不按批次过滤）
python scripts\dataImport\order_temu.py --all

# 仅统计，不写库
python scripts\dataImport\order_temu.py --dry-run

# 首次部署：将 shipped.line_hash 回填到 temu_order_item（仅填空值）
python scripts\dataImport\order_temu.py --backfill-line-hash
```

### amz_transaction.py — Amazon 交易明细

| 项 | 说明 |
|----|------|
| 源文件 | `transaction交易明细-*.xlsx` |
| 目标表 | `amz_transaction` |
| 表头行 | Excel 第 1 行（`header=0`，可通过 `--header-row` 参数调整） |
| 有效行 | 至少有 `amazon_order_id` 或 `group_id` |
| 去重 | `line_hash` → 唯一键 `uk_amz_txn_line_hash` |
| 默认日期 | **当天**（与其它脚本的 `DATE_PATH` 3 天前不同） |

```powershell
python scripts\dataImport\amz_transaction.py
python scripts\dataImport\amz_transaction.py --date 2026-06-09
python scripts\dataImport\amz_transaction.py --file "\\Betohow\...\transaction交易明细-*.xlsx"
python scripts\dataImport\amz_transaction.py --header-row 1  # 表头在第 2 行
```

### sku_mapping_import.py — SKU 映射（库模块）

由 `order_shipped.py` 在写入发货表前调用，不单独作为命令行入口。

| 项 | 说明 |
|----|------|
| 目标表 | `product_sku_mapping` |
| 导入范围 | 仅 `order_offline_type=一票一件` |
| 跳过规则 | `warehouse_sku` 以 `AMZN.GR` 开头（Amazon FBA 仓 SKU） |
| 每行产出 | platform 维度 + warehouse 维度各 1 条（`mapping_type=single`） |
| 组合 SKU | `platform_sku` 含 `+` 时整组映射到单个 `warehouse_sku` |
| 写入策略 | `INSERT IGNORE`（`line_hash` 去重） |

---

## 数据源路径

路径由 `config/path_config.py` 控制：

| 配置项 | 含义 |
|--------|------|
| `MODE_RUN` | `日报` / 月报，决定默认 `DATE_PATH` 格式 |
| `DATE_PATH` | 日报：3 天前 `YYYY-MM-DD`；月报：当月 `YYYY-MM` |
| `MODE_PATTERN` | `每天` / `每月`，决定共享盘子目录 |
| `ERP_ORDER_STA_PATH` | 订单统计、RMA 所在根路径 |
| `SECOND_RELISTING_PATH` | 二次上架明细根路径 |

### 日报示例（`MODE_PATTERN=每天`）

| 脚本 | 路径模式 |
|------|----------|
| 发货 / 退款 | `\\Betohow\...\每天\ERP订单、RMA下载\{YYYY-MM-DD}\` |
| 退件 | `\\Betohow\...\每天\鸿羽仓二次上架明细\{M.D}\` |
| transaction 交易明细 | `\\Betohow\...\每天\transaction交易明细\{YYYY-MM-DD}\` |

退件目录日期为 **`M.D`** 格式，由 `--date` 或 `DATE_PATH` 自动换算，例如 `2026-06-09` → `6.9`。

### 月报示例（`MODE_PATTERN=每月`）

| 脚本 | 路径模式 |
|------|----------|
| 发货 / 退款 | `...\每月\ERP订单、RMA下载\{YYYY-MM}\` |
| 退件 | `...\每月\鸿羽仓二次上架明细\{M.D}\`（仍由具体日期换算） |
| transaction 交易明细 | `...\每月\transaction交易明细\{YYYY-MM}\` |

未指定 `--date`、`--dir`、`--file` 时，各脚本使用 `DATE_PATH` 解析默认目录。

---

## 写入策略

- **UPSERT**：`INSERT ... ON DUPLICATE KEY UPDATE`，以 `line_hash` 唯一键去重，重复导入会更新行内容（发货、退款、退件、Temu 明细表）。
- **INSERT IGNORE**：`platform_shop`、`product_sku_mapping` 仅新增，已存在则跳过。
- **事务**：每个 Excel 文件处理完后 `commit`；单文件失败会 `rollback` 该文件所在事务。
- **line_hash**：对约定字段子集做「键排序 JSON + SHA-256」，算法见 `import_common.stable_line_hash`。

---

## 子脚本通用参数

| 参数 | 说明 |
|------|------|
| `--date YYYY-MM-DD` | 覆盖日期子目录 |
| `--mode 每天\|每月` | 覆盖 `path_config.MODE_PATTERN` |
| `--dir DIR` | 直接指定 Excel 所在目录 |
| `--file FILE` | 指定单个 Excel 文件 |
| `--import-batch BATCH` | 写入/读取批次号（见上表） |

`order_temu.py` 另有：`--all`、`--dry-run`、`--backfill-line-hash`。

---

## 退出码

| 码 | 含义 |
|----|------|
| `0` | 成功 |
| `1` | 未找到 Excel、`run_batch` 有步骤失败，或参数错误 |
| `2` | 目录不存在等路径错误 / 数据库异常 |

---

## 排查建议

1. **目录不存在**：检查共享盘是否挂载、`path_config.py` 中 `DATE_PATH` 是否与 ERP 实际落盘日期一致。
2. **未找到文件**：确认文件名前缀（`订单统计`、`RMA`、`二次上架明细`）及退件目录是否为 `M.D` 格式。
3. **发货导入中止（币种）**：订单统计 Excel A3 须为 EUR，与报表脚本要求一致。
4. **退件店铺为空**：先确认 `order_shipped` 已导入，且 `orig_sales_order_no` 能在 `sales_order_shipped.order_no` 中匹配。
5. **Temu 价格未更新**：确认 `order_shipped` 已同步 `temu_order_item`；检查 `import_batch` 是否一致；无 `line_hash` 时执行 `--backfill-line-hash`。
6. **按批次核对**：`run_batch.py` 日志会打印 `import_batch=...`，可在库中按 `import_batch` / `report_hash` 过滤查询。

---

## 相关文档

- 项目总览：`report/README.md`
- 数据库连接：`report/database/db_connection.py`、`report/config/db_config.json`
- 汇率配置：`report/config/common.py`
- 表结构：`report/database/tables/sales_order_*.sql`、`temu_order_item.sql`、`product_sku_mapping.sql`
