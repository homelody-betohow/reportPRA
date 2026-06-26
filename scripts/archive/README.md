# archive — 订单 SKU 利润归档计算

在 `dataImport` 完成 Excel 导入后，从 `sales_order_shipped` 生成并补全 `sales_order_sku_profit`（订单 SKU 毛利表）。

**前置条件**：须先执行 `scripts/dataImport/run_batch.py`，确保本批 `import_batch` 的发货数据已入库。

**推荐入口**：使用 `run_batch.py` 一键按批次顺序执行；也可单独运行各子脚本。批次号默认从 `scripts/dataImport/run_batch.lock` 读取。

---

## 目录说明

| 文件 | 作用 |
|------|------|
| `run_batch.py` | 读取 `import_batch`，按固定顺序调用五个利润脚本 |
| `profit_001_order_sku.py` | 发货表 → 利润表初始化（UPSERT，`calc_node=init`） |
| `profit_002_order_market.py` | 从 `platform_shop` 回填 `market_region`、`market_code` |
| `profit_002_order_price.py` | 从 `temu_order_item` 覆盖 Temu 订单价格（RPA 明细） |
| `profit_003_order_first.py` | 按 SKU 主数据重算头程/关税（`calc_node=first_leg`） |
| `profit_004_order_delivery_amz.py` | 从 `amz_transaction` 汇总 Amazon FBA 派送费 |

公共依赖：

- `scripts/dataImport/batch_lock.py` — 读取 `run_batch.lock` 中的 `import_batch`
- `scripts/dataImport/import_common.py` — `profit_001` 使用的 UPSERT 工具
- `config/db_config.json` — 数据库连接
- `config/common.py` — 汇率、`BTH_ALL_SKU_DETAIL_PATH`（头程主数据 Excel）

表结构 DDL 见 `docs/database/025_sales_order_sku_profit.sql` 等。

---

## 执行顺序与依赖

```
dataImport/run_batch.py          （前置：写入 sales_order_shipped）
    │
archive/run_batch.py
    ├─ 1. profit_001_order_sku.py      初始化利润表
    ├─ 2. profit_002_order_market.py   补全市场区域/编码
    ├─ 3. profit_002_order_price.py    覆盖 Temu 价格（可选但建议）
    ├─ 4. profit_003_order_first.py    重算头程/关税
    └─ 5. profit_004_order_delivery_amz.py  Amazon FBA 派送费
```

**待实现**（脚本占位，未接入 `run_batch`）：

- `profit_004_order_delivery_mano.py` — Mano MMF 派送费
- `profit_005_order_vat.py` — VAT

- **001 必须先跑**：后续脚本均依赖 `sales_order_sku_profit` 中已存在的 `line_hash` 行。
- **002_market 与 002_price 互不依赖**，顺序可互换；均通过 `sales_order_shipped.line_hash` 关联本批数据。
- **003 应最后执行**：需要利润表已有对应行，且会更新 `calc_node` 为 `first_leg`。
- **与 dataImport 的分工**：
  - `dataImport/order_temu.py` 更新 **`sales_order_shipped`** 的价格（来源：全部 `temu_order_item`）
  - `profit_002_order_price.py` 更新 **`sales_order_sku_profit`** 的价格（仅 `temu_order_item.file_name` 非空的 RPA/手工明细）

---

## 快速开始

在 **项目根目录**下执行：

```powershell
cd <项目根目录>

# 前置：Excel 导入（若尚未执行）
python scripts\dataImport\run_batch.py

# 一键利润归档（推荐）
python scripts\archive\run_batch.py
```

`run_batch.py` 会：

1. 从 `dataImport/run_batch.lock` 读取 `import_batch`（或用 `--batch` 覆盖）
2. 将同一批次号传给五个子脚本
3. 某步失败则停止（可用 `--continue-on-error` 继续）

### 常用参数

```powershell
# 手动指定批次号
python scripts\archive\run_batch.py --batch 20260616_203140

# 头程主数据用 Excel
python scripts\archive\run_batch.py --pricing-source excel

# 仅跑头程运费步骤（传给 profit_003 --step 1）
python scripts\archive\run_batch.py --step 1

# 某步失败后仍继续
python scripts\archive\run_batch.py --continue-on-error

# dry-run（001 仍会写库；002/003 仅统计）
python scripts\archive\run_batch.py --dry-run
```

### 单独运行子脚本

```powershell
python scripts\archive\profit_001_order_sku.py
python scripts\archive\profit_002_order_market.py
python scripts\archive\profit_002_order_price.py
python scripts\archive\profit_003_order_first.py

# 指定批次
python scripts\archive\profit_001_order_sku.py --batch 20260616_203140
```

---

## 批次号（import_batch）

| 脚本 | 命令行参数 | 数据库用途 |
|------|------------|------------|
| `run_batch.py` | `--batch` / `--import-batch` | 下发给子脚本 |
| 各子脚本 | `--batch BATCH` | 过滤 `sales_order_shipped.import_batch` |
| 默认 | （无参数） | 读取 `dataImport/run_batch.lock` 的 `import_batch` |

`profit_001` 写入时，`report_hash` = 对应发货行的 `import_batch`。

`profit_002_*` 另支持 `--all`：不按批次过滤，更新全表可匹配行（与 `--batch` 互斥）。

---

## 各脚本说明

### profit_001_order_sku.py — 利润表初始化

| 项 | 说明 |
|----|------|
| 数据源 | `sales_order_shipped`（按 `import_batch`） |
| 目标表 | `sales_order_sku_profit` |
| 关联键 | `line_hash`（与发货表一致） |
| 写入策略 | UPSERT（`uk_sosp_line_hash`） |
| calc_node | 写入 `init` |
| 过滤规则 | `order_type=重发订单` 始终写入；其他类型仅 `order_total_base > 0` 时写入 |
| 字段映射 | `warehouse_sku_qty` → `shipped_qty`；`warehouse_sku` → `product_sku`；毛利/净利暂相同，退款字段恒 0 |

```powershell
python scripts\archive\profit_001_order_sku.py
python scripts\archive\profit_001_order_sku.py --batch 20260616_203140
```

### profit_002_order_market.py — 市场信息回填

| 项 | 说明 |
|----|------|
| 数据源 | `platform_shop` |
| 目标表 | `sales_order_sku_profit` |
| 关联键 | `platform` + `platform_site` + `shop_name_en` |
| 更新字段 | `market_region`、`market_code` |
| 批次范围 | 经 `sales_order_shipped.line_hash` 关联本批发货行 |

```powershell
python scripts\archive\profit_002_order_market.py
python scripts\archive\profit_002_order_market.py --all
python scripts\archive\profit_002_order_market.py --dry-run
```

> 运行前须确认 `sales_order_sku_profit` 表已存在 `market_region`、`market_code` 列，否则脚本报错退出。

### profit_002_order_price.py — Temu 价格覆盖

| 项 | 说明 |
|----|------|
| 数据源 | `temu_order_item`（`file_name` 非空，即 RPA/手工导入的订单详情） |
| 目标表 | `sales_order_sku_profit` |
| 平台范围 | `platform = semitemu` |
| 关联键 | 优先 `line_hash`；回退 `ref_no`+`platform_sku` = `order_no`+`sku_id` |
| 性能 | 分批 IN 查 temu + 按 `id` 分批 UPDATE，避免大表 TRIM/COLLATE JOIN |
| 更新字段 | `order_total_pay` ← `sales_revenue`；`order_goods_base` ← `order_payment` |
| 跳过规则 | `order_type = 重发订单` 的行不更新 |

```powershell
python scripts\archive\profit_002_order_price.py
python scripts\archive\profit_002_order_price.py --all
python scripts\archive\profit_002_order_price.py --dry-run
```

### profit_003_order_first.py — 头程/关税重算

按 SKU 主数据中的「单个头程/关税（RMB/件）」分两步重算利润表头程费用，**不修改** `sales_order_shipped`。

| 项 | 说明 |
|----|------|
| 数据源 | `sales_order_shipped`（范围）+ SKU 主数据（计价） |
| 目标表 | `sales_order_sku_profit` |
| 步骤 1 | 更新 `first_leg_shipping_base`，`calc_node` → `first_leg_shipping` |
| 步骤 2 | 更新 `first_leg_tax_base`；头程已齐则 `calc_node` → `first_leg`，否则 → `first_leg_tax` |
| 计算公式 | `*_base = 单价(RMB/件) × warehouse_sku_qty ÷ RMB_di_EUR` |
| 分销行 | `distribution_lev` 非 0（步骤 1 处理，由 `profit_002_order_market` 预先标记）：仅标记 `calc_node=first_leg` |
| 主数据来源 | 脚本顶部 `pricing_schedule` 或 `--pricing-source`：`excel` / `db` |

**主数据切换**（脚本顶部变量，推荐直接改）：

```python
pricing_schedule = "excel"   # BTH全部SKU明细 Excel
# pricing_schedule = "db"    # product_sku_pricing 表
```

Excel 路径默认取 `config/common.py` 的 `BTH_ALL_SKU_DETAIL_PATH`（共享盘最新 `BTH全部SKU明细-*.xlsx`）。

**市场分组**（决定取哪组头程/关税列）：US / UK / CA / JP / AU / EU，依据 `platform_site`、`shop_alias`、`warehouse_name` 推断。

```powershell
python scripts\archive\profit_003_order_first.py
python scripts\archive\profit_003_order_first.py --step 1
python scripts\archive\profit_003_order_first.py --step 2
python scripts\archive\profit_003_order_first.py --dry-run
python scripts\archive\profit_003_order_first.py --pricing-source excel
python scripts\archive\profit_003_order_first.py --excel-file "\\Betohow\...\BTH全部SKU明细-v2026.06.02.xlsx"
```

### profit_004_order_delivery_amz.py — Amazon FBA 派送费

| 项 | 说明 |
|----|------|
| 数据源 | `amz_transaction`（按 `order_no` + `platform_sku` 汇总） |
| 目标表 | `sales_order_sku_profit` |
| 更新字段 | `delivery_shipping_base` |
| 前置 | 须先导入 `amz_transaction`（`dataImport/amz_transaction.py`） |

```powershell
python scripts\archive\profit_004_order_delivery_amz.py
python scripts\archive\profit_004_order_delivery_amz.py --batch 20260616_203140
python scripts\archive\profit_004_order_delivery_amz.py --all
python scripts\archive\profit_004_order_delivery_amz.py --dry-run
```

---

## calc_node 流转

| 阶段 | 脚本 | calc_node |
|------|------|-----------|
| 初始化 | `profit_001` | `init` |
| 头程运费 | `profit_003` 步骤 1 | `first_leg_shipping` |
| 头程关税 | `profit_003` 步骤 2 | `first_leg`（或中间态 `first_leg_tax`） |

可用 `calc_node` 在库中筛选各行已完成的计算阶段。

---

## 通用参数

| 参数 | 适用脚本 | 说明 |
|------|----------|------|
| `--batch` / `--import-batch` | `run_batch`、全部子脚本 | 指定 `import_batch` |
| `--continue-on-error` | `run_batch` | 某步失败后继续 |
| `--no-color` | `run_batch` | 禁用彩色日志 |
| `--all` | `002_market`、`002_price` | 全表更新，忽略批次 |
| `--dry-run` | `002_market`、`002_price`、`003`、`004_amz` | 仅统计/计算，不写库 |
| `--step 0\|1\|2` | `003` | `0`=两步都跑（默认），`1`=仅头程运费，`2`=仅头程关税 |
| `--pricing-source db\|excel` | `003` | 头程主数据来源 |
| `--excel-file PATH` | `003` | Excel 模式下的 SKU 明细文件 |

---

## 退出码

| 码 | 含义 |
|----|------|
| `0` | 成功 |
| `1` | 无法获取批次号、参数冲突等 |
| `2` | 数据库异常 / 运行时错误 |

---

## 排查建议

1. **无法获取批次号**：先跑 `dataImport/run_batch.py`，或手动 `--batch`；确认 `run_batch.lock` 存在且 `import_batch` 非空。
2. **001 写入 0 条**：检查本批 `sales_order_shipped` 是否有数据；非重发订单是否 `order_total_base <= 0` 被过滤。
3. **市场字段更新失败**：确认 `platform_shop` 已由 `order_shipped` 导入；利润表是否已加 `market_region`、`market_code` 列。
4. **Temu 价格未覆盖**：确认 `temu_order_item.file_name` 非空（RPA 明细）；`ref_no`/`platform_sku` 能否与 Temu 明细匹配。
5. **头程/关税为 0 或未更新**：检查 `warehouse_sku` 是否在主数据中存在；`warehouse_sku_qty` 是否 > 0；`distribution_lev` 非 0 的分销行仅打标不计算金额（须先跑 `profit_002_order_market` 标记分销等级）。
6. **按批次核对**：各脚本日志会打印批次号；库中可按 `report_hash` 或关联 `sales_order_shipped.import_batch` 查询。

---

## 相关文档

- 数据导入：`scripts/dataImport/README.md`
- 项目总览：`README.md`、`docs/项目说明.md`
- 汇率与 Excel 路径：`config/common.py`
- 表结构：`docs/database/025_sales_order_sku_profit.sql`、`005_platform_shop.sql`、`003_product_sku_pricing.sql`
