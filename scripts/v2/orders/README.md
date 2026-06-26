# `python/v2/orders` — 订单相关 Excel → MySQL

本目录存放**订单/退款/Amazon 交易/TEMU 费用**等导入脚本，以及共用工具 `excel_common.py`。脚本依赖仓库 `python/v2` 下的 `db.py`、`config/` 等，并通常通过 **`.env`** 配置数据库（与 Docker 内运行方式见上级 `python/v2/README.md`）。

---

## 一键入口：`run_import.py`

按固定顺序、**单连接**执行多条导入，末尾一次 `commit`（异常则 `rollback`）。

| 顺序 | 脚本 | 作用摘要 |
|------|------|----------|
| 1 | `import_order_shipped` | `订单统计*.xlsx` → `sales_order_shipped` |
| 2 | `import_order_refund` | `RMA*.xlsx` → `sales_order_refund` |
| 3 | `import_amz_transaction` | `transaction交易明细*.xlsx` → `amz_transaction` |
| 4 | `import_temu_fee` | TEMU Excel → `temu_order_detail`，并同步 `sales_order_shipped` 费用字段 |

常用命令（在仓库 **`python`** 目录下执行时路径写成 `v2/orders/...`）：

```bash
python v2/orders/run_import.py
python v2/orders/run_import.py --order-dir path/to/excel/daily/order
python v2/orders/run_import.py --no-temu
python v2/orders/run_import.py --temu-only
python v2/orders/run_import.py --temu-write-order-days 0   # TEMU 全量；默认近 30 天
```

更多参数见 `run_import.py` 顶部文档字符串：`--shipped-only` / `--refund-only` / `--amz-only` / `--temu-only`、`--temu-file`、`--temu-no-detail-table`、`--temu-no-mail` 等。

> **说明**：`import_order_returned.py` 不在 `run_import` 流水线中，需单独执行（见下文）。

---

## 各脚本说明

### `import_order_shipped.py`

- **数据**：订单统计 Excel（文件名形如 `订单统计*.xlsx`）。
- **落库**：`sales_order_shipped`；同批还会维护 `platform_shop_config`（按 `shop_hash`，无则插入）。
- **特点**：`line_hash` 与列映射见脚本内 `_SHIPPED_MAP`、`LINE_HASH_KEYS`。
- **其它**：支持 `--backfill-mapping-from-shipped` 从已落库发货表补 `product_sku_mapping`（见脚本内示例）。

```bash
python v2/orders/import_order_shipped.py
python v2/orders/import_order_shipped.py --file path/to/订单统计.xlsx
```

---

### `import_order_refund.py`

- **数据**：`RMA*.xlsx`。
- **落库**：`sales_order_refund`。
- **特点**：`line_hash` 仅使用 `LINE_HASH_KEYS` 子集（见脚本注释）。

```bash
python v2/orders/import_order_refund.py
```

---

### `import_amz_transaction.py`

- **数据**：Amazon 交易明细目录下 `transaction交易明细*.xlsx`（已发放 / 已推迟等，由文件名区分 `source_kind`）。
- **落库**：`amz_transaction`（全量明细行 UPSERT）。
- **默认目录**：`python/excel/daily/amazon/`（可用 `--dir` / `--file` 覆盖）。

```bash
python v2/orders/import_amz_transaction.py
python v2/orders/import_amz_transaction.py --dir path/to/excel/daily/amazon
```

---

### `import_temu_fee.py`

- **数据**：`TEMU-订单详情.xlsx`（多 sheet，紫鸟 / RMB / USD 等）。
- **步骤 1**：Excel → `temu_order_detail`（按 `line_hash` UPSERT）。
- **步骤 2**：2a 从 `sales_order_shipped`（`platform=semitemu`）回写明细空字段；2b 用明细补全发货表币种与费用等。
- **表结构**：`docs/database/030_temu_order_detail.sql`。
- **常用参数**：`--only-step1` / `--only-step2`、`--dry-run`、`--write-order-days N`（仅近 N 天写入；不传则全量）、邮件相关见脚本内说明。

```bash
python v2/orders/import_temu_fee.py
python v2/orders/import_temu_fee.py --write-order-days 30
```

`run_import` 调用本脚本时，**默认**传入与 `--write-order-days 30` 等价的逻辑（可通过 `run_import` 的 `--temu-write-order-days` 修改，`0` 表示全量）。

---

### `import_order_returned.py`（未纳入 `run_import`）

- **数据**：鸿羽「二次上架/退件」明细（默认 `*二次上架*.xls` / `.xlsx`，sheet `ReturnOrders`）。
- **落库**：`sales_order_returned`。
- **特点**：可用 `orig_order_no` 在 `sales_order_shipped` 上回填部分字段（见脚本内文档）；**不参与** `run_import.py` 的默认顺序（登记在 `_IMPORT_STEMS_EXCLUDE_ON_DISK` 中排除校验）。

```bash
python v2/orders/import_order_returned.py
```

---

### `excel_common.py`（库模块，非 CLI）

订单目录内各 `import_*.py` 共用的工具：

- 默认订单 Excel 目录：`default_order_excel_dir()` → `python/excel/daily/order`
- 单元格解析：`cell_str`、`cell_dt`、`cell_decimal` 等
- 行级 UPSERT：`upsert_rows`（`INSERT ... ON DUPLICATE KEY UPDATE`）
- `line_hash`：`row_subset_for_line_hash`、`stable_line_hash`

---

## 环境与依赖

- **数据库**：`DB_HOST`、`DB_PORT`、`DB_USER`、`DB_PASS`、`DB_NAME`（及 Docker 下约定）见 `v2/db.py`。
- **Python 包**：各脚本普遍使用 `pandas`、`openpyxl` 或 `mysql-connector-python` 等，以仓库 `python/requirements.txt` 为准。

若只关心「如何跑一键导入」，优先阅读 **`run_import.py` 顶部说明** 与本文 **`run_import.py` 一节** 即可。
