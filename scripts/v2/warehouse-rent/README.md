# `warehouse-rent`：仓租数据（Excel → MySQL）

本目录负责把 `python/excel/daily/` 下的“仓租相关 Excel”导入 MySQL：

- **明细表**：`warehouse_rent_detail`（一行=一条计费明细）
- **汇总表**：`warehouse_rent_daily`（一行=日 + 仓库 + 币种 + SKU 的汇总）

适用场景：

- 报表/分析直接查库（不用每次再读 Excel）
- 多服务商（HY/4PX/AMZ_FBA…）统一落库，便于横向对比

---

## 0) 当前支持的文件

### HY（鸿羽）

- **输入文件名**：`鸿羽*仓-仓租明细*.xlsx`
- **读取 sheet**：`bizWarehouseRentByMonthDetail`（不使用 `仓租` sheet）
- **仓库字段**：
  - `warehouse_code`：来自 Excel 列 `仓库代码(Warehouse Code)`（如 `DEHY`）
  - `warehouse_name`：从文件名截取（如 `鸿羽1仓-仓租明细...` → `鸿羽1仓`）

### 4PX

- **输入文件名**：`4PX*仓-仓租明细*.xlsx`
- **读取 sheet**：第 0 个 sheet（通常显示为 `sheet0`）
- **仓库字段**：
  - `warehouse_name`：优先来自 Excel 列 `计费仓库`（如 `法国巴黎2仓`），为空时用文件名兜底
  - `warehouse_code`：若 Excel 没有“仓库代码”，则由 `warehouse_name` 生成稳定 code（如 `4PX_FR_PARIS2`）

### AMZ_FBA（Amazon FBA 仓租）

- **输入文件名**：`FBA仓租明细*.xlsx`
- **读取 sheet**：第 1 个 sheet（你的样例为 `SellerSku利润报表` 结构）
- **口径**：`FBA仓租费 = abs(仓储费用（已分摊） + 长期仓储费（已分摊）)`
- **过滤**：`sellerSku` 为空的行直接跳过；金额为 0 的行跳过
- **仓库字段（你当前规则）**：
  - `warehouse_code = 站点 + '_' + 店铺`
  - `warehouse_name` 相同（直接等于 `warehouse_code`，便于按账号/站点汇总）

---

## 1) 目标数据表

### `warehouse_rent_detail`（明细）

- **一行 = 一条计费明细**
- **幂等去重**：唯一键 `(provider, line_hash)`（重复运行脚本不会重复插入）
- **追溯**：保留 `doc_no` + `raw_row_json`

建表 SQL：`docs/database/010_warehouse_rent_detail.sql`

### `warehouse_rent_daily`（日汇总，含 SKU 维度）

- **一行 = 日 + 仓库 + 币种 + SKU 的汇总结果**
- upsert 更新：`INSERT ... ON DUPLICATE KEY UPDATE`（脚本已实现）

建表 SQL：`docs/database/010_warehouse_rent_daily .sql`

---

## 2) 脚本说明（入口 + 各服务商导入器 + 汇总）

### `import_all_detail.py`（入口：一键导入 + 汇总）

做的事情：

- 扫描 `daily_dir` 目录
- 找到 HY/4PX/AMZ_FBA 的 Excel 文件
- 调用对应导入器写入 `warehouse_rent_detail`
- 从明细聚合 upsert 到 `warehouse_rent_daily`
- 全流程带日志、事务提交/回滚

### `import_provider_hy_detail.py`（HY 明细导入器）

- 输入：单个 HY Excel 文件
- 读取：sheet `bizWarehouseRentByMonthDetail`
- 输出：写入 `warehouse_rent_detail`

### `import_provider_4px_detail.py`（4PX 明细导入器）

- 输入：单个 4PX Excel 文件
- 读取：第 0 个 sheet（`sheet0`）
- 输出：写入 `warehouse_rent_detail`

### `import_provider_amazon_fba_detail.py`（AMZ_FBA 明细导入器）

- 输入：单个 `FBA仓租明细*.xlsx`
- 读取：第 1 个 sheet（SellerSku 利润报表结构）
- 输出：写入 `warehouse_rent_detail`
- 支持独立运行：

```bash
python "python/v2/warehouse-rent/import_provider_amazon_fba_detail.py" --xlsx "python/excel/daily/FBA仓租明细3.1-3.31.xlsx"
```

### `summary_detail_to_daily.py`（从明细生成日汇总）

- 输入：数据库 `warehouse_rent_detail`
- 输出：upsert 写入 `warehouse_rent_daily`
- 可按 provider 分开汇总（HY / 4PX / AMZ_FBA）

---

## 3) 运行方式

### 方式 A：直接运行 Python（宿主机）

先安装依赖：

```bash
pip install -r python/requirements.txt
```

运行入口脚本：

```bash
python "python/v2/warehouse-rent/import_all_detail.py" --daily-dir "python/excel/daily"
```

### 方式 B：Windows 双击运行（推荐给日常使用）

脚本：`script/run-warehouse-rent-import.bat`

特点：

- 会设置 `PYTHONWARNINGS`，静默 openpyxl 的 “Workbook contains no default style” 警告（不改 Python 代码）
- 支持传参指定 daily 目录（不传则用脚本默认）

示例：

```bat
script\run-warehouse-rent-import.bat
```

或：

```bat
script\run-warehouse-rent-import.bat "e:\gitea\rps-task\rpa-task\python\excel\daily"
```

---

## 4) 文件扫描规则（入口脚本）

入口脚本默认扫描 `python/excel/daily/`，glob 规则如下：

- HY：匹配 `鸿羽*仓-仓租明细*.xlsx`
- 4PX：匹配 `4PX*仓-仓租明细*.xlsx`
- AMZ_FBA：匹配 `FBA仓租明细*.xlsx`

如果后续服务商的文件名规律不同，建议在 `import_all_detail.py` 里新增一条 glob 规则，并新增对应导入器。

---

## 5) 如何新增一个新的仓库/服务商（扩展指南）

假设要新增 `ABC` 服务商（示例）：

1. **新增导入器脚本**（建议命名 `import_provider_abc_detail.py`）
   - 实现函数：`import_abc_file(cur, xlsx: Path) -> int`
   - 负责：读取 Excel → 组装 rows → `INSERT INTO warehouse_rent_detail ... ON DUPLICATE KEY UPDATE`
   - 必须写入：`provider='ABC'`、`line_hash`、`doc_no`（如果有）、`raw_row_json`

2. **在入口脚本中接入**
   - 增加文件扫描规则（glob）
   - 导入你的 `import_abc_file`
   - 导入循环里调用，并累加 counts
   - 汇总阶段调用：`upsert_daily_from_details(cur, "ABC")`

3. **`line_hash` 设计建议**
   - 原则：选择能稳定区分“同一条计费明细”的字段，拼成 dict 后 `json.dumps` 再 sha256
   - 典型字段：`doc_no/warehouse/charge_date/sku/库龄段/金额/币种/费用名/入库单号...`
   - 目的：支持重复导入时幂等去重

4. **字段缺失的处理**
   - 如果某些字段该服务商没有（例如没有 `warehouse_code`），可写 `NULL`
   - 尽量把“原始行”放进 `raw_row_json`，后续补字段时还能回溯

---

## 6) 常见问题

### Q1：为什么会看到 openpyxl 的 “Workbook contains no default style” 警告？

这是某些导出的 xlsx 不包含默认样式导致的提示，一般不影响读取的数据。你如果用 `script/run-warehouse-rent-import.bat` 运行，会自动静默此警告。

### Q2：重复运行会不会重复入库？

不会。明细表 `warehouse_rent_detail` 通过 `(provider, line_hash)` 唯一键去重；汇总表 `warehouse_rent_daily` 通过唯一键做 upsert 更新。

