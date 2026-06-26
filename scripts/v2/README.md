# v2：Excel → MySQL（方案A：永远用最新）

本目录实现“**方案A**”：

- 用 `docs/database/001_eu_pricing_tables.sql` 建表
- 每次做报表前：**TRUNCATE 清空表** → 从 `python/excel/base/欧洲平台定价表.xlsx` **全量导入** → 报表脚本直接查 MySQL（永远最新）

## 1. 安装依赖

在 **docker 的 `python` 容器内**运行（推荐，与你的部署一致）：

```bash
docker compose exec python pip install -r python/requirements.txt
```

（如果你确实要在宿主机运行，也可以 `pip install -r python/requirements.txt`）

## 2. 配置数据库

使用仓库根目录的 `.env`（已包含 `DB_NAME/DB_USER/DB_PASS/MYSQL_PORT`）。

说明：

- 在 **容器内**运行时，脚本会自动默认：`DB_HOST=mysql`、`DB_PORT=3306`（无需你额外配置）
- 在 **宿主机**运行时，脚本会自动默认：`DB_HOST=127.0.0.1`、`DB_PORT=${MYSQL_PORT}`

## 3. 执行导入

在 **docker 的 `python` 容器内**执行（推荐）：

```bash
docker compose exec python python -m python.v2.import_eu_pricing_to_mysql
```

自定义 Excel 路径：

```bash
docker compose exec python python -m python.v2.import_eu_pricing_to_mysql --xlsx "python/excel/base/欧洲平台定价表.xlsx"
```

## 说明

- 当前脚本已导入：`汇率表 / MANO尾程 仓租 / 欧洲平台定价表 / TEMU / MANO-UK / RDC / Conforama`
- `基础表` 结构是多行表头且列非常多，后续若你要用到它的数据，我再把它的“规范化导入”补齐。

