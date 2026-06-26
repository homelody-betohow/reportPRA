# 报表系统（rpa-task）

RPA 报表数据导入与利润计算：从共享盘 Excel 导入 MySQL，分步计算订单 SKU 利润。

## 快速开始

```powershell
pip install -r requirements.txt
# 配置 config/db_config.json 后：
python database\db_connection.py
python scripts\dataImport\run_batch.py
python scripts\archive\run_batch.py
```

## 目录结构

```
rpa-task/
├── config/             # 数据库、路径、汇率
├── database/           # db_connection.py
├── scripts/
│   ├── dataImport/     # Excel → MySQL
│   └── archive/        # 利润计算
├── docs/               # 文档与表结构 DDL
└── docker-compose.yml  # MySQL / Redis / Python（可选）
```

## 文档

| 文档 | 说明 |
|------|------|
| [docs/快速开始.md](docs/快速开始.md) | 安装、配置、运行 |
| [docs/项目说明.md](docs/项目说明.md) | 架构与流水线 |
| [docs/Docker部署.md](docs/Docker部署.md) | Docker 与定时邮件 |
| [docs/README.md](docs/README.md) | 文档索引 |
| [scripts/dataImport/README.md](scripts/dataImport/README.md) | 导入脚本参数说明 |
| [scripts/archive/README.md](scripts/archive/README.md) | 利润脚本参数说明 |

## 数据库

- 表结构 DDL：`docs/database/*.sql`
- 连接配置：`config/db_config.json`（复制 `db_config.example.json`）
- Navicat / 局域网连接：[docs/MySQL-连接说明.md](docs/MySQL-连接说明.md)
