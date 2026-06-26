# 报表系统（report 项目根）

本目录是**报表重构升级**的项目根，与 **`../A_报表/` 现网流水线并行、互不影响**。

- **日常生成报表**：仍在 `py-project` 下执行 `python A_报表/.../runAll_*.py`（不依赖本目录）。
- **本目录**：连接局域网 MySQL、映射查询、配置向导、后续自动化运行器等（按需启用）。

详见 [兼容性说明.md](兼容性说明.md)。

> `A_报表\report\` 仅保留跳转说明，代码都在 **`d:\py-project\report`**。

## 目录结构

```
py-project/
├── report/                 ← 本项目根（你在这里）
│   ├── bootstrap.py        # 路径引导
│   ├── config/             # 配置与配置向导
│   │   └── db_config.json  # 数据库连接（唯一生效）
│   ├── database/           # 连接、映射、表结构
│   │   └── tables/         # 各表 DDL（从线上导出，每表一文件）
│   └── scripts/            # 工具脚本
└── A_报表/                 ← 原有报表流水线（B/C/D/F/G…）
```

## 快速开始

### 1. 安装依赖

```powershell
cd d:\py-project\report
pip install -r requirements.txt
```

### 2. 配置数据库

编辑 `config/db_config.json`，填写局域网 MySQL 连接信息（须能访问公司内网）：

| 项 | 说明 |
|----|------|
| host | MySQL 服务器 IP |
| port | 端口 |
| user / password | 账号密码（向管理员申请） |
| database | 库名（当前为 `rpa-report`） |

### 3. 测试连接

```powershell
python database\db_connection.py
```

### 4. 配置向导（可选）

```powershell
python config\run_config.py
```

### 5. 迁移 Excel 映射表到数据库（可选）

```powershell
python database\migrate_excel_to_db.py
```

## 数据库说明

- 使用**局域网共享 MySQL**，库表由服务器统一维护，本机无需安装 MySQL。
- 各表现有结构见 `database/tables/*.sql`。
- 库名含连字符时，手写 SQL 须加反引号：`` USE `rpa-report`; ``
- 修改连接信息：只改 `config/db_config.json` 即可。

**Navicat 连接**：连接编码设为 **UTF-8 / utf8mb4**，详见 [docs/Navicat中文备注乱码.md](docs/Navicat中文备注乱码.md)。

## 文档

- [快速开始.md](快速开始.md)
- [部署指南.md](部署指南.md)
- [README_重构方案.md](README_重构方案.md)
- [项目总结.md](项目总结.md)

## 运行原有报表（示例）

```powershell
cd d:\py-project\report
python -c "import bootstrap; bootstrap.bootstrap(__file__)"
# 或直接在上级目录运行（与之前相同）：
cd d:\py-project
python A_报表\F_测评\runAll_F.py
```
