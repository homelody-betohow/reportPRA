# Docker 部署

项目通过 `docker-compose.yml` 提供三个服务：

| 服务 | 说明 |
|------|------|
| **mysql** | MySQL 8.0，数据目录 `docker/mysql/data/` |
| **redis** | 可选缓存 |
| **python** | 挂载项目根目录，内置 cron 定时任务 |

## 启动

```powershell
# 推荐：自动创建 .env 并等待 MySQL 就绪
.\scripts\docker-up.ps1

# 或手动
cp .env.example .env   # 首次
docker compose up -d
```

## 环境变量

复制 `.env.example` 为 `.env` 后修改：

| 变量 | 说明 |
|------|------|
| `MYSQL_PORT` | 宿主机映射端口（默认 3306） |
| `DB_NAME` / `DB_USER` / `DB_PASS` | 业务库与账号 |
| `MYSQL_ROOT_PASSWORD` | root 密码 |
| `SMTP_*` / `MAIL_*` | 心跳邮件（见下文） |

脚本直连数据库时使用 `config/db_config.json`，须与 Docker 中 MySQL 地址、端口一致。

## 初始化数据库

Docker 首次启动**不会**自动导入业务表 DDL。请在 MySQL 中按顺序执行 `docs/database/*.sql`（至少包含订单、利润相关表），或在容器已启动时运行 `.\scripts\init-db.ps1`。

Navicat 连接说明见 [MySQL-连接说明.md](MySQL-连接说明.md)。

## Python 容器与定时任务

- 镜像定义：`docker/python/Dockerfile`
- 入口：`docker/python/entrypoint.sh`（启动 cron）
- 定时配置：`docker/python/crontab`（每天 8:00）
- 心跳脚本：`scripts/jobs/heartbeat.py`

### 手动测试心跳

```bash
docker compose up -d --build python
docker compose exec python python /app/scripts/jobs/heartbeat.py
```

### 查看 cron 日志

```bash
docker compose logs -f python
```

邮件配置与排障详见 [邮件通知心跳检测说明.md](邮件通知心跳检测说明.md)。

## 停止服务

```powershell
.\scripts\docker-down.ps1
# 或
docker compose down
```

## 注意事项

- `docker/mysql/data/` 须为空目录才能完成首次 MySQL 初始化，勿提交到 Git。
- Windows 下 shell 脚本使用 LF 换行（见 `.gitattributes`），避免 cron 报 `bash\r` 错误。
- 构建 Python 镜像时依赖根目录 `requirements.txt`（非 `python/requirements.txt`）。
