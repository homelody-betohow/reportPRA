# MySQL 局域网访问与 Navicat 连接说明

本文说明本项目 Docker 中 MySQL 的监听方式，以及在本机用 Navicat、在内网其他电脑上连接的方式与安全注意点。

---

## 1. 项目内配置检查结论

| 项目 | 位置 | 说明 |
|------|------|------|
| 端口映射 | `docker-compose.yml` → `mysql.ports` | 使用 `0.0.0.0:${MYSQL_PORT:-3306}:3306`，将容器 3306 映射到**宿主机所有网卡**，局域网可通过宿主机 IP 访问。 |
| 监听地址 | `docker/mysql/conf.d/my.cnf` | `bind-address = 0.0.0.0`，mysqld 接受来自任意接口的连接（仍需用户权限与防火墙放行）。 |
| 远程用户 | `docker/mysql/init/01-grant-remote.sh` | 仅在**数据目录首次初始化**时执行：为 `MYSQL_USER` 创建 `'%'` 主机账号，并授权业务库。 |

**结论：** 当前配置已支持「宿主机监听 + MySQL 对外监听 + 业务用户远程授权」；若仍连不上，多为防火墙、网段、端口或「库已初始化过未执行 init」导致，见下文「常见问题」。

---

## 2. 连接前准备

1. **启动容器**

   ```bash
   cd /path/to/rpa-task
   docker compose up -d mysql
   ```

2. **确认端口**  
   宿主机实际端口以 `.env` 中的 `MYSQL_PORT` 为准（未设置时默认为 `3306`）。若改为例如 `13306`，则所有连接里的端口都填 `13306`。

3. **账号与库名**（与 `.env` / `.env.example` 一致，请使用你自己修改后的值）

   - 业务用户：`DB_USER` / `DB_PASS`
   - 业务库：`DB_NAME`
   - 管理员：`root` / `MYSQL_ROOT_PASSWORD`（**不建议**对局域网开放 root，仅内网调试且需强密码）

4. **宿主机局域网 IP**  
   - macOS：系统设置 → 网络，或终端执行 `ipconfig getifaddr en0`（有线可能是 `en1` 等）。  
   - 内网其他机器连接时使用该 IP，不要用 `127.0.0.1`（那代表「对方自己本机」）。

5. **防火墙**  
   - **macOS**：系统设置 → 网络 → 防火墙，允许 `Docker` / `com.docker.backend` 或按需放行 `MYSQL_PORT`。  
   - **路由器**：一般无需改；确保客户端与宿主机在同一局域网或可路由网段。

---

## 3. 本机使用 Navicat 连接 MySQL

适用于：Navicat 与 Docker **在同一台电脑**上运行。

| 字段 | 填写方式 |
|------|----------|
| 连接名 | 自定义，如 `rpa-task-local` |
| 主机 | `127.0.0.1` 或 `localhost` |
| 端口 | `.env` 中的 `MYSQL_PORT`（默认 `3306`） |
| 用户名 | **优先**填 `.env` 的 `DB_USER`（如 `rpa-bth`） |
| 密码 | 填对应 `DB_PASS`（**不要用错成 root 的密码**） |

连接成功后，选择数据库为 `DB_NAME`（默认示例为 `rpa-bth`）。

**SSL：** 本地开发一般选「不使用 SSL」或关闭 SSL（除非你已经为 MySQL 配置了证书）。

---

## 4. 内网其他机器连接 MySQL

适用于：办公室/家里另一台电脑、手机热点下的笔记本等（与运行 Docker 的电脑在同一二层或可互通网络）。

| 字段 | 填写方式 |
|------|----------|
| 主机 | 运行 Docker 的那台电脑的 **局域网 IP**（如 `192.168.1.100`） |
| 端口 | 与宿主机相同的 `MYSQL_PORT`（默认 `3306`） |
| 用户名 | 建议使用 `DB_USER`（业务账号） |
| 密码 | `DB_PASS` |

Navicat、MySQL Workbench、命令行客户端均可使用相同参数：

```bash
mysql -h 192.168.1.100 -P 3306 -u rpa-bth -p
```

（将 IP、端口、用户名换成你的实际值。）

---

## 5. 常见问题

### 5.1 `01-grant-remote.sh` 未执行 / 业务用户不能远程登录

初始化脚本**只在 MySQL 数据目录为空、第一次创建库时**运行。若你早年启动过容器且 `./docker/mysql/data` 里已有数据，后来才加入 init 脚本，则需要**手动**在容器内执行授权（示例）：

```bash
docker compose exec mysql mysql -uroot -p
```

进入后（请替换用户名、密码、库名）：

```sql
CREATE USER IF NOT EXISTS 'rpa-bth'@'%' IDENTIFIED BY '你的DB_PASS';
GRANT ALL PRIVILEGES ON `rpa-bth`.* TO 'rpa-bth'@'%';
FLUSH PRIVILEGES;
```

### 5.2 本机能连，内网其他机器不能连

- 检查宿主机防火墙是否放行 `MYSQL_PORT`。  
- 检查客户端 ping / `telnet 宿主机IP 端口` 或 `nc -zv 宿主机IP 端口` 是否通。  
- 确认使用的是**宿主机局域网 IP**，不是容器的 IP。

### 5.3 端口冲突

若本机已有 MySQL 占用 3306，可在 `.env` 中设置 `MYSQL_PORT=13306`，然后 `docker compose up -d`；连接时端口一律改为 `13306`。

### 5.4 Navicat 报错 `1045 - Access denied for user 'root'@'172.18.0.1'`

说明：**能出现这条报错，说明端口、网络已经打通**（Navicat 已经连到某个 MySQL），失败在**账号或密码**。

1. **`172.18.0.1` 是什么**  
   在 macOS / Windows 上用 `127.0.0.1` 连接 Docker 映射端口时，容器里的 MySQL 往往把客户端地址看成 Docker 网桥网关（常见为 `172.18.0.1`），这是正常现象，不必改成填这个 IP。

2. **优先改用业务账号（推荐）**  
   - 用户名：`.env` 里的 `DB_USER`（默认示例 `rpa-bth`）  
   - 密码：`.env` 里的 `DB_PASS`（默认示例 `rpa-bth_secret`）  
   - 端口：必须与 `.env` 的 `MYSQL_PORT` 一致（例如你设为 **3309** 就填 **3309**，不要仍填 3306）。  
   业务用户在首次初始化时会创建 `用户名@'%'`，适合从宿主机 Navicat 连接。

3. **仍要用 `root` 时**  
   - 密码必须与 `.env` 中的 **`MYSQL_ROOT_PASSWORD` 完全一致**（复制时注意首尾空格）。  
   - 若曾改过 `.env` 里的密码但数据目录是旧的，库里的 root 密码可能仍是旧值，需要在容器内重置（见下）。

4. **确认没有连错实例**  
   本机若还装了其他 MySQL，或 `MYSQL_PORT` 不是 3306，Navicat 端口填错会连到**另一套**库，也会出现 1045。请以 `.env` 的 `MYSQL_PORT` 为准。

**命令行自测（把端口、密码换成你的 `.env`）：**

```bash
# 方式 A：从宿主机测（与 Navicat 一致）
mysql -h 127.0.0.1 -P 3309 -u rpa-bth -prpa-bth_secret -e "SELECT 1"

# 方式 B：进容器用 root（验证 root 密码是否为当前 MYSQL_ROOT_PASSWORD）
docker compose exec mysql mysql -uroot -proot_secret -e "SELECT user, host FROM mysql.user WHERE user IN ('root','rpa-bth');"
```

若方式 B 成功但方式 A 失败，重点检查端口、防火墙或是否连到别的进程。

**仅在需要时：把 root 密码改成与当前 `.env` 一致（在容器内已能 root 登录的前提下执行）：**

```sql
ALTER USER 'root'@'%' IDENTIFIED BY '与.env中MYSQL_ROOT_PASSWORD一致';
ALTER USER 'root'@'localhost' IDENTIFIED BY '同上';
FLUSH PRIVILEGES;
```

若不存在 `root@'%'`，可先执行：

```sql
CREATE USER 'root'@'%' IDENTIFIED BY '与.env中MYSQL_ROOT_PASSWORD一致';
GRANT ALL PRIVILEGES ON *.* TO 'root'@'%' WITH GRANT OPTION;
FLUSH PRIVILEGES;
```

### 5.5 数据目录是旧的：Navicat 报 `1045` 且用户为 `rpa-bth@172.18.0.1`

常见原因：`./docker/mysql/data` 是**以前别的项目**留下的（例如只有 `yii_link` 等用户），**从未按当前 `.env` 创建过 `DB_USER`**。此时改 Navicat 密码也无效，因为库里根本没有这条账号（或库名也不存在）。

**先确认（在宿主机执行，把 root 密码换成你的 `MYSQL_ROOT_PASSWORD`）：**

```bash
docker compose exec mysql mysql -uroot -p -e "SELECT user, host FROM mysql.user; SHOW DATABASES;"
```

若结果里**没有**你的 `DB_USER`，或没有 `DB_NAME` 对应的数据库，在 `mysql -uroot -p` 里执行（三处替换为你的 `.env`：`DB_NAME`、`DB_USER`、`DB_PASS`）：

```sql
CREATE DATABASE IF NOT EXISTS `rpa-report` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
CREATE USER IF NOT EXISTS 'rpa-bth'@'%' IDENTIFIED BY '与.env中DB_PASS完全一致';
GRANT ALL PRIVILEGES ON `rpa-report`.* TO 'rpa-bth'@'%';
FLUSH PRIVILEGES;
```

**注意：** `GRANT` 里反引号包起来的是**数据库名**（`DB_NAME`），不是用户名。文档旧示例写成 `` `rpa-bth`.* `` 容易误导；若你的库名是 `rpa-report`，必须写成 `` `rpa-report`.* ``。

Navicat 里：**用户名 = `DB_USER`，密码 = 与上面 `IDENTIFIED BY` 相同**，端口 = `MYSQL_PORT`。

若用户已存在但密码改过：

```sql
ALTER USER 'rpa-bth'@'%' IDENTIFIED BY '新的DB_PASS';
FLUSH PRIVILEGES;
```

### 5.6 MySQL 容器反复重启：日志里 `initialize ... data directory has files in it`

**典型日志：**

- `Initializing database files`
- `[ERROR] --initialize specified but the data directory has files in it. Aborting.`
- `[ERROR] The designated data directory /var/lib/mysql/ is unusable`

**原因简述：** 官方镜像在「还没有完整数据目录」时会执行 `mysqld --initialize`，但 **`/var/lib/mysql` 里不能有任何文件**（哪怕只有一个 `.gitkeep`）。若目录里已有残留文件、或只放了占位文件，初始化会失败，容器退出后又被拉起，形成**重启循环**。

本项目已将 `docker/mysql/data/` 加入 `.gitignore`，且**不应**在该目录里放 `.gitkeep` 等占位文件（需要空目录时由 Docker 在首次挂载时创建）。

**处理步骤（需要全新库时）：**

```bash
cd /path/to/rpa-task
docker compose stop mysql
# 如需保留旧数据请先备份整个 docker/mysql/data
rm -rf docker/mysql/data/*
docker compose up -d mysql
```

确保 **`docker/mysql/data` 在首次初始化前为空**（或已是完整、未损坏的一套 MySQL 数据文件）。

**另一条常见警告：** `World-writable config file '/etc/mysql/conf.d/my.cnf' is ignored.`  
MySQL 会忽略「全局可写」的配置文件。在宿主机执行：

```bash
chmod 644 docker/mysql/conf.d/my.cnf
```

然后 `docker compose restart mysql`。

---

## 6. 安全建议（生产或敏感数据必读）

- **不要使用弱密码**；定期修改 `MYSQL_ROOT_PASSWORD`、`DB_PASS`。  
- **不要将 root 提供给局域网**，仅使用受限的 `DB_USER` 与单库权限。  
- 仅在可信内网开放端口；公网暴露 MySQL 风险极高，需 VPN、SSH 隧道或云安全组严格限制来源 IP。  
- 需要加密传输时，再为 MySQL 配置 SSL/TLS（本项目默认未强制 SSL）。

---

## 7. 相关文件索引

- `docker-compose.yml` — MySQL 服务、`ports` 映射  
- `docker/mysql/conf.d/my.cnf` — `bind-address` 等  
- `docker/mysql/init/01-grant-remote.sh` — 首次初始化时的远程授权  
- `.env.example` — `MYSQL_PORT`、`DB_*`、`MYSQL_ROOT_PASSWORD` 说明  

文档版本：与仓库 `docker-compose.yml` 当前配置一致时可沿用上述步骤。
