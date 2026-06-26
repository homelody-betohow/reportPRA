from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import mysql.connector


def _load_dotenv_once() -> None:
    """若已安装 python-dotenv，则尝试加载首个存在的 .env（与导入脚本共用同一套变量）。"""
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    here = Path(__file__).resolve().parent  # python/v2
    for env_path in (
        here.parent / ".env",  # python/.env
        here / ".env",  # python/v2/.env
        here.parent.parent / ".env",  # 仓库根目录 .env
    ):
        if env_path.is_file():
            load_dotenv(env_path)
            return


@dataclass(frozen=True)
class DbConfig:
    host: str
    port: int
    user: str
    password: str
    database: str


def load_db_config() -> DbConfig:
    """
    读取 .env / 环境变量（与 v2 下各导入脚本共用；连接使用 connect()）。

    约定：
    - 宿主机直连：DB_HOST=127.0.0.1, DB_PORT=${MYSQL_PORT}
    - Docker python 容器内：DB_HOST=mysql, DB_PORT=3306
    """
    _load_dotenv_once()
    # 在 docker-compose 的 python 容器里，应该用服务名 mysql + 容器内端口 3306
    # 在宿主机本地运行时，默认走 127.0.0.1 + MYSQL_PORT（端口映射）
    in_docker = os.path.exists("/.dockerenv")
    default_host = "mysql" if in_docker else "127.0.0.1"
    default_port = "3306" if in_docker else os.getenv("MYSQL_PORT", "3306")

    host = os.getenv("DB_HOST", default_host)
    port = int(os.getenv("DB_PORT", default_port))
    user = os.getenv("DB_USER", "")
    password = os.getenv("DB_PASS", "")
    database = os.getenv("DB_NAME", "")
    if not (user and password and database):
        raise RuntimeError(
            "缺少数据库配置：请在 .env 或环境变量里设置 DB_USER / DB_PASS / DB_NAME（以及可选 DB_HOST / DB_PORT）"
        )
    return DbConfig(host=host, port=port, user=user, password=password, database=database)


def connect(cfg: DbConfig) -> mysql.connector.MySQLConnection:
    return mysql.connector.connect(
        host=cfg.host,
        port=cfg.port,
        user=cfg.user,
        password=cfg.password,
        database=cfg.database,
        use_pure=True,
        autocommit=False,
        charset="utf8mb4",
        use_unicode=True,
    )

