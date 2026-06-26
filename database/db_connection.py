"""
MySQL 连接配置与连接池（pymysql + DBUtils）。

配置来源：config/db_config.json（复制 db_config.example.json 后填写）。

用法：
    from database.db_connection import DatabaseConfig, get_db_manager

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    try:
        ...
        conn.commit()
    finally:
        conn.close()

测试连接：
    python database/db_connection.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pymysql
from dbutils.pooled_db import PooledDB

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_CONFIG_PATH = _PROJECT_ROOT / "config" / "db_config.json"

_manager: DbManager | None = None


class DatabaseConfig:
    """从 config/db_config.json 读取 MySQL 连接参数。"""

    def __init__(self, config_path: Path | str | None = None) -> None:
        path = Path(config_path) if config_path else _DEFAULT_CONFIG_PATH
        if not path.is_file():
            raise FileNotFoundError(
                f"数据库配置文件不存在：{path}\n"
                "请复制 config/db_config.example.json 为 config/db_config.json 并填写连接信息。"
            )
        data = json.loads(path.read_text(encoding="utf-8"))
        self.host: str = str(data["host"])
        self.port: int = int(data["port"])
        self.user: str = str(data["user"])
        self.password: str = str(data["password"])
        self.database: str = str(data["database"])
        self.charset: str = str(data.get("charset", "utf8mb4"))


class DbManager:
    def __init__(self, cfg: DatabaseConfig) -> None:
        self._cfg = cfg
        self._pool = PooledDB(
            creator=pymysql,
            maxconnections=10,
            mincached=1,
            maxcached=5,
            blocking=True,
            host=cfg.host,
            port=cfg.port,
            user=cfg.user,
            password=cfg.password,
            database=cfg.database,
            charset=cfg.charset,
            use_unicode=True,
            autocommit=False,
        )

    def get_connection(self):
        return self._pool.connection()


def get_db_manager(cfg: DatabaseConfig | None = None) -> DbManager:
    global _manager
    if _manager is None:
        _manager = DbManager(cfg or DatabaseConfig())
    return _manager


def _test_connection() -> int:
    cfg = DatabaseConfig()
    print(f"连接 {cfg.host}:{cfg.port}/{cfg.database} ...")
    conn = pymysql.connect(
        host=cfg.host,
        port=cfg.port,
        user=cfg.user,
        password=cfg.password,
        database=cfg.database,
        charset=cfg.charset,
        use_unicode=True,
    )
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1")
            cur.fetchone()
        print("[OK] 数据库连接成功")
        return 0
    finally:
        conn.close()


if __name__ == "__main__":
    try:
        raise SystemExit(_test_connection())
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        raise SystemExit(1)
