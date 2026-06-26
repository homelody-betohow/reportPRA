from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path
from typing import Any

LOCK_FILE = Path(__file__).resolve().parent / "run_batch.lock"


def read_lock() -> dict[str, Any] | None:
    """读取 run_batch.lock，失败或格式无效时返回 None。"""
    if not LOCK_FILE.is_file():
        return None
    try:
        data = json.loads(LOCK_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return data if isinstance(data, dict) else None


def read_import_batch_from_lock() -> str | None:
    """从 run_batch.lock 读取 import_batch。"""
    lock = read_lock()
    if not lock:
        return None
    batch = str(lock.get("import_batch", "")).strip()
    return batch or None


def make_import_batch(override: str | None) -> str:
    """生成批次号（YYYYMMDD_HHMMSS），override 非空则直接使用。"""
    if override and override.strip():
        return override.strip()
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def write_lock(import_batch: str) -> None:
    """写入/更新 run_batch.lock。"""
    payload = {
        "run_date": date.today().isoformat(),
        "import_batch": import_batch,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    LOCK_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def resolve_import_batch(override: str | None) -> tuple[str, bool]:
    """
    解析本批 import_batch：
    - override 指定则直接用，不写锁
    - 否则若锁文件存在且 run_date=今天，则复用锁
    - 否则生成新批次并写入锁文件

    返回 (import_batch, lock_written)
    """
    if override and override.strip():
        return override.strip(), False

    today = date.today().isoformat()
    lock = read_lock()
    if lock is not None:
        lock_date = str(lock.get("run_date", "")).strip()
        lock_batch = str(lock.get("import_batch", "")).strip()
        if lock_date == today and lock_batch:
            return lock_batch, False

    batch = make_import_batch(None)
    write_lock(batch)
    return batch, True
