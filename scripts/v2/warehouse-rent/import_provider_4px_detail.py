from __future__ import annotations

import hashlib
import json
import math
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from logger import get_logger, setup_stdout_utf8


_logger = get_logger("4PX")


def _derive_warehouse_name_from_filename(xlsx: Path) -> str | None:
    """
    例：4PX法国仓-仓租明细-5.1-5.5.xlsx -> 4PX法国仓
    """
    stem = xlsx.stem.strip()
    for marker in ("-仓租明细", "仓租明细"):
        if marker in stem:
            name = stem.split(marker, 1)[0].strip("- _")
            return name or None
    return None


def _derive_warehouse_code(provider: str, warehouse_name: str) -> str:
    """
    生成稳定的 4PX 仓库代码（用于保证字段非空）。
    优先生成形如：4PX_FR_PARIS2；否则退化为 4PX_{slug}
    """
    # 简单国家识别
    country = "FR" if "法国" in warehouse_name else ("DE" if "德国" in warehouse_name else "")
    # 抽取“巴黎2/巴黎1”之类的数字后缀
    m = re.search(r"巴黎\s*(\d+)", warehouse_name)
    city = f"PARIS{m.group(1)}" if m else ""
    if country and city:
        return f"{provider}_{country}_{city}"
    base = re.sub(r"[^0-9A-Za-z]+", "_", warehouse_name).strip("_")
    base = re.sub(r"_+", "_", base)
    return f"{provider}_{base}" if base else provider


def _is_nan(v: Any) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v))


def _to_str(v: Any) -> str | None:
    if _is_nan(v):
        return None
    s = str(v).strip()
    return s if s else None


def _to_float(v: Any) -> float | None:
    if _is_nan(v):
        return None
    try:
        return float(v)
    except Exception:
        return None


def _to_date(v: Any) -> datetime.date | None:
    try:
        ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def _json_dumps(v: Any) -> str:
    return json.dumps(v, ensure_ascii=False, separators=(",", ":"))


def _sha256_hex(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _read_excel(xlsx: Path, sheet_name: int = 0) -> pd.DataFrame:
    df = pd.read_excel(xlsx, sheet_name=sheet_name, dtype=object)
    df = df.dropna(how="all")
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    return df


def _insert_many(cur, sql: str, rows: Iterable[tuple[Any, ...]], chunk: int = 1000) -> int:
    rows_list = list(rows)
    total = 0
    for i in range(0, len(rows_list), chunk):
        batch = rows_list[i : i + chunk]
        if not batch:
            continue
        cur.executemany(sql, batch)
        total += len(batch)
    return total


def import_4px_file(cur, xlsx: Path) -> int:
    """
    导入 4PX 仓租明细：读取 sheet0（第 0 个 sheet）。
    目标表：warehouse_rent_detail
    """
    provider = "4PX"
    setup_stdout_utf8()
    file_warehouse_name = _derive_warehouse_name_from_filename(xlsx)
    if file_warehouse_name:
        _logger.info(f"从文件名解析 warehouse_name（兜底用）：{file_warehouse_name}")

    # 同一文件重导：先删旧明细，避免 line_hash 规则变更后残留或重复去重导致合计偏差
    cur.execute(
        """
        DELETE FROM warehouse_rent_detail
        WHERE provider='4PX'
          AND JSON_UNQUOTE(JSON_EXTRACT(raw_row_json, '$.source_file')) = %s
        """,
        (xlsx.name,),
    )
    if cur.rowcount:
        _logger.warn(f"已删除旧数据：source_file={xlsx.name} rows={int(cur.rowcount)}")

    _logger.warn(f"读取 Excel：{xlsx} sheet=0")
    df = _read_excel(xlsx, 0)
    _logger.info(f"读取完成：行数={len(df)} 列数={len(df.columns)}")

    rows = []
    for idx, r in df.iterrows():
        source_row_no = int(idx) + 2  # 含表头的近似 Excel 行号
        doc_no = _to_str(r.get("仓租单号"))
        warehouse_name = _to_str(r.get("计费仓库")) or file_warehouse_name
        warehouse_code = _to_str(r.get("仓库代码"))
        if not warehouse_code and warehouse_name:
            warehouse_code = _derive_warehouse_code(provider, warehouse_name)
        sku = _to_str(r.get("SKU"))
        charge_date = _to_date(r.get("仓租日期"))
        currency = _to_str(r.get("币种")) or "EUR"
        receivable = _to_float(r.get("应收金额"))

        if not charge_date or receivable is None:
            continue

        line_sig = {
            "provider": provider,
            "doc_no": doc_no,
            "warehouse_code": warehouse_code,
            "charge_date": charge_date.isoformat(),
            "sku": sku,
            "aging_bucket": _to_str(r.get("库龄段")),
            "fee_name": _to_str(r.get("费用名称")),
            "amount": round(float(receivable), 6),
            "currency": currency,
            "source_row_no": source_row_no,
        }
        line_hash = _sha256_hex(_json_dumps(line_sig))

        raw_row = {k: ("" if _is_nan(v) else v) for k, v in r.to_dict().items()}
        rows.append(
            (
                provider,
                line_hash,
                doc_no,
                charge_date,
                warehouse_code,
                warehouse_name,
                currency,
                sku,
                None,  # barcode
                None,  # product_name
                _to_float(r.get("SKU数量")),
                None,  # volume_m3
                None,  # weight_kg
                None,  # aging_days
                None,  # rent_free_days
                None,  # toll_days
                None,  # receiving_no
                None,  # putaway_at
                _to_str(r.get("库龄段")),
                _to_str(r.get("服务类别")),
                _to_str(r.get("服务产品")),
                _to_str(r.get("计费类型")),
                _to_str(r.get("费用名称")),
                float(receivable),
                _to_float(r.get("计费金额")),
                _to_float(r.get("优惠金额")),
                _json_dumps(
                    {"source_file": xlsx.name, "source_sheet": "sheet0", "source_row_no": source_row_no, **raw_row}
                ),
            )
        )

    _logger.info(f"准备写入 MySQL 明细：rows={len(rows)}（将按 line_hash 去重）")
    sql = """
        INSERT INTO warehouse_rent_detail (
          provider, line_hash, doc_no,
          charge_date, warehouse_code, warehouse_name, currency,
          sku, barcode, product_name,
          qty, volume_m3, weight_kg,
          aging_days, rent_free_days, toll_days, receiving_no, putaway_at,
          aging_bucket, service_category, service_product, fee_type, fee_name,
          amount, billed_amount, discount_amount,
          raw_row_json
        ) VALUES (
          %s,%s,%s,
          %s,%s,%s,%s,
          %s,%s,%s,
          %s,%s,%s,
          %s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,
          %s,%s,%s,
          %s
        )
        ON DUPLICATE KEY UPDATE
          line_hash = line_hash
    """
    inserted = _insert_many(cur, sql, rows, chunk=1000)
    _logger.info(f"MySQL executemany 完成：批次数量={len(rows)}（executemany 行数={inserted}）")
    return inserted

