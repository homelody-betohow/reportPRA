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


_logger = get_logger("HY")


def _derive_warehouse_name_from_filename(xlsx: Path) -> str | None:
    """
    例：鸿羽1仓-仓租明细5.1-5.5.xlsx -> 鸿羽1仓
    """
    stem = xlsx.stem.strip()
    for marker in ("-仓租明细", "仓租明细"):
        if marker in stem:
            name = stem.split(marker, 1)[0].strip("- _")
            return name or None
    return None


def _derive_warehouse_code(provider: str, warehouse_name: str) -> str:
    """
    在源文件缺少 warehouse_code 时，生成一个稳定的 code（用于保证字段非空）。
    规则：PROVIDER + '_' + 仅保留字母数字，其他转为下划线，连续下划线压缩。
    """
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


def _to_int(v: Any) -> int | None:
    if _is_nan(v):
        return None
    try:
        return int(float(v))
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


def _to_dt(v: Any) -> datetime | None:
    try:
        ts = pd.to_datetime(v, errors="coerce")
        if pd.isna(ts):
            return None
        return datetime.fromisoformat(ts.to_pydatetime().isoformat(sep=" "))
    except Exception:
        return None


def _json_dumps(v: Any) -> str:
    return json.dumps(v, ensure_ascii=False, separators=(",", ":"))


def _sha256_hex(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _read_excel(xlsx: Path, sheet_name: str) -> pd.DataFrame:
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


def import_hy_file(cur, xlsx: Path) -> int:
    """
    导入 鸿羽（HY）仓租明细：只使用 sheet `bizWarehouseRentByMonthDetail`。
    目标表：warehouse_rent_detail
    """
    provider = "HY"
    sheet = "bizWarehouseRentByMonthDetail"
    setup_stdout_utf8()
    file_warehouse_name = _derive_warehouse_name_from_filename(xlsx)
    if file_warehouse_name:
        _logger.info(f"从文件名解析 warehouse_name：{file_warehouse_name}")

    # Re-import safety: remove existing rows for this source_file.
    # This also fixes historical cases where an older line_hash algorithm
    # deduplicated legitimate duplicate lines and caused sum mismatches.
    cur.execute(
        """
        DELETE FROM warehouse_rent_detail
        WHERE provider='HY'
          AND JSON_UNQUOTE(JSON_EXTRACT(raw_row_json, '$.source_file')) = %s
        """,
        (xlsx.name,),
    )
    if cur.rowcount:
        _logger.warn(f"已删除旧数据：source_file={xlsx.name} rows={int(cur.rowcount)}")

    _logger.warn(f"读取 Excel：{xlsx} sheet={sheet}")
    df = _read_excel(xlsx, sheet)
    _logger.info(f"读取完成：行数={len(df)} 列数={len(df.columns)}")

    col = {
        "doc_no": "编码（Code）",
        "warehouse_code": "仓库代码(Warehouse Code)",
        "sku": "产品代码（SKU）",
        "barcode": "自定义编码（Barcode）",
        "product_name": "产品名称(Product Name)",
        "qty": "数量(Quantity)",
        "charge_date": "计费时间(Charge Date)",
        "aging_days": "库龄(Library of age)",
        "rent_free_days": "免租天数(Rent Free Days)",
        "toll_days": "收费天数(Toll Days)",
        "volume_m3": "体积/m³(Volume)",
        "weight_kg": "重量/kg(Weight)",
        "amount": "产品金额（Product amount）",
        "currency": "币种（Currency）",
        "receiving_no": "入库单(Receiving)",
        "putaway_at": "上架时间(Putaway Date)",
    }

    rows = []
    for idx, r in df.iterrows():
        source_row_no = int(idx) + 2  # 粗略 Excel 行号（含表头）
        doc_no = _to_str(r.get(col["doc_no"]))
        warehouse_code = _to_str(r.get(col["warehouse_code"]))
        warehouse_name = file_warehouse_name
        if not warehouse_code and warehouse_name:
            warehouse_code = _derive_warehouse_code(provider, warehouse_name)
        sku = _to_str(r.get(col["sku"]))
        charge_date = _to_date(r.get(col["charge_date"]))
        amount = _to_float(r.get(col["amount"]))
        currency = _to_str(r.get(col["currency"])) or "EUR"

        if not charge_date or amount is None:
            continue

        # line_hash: include source_row_no to avoid dropping legitimate duplicate lines
        # (Excel may repeat identical values across multiple rows that should be counted).
        line_sig = {
            "provider": provider,
            "doc_no": doc_no,
            "warehouse_code": warehouse_code,
            "charge_date": charge_date.isoformat(),
            "sku": sku,
            "receiving_no": _to_str(r.get(col["receiving_no"])),
            "amount": round(float(amount), 6),
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
                _to_str(r.get(col["barcode"])),
                _to_str(r.get(col["product_name"])),
                _to_float(r.get(col["qty"])),
                _to_float(r.get(col["volume_m3"])),
                _to_float(r.get(col["weight_kg"])),
                _to_int(r.get(col["aging_days"])),
                _to_int(r.get(col["rent_free_days"])),
                _to_int(r.get(col["toll_days"])),
                _to_str(r.get(col["receiving_no"])),
                _to_dt(r.get(col["putaway_at"])),
                None,  # aging_bucket
                None,  # service_category
                None,  # service_product
                None,  # fee_type
                None,  # fee_name
                float(amount),
                None,  # billed_amount
                None,  # discount_amount
                _json_dumps({"source_file": xlsx.name, "source_sheet": sheet, "source_row_no": source_row_no, **raw_row}),
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

