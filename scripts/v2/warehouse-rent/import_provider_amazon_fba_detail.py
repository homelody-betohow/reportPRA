from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from logger import get_logger, setup_stdout_utf8

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None  # type: ignore[assignment]

# Allow running as a standalone script: import `python/v2/db.py`
_THIS_DIR = Path(__file__).resolve().parent
_V2_DIR = _THIS_DIR.parent
sys.path.insert(0, str(_V2_DIR))
from db import connect, load_db_config  # noqa: E402


_logger = get_logger("AMZ-FBA")


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


def _json_dumps(v: Any) -> str:
    return json.dumps(v, ensure_ascii=False, separators=(",", ":"))


def _sha256_hex(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _read_excel_first_sheet(xlsx: Path) -> tuple[str, pd.DataFrame]:
    xl = pd.ExcelFile(xlsx)
    sheet = xl.sheet_names[0]
    df = pd.read_excel(xlsx, sheet_name=sheet, dtype=object)
    df = df.dropna(how="all")
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    return sheet, df


def _derive_charge_date_from_filename(xlsx: Path) -> date:
    """
    例：FBA仓租明细3.1-3.31.xlsx -> 取当年 03-31 作为 charge_date
    注意：文件名不含年份时，默认使用当前年份。
    """
    year = datetime.now().year
    m = re.search(r"(\d{1,2})\.(\d{1,2})\s*-\s*(\d{1,2})\.(\d{1,2})", xlsx.stem)
    if not m:
        return date(year, 1, 1)
    m1, d1, m2, d2 = (int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4)))
    # 取区间结束日
    return date(year, m2, d2)


def _extract_warehouse_sku(s: Any) -> str | None:
    """
    复用 v1 的清洗规则：
    - amzn.gr. 特殊格式
    - 默认去掉 # 及 BCFBAFL/FBFBAFL 后缀
    """
    if _is_nan(s):
        return None
    s2 = str(s)
    if "amzn.gr." in s2:
        return s2.split("amzn.gr.")[-1].split("-")[0].split("_")[0].strip() or None
    return s2.split("#")[0].split("BCFBAFL")[0].split("FBFBAFL")[0].strip() or None


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """
    在列名里找最可能匹配的列。
    优先 exact match，其次包含匹配。
    """
    cols = list(df.columns)
    for c in candidates:
        if c in cols:
            return c
    for c in candidates:
        for col in cols:
            if c and c in str(col):
                return str(col)
    return None


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


def import_amazon_fba_file(cur, xlsx: Path) -> int:
    """
    导入 Amazon FBA 仓租（月度/账期级别）。

    数据来源：FBA仓租明细*.xlsx（实为 SellerSku 利润报表结构）
    口径：FBA仓租费 = abs(仓储费用（已分摊） + 长期仓储费（已分摊）)
    """
    provider = "AMZ_FBA"
    setup_stdout_utf8()
    sheet, df = _read_excel_first_sheet(xlsx)
    _logger.warn(f"读取 Excel：{xlsx} sheet={sheet}")
    _logger.info(f"读取完成：行数={len(df)} 列数={len(df.columns)}")

    # 同一文件重导：先删旧明细，避免 line_hash 规则变更后残留或重复去重导致合计偏差
    cur.execute(
        """
        DELETE FROM warehouse_rent_detail
        WHERE provider='AMZ_FBA'
          AND JSON_UNQUOTE(JSON_EXTRACT(raw_row_json, '$.source_file')) = %s
        """,
        (xlsx.name,),
    )
    if cur.rowcount:
        _logger.warn(f"已删除旧数据：source_file={xlsx.name} rows={int(cur.rowcount)}")

    col_seller_sku = _find_col(df, ["sellerSku", "SellerSku", "seller_sku"])
    col_asin = _find_col(df, ["ASIN", "asin"])
    col_wh_sku = _find_col(df, ["仓库sku", "仓库SKU", "warehouse_sku"])
    col_site = _find_col(df, ["站点", "site"])
    col_store = _find_col(df, ["店铺", "店铺/账号", "平台账号", "account", "store"])
    col_store_fee = _find_col(df, ["仓储费用（已分摊）", "仓储费用", "仓储费"])
    col_lts_fee = _find_col(df, ["长期仓储费（已分摊）", "长期仓储费", "长期仓储"])
    col_currency = _find_col(df, ["币种", "Currency", "currency"])

    if not col_store_fee or not col_lts_fee:
        raise RuntimeError(
            f"找不到 FBA 仓租费用列：需要包含 '仓储费用' 与 '长期仓储费' 的列。当前列数={len(df.columns)}"
        )

    charge_date = _derive_charge_date_from_filename(xlsx)
    doc_no = xlsx.name[:-5]  # 去掉 .xlsx 后缀

    rows = []
    total_amount = 0.0
    currency = "EUR"
    skipped_empty_seller_sku = 0
    skipped_empty_asin = 0
    skipped_zero_amount = 0
    for idx, r in df.iterrows():
        source_row_no = int(idx) + 2  # 含表头的近似 Excel 行号
        seller_sku = _to_str(r.get(col_seller_sku)) if col_seller_sku else None
        if not seller_sku:
            skipped_empty_seller_sku += 1
            continue

        asin = _to_str(r.get(col_asin)) if col_asin else None
        if not asin:
            skipped_empty_asin += 1
            continue

        wh_sku_raw = r.get(col_wh_sku) if col_wh_sku else None
        if _is_nan(wh_sku_raw) and seller_sku:
            wh_sku_raw = seller_sku

        sku = _extract_warehouse_sku(wh_sku_raw)
        site = _to_str(r.get(col_site)) if col_site else None
        store = _to_str(r.get(col_store)) if col_store else None

        store_fee = _to_float(r.get(col_store_fee))
        lts_fee = _to_float(r.get(col_lts_fee))
        if store_fee is None and lts_fee is None:
            continue
        # 计算 FBA 仓租费 = 仓储费用（已分摊） + 长期仓储费（已分摊）
        amount = abs(float(store_fee or 0.0) + float(lts_fee or 0.0))
        if amount == 0:
            skipped_zero_amount += 1
            continue

        currency = (_to_str(r.get(col_currency)) if col_currency else None) or "EUR"

        # Ensure warehouse_code/name are not empty
        # Rule: warehouse_code = site + store; warehouse_name same (easier for grouping).
        # Example: "DE_BAITARANE TRADING LTD_DE"
        if site and store:
            warehouse_code = f"{site}-{store}"
            warehouse_name = warehouse_code
        elif site:
            warehouse_code = f"{site}_UNKNOWN_STORE"
            warehouse_name = warehouse_code
        elif store:
            warehouse_code = f"UNKNOWN_SITE_{store}"
            warehouse_name = warehouse_code
        else:
            warehouse_code = "UNKNOWN_SITE_UNKNOWN_STORE"
            warehouse_name = warehouse_code

        line_sig = {
            "provider": provider,
            "warehouse_code": warehouse_code,
            "charge_date": charge_date.isoformat(),
            "seller_sku": seller_sku,
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
                asin,  # barcode
                seller_sku,  # product_name
                None,  # qty
                None,  # volume_m3
                None,  # weight_kg
                None,  # aging_days
                None,  # rent_free_days
                None,  # toll_days
                None,  # receiving_no
                None,  # putaway_at
                None,  # aging_bucket
                None,  # service_category
                None,  # service_product
                None,  # fee_type
                "FBA仓租费",  # fee_name
                float(amount),
                None,  # billed_amount
                None,  # discount_amount
                _json_dumps(
                    {
                        "source_file": xlsx.name,
                        "source_sheet": sheet,
                        "source_row_no": source_row_no,
                        "charge_date_from_filename": charge_date.isoformat(),
                        "sellerSku": seller_sku,
                        "site": site,
                        "store": store,
                        "store_fee": store_fee,
                        "lts_fee": lts_fee,
                        **raw_row,
                    }
                ),
            )
        )
        total_amount += float(amount)

    _logger.info(
        f"数据统计：总行数={len(df)}；sellerSku为空跳过={skipped_empty_seller_sku}；ASIN为空跳过={skipped_empty_asin}；金额为0跳过={skipped_zero_amount}；有效明细行={len(rows)}"
    )
    _logger.info(f"本次文件仓租合计（用于核对）：{round(total_amount, 6)} {currency}")
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


def main() -> None:
    setup_stdout_utf8()
    if load_dotenv:
        load_dotenv()

    parser = argparse.ArgumentParser(description="导入 Amazon FBA 仓租明细（FBA仓租明细*.xlsx）到 MySQL")
    parser.add_argument(
        "--xlsx",
        default=r"python/excel/daily/FBA仓租明细3.1-3.31.xlsx",
        help="Excel 路径",
    )
    args = parser.parse_args()

    xlsx = Path(args.xlsx)
    if not xlsx.exists():
        raise FileNotFoundError(f"找不到 Excel：{xlsx}")

    cfg = load_db_config()
    conn = connect(cfg)
    try:
        cur = conn.cursor()
        _logger.info("开始导入 AMZ_FBA 明细（独立运行）")
        n = import_amazon_fba_file(cur, xlsx)
        conn.commit()
        _logger.info(f"导入完成（已提交）：rows={n}")
    except Exception:
        _logger.error("发生异常，准备回滚")
        conn.rollback()
        raise
    finally:
        _logger.info("关闭数据库连接")
        conn.close()


if __name__ == "__main__":
    main()

