from __future__ import annotations

import argparse
import math
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from dotenv import load_dotenv

from .db import connect, load_db_config


def _is_nan(v: Any) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v))


def _to_int(v: Any) -> int | None:
    if _is_nan(v):
        return None
    try:
        return int(float(v))
    except Exception:
        return None


def _to_float(v: Any) -> float | None:
    if _is_nan(v):
        return None
    try:
        return float(v)
    except Exception:
        return None


def _to_str(v: Any) -> str | None:
    if _is_nan(v):
        return None
    s = str(v).strip()
    return s if s else None


def _read_sheet(xlsx_path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=sheet, dtype=object)
    df = df.dropna(how="all")
    # 清理列名
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    return df


def _promote_first_row_as_header(df: pd.DataFrame) -> pd.DataFrame:
    """
    这个 Excel 的多个 sheet 是“第一行存的是字段名”，而不是 pandas 的 columns。
    例如 sheet: 欧洲平台定价表 / TEMU / MANO-UK / RDC / Conforama
    """
    if df.empty:
        return df
    header = [("" if _is_nan(v) else str(v)).strip() for v in df.iloc[0].tolist()]
    df2 = df.iloc[1:].copy()
    df2.columns = header
    df2 = df2.dropna(how="all")
    return df2


def _truncate_tables(cur) -> None:
    tables = [
        "excel_eu_pricing_exchange_rate",
        "excel_eu_pricing_base",
        "excel_mano_tail_fee",
        "excel_eu_platform_pricing",
        "excel_temu_pricing",
        "excel_platform_sku_pricing",
    ]
    for t in tables:
        cur.execute(f"TRUNCATE TABLE {t}")


def _apply_schema_if_needed(cur) -> None:
    """
    如果目标表不存在，则执行 docs/database/001_eu_pricing_tables.sql 创建表结构。
    这样容器里一键跑脚本即可，无需手动先建表。
    """
    cur.execute("SHOW TABLES LIKE 'excel_eu_pricing_exchange_rate'")
    if cur.fetchone():
        return

    schema_path = Path("docs/database/001_eu_pricing_tables.sql")
    sql_text = schema_path.read_text(encoding="utf-8")

    # 简单切分 SQL（本项目的 schema 文件是 CREATE TABLE + KEY，足够安全）
    buf: list[str] = []
    statements: list[str] = []
    for raw_line in sql_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("--"):
            continue
        buf.append(raw_line)
        if ";" in raw_line:
            joined = "\n".join(buf)
            parts = joined.split(";")
            for p in parts[:-1]:
                s = p.strip()
                if s:
                    statements.append(s)
            buf = [parts[-1]] if parts[-1].strip() else []
    tail = "\n".join(buf).strip()
    if tail:
        statements.append(tail)

    for stmt in statements:
        cur.execute(stmt)


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


def import_exchange_rate(cur, xlsx_path: Path, import_batch_id: str) -> int:
    df = _read_sheet(xlsx_path, "汇率表")
    rows = []
    for idx, r in df.iterrows():
        rows.append(
            (
                import_batch_id,
                int(idx) + 2,  # Excel 行号（粗略，含表头）
                _to_str(r.get("平台国家")),
                _to_float(r.get("平台对应人民币汇率")),
                _to_str(r.get("币种符号")),
                _to_str(r.get("币种")),
                _to_float(r.get("对人民币汇率")),
                _to_str(r.get("发货仓库")),
            )
        )

    sql = """
        INSERT INTO excel_eu_pricing_exchange_rate
        (import_batch_id, source_row_no, platform_country, platform_rmb_rate, currency_symbol, currency_name, rmb_rate, shipping_warehouse)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
    """
    return _insert_many(cur, sql, rows)


def import_mano_tail_fee(cur, xlsx_path: Path, import_batch_id: str) -> int:
    df = _read_sheet(xlsx_path, "MANO尾程 仓租")
    # 第一行是“≤(g) / 法国/西班牙/意大利...”的表头行
    # 我们按列分组：每 4 列一个尺寸段：(max_weight_g, FR, ES, IT)
    # size_group 直接用列名（如 ≤50 x 40 x 30 cm）
    cols = list(df.columns)
    rows = []
    # 从第 1 行开始（跳过 header 行）
    for ridx in range(1, len(df)):
        row = df.iloc[ridx]
        # 每组 4 列
        for base in (0, 4, 8, 12):
            if base >= len(cols):
                continue
            size_group = _to_str(cols[base])
            max_weight_g = _to_int(row.iloc[base])
            if not size_group or max_weight_g is None:
                continue
            fr = _to_float(row.iloc[base + 1]) if base + 1 < len(cols) else None
            es = _to_float(row.iloc[base + 2]) if base + 2 < len(cols) else None
            it = _to_float(row.iloc[base + 3]) if base + 3 < len(cols) else None
            if fr is not None:
                rows.append((import_batch_id, ridx + 1, size_group, max_weight_g, "FR", fr))
            if es is not None:
                rows.append((import_batch_id, ridx + 1, size_group, max_weight_g, "ES", es))
            if it is not None:
                rows.append((import_batch_id, ridx + 1, size_group, max_weight_g, "IT", it))

    sql = """
        INSERT INTO excel_mano_tail_fee
        (import_batch_id, source_row_no, size_group, max_weight_g, country_code, fee_eur)
        VALUES (%s,%s,%s,%s,%s,%s)
    """
    return _insert_many(cur, sql, rows)


def import_eu_platform_pricing(cur, xlsx_path: Path, import_batch_id: str) -> int:
    df_raw = _read_sheet(xlsx_path, "欧洲平台定价表")
    df = _promote_first_row_as_header(df_raw)

    rows = []
    for idx, r in df.iterrows():
        rows.append(
            (
                import_batch_id,
                int(idx) + 2,
                _to_int(r.get("序号")),
                _to_str(r.get("商品ID")),
                _to_str(r.get("SKU")),
                _to_str(r.get("名称")),
                _to_str(r.get("头程方式")),
                _to_str(r.get("发货仓库")),
                _to_str(r.get("是否调拨")),
                _to_str(r.get("是否含税")),
                _to_float(r.get("RMA占比")),
                _to_float(r.get("广告占比")),
                _to_float(r.get("管理成本")),
                _to_int(r.get("秒杀天数")),
                _to_int(r.get("测评量")),
                _to_float(r.get("正常销价")),
                _to_int(r.get("正常量")),
                _to_float(r.get("促销价格")),
                _to_int(r.get("促销量")),
                _to_float(r.get("秒杀价价格")),
                _to_int(r.get("秒杀数量")),
                _to_int(r.get("总销量")),
                _to_float(r.get("总毛利率")),
                _to_float(r.get("平台费")),
                _to_float(r.get("销售税")),
                _to_float(r.get("提现费")),
                _to_float(r.get("测评花费")),
                _to_float(r.get("秒杀花费")),
                _to_float(r.get("促销费")),
                _to_float(r.get("采购价")),
                _to_float(r.get("头程关税")),
                _to_float(r.get("尾程派送费")),
                _to_float(r.get("销售仓租费")),
                _to_float(r.get("调拨费")),
                _to_float(r.get("平均平台费")),
                _to_float(r.get("平均销售税")),
                _to_float(r.get("平均提现费")),
                _to_float(r.get("平均采购价")),
                _to_float(r.get("平均头程关税")),
                _to_float(r.get("平均尾程派送费")),
                _to_float(r.get("平均调拨费")),
            )
        )

    sql = """
        INSERT INTO excel_eu_platform_pricing (
          import_batch_id, source_row_no,
          seq_no, product_id, sku, name,
          first_leg_method, shipping_warehouse, is_transfer, is_tax_included,
          rma_ratio, ad_ratio, mgmt_cost, seckill_days, review_qty,
          normal_price, normal_qty, promo_price, promo_qty, seckill_price, seckill_qty,
          total_qty, total_gross_margin_rate,
          platform_fee, sales_tax, withdrawal_fee, review_cost, seckill_cost, promo_cost,
          purchase_price, first_leg_tariff, last_mile_fee, warehouse_rent_fee, transfer_fee,
          avg_platform_fee, avg_sales_tax, avg_withdrawal_fee, avg_purchase_price, avg_first_leg_tariff,
          avg_last_mile_fee, avg_transfer_fee
        ) VALUES (
          %s,%s,
          %s,%s,%s,%s,
          %s,%s,%s,%s,
          %s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,%s,
          %s,%s,
          %s,%s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,
          %s,%s
        )
    """
    return _insert_many(cur, sql, rows, chunk=500)


def import_temu(cur, xlsx_path: Path, import_batch_id: str) -> int:
    df_raw = _read_sheet(xlsx_path, "TEMU")
    df = _promote_first_row_as_header(df_raw)

    rows = []
    for idx, r in df.iterrows():
        rows.append(
            (
                import_batch_id,
                int(idx) + 2,
                _to_int(r.get("序号")),
                _to_str(r.get("商品ID")),
                _to_str(r.get("1个装SKU")),
                _to_str(r.get("2个装SKU")),
                _to_int(r.get("数量")),
                _to_str(r.get("发货仓库")),
                _to_str(r.get("头程方式")),
                _to_float(r.get("RMA占比")),
                _to_float(r.get("正常销价（欧元）")),
                _to_float(r.get("正常售价（人民币）")),
                _to_float(r.get("总毛利率")),
                _to_float(r.get("采购价")),
                _to_float(r.get("头程关税")),
                _to_float(r.get("尾程派送费")),
            )
        )

    sql = """
        INSERT INTO excel_temu_pricing (
          import_batch_id, source_row_no,
          seq_no, product_id, sku_1pack, sku_2pack, qty,
          shipping_warehouse, first_leg_method, rma_ratio,
          normal_price_eur, normal_price_cny, total_gross_margin_rate,
          purchase_price, first_leg_tariff, last_mile_fee
        ) VALUES (
          %s,%s,
          %s,%s,%s,%s,%s,
          %s,%s,%s,
          %s,%s,%s,
          %s,%s,%s
        )
    """
    return _insert_many(cur, sql, rows, chunk=1000)


def import_platform_sku_pricing(cur, xlsx_path: Path, import_batch_id: str, sheet: str, platform_site: str) -> int:
    df_raw = _read_sheet(xlsx_path, sheet)
    df = _promote_first_row_as_header(df_raw)

    rows = []
    for idx, r in df.iterrows():
        rows.append(
            (
                import_batch_id,
                int(idx) + 2,
                platform_site,
                _to_int(r.get("序号")),
                _to_str(r.get("SKU")),
                _to_str(r.get("名称")),
                _to_str(r.get("规格")),
                _to_str(r.get("品类")),
                _to_str(r.get("供应商")),
                _to_str(r.get("是否泛欧")),
                _to_str(r.get("头程方式")),
                _to_str(r.get("发货仓库")),
                _to_str(r.get("是否调拨")),
                _to_float(r.get("RMA占比")),
                _to_float(r.get("广告占比")),
                _to_float(r.get("站外占比")),
                _to_float(r.get("测评占比")),
                _to_float(r.get("秒杀花费")),
                _to_float(r.get("正常销价")),
                _to_int(r.get("正常量")),
                _to_int(r.get("测评量")),
                _to_int(r.get("秒杀次数")),
                _to_float(r.get("促销价格")),
                _to_int(r.get("促销量")),
                _to_int(r.get("总销量")),
                _to_float(r.get("总毛利率")),
                _to_float(r.get("平台费")),
                _to_float(r.get("销售税")),
                _to_float(r.get("提现费")),
                _to_float(r.get("测评花费")),
                _to_float(r.get("采购价")),
                _to_str(r.get("运营模式")),
                _to_float(r.get("代运营佣金")),
                _to_float(r.get("头程关税")),
                _to_float(r.get("尾程派送费")),
                _to_float(r.get("销售仓租费")),
                _to_float(r.get("调拨费")),
            )
        )

    sql = """
        INSERT INTO excel_platform_sku_pricing (
          import_batch_id, source_row_no, platform_site,
          seq_no, sku, name, spec, category, supplier,
          is_pan_eu, first_leg_method, shipping_warehouse, is_transfer,
          rma_ratio, ad_ratio, offsite_ratio, review_ratio,
          seckill_cost,
          normal_price, normal_qty, review_qty, seckill_times,
          promo_price, promo_qty, total_qty,
          total_gross_margin_rate, platform_fee, sales_tax, withdrawal_fee, review_cost,
          purchase_price, operation_mode, managed_service_commission,
          first_leg_tariff, last_mile_fee, warehouse_rent_fee, transfer_fee
        ) VALUES (
          %s,%s,%s,
          %s,%s,%s,%s,%s,%s,
          %s,%s,%s,%s,
          %s,%s,%s,%s,
          %s,
          %s,%s,%s,%s,
          %s,%s,%s,
          %s,%s,%s,%s,%s,
          %s,%s,%s,
          %s,%s,%s,%s
        )
    """
    return _insert_many(cur, sql, rows, chunk=1000)


def run_import(xlsx_path: Path, import_batch_id: str, truncate: bool = True, apply_schema: bool = True) -> None:
    cfg = load_db_config()
    conn = connect(cfg)
    try:
        cur = conn.cursor()
        if apply_schema:
            _apply_schema_if_needed(cur)
        if truncate:
            _truncate_tables(cur)

        counts: dict[str, int] = {}
        counts["exchange_rate"] = import_exchange_rate(cur, xlsx_path, import_batch_id)
        counts["mano_tail_fee"] = import_mano_tail_fee(cur, xlsx_path, import_batch_id)
        counts["eu_platform_pricing"] = import_eu_platform_pricing(cur, xlsx_path, import_batch_id)
        counts["temu"] = import_temu(cur, xlsx_path, import_batch_id)
        counts["mano_uk"] = import_platform_sku_pricing(cur, xlsx_path, import_batch_id, "MANO-UK", "MANO-UK")
        counts["rdc_fr"] = import_platform_sku_pricing(cur, xlsx_path, import_batch_id, "RDC", "RDC-FR")
        counts["conforama_fr"] = import_platform_sku_pricing(cur, xlsx_path, import_batch_id, "Conforama", "Conforama-FR")

        # 基础表暂不导入：结构是多行表头 + 大量动态列，后续如果你需要我再补齐规范化导入
        # counts["base"] = import_base(...)

        conn.commit()
        print("导入完成（已提交）:", counts)
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def main() -> None:
    load_dotenv()  # 读取仓库根目录的 .env

    parser = argparse.ArgumentParser(description="方案A：清空表并从欧洲平台定价表.xlsx全量导入MySQL（后续报表直接查库）")
    parser.add_argument(
        "--xlsx",
        default=r"python/excel/base/欧洲平台定价表.xlsx",
        help="Excel 路径",
    )
    parser.add_argument(
        "--batch",
        default=datetime.now().strftime("%Y%m%d_%H%M%S"),
        help="import_batch_id（A方案不依赖它查询，但会写入用于排查）",
    )
    parser.add_argument(
        "--no-truncate",
        action="store_true",
        help="不清空表，直接追加（调试用）",
    )
    parser.add_argument(
        "--no-apply-schema",
        action="store_true",
        help="不自动执行 docs/database/001_eu_pricing_tables.sql（调试用）",
    )
    args = parser.parse_args()

    xlsx_path = Path(args.xlsx)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"找不到 Excel：{xlsx_path}")

    run_import(
        xlsx_path=xlsx_path,
        import_batch_id=args.batch,
        truncate=not args.no_truncate,
        apply_schema=not args.no_apply_schema,
    )


if __name__ == "__main__":
    main()

