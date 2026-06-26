from __future__ import annotations

"""
从 sales_order_shipped 表中读取指定 import_batch 的数据，
按 order_type、platform 排序后写入 sales_order_sku_profit 表。

功能说明：
1. 读取 run_batch.lock 文件获取 import_batch 批次号
2. 从 sales_order_shipped 表中查询该批次的订单数据
3. 按 order_type、platform 排序
4. order_type 不等于「重发订单」时，仅 order_total_base > 0 的行才写入
5. 转换并 UPSERT 写入 sales_order_sku_profit 表（line_hash 冲突则更新；report_hash = shipped.import_batch）

用法：
  cd d:\\py-project\\report
  python scripts\\archive\\profit_001_order_sku.py
  python scripts\\archive\\profit_001_order_sku.py --batch 20260616_203140
"""

import argparse
import re
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_ARCHIVE_DIR = Path(__file__).resolve().parent
_DATA_IMPORT_DIR = _REPORT_ROOT / "scripts" / "dataImport"

for _p in (_REPORT_ROOT, _ARCHIVE_DIR, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import read_import_batch_from_lock  # noqa: E402  # pyright: ignore[reportMissingImports]
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402
from scripts.dataImport.import_common import upsert_rows  # noqa: E402

SOURCE_TABLE = "sales_order_shipped"
TARGET_TABLE = "sales_order_sku_profit"
RESEND_ORDER_TYPE = "重发订单"
CALC_NODE = "init"
_ZERO_DEC = Decimal("0")

# AMZN.GR 泛欧仓 SKU：不区分大小写（参考 H1_映射_计算.py / v2 run_sku_clean.py）
_AMZN_GR_PREFIX_RE = re.compile(r"^AMZN\.GR", re.IGNORECASE)
_AMZN_GR_INFIX_RE = re.compile(r"amzn\.gr\.", re.IGNORECASE)


def _is_amzn_gr_warehouse_sku(s: str) -> bool:
    """warehouse_sku 是否为 Amazon FBA 泛欧 SKU 前缀（如 AMZN.GR.U02033010_...）。"""
    return bool(s) and bool(_AMZN_GR_PREFIX_RE.match(s))


def extract_sku_value(s: Any) -> str:
    """
    从 SKU 字符串提取内部产品编码（参考 H1_映射_计算.py extract_values）。

    - 含 amzn.gr.：取其后第一段，再按 - / _ 截断
    - 其他：去掉 # 尾缀与 BCFBAFL 后缀
    """
    text = _str_or_empty(s)
    if not text:
        return ""
    text = text.split("#", 1)[0]
    if _AMZN_GR_INFIX_RE.search(text):
        tail = _AMZN_GR_INFIX_RE.split(text)[-1]
        return tail.split("-", 1)[0].split("_", 1)[0]
    return text.split("BCFBAFL", 1)[0]


def normalize_skus_for_amzn_gr(warehouse_sku: str, platform_sku: Any) -> tuple[str, Any]:
    """
    当 warehouse_sku 以 AMZN.GR 开头时，截取 warehouse_sku 与 platform_sku。
    非 AMZN.GR 行原样返回。
    """
    wh = _str_or_empty(warehouse_sku)
    if not _is_amzn_gr_warehouse_sku(wh):
        return wh, platform_sku
    cleaned_wh = extract_sku_value(wh)
    if platform_sku is not None and str(platform_sku).strip():
        cleaned_ps: Any = extract_sku_value(platform_sku)
    else:
        cleaned_ps = platform_sku   
    return cleaned_wh, cleaned_ps

# sales_order_sku_profit 表字段（按表结构顺序）
TARGET_COLUMNS = [
    "line_hash",
    "platform",
    "shop_name_en",
    "platform_site",
    "order_type",
    "ref_no",
    "order_no",
    "product_sku",
    "warehouse_sku",
    "platform_sku",
    "warehouse_name",
    "warehouse_type",
    "shipping_method",
    "pay_currency",
    "base_currency",
    "pay_time",
    "ship_time",
    "shipped_qty",
    "order_total_pay",
    "order_total_base",
    "order_goods_base",
    "platform_shipping_base",
    "payment_fee_base",
    "platform_fee_base",
    "fba_fee_base",
    "platform_subsidy_base",
    "vat_fee_base",
    "tax_base",
    "other_fee_base",
    "purchase_cost_base",
    "purchase_shipping_base",
    "purchase_tax_base",
    "first_leg_shipping_base",
    "first_leg_tax_base",
    "packaging_fee_base",
    "delivery_shipping_base",
    "total_fee_base",
    "total_cost_base",
    "gross_profit_base",
    "gross_margin_rate",
    "refund_qty",
    "refund_amount_base",
    "net_profit_base",
    "net_margin_rate",
    "distribution_lev",
    "report_hash",
    "calc_node",
    "source_note",
]


def _log(level: str, msg: str) -> None:
    """日志输出"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def _dec_or_zero(v: Any) -> Decimal:
    """转换为 Decimal，空值返回 0"""
    if v is None:
        return _ZERO_DEC
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v))
    except Exception:
        return _ZERO_DEC


def _int_or_zero(v: Any) -> int:
    """转换为 int，空值返回 0"""
    if v is None:
        return 0
    try:
        return int(v)
    except Exception:
        return 0


def _str_or_empty(v: Any) -> str:
    """转换为字符串，空值返回空字符串"""
    if v is None:
        return ""
    return str(v).strip()


def should_write_row(row: dict[str, Any]) -> bool:
    """
    判断是否应写入 profit 表。
    重发订单始终写入；其他 order_type 仅当 order_total_base > 0 时写入。
    """
    order_type = _str_or_empty(row.get("order_type"))
    if order_type == RESEND_ORDER_TYPE:
        return True
    return _dec_or_zero(row.get("order_total_base")) > _ZERO_DEC


def fetch_shipped_data(conn, import_batch: str) -> list[dict[str, Any]]:
    """
    从 sales_order_shipped 表中查询指定批次的数据，按 order_type、platform 排序
    
    Args:
        conn: 数据库连接
        import_batch: 导入批次号
    
    Returns:
        订单数据列表（字典格式）
    """
    sql = f"""
        SELECT 
            line_hash,
            platform,
            shop_name_en,
            platform_site,
            order_type,
            ref_no,
            order_no,
            warehouse_sku,
            platform_sku,
            warehouse_name,
            warehouse_type,
            shipping_method,
            pay_currency,
            base_currency,
            pay_time,
            ship_time,
            warehouse_sku_qty,
            order_total_pay,
            order_total_base,
            order_goods_base,
            platform_shipping_base,
            payment_fee_base,
            platform_fee_base,
            fba_fee_base,
            platform_subsidy_base,
            tax_base,
            other_fee_base,
            purchase_cost_base,
            purchase_shipping_base,
            purchase_tax_base,
            first_leg_shipping_base,
            first_leg_tax_base,
            packaging_fee_base,
            delivery_shipping_base,
            total_fee_base,
            total_cost_base,
            gross_profit_base,
            gross_margin_rate,
            distribution_lev,
            import_batch
        FROM {SOURCE_TABLE}
        WHERE import_batch = %s
        ORDER BY order_type, platform
    """
    
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        _log("INFO", f"查询 {SOURCE_TABLE} 表，批次：{import_batch}")
        cur.execute(sql, (import_batch,))
        rows = cur.fetchall()
        _log("INFO", f"查询完成：找到 {len(rows)} 条记录")
        return rows
    finally:
        cur.close()


def transform_row(row: dict[str, Any]) -> tuple[Any, ...]:
    """
    将 sales_order_shipped 的一行数据转换为 sales_order_sku_profit 格式
    
    Args:
        row: shipped 表的一行数据（字典）
    
    Returns:
        profit 表的一行数据（元组）
    """
    # AMZN.GR 泛欧仓 SKU：截取 warehouse_sku / platform_sku 内部产品编码
    warehouse_sku, platform_sku = normalize_skus_for_amzn_gr(
        _str_or_empty(row.get("warehouse_sku")),
        row.get("platform_sku"),
    )

    # product_sku: 默认使用清洗后的 warehouse_sku
    product_sku = warehouse_sku
    
    # shipped_qty: 仓库 SKU 销量
    shipped_qty = _int_or_zero(row.get("warehouse_sku_qty"))
    
    # vat_fee_base: shipped 表无 vat_fee 字段，设为 0
    vat_fee_base = _ZERO_DEC
    
    # 发货毛利与净利润：当前不考虑退款，两者相同
    gross_profit_base = _dec_or_zero(row.get("gross_profit_base"))
    gross_margin_rate = row.get("gross_margin_rate")  # 可能为 None
    net_profit_base = gross_profit_base
    net_margin_rate = gross_margin_rate
    
    # 退款相关：当前恒为 0
    refund_qty = 0
    refund_amount_base = _ZERO_DEC
    
    # 分销等级和计算节点
    distribution_lev = _int_or_zero(row.get("distribution_lev"))
    calc_node = CALC_NODE
    
    # 来源说明
    source_note = ""
    
    # report_hash：与 sales_order_shipped.import_batch 一致
    report_hash = _str_or_empty(row.get("import_batch")) or None

    # line_hash：与 sales_order_shipped 保持一致，作为两表关联键
    return (
        _str_or_empty(row.get("line_hash")),
        _str_or_empty(row.get("platform")),
        row.get("shop_name_en"),
        row.get("platform_site"),
        row.get("order_type"),
        _str_or_empty(row.get("ref_no")),
        _str_or_empty(row.get("order_no")),
        product_sku,
        warehouse_sku,
        platform_sku,
        row.get("warehouse_name"),
        row.get("warehouse_type"),
        row.get("shipping_method"),
        row.get("pay_currency"),
        row.get("base_currency"),
        row.get("pay_time"),
        row.get("ship_time"),
        shipped_qty,
        _dec_or_zero(row.get("order_total_pay")),
        _dec_or_zero(row.get("order_total_base")),
        _dec_or_zero(row.get("order_goods_base")),
        _dec_or_zero(row.get("platform_shipping_base")),
        _dec_or_zero(row.get("payment_fee_base")),
        _dec_or_zero(row.get("platform_fee_base")),
        _dec_or_zero(row.get("fba_fee_base")),
        _dec_or_zero(row.get("platform_subsidy_base")),
        vat_fee_base,
        _dec_or_zero(row.get("tax_base")),
        _dec_or_zero(row.get("other_fee_base")),
        _dec_or_zero(row.get("purchase_cost_base")),
        _dec_or_zero(row.get("purchase_shipping_base")),
        _dec_or_zero(row.get("purchase_tax_base")),
        _dec_or_zero(row.get("first_leg_shipping_base")),
        _dec_or_zero(row.get("first_leg_tax_base")),
        _dec_or_zero(row.get("packaging_fee_base")),
        _dec_or_zero(row.get("delivery_shipping_base")),
        _dec_or_zero(row.get("total_fee_base")),
        _dec_or_zero(row.get("total_cost_base")),
        gross_profit_base,
        gross_margin_rate,
        refund_qty,
        refund_amount_base,
        net_profit_base,
        net_margin_rate,
        distribution_lev,
        report_hash,
        calc_node,
        source_note,
    )


def write_to_profit_table(conn, rows: list[tuple[Any, ...]]) -> int:
    """
    将数据 UPSERT 写入 sales_order_sku_profit 表（line_hash 冲突则更新）

    Args:
        conn: 数据库连接
        rows: 数据行列表

    Returns:
        累计 UPSERT 行数
    """
    if not rows:
        _log("WARN", "没有数据需要写入")
        return 0

    _log("INFO", f"准备写入 {TARGET_TABLE}：共 {len(rows)} 条记录（UPSERT，line_hash 冲突则更新）")
    n_upsert = upsert_rows(
        conn,
        table=TARGET_TABLE,
        columns=TARGET_COLUMNS,
        rows=rows,
        chunk_size=300,
    )
    _log("INFO", f"写入完成：UPSERT {n_upsert} 条")
    return n_upsert


def process_batch(conn, import_batch: str) -> tuple[int, int]:
    """
    处理指定批次的数据：从 shipped 表读取并写入 profit 表

    Args:
        conn: 数据库连接
        import_batch: 导入批次号

    Returns:
        (查询行数, UPSERT 行数)
    """
    # 1. 从 shipped 表查询数据（已按 order_type、platform 排序）
    shipped_rows = fetch_shipped_data(conn, import_batch)
    if not shipped_rows:
        _log("WARN", f"批次 {import_batch} 无数据")
        return 0, 0

    # 2. 过滤：非重发订单须 order_total_base > 0
    eligible_rows = [row for row in shipped_rows if should_write_row(row)]
    n_filtered = len(shipped_rows) - len(eligible_rows)
    if n_filtered:
        _log(
            "INFO",
            f"过滤非「{RESEND_ORDER_TYPE}」且 order_total_base<=0：跳过 {n_filtered} 条",
        )
    if not eligible_rows:
        _log("WARN", f"批次 {import_batch} 过滤后无待写入数据")
        return len(shipped_rows), 0

    # 3. 转换数据格式（line_hash 原样沿用 shipped 表）
    _log("INFO", "开始转换数据格式...")
    profit_rows = [transform_row(row) for row in eligible_rows]
    _log("INFO", f"转换完成：{len(profit_rows)} 条记录")

    # 4. UPSERT 写入 profit 表
    n_upsert = write_to_profit_table(conn, profit_rows)

    return len(shipped_rows), n_upsert


def parse_args() -> argparse.Namespace:
    """解析命令行参数"""
    ap = argparse.ArgumentParser(
        description="从 sales_order_shipped 表读取数据并写入 sales_order_sku_profit 表"
    )
    ap.add_argument(
        "--batch",
        type=str,
        default=None,
        metavar="BATCH",
        help="指定导入批次号（默认从 run_batch.lock 文件读取）",
    )
    return ap.parse_args()


def main() -> int:
    """主函数"""
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    
    args = parse_args()
    
    # 获取批次号
    if args.batch:
        import_batch = args.batch
        _log("INFO", f"使用命令行指定的批次号：{import_batch}")
    else:
        import_batch = read_import_batch_from_lock()
        if import_batch:
            _log("INFO", f"从 run_batch.lock 读取批次号：{import_batch}")
        if not import_batch:
            _log("ERROR", "无法获取批次号，请使用 --batch 参数指定或确保 run_batch.lock 文件存在")
            return 1
    
    _log("INFO", f"任务：{SOURCE_TABLE} -> {TARGET_TABLE}")
    _log("INFO", f"批次号：{import_batch}")
    
    # 连接数据库
    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    
    try:
        # 处理数据
        n_read, n_upsert = process_batch(conn, import_batch)

        # 提交事务
        conn.commit()

        _log(
            "INFO",
            f"任务完成：读取 {n_read} 条，UPSERT {n_upsert} 条",
        )
        _log("INFO", "=" * 80)
        
        return 0
    
    except Exception as e:
        conn.rollback()
        _log("ERROR", f"任务失败：{e}")
        import traceback
        traceback.print_exc()
        return 2
    
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
