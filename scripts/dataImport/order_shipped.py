from __future__ import annotations

"""
从 ERP 共享目录读取「订单统计-*.xlsx」，写入 MySQL 表 sales_order_shipped。
platform=semitemu 且订单类型=销售订单 的行同步 UPSERT 到 temu_order_item（唯一键 order_no + sku_id）。
temu_order_item.file_name 已有值的行不更新业务字段（保护 RPA/手工导入的订单详情），
但仍会同步 sales_order_shipped.line_hash 到 temu_order_item.line_hash。

路径：config.path_config.ERP_ORDER_STA_PATH + MODE_PATTERN + 日期子目录
  每天 -> .../每天/ERP订单、RMA下载/2026-05-26/订单统计-*.xlsx
  每月 -> .../每月/ERP订单、RMA下载/2026-01/订单统计-*.xlsx

line_hash：对 LINE_HASH_KEYS 子集做 stable_line_hash（键排序 JSON + SHA-256），
与线上去重键 uk_order_line_hash 一致。

币种校验：导入前读取 Excel A3 单元格，解析币种代码须为 EUR（与 A2 脚本一致），
否则中止导入；校验通过后读取 A2 单元格并以彩色输出表格元信息。

import_batch：未指定 --import-batch 时，默认取 run_batch.lock 的 import_batch。
导入时检测并输出 line_hash 重复分组（含 Excel 行号与关键字段）。

item_id：若含中文逗号「，」或英文逗号「,」，按逗号拆分为数组、排序后用英文逗号拼接写回
（在计算 line_hash 之前处理，保证同一组 ID 顺序不同也能对齐）。

分销仓：warehouse_name 含「分销」且 distribution_lev=0 时写入 distribution_lev=1；
此类行不写入 product_sku_mapping。

用法：
  cd d:\\py-project\\report
  python scripts\\dataImport\\order_shipped.py
  python scripts\\dataImport\\order_shipped.py --date 2026-05-26
  python scripts\\dataImport\\order_shipped.py --file "\\\\Betohow\\...\\订单统计-5.1-5.26.xlsx"

未指定 --date / --dir / --file 时，日期子目录默认取 path_config.DATE_PATH
（日报=3 天前 YYYY-MM-DD，月报=YYYY-MM，与 MODE_RUN 一致）。
"""

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

import pandas as pd
import pymysql.cursors

_REPORT_ROOT = Path(__file__).resolve().parents[2]
_DATA_IMPORT_DIR = Path(__file__).resolve().parent
for _p in (_REPORT_ROOT, _DATA_IMPORT_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from batch_lock import resolve_import_batch  # noqa: E402
from config.path_config import DATE_PATH, ERP_ORDER_STA_PATH, MODE_PATTERN  # noqa: E402
from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402
from import_common import (  # noqa: E402
    cell_decimal,
    cell_dt,
    cell_int,
    cell_margin_rate,
    cell_str,
    cell_str_or_empty,
    insert_ignore_rows,
    row_subset_for_line_hash,
    stable_line_hash,
    upsert_rows,
)
from sku_mapping_import import (  # noqa: E402
    SHOP_HASH_KEYS,
    upsert_product_sku_mapping,
)

TABLE = "sales_order_shipped"
TEMU_ORDER_ITEM_TABLE = "temu_order_item"
PLATFORM_SHOP_TABLE = "platform_shop"
PLATFORM_SEMITEMU = "semitemu"
ORDER_TYPE_SALES = "销售订单"
SOURCE_TYPE = "Excel"
_DISTRIBUTION_WH_MARK = "分销"
EXPECTED_CURRENCY = "EUR"  # 订单统计 Excel A3 单元格须为该币种（与 A2 脚本一致）
_ZERO_DEC = Decimal("0")
_TWO_PLACES = Decimal("0.01")
_ITEM_ID_SEP_RE = re.compile(r"[,，]")

# ANSI 终端颜色
_RESET = "\033[0m"
_ANSI = {
    "RED": "\033[91m",
    "GREEN": "\033[92m",
    "YELLOW": "\033[93m",
    "CYAN": "\033[96m",
    "BOLD": "\033[1m",
}


def _enable_windows_ansi() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        for handle_id in (-11, -12):
            handle = kernel32.GetStdHandle(handle_id)
            mode = ctypes.c_uint32()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        pass


def _use_color() -> bool:
    if os.getenv("NO_COLOR"):
        return False
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    return True


def _c(text: str, *styles: str) -> str:
    if not _use_color():
        return text
    codes = "".join(_ANSI.get(s, s) for s in styles)
    return f"{codes}{text}{_RESET}"


def _cell_display(cell: Any) -> str:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    return str(cell).strip()


# Temu 订单项表头
TEMU_ORDER_ITEM_COLUMNS: tuple[str, ...] = (
    "line_hash",
    "order_no",
    "sub_order_no",
    "shop_account",
    "shop_name_en",
    "shop_site_name",
    "country_region",
    "order_status",
    "shipping_method",
    "created_time",
    "confirmed_time",
    "required_ship_deadline",
    "shipped_time",
    "estimated_delivery_time",
    "delivered_time",
    "package_no",
    "tracking_no",
    "logistics_company",
    "sku_id",
    "skc_id",
    "spu_id",
    "warehouse_sku",
    "product_name",
    "variant_name",
    "quantity",
    "warehouse_name",
    "declared_price",
    "order_payment",
    "second_payment",
    "sales_revenue",
    "sales_return",
    "shipping_income",
    "shipping_deduction",
    "expected_income",
    "currency",
    "receiver_name",
    "receiver_phone",
    "receiver_email",
    "receiver_address",
    "raw_json",
)

# 店铺唯一业务维度；shop_hash = stable_line_hash(下列三列)
PLATFORM_SHOP_INSERT_COLUMNS: tuple[str, ...] = (
    "shop_hash",
    "shop_name_en",
    "shop_name_cn",
    "shop_alias",
    "store_account",
    "store_secret",
    "platform",
    "platform_site",
    "currency",
    "fx_rate",
    "ops_owner",
)

# Excel 表头在第 5 行 -> pandas header=4
_SHIPPED_MAP: list[tuple[str, str, str]] = [
    ("平台", "platform", "s64"),
    ("店铺英文名", "shop_name_en", "s128"),
    ("店铺别名", "shop_alias", "s128"),
    ("站点", "platform_site", "s64"),
    ("仓库", "warehouse_name", "s255"),
    ("仓库属性", "warehouse_type", "s64"),
    ("创建时间", "order_created_at", "dt"),
    ("付款时间", "pay_time", "dt"),
    ("发货时间", "ship_time", "dt"),
    ("审核时间", "audit_time", "dt"),
    ("订单销售状态", "order_sales_status", "s64"),
    ("订单类型", "order_type", "s64"),
    ("订单号", "order_no", "s128"),
    ("参考号", "ref_no", "s128e"),
    ("原平台sku", "platform_sku_orig", "s255"),
    ("平台sku", "platform_sku", "s255"),
    ("仓库sku", "warehouse_sku", "s128"),
    ("仓库sku销量", "warehouse_sku_qty", "int"),
    ("产品名称", "product_name", "s512"),
    ("产品款式", "product_style", "s512"),
    ("品牌", "brand", "s128"),
    ("一级品类", "category_lv1", "s128"),
    ("二级品类", "category_lv2", "s128"),
    ("三级品类", "category_lv3", "s128"),
    ("产品销售状态", "product_sales_status", "s64"),
    ("仓库单号", "warehouse_order_no", "s128"),
    ("服务商单号", "provider_order_no", "s128"),
    ("订单下架类型", "order_offline_type", "s64"),
    ("是否发货", "is_shipped", "s16"),
    ("是否售后", "is_after_sale", "s16"),
    ("跟踪单号", "tracking_no", "s255"),
    ("交易号", "transaction_no", "s255"),
    ("运输方式", "shipping_method", "s255"),
    ("产品重量", "product_weight", "dec"),
    ("发货批次号", "ship_batch_no", "s128"),
    ("订单付款币种", "pay_currency", "s16"),
    ("汇率（转本位币汇率）", "fx_rate_to_base", "dec"),
    ("销售单价（付款币种）", "unit_price_pay", "dec"),
    ("订单总金额（付款币种）", "order_total_pay", "dec"),
    ("订单商品金额（付款币种）", "order_goods_pay", "dec"),
    ("平台运费（付款币种）", "platform_shipping_pay", "dec"),
    ("支付手续费（付款币种）", "payment_fee_pay", "dec"),
    ("平台手续费（付款币种）", "platform_fee_pay", "dec"),
    ("fba费用（付款币种）", "fba_fee_pay", "dec"),
    ("平台补贴费（付款币种）", "platform_subsidy_pay", "dec"),
    ("税费（付款币种）", "tax_pay", "dec"),
    ("其他费用（付款币种）", "other_fee_pay", "dec"),
    ("采购成本（付款币种）", "purchase_cost_pay", "dec"),
    ("采购运费（付款币种）", "purchase_shipping_pay", "dec"),
    ("采购税费（付款币种）", "purchase_tax_pay", "dec"),
    ("头程运费（付款币种）", "first_leg_shipping_pay", "dec"),
    ("头程税费（付款币种）", "first_leg_tax_pay", "dec"),
    ("包材费用（付款币种）", "packaging_fee_pay", "dec"),
    ("派送运费（付款币种）", "delivery_shipping_pay", "dec"),
    ("币种", "base_currency", "s16"),
    ("销售单价", "unit_price_base", "dec"),
    ("订单总金额", "order_total_base", "dec"),
    ("订单商品金额", "order_goods_base", "dec"),
    ("平台运费", "platform_shipping_base", "dec"),
    ("支付手续费", "payment_fee_base", "dec"),
    ("平台手续费", "platform_fee_base", "dec"),
    ("fba费用", "fba_fee_base", "dec"),
    ("平台补贴费", "platform_subsidy_base", "dec"),
    ("税费", "tax_base", "dec"),
    ("其他费用", "other_fee_base", "dec"),
    ("采购成本", "purchase_cost_base", "dec"),
    ("采购运费", "purchase_shipping_base", "dec"),
    ("采购税费", "purchase_tax_base", "dec"),
    ("头程运费", "first_leg_shipping_base", "dec"),
    ("头程税费", "first_leg_tax_base", "dec"),
    ("包材费用", "packaging_fee_base", "dec"),
    ("派送运费", "delivery_shipping_base", "dec"),
    ("总费用", "total_fee_base", "dec"),
    ("总成本", "total_cost_base", "dec"),
    ("毛利", "gross_profit_base", "dec"),
    ("毛利率", "gross_margin_rate", "margin"),
    ("ITEM_ID", "item_id", "s128"),
    ("ASIN", "asin", "s64"),
    ("原销售订单号", "orig_sales_order_no", "s128"),
    ("销售负责人", "sales_owner", "s128"),
    ("采购负责人", "purchase_owner", "s128"),
    ("开发负责人", "dev_owner", "s128"),
    ("附属销售员", "sub_sales", "s128"),
    ("平台sku负责人", "platform_sku_owner", "s128"),
    ("店铺负责人", "shop_owner", "s128"),
    ("买家ID", "buyer_id", "s128"),
    ("买家邮箱", "buyer_email", "s255"),
    ("电话", "phone", "s128"),
    ("买家姓名", "buyer_name", "s255"),
    ("收件人", "consignee", "s255"),
    ("国家", "country", "s128"),
    ("州/省", "state", "s128"),
    ("城市", "city", "s128"),
    ("联系地址", "address_line", "s1024"),
    ("邮编", "postal_code", "s64"),
    ("销售订单号", "sales_order_no", "s128"),
    ("客服备注", "cs_remark", "s2048"),
    ("系统备注", "system_remark", "s2048"),
    ("财务备注", "finance_remark", "s2048"),
    ("shopify支付方式", "shopify_pay_method", "s128"),
    ("shopify支付交易号", "shopify_pay_txn_no", "s255"),
    ("线下费用汇总", "offline_fee_summary", "dec"),
]
# ======================== 订单统计 唯一性逻辑键 =============================================== 
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
# 注意：这里不可随意修改，否则历史行的 line_hash 与库内不一致，需按业务重导或清表
LINE_HASH_KEYS: tuple[str, ...] = (
    "platform",
    "platform_site",
    "shop_name_en",
    "platform_sku",
    "ref_no",
    "order_no",
    "item_id",
    "warehouse_sku"
)
# ================================================================================ 

def _log(level: str, msg: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] [{level}] {msg}")


def _excel_row_no(pandas_idx: int) -> int:
    """表头在第 5 行（header=4），数据行从 Excel 第 6 行起。"""
    return int(pandas_idx) + 6


def _log_duplicate_line_hashes(hash_groups: dict[str, list[dict[str, Any]]]) -> None:
    """输出本批导入中 line_hash 重复的分组明细。"""
    dups = {h: rows for h, rows in hash_groups.items() if len(rows) > 1}
    if not dups:
        _log("INFO", "line_hash 检查：本批无重复")
        return

    extra = sum(len(rows) - 1 for rows in dups.values())
    unique = len(hash_groups)
    total = unique + extra
    _log(
        "WARN",
        f"line_hash 重复：有效行={total}，唯一={unique}，重复组={len(dups)}，多出行数={extra}",
    )
    for i, (h, rows) in enumerate(
        sorted(dups.items(), key=lambda x: x[1][0].get("_excel_row", 0)),
        1,
    ):
        _log("WARN", f"  重复组 {i}/{len(dups)} line_hash={h}")
        for d in rows:
            key_bits = ", ".join(f"{k}={d.get(k)!r}" for k in LINE_HASH_KEYS)
            excel_row = d.get("_excel_row", "?")
            _log(
                "WARN",
                f"    Excel行{excel_row}: {key_bits}；"
                f"unit_price_pay={d.get('unit_price_pay')}，"
                f"order_goods_base={d.get('order_goods_base')}，"
                f"gross_profit_base={d.get('gross_profit_base')}",
            )


def _convert(v: Any, kind: str) -> Any:
    if kind == "dt":
        return cell_dt(v)
    if kind == "int":
        return cell_int(v)
    if kind == "dec":
        return cell_decimal(v)
    if kind == "margin":
        return cell_margin_rate(v)
    if kind.startswith("s") and kind.endswith("e"):
        return cell_str_or_empty(v, int(kind[1:-1]))
    if kind.startswith("s"):
        return cell_str(v, int(kind[1:]))
    return cell_str(v)


def _normalize_item_id(v: Any) -> str | None:
    """
    规范化 item_id：含逗号时拆分为多个 ID，排序后用英文逗号拼接。
    用于消除「id1，id2」与「id2,id1」导致的 line_hash 不一致。
    """
    s = cell_str(v, 128)
    if not s:
        return None
    if not _ITEM_ID_SEP_RE.search(s):
        return s
    parts = [p.strip() for p in _ITEM_ID_SEP_RE.split(s) if p.strip()]
    if not parts:
        return None
    if len(parts) == 1:
        return cell_str(parts[0], 128)
    return cell_str(",".join(sorted(parts)), 128)


def _row_dict(series: pd.Series) -> dict[str, Any]:
    out: dict[str, Any] = {col: None for _, col, _ in _SHIPPED_MAP}
    for zh, col, kind in _SHIPPED_MAP:
        if zh not in series.index:
            continue
        out[col] = _convert(series[zh], kind)
    out["item_id"] = _normalize_item_id(out.get("item_id"))
    out["source_type"] = SOURCE_TYPE
    return out


def _parse_currency_code(cell: Any) -> str:
    """
    从 A3 类单元格解析 ISO 币种代码（大写）。
    支持：EUR、币种:EUR、Currency: eur、币种： EUR 等常见写法。
    """
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return ""
    s = str(cell).strip()
    if not s:
        return ""
    compact = re.sub(r"\s+", "", s.upper())
    if re.fullmatch(r"[A-Z]{3}", compact):
        return compact
    for sep in (":", "："):
        if sep not in s:
            continue
        tail = re.sub(r"\s+", "", s.split(sep)[-1].strip().upper())
        m = re.match(r"^([A-Z]{3})\b", tail)
        if m:
            return m.group(1)
    return ""


def _read_excel_a2_a3_cells(xlsx: Path) -> tuple[Any, Any]:
    """读取订单统计 Excel A2、A3 单元格，返回 (A2 原始值, A3 原始值)。"""
    header_df = pd.read_excel(
        xlsx,
        sheet_name=0,
        header=None,
        usecols=[0],
        nrows=3,
        engine="openpyxl",
    )
    a2_raw = header_df.iloc[1, 0] if len(header_df) > 1 else None
    a3_raw = header_df.iloc[2, 0] if len(header_df) > 2 else None
    return a2_raw, a3_raw


def _print_excel_header_info(
    xlsx: Path,
    *,
    a2_raw: Any,
    a3_raw: Any,
    currency_code: str,
) -> None:
    """币种校验通过后，彩色输出 A2 / A3 表格元信息。"""
    a2_text = _cell_display(a2_raw) or "(空)"
    a3_text = _cell_display(a3_raw) or "(空)"
    sep = _c("=" * 60, "CYAN")
    print(sep, flush=True)
    print(
        f"{_c('[源文件]', 'BOLD', 'CYAN')} {_c(xlsx.name, 'BOLD')}",
        flush=True,
    )
    print(
        f"  {_c('币种校验：', 'YELLOW')}："
        f"{_c(a3_text, 'GREEN')} "
        f"→ 解析={_c(currency_code, 'BOLD', 'GREEN')} "
        f"{_c('✓ 校验通过', 'BOLD', 'GREEN')}",
        flush=True,
    )
    print(
        # f"  {_c('A2', 'YELLOW')}：{_c(a2_text, 'BOLD', 'CYAN')}",
        f"  {_c(a2_text, 'YELLOW')}",
        flush=True,
    )
    print(sep, flush=True)


def _validate_excel_currency(xlsx: Path) -> None:
    """
    校验 Excel A3 单元格币种为 EXPECTED_CURRENCY（默认 EUR）。
    校验通过后彩色输出 A2 单元格信息。
    """
    a2_raw, a3_raw = _read_excel_a2_a3_cells(xlsx)
    code = _parse_currency_code(a3_raw)
    if code != EXPECTED_CURRENCY:
        print(
            f"{_c('[币种校验失败]', 'BOLD', 'RED')} "
            f"文件={xlsx.name} "
            f"A3={_cell_display(a3_raw)!r} "
            f"解析={code!r} "
            f"要求={EXPECTED_CURRENCY}",
            flush=True,
        )
        raise RuntimeError(
            f"币种非 {EXPECTED_CURRENCY}（文件={xlsx.name}，A3={a3_raw!r}，解析={code!r}），导入已中止"
        )
    _print_excel_header_info(
        xlsx,
        a2_raw=a2_raw,
        a3_raw=a3_raw,
        currency_code=code,
    )


def _read_shipped_frame(xlsx: Path) -> pd.DataFrame:
    _log("INFO", f"读取 Excel：{xlsx}（表头第 5 行）")
    df = pd.read_excel(xlsx, sheet_name=0, header=4, engine="openpyxl", dtype=object)
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    df = df.dropna(how="all")
    _log("INFO", f"读取完成：行数={len(df)} 列数={len(df.columns)}")
    return df


def _insert_columns(*, with_import_batch: bool = False) -> list[str]:
    cols = ["line_hash"]
    cols.extend(col for _, col, _ in _SHIPPED_MAP)
    if with_import_batch:
        cols.append("import_batch")
    cols.append("distribution_lev")
    cols.append("source_type")
    return cols


def _is_blank_str(v: Any) -> bool:
    if v is None:
        return True
    return not str(v).strip()


def _is_distribution_warehouse(warehouse_name: Any) -> bool:
    """仓库名含「分销」时视为分销渠道。"""
    if warehouse_name is None:
        return False
    name = str(warehouse_name).strip()
    return bool(name) and _DISTRIBUTION_WH_MARK in name


def _resolve_distribution_lev(d: dict[str, Any]) -> int:
    """distribution_lev 为 0 且仓库名含「分销」时标记为 1。"""
    lev = d.get("distribution_lev")
    if lev is None:
        current = 0
    else:
        try:
            current = int(lev)
        except (TypeError, ValueError):
            current = 0
    if current == 0 and _is_distribution_warehouse(d.get("warehouse_name")):
        return 1
    return current


def _shop_triple(d: dict[str, Any]) -> tuple[str, str, str] | None:
    if _is_blank_str(d.get("platform")) or _is_blank_str(d.get("shop_name_en")):
        return None
    plat = str(d["platform"]).strip()
    site = d.get("platform_site")
    site_s = "" if site is None else str(site).strip()
    shop_en = str(d["shop_name_en"]).strip()
    return plat, site_s, shop_en


def _lookup_existing_shops(
    conn, triples: set[tuple[str, str, str]]
) -> set[tuple[str, str, str]]:
    """按 (platform, platform_site, shop_name_en) 批量查询已存在店铺。"""
    if not triples:
        return set()
    existing: set[tuple[str, str, str]] = set()
    items = list(triples)
    chunk = 200
    cur = conn.cursor(pymysql.cursors.Cursor)
    try:
        for i in range(0, len(items), chunk):
            part = items[i : i + chunk]
            placeholders = ",".join(["(%s,%s,%s)"] * len(part))
            sql = (
                f"SELECT `platform`, `platform_site`, `shop_name_en` FROM `{PLATFORM_SHOP_TABLE}` "
                f"WHERE (`platform`, `platform_site`, `shop_name_en`) IN ({placeholders})"
            )
            params: list[str] = []
            for p, s, n in part:
                params.extend([p, s, n])
            cur.execute(sql, params)
            for plat, site, name in cur.fetchall():
                existing.add(
                    (str(plat).strip(), "" if site is None else str(site).strip(), str(name).strip())
                )
    finally:
        cur.close()
    return existing


def _build_platform_shop_row(d: dict[str, Any]) -> tuple[Any, ...]:
    triple = _shop_triple(d)
    if triple is None:
        raise ValueError("invalid shop row")
    plat, site_s, shop_en = triple
    shop_hash = stable_line_hash(row_subset_for_line_hash(d, SHOP_HASH_KEYS))

    alias = d.get("shop_alias")
    alias_s = "" if alias is None else str(alias).strip()

    cur = d.get("base_currency") or d.get("pay_currency")
    currency_s = "" if cur is None else str(cur).strip()

    fx = d.get("fx_rate_to_base")
    fx_v = _ZERO_DEC if fx is None else fx

    ops = d.get("shop_owner")
    ops_s = "" if ops is None else str(ops).strip()

    return (
        shop_hash,
        shop_en,
        alias_s,  # shop_name_cn
        alias_s,  # shop_alias
        "",  # store_account
        "",  # store_secret
        plat,
        site_s,
        currency_s,
        fx_v,
        ops_s,
    )


def _is_semitemu(d: dict[str, Any]) -> bool:
    p = d.get("platform")
    if p is None:
        return False
    return str(p).strip().lower() == PLATFORM_SEMITEMU


def _is_sales_order(d: dict[str, Any]) -> bool:
    ot = d.get("order_type")
    if ot is None:
        return False
    return str(ot).strip() == ORDER_TYPE_SALES


def _temu_dec2(v: Any) -> Decimal | None:
    if v is None:
        return None
    if not isinstance(v, Decimal):
        v = cell_decimal(v)
    if v is None:
        return None
    return v.quantize(_TWO_PLACES)


def _temu_sku_id(d: dict[str, Any]) -> str | None:
    for key in ("platform_sku", "platform_sku_orig", "item_id", "warehouse_sku"):
        s = cell_str(d.get(key), max_len=100)
        if s:
            return s
    return None


def _temu_receiver_address(d: dict[str, Any]) -> str | None:
    parts: list[str] = []
    for key in ("address_line", "city", "state", "postal_code", "country"):
        s = cell_str(d.get(key))
        if s:
            parts.append(s)
    if not parts:
        return None
    return ", ".join(parts)


def _json_default(obj: Any) -> Any:
    if isinstance(obj, datetime):
        return obj.isoformat(sep=" ", timespec="seconds")
    if isinstance(obj, Decimal):
        return format(obj.normalize(), "f")
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")


def _build_temu_order_item_dict(d: dict[str, Any]) -> dict[str, Any] | None:
    order_no = cell_str(d.get("ref_no"), max_len=50)
    sku_id = _temu_sku_id(d)
    if not order_no or not sku_id:
        return None

    receiver_name = cell_str(d.get("consignee"), max_len=100)
    if not receiver_name:
        receiver_name = cell_str(d.get("buyer_name"), max_len=100)

    currency = (
        cell_str(d.get("pay_currency"), max_len=10)
        or cell_str(d.get("base_currency"), max_len=10)
        or "USD"
    )

    raw_subset = {k: v for k, v in d.items() if k not in ("line_hash", "import_batch")}
    # raw_json = json.dumps(raw_subset, ensure_ascii=False, default=_json_default)

    return {
        "line_hash": cell_str(d.get("line_hash"), max_len=64),
        "order_no": order_no,
        "sub_order_no": "",
        "shop_account": "",
        "shop_site_name": "",
        "shop_name_en": cell_str(d.get("shop_name_en"), max_len=100),
        "country_region": cell_str(d.get("country"), max_len=50),
        "order_status": cell_str(d.get("order_sales_status"), max_len=50),
        "shipping_method": cell_str(d.get("shipping_method"), max_len=100),
        "created_time": d.get("order_created_at"),
        "confirmed_time": d.get("pay_time"),
        "required_ship_deadline": None,
        "shipped_time": d.get("ship_time"),
        "estimated_delivery_time": None,
        "delivered_time": None,
        "package_no": None,
        "tracking_no": cell_str(d.get("tracking_no"), max_len=100),
        "logistics_company": None,
        "sku_id": cell_str(d.get("platform_sku"), max_len=100),
        "skc_id": None,
        "spu_id": None,
        "warehouse_sku": cell_str(d.get("warehouse_sku"), max_len=100),
        "product_name": cell_str(d.get("product_name"), max_len=500),
        "variant_name": None,
        "quantity": d.get("warehouse_sku_qty") if d.get("warehouse_sku_qty") is not None else 1,
        "warehouse_name": cell_str(d.get("warehouse_name"), max_len=100),
        "declared_price": _temu_dec2(d.get("unit_price_pay")),
        "order_payment": _temu_dec2(d.get("order_goods_pay")),
        "second_payment": None,
        "sales_revenue": _temu_dec2(d.get("order_total_pay")),
        "sales_return": None,
        "shipping_income": _temu_dec2(d.get("platform_shipping_pay")),
        "shipping_deduction": None,
        "expected_income": _temu_dec2(d.get("order_total_pay")),
        "currency": currency,
        "receiver_name": receiver_name,
        "receiver_phone": cell_str(d.get("phone"), max_len=100),
        "receiver_email": cell_str(d.get("buyer_email"), max_len=200),
        "receiver_address": _temu_receiver_address(d),
        "raw_json": None,
    }


def _lookup_temu_file_name_locked(
    conn,
    keys: set[tuple[str, str]],
) -> set[tuple[str, str]]:
    """返回 file_name 非空的 (order_no, sku_id)。"""
    if not keys:
        return set()
    locked: set[tuple[str, str]] = set()
    items = sorted(keys)
    chunk = 200
    cur = conn.cursor(pymysql.cursors.Cursor)
    try:
        for i in range(0, len(items), chunk):
            part = items[i : i + chunk]
            placeholders = ",".join(["(%s,%s)"] * len(part))
            sql = (
                f"SELECT `order_no`, `sku_id` FROM `{TEMU_ORDER_ITEM_TABLE}` "
                f"WHERE (`order_no`, `sku_id`) IN ({placeholders}) "
                f"AND `file_name` IS NOT NULL AND TRIM(`file_name`) <> ''"
            )
            params: list[str] = []
            for order_no, sku_id in part:
                params.extend([order_no, sku_id])
            cur.execute(sql, params)
            for order_no, sku_id in cur.fetchall():
                locked.add((str(order_no).strip(), str(sku_id).strip()))
    finally:
        cur.close()
    return locked


def _upsert_temu_order_item_rows(
    conn,
    rows: list[tuple[Any, ...]],
    *,
    chunk_size: int = 300,
) -> int:
    """
    UPSERT temu_order_item；若已有行 file_name 非空则保留原业务字段不覆盖。
    line_hash 始终与 sales_order_shipped 同步（不受 file_name 保护影响）。
    """
    if not rows:
        return 0
    columns = list(TEMU_ORDER_ITEM_COLUMNS)
    cols_sql = ", ".join(f"`{c}`" for c in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    file_guard = "(TRIM(COALESCE(`file_name`, '')) = '')"
    updates = ", ".join(
        f"`{c}`=IF({file_guard}, VALUES(`{c}`), `{c}`)"
        for c in columns
        if c != "line_hash"
    )
    updates += ", `line_hash`=VALUES(`line_hash`)"
    sql = (
        f"INSERT INTO `{TEMU_ORDER_ITEM_TABLE}` ({cols_sql}) VALUES ({placeholders}) "
        f"ON DUPLICATE KEY UPDATE {updates}"
    )
    cur = conn.cursor()
    n = 0
    buf: list[tuple[Any, ...]] = []
    try:
        for row in rows:
            buf.append(row)
            if len(buf) >= chunk_size:
                cur.executemany(sql, buf)
                n += len(buf)
                buf.clear()
        if buf:
            cur.executemany(sql, buf)
            n += len(buf)
    finally:
        cur.close()
    return n


def _sync_temu_line_hash_for_locked_rows(
    conn,
    rows: list[tuple[str, str, str]],
    *,
    chunk_size: int = 300,
) -> int:
    """file_name 已锁定的行仅同步 line_hash（order_no + sku_id 定位）。"""
    if not rows:
        return 0
    sql = (
        f"UPDATE `{TEMU_ORDER_ITEM_TABLE}` SET `line_hash`=%s "
        f"WHERE `order_no`=%s AND `sku_id`=%s "
        f"AND `file_name` IS NOT NULL AND TRIM(`file_name`) <> ''"
    )
    cur = conn.cursor()
    n = 0
    buf: list[tuple[str, str, str]] = []
    try:
        for line_hash, order_no, sku_id in rows:
            if not line_hash:
                continue
            buf.append((line_hash, order_no, sku_id))
            if len(buf) >= chunk_size:
                cur.executemany(sql, buf)
                n += cur.rowcount
                buf.clear()
        if buf:
            cur.executemany(sql, buf)
            n += cur.rowcount
    finally:
        cur.close()
    return n


def upsert_temu_order_items(conn, dicts: list[dict[str, Any]]) -> tuple[int, int]:
    """
    platform=semitemu 且订单类型=销售订单 的行写入 temu_order_item（UPSERT，唯一键 order_no + sku_id）。
    若库内该行 file_name 已有值，则跳过业务字段更新，但仍同步 line_hash。
    返回 (UPSERT 行数, 跳过行数)。
    """
    temu_dicts = [d for d in dicts if _is_semitemu(d) and _is_sales_order(d)]
    if not temu_dicts:
        return 0, 0

    pending: list[tuple[str, str, str, tuple[Any, ...]]] = []
    skipped = 0
    for d in temu_dicts:
        item = _build_temu_order_item_dict(d)
        if item is None:
            skipped += 1
            continue
        order_no = item["order_no"]
        sku_id = item["sku_id"]
        line_hash = cell_str(item.get("line_hash"), max_len=64) or ""
        pending.append(
            (
                order_no,
                sku_id,
                line_hash,
                tuple(item[c] for c in TEMU_ORDER_ITEM_COLUMNS),
            )
        )

    if not pending:
        _log("WARN", f"{TEMU_ORDER_ITEM_TABLE}：semitemu 行均无 order_no/sku_id，跳过={skipped}")
        return 0, skipped

    locked = _lookup_temu_file_name_locked(
        conn, {(order_no, sku_id) for order_no, sku_id, _, _ in pending}
    )
    rows: list[tuple[Any, ...]] = []
    locked_line_hash_rows: list[tuple[str, str, str]] = []
    skipped_file_name = 0
    for order_no, sku_id, line_hash, row in pending:
        if (order_no, sku_id) in locked:
            skipped_file_name += 1
            if line_hash:
                locked_line_hash_rows.append((line_hash, order_no, sku_id))
            continue
        rows.append(row)

    skipped += skipped_file_name
    if skipped_file_name:
        _log(
            "INFO",
            f"{TEMU_ORDER_ITEM_TABLE}：file_name 已有值跳过业务字段={skipped_file_name}",
        )

    n = 0
    if rows:
        _log(
            "INFO",
            f"准备写入 {TEMU_ORDER_ITEM_TABLE}：有效行={len(rows)} 跳过={skipped}",
        )
        n = _upsert_temu_order_item_rows(conn, rows)
    elif skipped_file_name:
        _log(
            "INFO",
            f"{TEMU_ORDER_ITEM_TABLE}：无业务字段可写行，仅同步 line_hash（file_name 保护）",
        )

    if locked_line_hash_rows:
        n_hash = _sync_temu_line_hash_for_locked_rows(conn, locked_line_hash_rows)
        _log(
            "INFO",
            f"{TEMU_ORDER_ITEM_TABLE}：file_name 保护行 line_hash 同步={n_hash}",
        )

    if not rows and not locked_line_hash_rows:
        _log(
            "WARN",
            f"{TEMU_ORDER_ITEM_TABLE}：有效行均被跳过，跳过合计={skipped}",
        )
        return 0, skipped

    return n, skipped


def ensure_platform_shops(conn, dicts: list[dict[str, Any]]) -> int:
    """
    导入订单前补全 platform_shop：
    唯一维度 (platform, platform_site, shop_name_en)，不存在则 INSERT IGNORE。
    返回本批新增店铺数。
    """
    seen: set[tuple[str, str, str]] = set()
    candidates: list[dict[str, Any]] = []
    for d in dicts:
        triple = _shop_triple(d)
        if triple is None or triple in seen:
            continue
        seen.add(triple)
        candidates.append(d)

    if not candidates:
        return 0

    existing = _lookup_existing_shops(conn, seen)
    missing = [d for d in candidates if _shop_triple(d) not in existing]
    if not missing:
        _log("INFO", f"platform_shop：{len(seen)} 个店铺均已存在，无需新增")
        return 0

    rows = [_build_platform_shop_row(d) for d in missing]
    _log(
        "INFO",
        f"platform_shop：待新增 {len(rows)} 条（已存在 {len(existing)}）",
    )
    n_new = insert_ignore_rows(
        conn,
        table=PLATFORM_SHOP_TABLE,
        columns=list(PLATFORM_SHOP_INSERT_COLUMNS),
        rows=rows,
    )
    _log("INFO", f"platform_shop：实际新增 {n_new} 条")
    return n_new


def import_file(conn, xlsx: Path, *, import_batch: str | None = None) -> tuple[int, int, int, int]:
    """返回 (UPSERT 行数, 跳过行数, Excel 总行数, temu_order_item UPSERT 行数)。"""
    _validate_excel_currency(xlsx)
    df = _read_shipped_frame(xlsx)
    insert_cols = _insert_columns(with_import_batch=import_batch is not None)
    dicts: list[dict[str, Any]] = []
    skipped = 0

    for idx, series in df.iterrows():
        d = _row_dict(series)
        order_no = d.get("order_no")
        wh_sku = d.get("warehouse_sku")
        if not order_no or not str(order_no).strip() or not wh_sku or not str(wh_sku).strip():
            skipped += 1
            continue
        if d.get("ref_no") is None:
            d["ref_no"] = ""
        d["_excel_row"] = _excel_row_no(idx)
        d["distribution_lev"] = _resolve_distribution_lev(d)
        dicts.append(d)

    if not dicts:
        _log("WARN", f"无有效行：Excel 行数={len(df)} 跳过={skipped}")
        return 0, skipped, len(df), 0

    ensure_platform_shops(conn, dicts)
    mapping_dicts = [d for d in dicts if not _is_distribution_warehouse(d.get("warehouse_name"))]
    dist_mapping_skipped = len(dicts) - len(mapping_dicts)
    if dist_mapping_skipped:
        _log("INFO", f"product_sku_mapping：分销仓跳过 {dist_mapping_skipped} 条")
    upsert_product_sku_mapping(conn, mapping_dicts, source_type=SOURCE_TYPE, log_fn=_log)

    rows: list[tuple[Any, ...]] = []
    hash_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for d in dicts:
        h_in = row_subset_for_line_hash(d, LINE_HASH_KEYS)
        d["line_hash"] = stable_line_hash(h_in)
        hash_groups[d["line_hash"]].append(d)
        if import_batch is not None:
            d["import_batch"] = import_batch
        rows.append(tuple(d[c] for c in insert_cols))

    _log_duplicate_line_hashes(hash_groups)

    _log(
        "INFO",
        f"准备写入 {TABLE}：有效行={len(rows)} 跳过={skipped} line_hash 键数={len(LINE_HASH_KEYS)}",
    )
    n = upsert_rows(conn, table=TABLE, columns=insert_cols, rows=rows)
    n_temu, _ = upsert_temu_order_items(conn, dicts)
    return n, skipped, len(df), n_temu


def erp_base_dir(mode: str | None = None) -> Path:
    pattern = mode or MODE_PATTERN
    return Path(ERP_ORDER_STA_PATH.format(MODE_PATTERN=pattern))


def resolve_date_dir(base: Path, mode: str, on_date: date) -> Path:
    if mode == "每月":
        return base / f"{on_date.year:04d}-{on_date.month:02d}"
    return base / on_date.isoformat()


def default_date_dir(base: Path) -> Path:
    """默认日期目录：path_config.DATE_PATH（已按日报/月报格式化）。"""
    return base / DATE_PATH


def resolve_work_dir(base: Path, mode: str, on_date: date | None) -> Path:
    if on_date is not None:
        return resolve_date_dir(base, mode, on_date)
    return default_date_dir(base)


def discover_shipped_files(directory: Path) -> list[Path]:
    files = sorted(directory.glob("订单统计*.xlsx"))
    if not files:
        files = sorted(directory.glob("*订单统计*.xlsx"))
    return [p for p in files if p.is_file() and not p.name.startswith("~$")]


def _resolve_import_batch(cli_batch: str | None) -> str | None:
    """
    import_batch 优先取命令行 --import-batch；
    未指定时：复用当日 run_batch.lock，若无锁则生成新批次并写入锁文件。
    """
    batch, lock_written = resolve_import_batch(cli_batch)
    if lock_written:
        _log("INFO", f"run_batch.lock 不存在/非当日，已生成并写入 import_batch：{batch}")
    else:
        if cli_batch and cli_batch.strip():
            _log("INFO", f"使用命令行 import_batch：{batch}")
        else:
            _log("INFO", f"复用 run_batch.lock 的 import_batch：{batch}")
    return batch


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="ERP 订单统计 Excel -> sales_order_shipped")
    ap.add_argument(
        "--date",
        type=date.fromisoformat,
        default=None,
        metavar="YYYY-MM-DD",
        help=f"覆盖日期子目录（默认 path_config.DATE_PATH={DATE_PATH}）",
    )
    ap.add_argument(
        "--mode",
        choices=("每天", "每月"),
        default=None,
        help=f"路径模式，默认 path_config.MODE_PATTERN（{MODE_PATTERN}）",
    )
    ap.add_argument("--dir", type=Path, default=None, help="直接指定 Excel 目录")
    ap.add_argument("--file", type=Path, default=None, help="指定单个 xlsx")
    ap.add_argument(
        "--import-batch",
        "--batch",
        dest="import_batch",
        default=None,
        metavar="BATCH",
        help="导入批次号写入 import_batch（默认读 run_batch.lock 的 import_batch）",
    )
    return ap.parse_args()


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    _enable_windows_ansi()

    args = parse_args()
    mode = args.mode or MODE_PATTERN
    import_batch = _resolve_import_batch(args.import_batch)

    if args.file:
        files = [args.file.resolve()]
        work_dir = args.file.parent
    elif args.dir:
        work_dir = args.dir.resolve()
        if not work_dir.is_dir():
            _log("ERROR", f"目录不存在：{work_dir}")
            return 2
        files = discover_shipped_files(work_dir)
    else:
        work_dir = resolve_work_dir(erp_base_dir(mode), mode, args.date)
        if not work_dir.is_dir():
            _log("ERROR", f"日期目录不存在：{work_dir}")
            return 2
        files = discover_shipped_files(work_dir)

    if not files:
        _log("ERROR", f"未找到 订单统计-*.xlsx：{work_dir}")
        return 1

    _log("INFO", f"任务：导入 -> {TABLE}")
    _log("INFO", f"模式={mode} 目录={work_dir} 文件数={len(files)}")

    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    total_upsert = 0
    total_skip = 0
    total_temu = 0
    try:
        for fp in files:
            n, skipped, n_excel, n_temu = import_file(conn, fp, import_batch=import_batch)
            conn.commit()
            _log(
                "INFO",
                f"已提交：{fp.name} Excel行={n_excel} UPSERT={n} "
                f"temu_order_item={n_temu} 跳过={skipped}",
            )
            total_upsert += n
            total_skip += skipped
            total_temu += n_temu
        _log(
            "INFO",
            f"全部完成：UPSERT累计={total_upsert} temu_order_item累计={total_temu} 总跳过={total_skip}",
        )
        return 0
    except Exception:
        conn.rollback()
        _log("ERROR", "导入失败，已回滚")
        raise
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
