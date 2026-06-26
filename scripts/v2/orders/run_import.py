from __future__ import annotations

"""
一键导入：按固定顺序执行本目录下全部 import_*.py 对应的核心逻辑（模块内 import_file，单连接、末尾一次 commit）。

顺序见 _IMPORT_STEMS_ORDER（与磁盘上 import_*.py 必须一致，缺或多会在启动时报错）。

默认阶段：
  1) import_order_shipped  — 订单统计*.xlsx -> sales_order_shipped
  2) import_order_refund   — RMA*.xlsx -> sales_order_refund
  3) import_amz_transaction — transaction交易明细*.xlsx -> amz_transaction（目录见 --amz-dir）
  4) import_temu_fee     — 步骤1：TEMU Excel -> temu_order_detail；步骤2：明细补全 + sales_order_shipped 费用
     （默认按近 30 天写入，见 --temu-write-order-days）

用法（在 python 目录下）：
  python v2/orders/run_import.py
  python v2/orders/run_import.py --order-dir path/to/excel/daily/order
  python v2/orders/run_import.py --shipped-only
  python v2/orders/run_import.py --refund-only
  python v2/orders/run_import.py --amz-only
  python v2/orders/run_import.py --temu-only
  python v2/orders/run_import.py --no-temu
  python v2/orders/run_import.py --no-amz
  python v2/orders/run_import.py --amz-dir path/to/excel/daily/amazon
  python v2/orders/run_import.py --temu-file path/to/TEMU-订单详情.xlsx
  python v2/orders/run_import.py --temu-no-mail
  python v2/orders/run_import.py --temu-no-detail-table   # 仅步骤2（可无 TEMU xlsx）
  python v2/orders/run_import.py --temu-write-order-days 0   # TEMU 全量写入（不按近 N 天过滤；默认近 30 天）
"""

import argparse
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None  # type: ignore[assignment]

_THIS_DIR = Path(__file__).resolve().parent
_V2_DIR = _THIS_DIR.parent
_WR_DIR = _V2_DIR / "warehouse-rent"
for _p in (_V2_DIR, _THIS_DIR, _WR_DIR):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))


def _load_env_files() -> None:
    if not load_dotenv:
        return
    for env_path in (
        _V2_DIR.parent / ".env",  # python/.env
        _V2_DIR / ".env",  # python/v2/.env
        _V2_DIR.parent.parent / ".env",  # 仓库根 .env
    ):
        if env_path.is_file():
            load_dotenv(env_path)
            return


from db import connect, load_db_config  # noqa: E402
from excel_common import default_order_excel_dir  # noqa: E402
from import_amz_transaction import (  # noqa: E402
    default_amz_transaction_excel_dir,
    import_file as import_amz_txn_file,
)
from import_order_refund import import_file as import_refund_file  # noqa: E402
from import_order_shipped import import_file as import_shipped_file  # noqa: E402
from import_temu_fee import (  # noqa: E402
    default_temu_excel_path,
    import_file as import_temu_file,
    send_notification as send_temu_notification,
)
from logger import get_logger, setup_stdout_utf8  # noqa: E402

_log = get_logger("ORDERS-RUN")

# 本目录 import_*.py 的执行顺序（后一步可能依赖前面已写入的数据，勿随意调换）
_IMPORT_STEMS_ORDER: tuple[str, ...] = (
    "import_order_shipped",
    "import_order_refund",
    "import_temu_fee",
    "import_amz_transaction",
)

# 与 _IMPORT_STEMS_ORDER 一一对应（新增脚本时两处一起改）
_STEP_LABEL_ZH: dict[str, str] = {
    "import_order_shipped": "订单发货 -> sales_order_shipped",
    "import_order_refund": "RMA 退款 -> sales_order_refund",
    "import_temu_fee": "TEMU：Excel->temu_order_detail；DB 同步 sales_order_shipped",
    "import_amz_transaction": "Amazon 交易明细 -> amz_transaction",
}


def _stems_on_disk() -> set[str]:
    return {p.stem for p in _THIS_DIR.glob("import_*.py") if p.is_file()}


# 存在于本目录但不参与「一键导入」流水线的 import_*.py（校验磁盘集合时排除）
_IMPORT_STEMS_EXCLUDE_ON_DISK: frozenset[str] = frozenset({"import_order_returned"})


def _validate_import_registry() -> None:
    """保证顺序表与目录中 import_*.py 一一对应，避免漏跑或顺序漂移。"""
    on_disk = _stems_on_disk() - _IMPORT_STEMS_EXCLUDE_ON_DISK
    ordered = set(_IMPORT_STEMS_ORDER)
    if on_disk != ordered:
        extra = sorted(on_disk - ordered)
        missing = sorted(ordered - on_disk)
        parts: list[str] = []
        if extra:
            parts.append(f"目录有但未列入 _IMPORT_STEMS_ORDER：{extra}")
        if missing:
            parts.append(f"顺序表有但文件不存在：{missing}")
        raise RuntimeError("run_import.py 的 import 登记与 orders 目录不一致 — " + "；".join(parts))
    label_stems = frozenset(_STEP_LABEL_ZH)
    if label_stems != ordered:
        raise RuntimeError(
            "_STEP_LABEL_ZH 的键必须与 _IMPORT_STEMS_ORDER 集合一致；"
            f"差集 label-order={sorted(label_stems - ordered)} order-label={sorted(ordered - label_stems)}"
        )


def _discover_shipped(base: Path) -> list[Path]:
    files = sorted(base.glob("订单统计*.xlsx"))
    if not files:
        files = sorted(base.glob("*订单统计*.xlsx"))
    return [p for p in files if p.is_file()]


def _discover_refund(base: Path) -> list[Path]:
    return sorted(p for p in base.glob("RMA*.xlsx") if p.is_file())


def _discover_amz_transaction(base: Path) -> list[Path]:
    return sorted(
        p
        for p in base.glob("transaction交易明细*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    )


def _resolve_temu_file(order_dir: Path, override: Path | None) -> Path | None:
    """
    定位 TEMU-订单详情.xlsx：
    - 用户显式 --temu-file 优先
    - 否则尝试 order_dir/TEMU-订单详情.xlsx
    - 再否则用 import_temu_fee.default_temu_excel_path()
    返回 None 表示文件不存在（调用方决定是 warn-跳过 还是报错）。
    """
    if override is not None:
        return override if override.is_file() else None
    candidate = order_dir / "TEMU-订单详情.xlsx"
    if candidate.is_file():
        return candidate
    fallback = default_temu_excel_path()
    return fallback if fallback.is_file() else None


def run_import(
    *,
    order_dir: Path,
    amz_dir: Path,
    do_shipped: bool,
    do_refund: bool,
    do_amz: bool,
    do_temu: bool,
    temu_file: Path | None = None,
    temu_send_mail: bool = True,
    temu_always_mail: bool = False,
    temu_save_detail: bool = True,
    temu_write_order_days: int | None = 30,
) -> dict[str, object]:
    _validate_import_registry()

    cfg = load_db_config()
    conn = connect(cfg)
    _log.info(f"连接数据库：host={cfg.host} port={cfg.port} database={cfg.database} user={cfg.user}")

    shipped_stats = {"files": 0, "upsert_rows": 0, "skipped": 0}
    refund_stats = {"files": 0, "upsert_rows": 0, "skipped": 0}
    amz_stats = {"files": 0, "upsert_rows": 0, "skipped": 0}
    temu_stats: dict[str, object] = {
        "ran": False,
        "file": "",
        "steps": "",
        "detail_rows_saved": 0,
        "detail_backfilled": 0,
        "matched": 0,
        "updated": 0,
        "unmatched_db": 0,
        "unmatched_excel": 0,
        "write_order_days": None,
    }
    temu_stats_obj = None  # 用于事务提交后发邮件

    try:
        _log.info(f"扫描订单目录：{order_dir.resolve()}")
        _log.info(f"扫描 Amazon 交易目录：{amz_dir.resolve()}")
        shipped_files = _discover_shipped(order_dir) if do_shipped else []
        refund_files = _discover_refund(order_dir) if do_refund else []
        amz_files = _discover_amz_transaction(amz_dir) if do_amz else []
        temu_xlsx = _resolve_temu_file(order_dir, temu_file) if do_temu else None
        _log.info(
            f"发现 订单统计*.xlsx：{len(shipped_files)}；RMA*.xlsx：{len(refund_files)}；"
            f"transaction交易明细*.xlsx：{len(amz_files)}；"
            f"TEMU 文件：{temu_xlsx if temu_xlsx else '（未找到/未启用）'}"
        )

        enabled: dict[str, bool] = {
            "import_order_shipped": do_shipped,
            "import_order_refund": do_refund,
            "import_amz_transaction": do_amz,
            "import_temu_fee": do_temu,
        }
        n_steps = sum(1 for s in _IMPORT_STEMS_ORDER if enabled[s])
        step_i = 0

        for stem in _IMPORT_STEMS_ORDER:
            if not enabled[stem]:
                continue
            step_i += 1
            label_zh = _STEP_LABEL_ZH[stem]
            _log.warn(f"【{step_i}/{n_steps}】{label_zh}（{stem}）")

            if stem == "import_order_shipped":
                if not shipped_files:
                    _log.warn("未找到 订单统计*.xlsx，本阶段跳过")
                for fp in shipped_files:
                    _log.info(f"导入发货：{fp.name}")
                    n, skipped, n_excel = import_shipped_file(conn, fp)
                    shipped_stats["files"] += 1
                    shipped_stats["upsert_rows"] += n
                    shipped_stats["skipped"] += skipped
                    _log.info(
                        f"完成：{fp.name} Excel行={n_excel} UPSERT累计={n} 跳过={skipped} "
                        f"（发货阶段累计 UPSERT={shipped_stats['upsert_rows']}）"
                    )

            elif stem == "import_order_refund":
                if not refund_files:
                    _log.warn("未找到 RMA*.xlsx，本阶段跳过")
                for fp in refund_files:
                    _log.info(f"导入退款：{fp.name}")
                    n, skipped, n_excel = import_refund_file(conn, fp)
                    refund_stats["files"] += 1
                    refund_stats["upsert_rows"] += n
                    refund_stats["skipped"] += skipped
                    _log.info(
                        f"完成：{fp.name} Excel行={n_excel} UPSERT累计={n} 跳过={skipped} "
                        f"（退款阶段累计 UPSERT={refund_stats['upsert_rows']}）"
                    )

            elif stem == "import_amz_transaction":
                if not amz_dir.is_dir():
                    _log.warn(f"Amazon 交易目录不存在，跳过：{amz_dir}")
                elif not amz_files:
                    _log.warn("未找到 transaction交易明细*.xlsx，本阶段跳过")
                for fp in amz_files:
                    _log.info(f"导入 Amazon 交易：{fp.name}")
                    n, skipped, n_excel = import_amz_txn_file(conn, fp)
                    amz_stats["files"] += 1
                    amz_stats["upsert_rows"] += n
                    amz_stats["skipped"] += skipped
                    _log.info(
                        f"完成：{fp.name} Excel行={n_excel} UPSERT累计={n} 跳过={skipped} "
                        f"（Amazon 交易累计 UPSERT={amz_stats['upsert_rows']}）"
                    )

            elif stem == "import_temu_fee":
                if temu_xlsx is None and temu_save_detail:
                    _log.warn(
                        "未找到 TEMU-订单详情.xlsx，本阶段跳过（如需指定路径请用 --temu-file）"
                    )
                else:
                    if temu_xlsx is not None:
                        _log.info(f"导入 TEMU：{temu_xlsx}")
                    elif not temu_save_detail:
                        _log.info("TEMU：未找到 Excel，仅执行步骤 2（temu_order_detail → shipped）")
                    wod = None if temu_write_order_days == 0 else temu_write_order_days
                    if wod is not None:
                        _log.info(
                            f"TEMU：write_order_days={wod}（等同单独执行 import_temu_fee.py --write-order-days {wod}）"
                        )
                    else:
                        _log.info("TEMU：write_order_days=全量（未按近 N 天过滤）")
                    temu_stats_obj = import_temu_file(
                        conn,
                        temu_xlsx,
                        save_detail_table=temu_save_detail,
                        write_order_days=wod,
                    )
                    temu_stats.update(
                        {
                            "ran": True,
                            "file": (temu_xlsx.name if temu_xlsx is not None else ""),
                            "steps": temu_stats_obj.steps_label,
                            "detail_rows_saved": temu_stats_obj.detail_rows_saved,
                            "detail_backfilled": temu_stats_obj.detail_backfilled,
                            "matched": temu_stats_obj.matched,
                            "updated": temu_stats_obj.updated,
                            "unmatched_db": len(temu_stats_obj.unmatched_db),
                            "unmatched_excel": len(temu_stats_obj.unmatched_excel),
                            "write_order_days": temu_stats_obj.write_order_days_applied,
                        }
                    )
                    _log.info(
                        f"完成：步骤={temu_stats_obj.steps_label} "
                        f"detail UPSERT={temu_stats_obj.detail_rows_saved} "
                        f"明细回写={temu_stats_obj.detail_backfilled} "
                        f"shipped 匹配={temu_stats_obj.matched} "
                        f"实际更新={temu_stats_obj.updated} "
                        f"DB未匹配={len(temu_stats_obj.unmatched_db)} "
                        f"明细未命中shipped={len(temu_stats_obj.unmatched_excel)}"
                    )

        conn.commit()
        summary = {
            "shipped": shipped_stats,
            "refund": refund_stats,
            "amz_transaction": amz_stats,
            "temu": temu_stats,
        }
        _log.info(f"全部完成（已提交）：{summary}")
    except Exception:
        _log.error("发生异常，准备回滚")
        conn.rollback()
        raise
    finally:
        _log.info("关闭数据库连接")
        conn.close()

    # 邮件通知放到事务关闭之后，避免 SMTP 卡顿拖长事务窗口
    if temu_stats_obj is not None:
        n_unmatched = (
            len(temu_stats_obj.unmatched_db) + len(temu_stats_obj.unmatched_excel)
        )
        if not temu_send_mail:
            _log.info("--temu-no-mail 已设置，跳过 TEMU 邮件通知")
        elif temu_always_mail or n_unmatched > 0:
            send_temu_notification(temu_stats_obj, dry_run=False)
        else:
            _log.info("TEMU 阶段无未匹配项，按默认策略不发邮件（如需总是发送请加 --temu-always-mail）")

    return summary


def main() -> int:
    setup_stdout_utf8()
    _load_env_files()

    parser = argparse.ArgumentParser(
        description=(
            "按固定顺序导入：订单统计、RMA、Amazon 交易明细、TEMU 费用（详见脚本顶部说明）"
        )
    )
    parser.add_argument(
        "--order-dir",
        type=Path,
        default=None,
        help=f"订单 Excel 目录，默认 {default_order_excel_dir()}",
    )
    parser.add_argument(
        "--amz-dir",
        type=Path,
        default=None,
        help=f"Amazon transaction Excel 目录，默认 {default_amz_transaction_excel_dir()}",
    )
    parser.add_argument(
        "--shipped-only",
        action="store_true",
        help="仅执行 import_order_shipped（订单统计*.xlsx）",
    )
    parser.add_argument(
        "--refund-only",
        action="store_true",
        help="仅执行 import_order_refund（RMA*.xlsx）",
    )
    parser.add_argument(
        "--amz-only",
        action="store_true",
        help="仅执行 import_amz_transaction（transaction交易明细*.xlsx）",
    )
    parser.add_argument(
        "--temu-only",
        action="store_true",
        help="仅执行 import_temu_fee（默认同路径 TEMU Excel：步骤1+2；配合 --temu-no-detail-table 可仅步骤2）",
    )
    parser.add_argument(
        "--no-temu",
        action="store_true",
        help="跳过 TEMU 费用回填阶段",
    )
    parser.add_argument(
        "--no-amz",
        action="store_true",
        help="跳过 Amazon 交易明细导入阶段",
    )
    parser.add_argument(
        "--temu-file",
        type=Path,
        default=None,
        help="覆盖 TEMU 文件路径；默认在 --order-dir 下找 TEMU-订单详情.xlsx",
    )
    parser.add_argument(
        "--temu-no-mail",
        action="store_true",
        help="TEMU 阶段不发邮件（即使有未匹配项）",
    )
    parser.add_argument(
        "--temu-always-mail",
        action="store_true",
        help="TEMU 阶段即便全部匹配也发一封成功邮件",
    )
    parser.add_argument(
        "--temu-no-detail-table",
        action="store_true",
        help="TEMU 仅执行步骤 2：从 temu_order_detail 补全并发货表费用（不写明细表；可无 TEMU xlsx）",
    )
    parser.add_argument(
        "--temu-write-order-days",
        type=int,
        default=30,
        metavar="N",
        help=(
            "传给 import_temu_fee：仅写入近 N 天（与单独脚本 --write-order-days 一致）。"
            "默认 30；传 0 表示全量、不按天过滤。"
        ),
    )
    args = parser.parse_args()

    only_flags = (args.shipped_only, args.refund_only, args.amz_only, args.temu_only)
    only_count = sum(only_flags)
    if only_count > 1:
        _log.error(
            "--shipped-only / --refund-only / --amz-only / --temu-only 四者最多只能指定一个"
        )
        return 2
    if args.no_temu and args.temu_only:
        _log.error("--no-temu 与 --temu-only 互相矛盾")
        return 2
    if args.no_amz and args.amz_only:
        _log.error("--no-amz 与 --amz-only 互相矛盾")
        return 2
    if args.temu_no_mail and args.temu_always_mail:
        _log.error("--temu-no-mail 与 --temu-always-mail 不能同时指定")
        return 2

    order_dir = args.order_dir or default_order_excel_dir()
    if not order_dir.is_dir():
        _log.error(f"找不到目录：{order_dir}")
        return 2

    amz_dir = args.amz_dir or default_amz_transaction_excel_dir()

    if only_count == 1:
        do_shipped = args.shipped_only
        do_refund = args.refund_only
        do_amz = args.amz_only
        do_temu = args.temu_only
    else:
        do_shipped = True
        do_refund = True
        do_amz = not args.no_amz
        do_temu = not args.no_temu

    # --temu-only 且要写明细时必须有 Excel；仅步骤 2 时可无文件
    if do_temu and args.temu_only and not args.temu_no_detail_table:
        resolved = _resolve_temu_file(order_dir, args.temu_file)
        if resolved is None:
            target = args.temu_file or (order_dir / "TEMU-订单详情.xlsx")
            _log.error(f"--temu-only 但找不到 TEMU 文件：{target}")
            return 2

    # --amz-only 时要求目录存在且能匹配到文件（与单独跑 import_amz_transaction 一致）
    if do_amz and args.amz_only:
        if not amz_dir.is_dir():
            _log.error(f"--amz-only 但目录不存在：{amz_dir}")
            return 2
        if not _discover_amz_transaction(amz_dir):
            _log.error(f"--amz-only 但在 {amz_dir} 未找到 transaction交易明细*.xlsx")
            return 2

    if args.temu_write_order_days < 0:
        _log.error("--temu-write-order-days 须为 >=0 的整数（0=全量）")
        return 2
    temu_wod: int | None = None if args.temu_write_order_days == 0 else args.temu_write_order_days

    run_import(
        order_dir=order_dir,
        amz_dir=amz_dir,
        do_shipped=do_shipped,
        do_refund=do_refund,
        do_amz=do_amz,
        do_temu=do_temu,
        temu_file=args.temu_file,
        temu_send_mail=not args.temu_no_mail,
        temu_always_mail=args.temu_always_mail,
        temu_save_detail=not args.temu_no_detail_table,
        temu_write_order_days=temu_wod,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
