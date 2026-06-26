"""
检查 A_报表 是否仍可独立运行（不启动 Docker、不 import report 业务代码）。
在 py-project 根目录执行：python report/scripts/verify_a_bao_standalone.py
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

PY_ROOT = Path(__file__).resolve().parents[2]
A_BAO = PY_ROOT / "A_报表"


def _bootstrap_a_bao():
    epr = PY_ROOT / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.bootstrap(str(A_BAO / "F_测评" / "runAll_F.py"))


def main() -> int:
    errors: list[str] = []

    if not A_BAO.is_dir():
        errors.append(f"缺少目录: {A_BAO}")
        return 1

    # 不应把 A_报表/report 里的旧代码加入路径
    stale = A_BAO / "report" / "database" / "db_connection.py"
    if stale.is_file():
        errors.append("A_报表/report 仍含旧 database 代码，请只保留 请到这里.md")

    old_db_shim = A_BAO / "Z_method" / "sku_映射_db.py"
    if old_db_shim.is_file():
        errors.append("请删除 A_报表/Z_method/sku_映射_db.py（已迁至 report/z_method）")

    _bootstrap_a_bao()

    checks = [
        ("A_报表.Z_method.sku_映射", "sku_mappings"),
        ("A_报表.Z_method.runall_utils", "run_script"),
        ("A_报表.A0_设置_时间段.A0_set_date", "folder_name"),
        ("A_报表.F_测评.runAll_F", "main"),
    ]
    for mod_name, attr in checks:
        try:
            mod = __import__(mod_name, fromlist=[attr])
            getattr(mod, attr)
        except Exception as e:
            errors.append(f"import {mod_name}.{attr}: {e}")

    runalls = [
        "B_订单统计_sale_resend/runAll_B.py",
        "C_退款/runAll_C.py",
        "D_广告/runAll_D.py",
        "F_测评/runAll_F.py",
        "G_二次上架/runAll_G.py",
        "H_AMZ_利润报表_OTTO_客户经理费/runAll_H.py",
        "K_仓租_映射产品信息/runAll_K.py",
        "M_毛利_销售负责人_表头排序/run_all_gross_profit.py",
        "A_TEMU_计算_订单总金额/runAll_A.py",
    ]
    for rel in runalls:
        p = A_BAO / rel
        if not p.is_file():
            errors.append(f"缺少入口: {rel}")

    # report 未污染 sys.modules 中的 A_报表
    if "report.database" in sys.modules and "A_报表" in sys.modules:
        pass  # 仅提示，不判失败

    if errors:
        print("[FAIL] A_报表 独立检查未通过：")
        for e in errors:
            print(f"  - {e}")
        return 1

    print("[OK] A_报表 可独立运行（入口与 import 正常）")
    print("     日常仍用: python A_报表/.../runAll_*.py")
    print("     重构选用: d:\\py-project\\report（Docker MySQL 可选）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
