from __future__ import annotations

"""
币种汇率配置加载器（业务侧维护）。

设计原则：
1. 业务人员只需修改同目录的 fx_rates.json，改完保存就生效，不动代码、不动 .env
2. 代码内置一份"出厂默认值"作为 fallback，万一 JSON 文件丢失/被改坏，脚本仍能跑
3. 加载时打印实际使用的汇率，让运行端人工核对（避免汇率配置错却无人察觉）
4. 接口对外暴露 ISO 4217 标准币种码（USD/PLN/HUF/CZK/SEK/RON/CAD/EUR），
   并提供 SYMBOL_TO_ISO 映射表，让脚本能从 TEMU 紫鸟 sheet 的字符串符号
   （€ / zł / Ft / Kč / kr / Lei）反查到 ISO 码。

JSON 文件结构：
{
  "rmb_per_eur": 7.3,          # 1 EUR 兑多少 RMB（脚本计算 RMB→EUR 时用 ÷ 这个数）
  "rates_to_eur": {            # 1 单位外币 = ? EUR
    "USD": 0.8540,
    "PLN": 0.237,
    ...
  }
}

下划线开头的字段（如 _comment / _updated_at）会被自动忽略，仅作为 JSON 文件内的备注。
"""

import json
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Mapping


# ============================================================
# 出厂默认值（与 v1/A0_set_date.py 保持一致）
# 仅当 fx_rates.json 不存在 / 不可解析时使用
# ============================================================
DEFAULT_RMB_PER_EUR = Decimal("7.3")

DEFAULT_RATES_TO_EUR: dict[str, Decimal] = {
    "USD": Decimal("0.8540"),
    "PLN": Decimal("0.237"),
    "HUF": Decimal("0.002611"),
    "CZK": Decimal("0.04133"),
    "SEK": Decimal("0.0934"),
    "RON": Decimal("0.196"),
    "CAD": Decimal("0.6179"),
    "EUR": Decimal("1"),
}


# ============================================================
# 币种符号 → ISO 4217 代码
# 给 TEMU 紫鸟 sheet 用：从 '58,50€' / '12,30 zł' 这种字符串里识别出币种
# 顺序很重要：长符号要排在前面（避免 'kr' 误吃 '€'）
# ============================================================
SYMBOL_TO_ISO: list[tuple[str, str]] = [
    ("€", "EUR"),
    ("zł", "PLN"),
    ("Ft", "HUF"),
    ("Kč", "CZK"),
    ("Lei", "RON"),
    ("kr", "SEK"),  # 注意：在多斯堪的纳维亚币种场景下需细化（DKK/NOK），TEMU 紫鸟现状用 SEK
    ("$", "USD"),
]


# ============================================================
# 数据结构
# ============================================================
@dataclass(frozen=True)
class FxRates:
    """加载后的汇率配置（不可变）。"""
    rmb_per_eur: Decimal
    rates_to_eur: Mapping[str, Decimal]
    source: str = "default"           # "json" | "default" | "json+fallback"
    json_path: Path | None = None
    updated_at: str | None = None     # JSON 文件里的 _updated_at（如果有）
    updated_by: str | None = None     # JSON 文件里的 _updated_by（如果有）
    issues: tuple[str, ...] = field(default_factory=tuple)  # 加载过程中的告警

    def to_eur(self, amount: Decimal | float | int, ccy_iso: str) -> Decimal | None:
        """把 amount（外币）转换为 EUR。币种未知或为 None 返回 None。"""
        if amount is None:
            return None
        rate = self.rates_to_eur.get(ccy_iso)
        if rate is None:
            return None
        try:
            d = amount if isinstance(amount, Decimal) else Decimal(str(amount))
        except InvalidOperation:
            return None
        return (d * rate).quantize(Decimal("0.000001"))

    def rmb_to_eur(self, amount: Decimal | float | int) -> Decimal | None:
        """RMB → EUR：amount ÷ rmb_per_eur。"""
        if amount is None:
            return None
        try:
            d = amount if isinstance(amount, Decimal) else Decimal(str(amount))
        except InvalidOperation:
            return None
        return (d / self.rmb_per_eur).quantize(Decimal("0.000001"))


# ============================================================
# JSON 文件解析
# ============================================================
def default_json_path() -> Path:
    """默认配置文件位置：与本模块同目录的 fx_rates.json。"""
    return Path(__file__).resolve().parent / "fx_rates.json"


def _to_decimal(v: object, name: str, issues: list[str]) -> Decimal | None:
    if v is None:
        return None
    if isinstance(v, bool):  # bool 是 int 的子类，单独排除
        issues.append(f"{name}={v!r} 不是数字")
        return None
    if isinstance(v, (int, float)):
        try:
            return Decimal(str(v))
        except InvalidOperation:
            issues.append(f"{name}={v!r} 无法解析为 Decimal")
            return None
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        try:
            return Decimal(s)
        except InvalidOperation:
            issues.append(f"{name}={v!r} 无法解析为 Decimal")
            return None
    issues.append(f"{name}={v!r} 类型未知（{type(v).__name__}）")
    return None


def load_rates(json_path: Path | None = None) -> FxRates:
    """
    加载汇率：
      1. 默认从同目录 fx_rates.json 读取；可传 json_path 指定其他位置
      2. JSON 中存在的字段覆盖出厂默认值；缺失的字段用默认值兜底（保证脚本永远能跑）
      3. 任何解析失败都记录到 issues 中，不抛异常（让 caller 决定是否容忍）
    """
    path = json_path or default_json_path()
    issues: list[str] = []

    if not path.is_file():
        issues.append(f"配置文件不存在：{path}（使用出厂默认值）")
        return FxRates(
            rmb_per_eur=DEFAULT_RMB_PER_EUR,
            rates_to_eur=dict(DEFAULT_RATES_TO_EUR),
            source="default",
            json_path=path,
            issues=tuple(issues),
        )

    try:
        with path.open("r", encoding="utf-8") as f:
            raw = json.load(f)
    except json.JSONDecodeError as e:
        issues.append(f"配置文件 JSON 格式错误：{e}（使用出厂默认值）")
        return FxRates(
            rmb_per_eur=DEFAULT_RMB_PER_EUR,
            rates_to_eur=dict(DEFAULT_RATES_TO_EUR),
            source="default",
            json_path=path,
            issues=tuple(issues),
        )

    if not isinstance(raw, dict):
        issues.append(f"配置文件根对象必须是 JSON object，实际是 {type(raw).__name__}（使用出厂默认值）")
        return FxRates(
            rmb_per_eur=DEFAULT_RMB_PER_EUR,
            rates_to_eur=dict(DEFAULT_RATES_TO_EUR),
            source="default",
            json_path=path,
            issues=tuple(issues),
        )

    # rmb_per_eur
    rmb = _to_decimal(raw.get("rmb_per_eur"), "rmb_per_eur", issues)
    if rmb is None or rmb <= 0:
        issues.append(f"rmb_per_eur 缺失或非正数，使用默认值 {DEFAULT_RMB_PER_EUR}")
        rmb = DEFAULT_RMB_PER_EUR

    # rates_to_eur：合并到默认值上（部分缺失时其他默认值兜底）
    merged: dict[str, Decimal] = dict(DEFAULT_RATES_TO_EUR)
    rates_in = raw.get("rates_to_eur") or {}
    if not isinstance(rates_in, dict):
        issues.append(f"rates_to_eur 必须是 JSON object，实际是 {type(rates_in).__name__}（忽略，全部用默认值）")
        rates_in = {}
    for ccy, val in rates_in.items():
        if not isinstance(ccy, str) or not ccy.strip():
            issues.append(f"币种代码非法：{ccy!r}（跳过）")
            continue
        d = _to_decimal(val, f"rates_to_eur[{ccy!r}]", issues)
        if d is None:
            continue
        if d <= 0:
            issues.append(f"rates_to_eur[{ccy!r}]={d} 不是正数（跳过）")
            continue
        merged[ccy.strip().upper()] = d

    has_overrides = bool(rates_in) or "rmb_per_eur" in raw
    source = "json" if has_overrides else "default"
    return FxRates(
        rmb_per_eur=rmb,
        rates_to_eur=merged,
        source=source,
        json_path=path,
        updated_at=raw.get("_updated_at") if isinstance(raw.get("_updated_at"), str) else None,
        updated_by=raw.get("_updated_by") if isinstance(raw.get("_updated_by"), str) else None,
        issues=tuple(issues),
    )


def format_summary(fx: FxRates) -> str:
    """生成单行汇率摘要，便于日志打印。"""
    parts = [f"RMB÷{fx.rmb_per_eur}"]
    for ccy in sorted(fx.rates_to_eur.keys()):
        if ccy == "EUR":
            continue
        parts.append(f"{ccy}={fx.rates_to_eur[ccy]}")
    return "  ".join(parts)
