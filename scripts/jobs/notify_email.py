"""
公共邮件通知模块（info / success / warning / error）。

从环境变量读取 SMTP 配置（与 heartbeat.py 相同）：
  SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_STARTTLS, SMTP_SSL,
  MAIL_FROM, MAIL_TO, MAIL_SUBJECT（可选，作为标题前缀）

用法（库）::

    from notify_email import send_notify

    send_notify(
        category="warning",
        subject="Excel 源文件检查失败",
        body="有 2 个步骤缺少源文件，run_batch 执行时可能失败。",
        details=[("模式", "每天"), ("日期", "2026-06-09")],
    )

用法（命令行测试）::

    python scripts/jobs/notify_email.py --category info --subject 测试 --body 你好
"""

from __future__ import annotations

import argparse
import os
import smtplib
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import EmailMessage
from typing import List, Literal, Optional, Sequence, Tuple

NotifyCategory = Literal["info", "success", "warning", "error"]

_CATEGORY_META: dict[NotifyCategory, dict[str, str]] = {
    "info": {
        "label": "信息",
        "badge": "ℹ️ 信息",
        "badge_bg": "#eff6ff",
        "badge_fg": "#1d4ed8",
        "header_from": "#1e3a8a",
        "header_to": "#1d4ed8",
        "subject_prefix": "[信息]",
    },
    "success": {
        "label": "成功",
        "badge": "✅ 成功",
        "badge_bg": "#ecfdf5",
        "badge_fg": "#065f46",
        "header_from": "#064e3b",
        "header_to": "#047857",
        "subject_prefix": "[成功]",
    },
    "warning": {
        "label": "警告",
        "badge": "⚠️ 警告",
        "badge_bg": "#fffbeb",
        "badge_fg": "#b45309",
        "header_from": "#92400e",
        "header_to": "#d97706",
        "subject_prefix": "[警告]",
    },
    "error": {
        "label": "错误",
        "badge": "❌ 错误",
        "badge_bg": "#fef2f2",
        "badge_fg": "#b91c1c",
        "header_from": "#7f1d1d",
        "header_to": "#dc2626",
        "subject_prefix": "[错误]",
    },
}


def _env_bool(name: str, default: bool = False) -> bool:
    v = os.getenv(name)
    if v is None:
        return default
    return v.strip().lower() in {"1", "true", "yes", "y", "on"}


def _env_int(name: str, default: int) -> int:
    v = os.getenv(name)
    if not v:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def _split_emails(value: str) -> List[str]:
    parts: List[str] = []
    for raw in value.replace(";", ",").split(","):
        s = raw.strip()
        if s:
            parts.append(s)
    return parts


def _esc(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


@dataclass(frozen=True)
class SmtpConfig:
    host: str
    port: int
    username: Optional[str]
    password: Optional[str]
    use_starttls: bool
    use_ssl: bool
    timeout_seconds: int = 30


@dataclass(frozen=True)
class MailConfig:
    mail_from: str
    mail_to: List[str]
    subject_prefix: str


@dataclass(frozen=True)
class SendResult:
    ok: bool
    message: str


def load_configs(*, mail_to: Sequence[str] | None = None) -> Tuple[SmtpConfig, MailConfig]:
    """从环境变量加载 SMTP / 邮件配置。"""
    host = os.getenv("SMTP_HOST", "").strip()
    port = _env_int("SMTP_PORT", 587)
    username = os.getenv("SMTP_USER") or None
    password = os.getenv("SMTP_PASS") or None
    use_starttls = _env_bool("SMTP_STARTTLS", True)
    use_ssl = _env_bool("SMTP_SSL", False)

    mail_from = os.getenv("MAIL_FROM", "").strip() or (username or "")
    mail_to_raw = os.getenv("MAIL_TO", "").strip()
    subject_prefix = os.getenv("MAIL_SUBJECT", "").strip() or "rpa-task 通知"

    if not host or not mail_from or (not mail_to_raw and not mail_to):
        raise ValueError("缺少必要配置：需要至少设置 SMTP_HOST、MAIL_FROM、MAIL_TO。")

    recipients = list(mail_to) if mail_to else _split_emails(mail_to_raw)
    if not recipients:
        raise ValueError("MAIL_TO 为空。")

    smtp = SmtpConfig(
        host=host,
        port=port,
        username=username,
        password=password,
        use_starttls=use_starttls,
        use_ssl=use_ssl,
        timeout_seconds=_env_int("SMTP_TIMEOUT", 30),
    )
    mail = MailConfig(mail_from=mail_from, mail_to=recipients, subject_prefix=subject_prefix)
    return smtp, mail


def _build_html(
    *,
    category: NotifyCategory,
    title: str,
    body: str,
    details: Sequence[Tuple[str, str]],
    subtitle: str,
) -> str:
    meta = _CATEGORY_META[category]
    rows = "\n".join(
        f"""
        <tr>
          <td style="padding:10px 12px;color:#6b7280;border-bottom:1px solid #eef2f7;white-space:nowrap;">{_esc(k)}</td>
          <td style="padding:10px 12px;color:#111827;border-bottom:1px solid #eef2f7;word-break:break-word;">{_esc(v)}</td>
        </tr>
        """.strip()
        for k, v in details
    )
    body_html = _esc(body).replace("\n", "<br/>")

    return f"""\
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{_esc(title)}</title>
  </head>
  <body style="margin:0;padding:0;background:#f6f7fb;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,'PingFang SC','Hiragino Sans GB','Microsoft YaHei',sans-serif;">
    <div style="max-width:720px;margin:0 auto;padding:24px;">
      <div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;">
        <div style="padding:16px 18px;background:linear-gradient(135deg,{meta['header_from']},{meta['header_to']});color:#fff;">
          <div style="font-size:16px;font-weight:700;letter-spacing:.2px;">{_esc(title)}</div>
          <div style="margin-top:6px;font-size:12px;opacity:.85;">{_esc(subtitle)}</div>
        </div>

        <div style="padding:16px 18px;">
          <div style="display:inline-block;padding:6px 10px;border-radius:999px;background:{meta['badge_bg']};color:{meta['badge_fg']};font-weight:700;font-size:12px;">
            {meta['badge']}
          </div>

          <div style="margin-top:14px;color:#111827;font-size:14px;line-height:1.7;word-break:break-word;">
            {body_html}
          </div>

          <table style="width:100%;border-collapse:collapse;margin-top:14px;font-size:13px;">
            <tbody>
              {rows}
            </tbody>
          </table>

          <div style="margin-top:14px;color:#6b7280;font-size:12px;line-height:1.6;">
            此邮件由 rpa-task 自动发送，请勿直接回复。
          </div>
        </div>
      </div>

      <div style="text-align:center;color:#9ca3af;font-size:11px;margin-top:12px;">
        Generated by rpa-task · notify_email
      </div>
    </div>
  </body>
</html>
"""


def _build_text(
    *,
    category: NotifyCategory,
    title: str,
    body: str,
    details: Sequence[Tuple[str, str]],
) -> str:
    meta = _CATEGORY_META[category]
    lines = [
        f"【rpa-task】{meta['label']}通知",
        "",
        title,
        "",
        body,
        "",
    ]
    if details:
        lines.append("--- 详情 ---")
        for k, v in details:
            lines.append(f"- {k}：{v}")
    return "\n".join(lines)


def build_message(
    *,
    category: NotifyCategory,
    mail: MailConfig,
    subject: str,
    body: str,
    details: Sequence[Tuple[str, str]] = (),
    subtitle: str = "rpa-task · notify_email",
) -> EmailMessage:
    meta = _CATEGORY_META[category]
    full_subject = f"{meta['subject_prefix']} {mail.subject_prefix} · {subject}".strip()

    msg = EmailMessage()
    msg["From"] = mail.mail_from
    msg["To"] = ", ".join(mail.mail_to)
    msg["Subject"] = full_subject
    msg.set_content(_build_text(category=category, title=full_subject, body=body, details=details))
    msg.add_alternative(
        _build_html(
            category=category,
            title=full_subject,
            body=body,
            details=details,
            subtitle=subtitle,
        ),
        subtype="html",
    )
    return msg


def send_mail(*, smtp: SmtpConfig, msg: EmailMessage) -> None:
    if smtp.use_ssl:
        with smtplib.SMTP_SSL(smtp.host, smtp.port, timeout=smtp.timeout_seconds) as server:
            if smtp.username and smtp.password:
                server.login(smtp.username, smtp.password)
            server.send_message(msg)
        return

    with smtplib.SMTP(smtp.host, smtp.port, timeout=smtp.timeout_seconds) as server:
        server.ehlo()
        if smtp.use_starttls:
            server.starttls()
            server.ehlo()
        if smtp.username and smtp.password:
            server.login(smtp.username, smtp.password)
        server.send_message(msg)


def send_notify(
    *,
    category: NotifyCategory,
    subject: str,
    body: str,
    details: Sequence[Tuple[str, str]] | None = None,
    mail_to: Sequence[str] | None = None,
    subtitle: str = "rpa-task · notify_email",
    log_prefix: str = "[notify]",
) -> SendResult:
    """
    发送分类通知邮件。配置缺失或发送失败时返回 SendResult(ok=False)，不抛异常。
    """
    extra = list(details or [])
    now = datetime.now(timezone.utc).astimezone()
    hostname = os.getenv("HOSTNAME", "").strip()
    extra.extend(
        [
            ("通知类型", _CATEGORY_META[category]["label"]),
            ("触发时间", now.isoformat(timespec="seconds")),
            ("容器", hostname or "-"),
        ]
    )

    try:
        smtp, mail = load_configs(mail_to=mail_to)
        msg = build_message(
            category=category,
            mail=mail,
            subject=subject,
            body=body,
            details=extra,
            subtitle=subtitle,
        )
        send_mail(smtp=smtp, msg=msg)
        result = SendResult(
            ok=True,
            message=f"邮件已发送：to={','.join(mail.mail_to)} subject={msg['Subject']!r}",
        )
        print(f"{log_prefix} {result.message}")
        return result
    except ValueError as e:
        result = SendResult(ok=False, message=f"邮件未发送：{e}")
        print(f"{log_prefix} {result.message}")
        return result
    except Exception as e:
        host = os.getenv("SMTP_HOST", "").strip()
        port = _env_int("SMTP_PORT", 587)
        use_starttls = _env_bool("SMTP_STARTTLS", True)
        use_ssl = _env_bool("SMTP_SSL", False)
        user = os.getenv("SMTP_USER", "").strip()
        result = SendResult(ok=False, message=f"邮件发送失败：{type(e).__name__}: {e}")
        print(f"{log_prefix} {result.message}")
        print(
            f"{log_prefix} 调试信息："
            f" SMTP_HOST={host!r} SMTP_PORT={port} SMTP_USER={user!r}"
            f" SMTP_STARTTLS={use_starttls} SMTP_SSL={use_ssl}"
        )
        return result


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="发送 rpa-task 分类通知邮件")
    ap.add_argument(
        "--category",
        choices=tuple(_CATEGORY_META.keys()),
        default="info",
        help="消息分类（默认 info）",
    )
    ap.add_argument("--subject", required=True, help="邮件主题（不含前缀）")
    ap.add_argument("--body", required=True, help="邮件正文")
    ap.add_argument(
        "--detail",
        action="append",
        default=[],
        metavar="KEY=VALUE",
        help="附加详情行，可重复",
    )
    return ap.parse_args(argv)


def _parse_detail_items(items: Sequence[str]) -> List[Tuple[str, str]]:
    parsed: List[Tuple[str, str]] = []
    for raw in items:
        if "=" not in raw:
            raise ValueError(f"--detail 格式应为 KEY=VALUE，收到：{raw!r}")
        key, value = raw.split("=", 1)
        parsed.append((key.strip(), value.strip()))
    return parsed


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        details = _parse_detail_items(args.detail)
    except ValueError as e:
        print(f"[notify] {e}")
        return 2

    result = send_notify(
        category=args.category,
        subject=args.subject,
        body=args.body,
        details=details,
        subtitle="rpa-task · notify_email CLI",
    )
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
