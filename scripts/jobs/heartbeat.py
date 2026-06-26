"""
心跳邮件（用于验证 cron 定时任务是否正常）

这个脚本会从环境变量读取 SMTP 配置，并发送一封简单通知邮件。
它被 `docker/python/crontab` 定时执行，输出会进入容器日志，便于排障。
"""

from __future__ import annotations

import os
import smtplib
from datetime import datetime, timezone
from email.message import EmailMessage
from dataclasses import dataclass
from typing import List, Optional, Tuple


def _env_bool(name: str, default: bool = False) -> bool:
    """把环境变量解析为布尔值。"""
    v = os.getenv(name)
    if v is None:
        return default
    return v.strip().lower() in {"1", "true", "yes", "y", "on"}


def _env_int(name: str, default: int) -> int:
    """把环境变量解析为整数；解析失败则使用默认值。"""
    v = os.getenv(name)
    if not v:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def _split_emails(value: str) -> List[str]:
    """把收件人字符串拆成列表；支持逗号/分号分隔。"""
    parts: List[str] = []
    for raw in value.replace(";", ",").split(","):
        s = raw.strip()
        if s:
            parts.append(s)
    return parts


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
    subject: str


def load_configs() -> Tuple[SmtpConfig, MailConfig]:
    """
    从环境变量加载 SMTP / 邮件配置。

    必填：
    - SMTP_HOST
    - MAIL_FROM
    - MAIL_TO（可多个，逗号分隔）
    """
    host = os.getenv("SMTP_HOST", "").strip()
    port = _env_int("SMTP_PORT", 587)
    username = os.getenv("SMTP_USER") or None
    password = os.getenv("SMTP_PASS") or None
    use_starttls = _env_bool("SMTP_STARTTLS", True)
    use_ssl = _env_bool("SMTP_SSL", False)

    mail_from = os.getenv("MAIL_FROM", "").strip() or (username or "")
    mail_to_raw = os.getenv("MAIL_TO", "").strip()
    subject = os.getenv("MAIL_SUBJECT", "").strip() or "rpa-task 定时任务通知"

    if not host or not mail_from or not mail_to_raw:
        raise ValueError("缺少必要配置：需要至少设置 SMTP_HOST、MAIL_FROM、MAIL_TO。")

    mail_to = _split_emails(mail_to_raw)
    if not mail_to:
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
    mail = MailConfig(mail_from=mail_from, mail_to=mail_to, subject=subject)
    return smtp, mail


def build_message(*, mail: MailConfig, body: str) -> EmailMessage:
    """构造邮件内容（纯文本 + HTML）。"""
    msg = EmailMessage()
    msg["From"] = mail.mail_from
    msg["To"] = ", ".join(mail.mail_to)
    msg["Subject"] = mail.subject
    # 纯文本版本：兼容所有客户端
    msg.set_content(body)
    return msg


def _build_html(*, title: str, lines: List[Tuple[str, str]]) -> str:
    """构造一个简洁的 HTML 邮件正文（尽量兼容常见邮箱客户端）。"""
    def esc(s: str) -> str:
        return (
            s.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&#39;")
        )

    rows = "\n".join(
        f"""
        <tr>
          <td style="padding:10px 12px;color:#6b7280;border-bottom:1px solid #eef2f7;white-space:nowrap;">{esc(k)}</td>
          <td style="padding:10px 12px;color:#111827;border-bottom:1px solid #eef2f7;word-break:break-word;">{esc(v)}</td>
        </tr>
        """.strip()
        for k, v in lines
    )

    return f"""\
<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{esc(title)}</title>
  </head>
  <body style="margin:0;padding:0;background:#f6f7fb;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,'PingFang SC','Hiragino Sans GB','Microsoft YaHei',sans-serif;">
    <div style="max-width:720px;margin:0 auto;padding:24px;">
      <div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;overflow:hidden;">
        <div style="padding:16px 18px;background:linear-gradient(135deg,#111827,#1f2937);color:#fff;">
          <div style="font-size:16px;font-weight:700;letter-spacing:.2px;">{esc(title)}</div>
          <div style="margin-top:6px;font-size:12px;opacity:.85;">rpa-task · cron heartbeat</div>
        </div>

        <div style="padding:16px 18px;">
          <div style="display:inline-block;padding:6px 10px;border-radius:999px;background:#ecfdf5;color:#065f46;font-weight:700;font-size:12px;">
            ✅ 心跳正常（任务已触发）
          </div>

          <table style="width:100%;border-collapse:collapse;margin-top:14px;font-size:13px;">
            <tbody>
              {rows}
            </tbody>
          </table>

          <div style="margin-top:14px;color:#6b7280;font-size:12px;line-height:1.6;">
            说明：此邮件由容器内 cron 定时发送，用于验证定时任务与邮件通道是否正常。若你不需要，请关闭或调整 `docker/python/crontab`。
          </div>
        </div>
      </div>

      <div style="text-align:center;color:#9ca3af;font-size:11px;margin-top:12px;">
        Generated by rpa-task
      </div>
    </div>
  </body>
</html>
"""


def send_mail(
    *,
    smtp: SmtpConfig,
    msg: EmailMessage,
) -> None:
    """
    发送邮件（支持三种常见组合）：
    - 587 + STARTTLS（use_starttls=True, use_ssl=False）最常见
    - 465 + SSL（use_ssl=True）
    - 25/587 明文（两者都 False，不推荐公网）
    """
    if smtp.use_ssl:
        with smtplib.SMTP_SSL(smtp.host, smtp.port, timeout=smtp.timeout_seconds) as server:
            if smtp.username and smtp.password:
                server.login(smtp.username, smtp.password)
            server.send_message(msg)
        return

    with smtplib.SMTP(smtp.host, smtp.port, timeout=smtp.timeout_seconds) as server:
        # 有些服务商要求先 EHLO 再 STARTTLS
        server.ehlo()
        if smtp.use_starttls:
            server.starttls()
            server.ehlo()
        if smtp.username and smtp.password:
            server.login(smtp.username, smtp.password)
        server.send_message(msg)


def main() -> None:
    now = datetime.now(timezone.utc).astimezone()
    hostname = os.getenv("HOSTNAME", "")
    local_time = now.isoformat(timespec="seconds")

    try:
        smtp, mail = load_configs()
        # 纯文本（兜底）+ HTML（美化）
        text_body = (
            "【rpa-task】cron 心跳通知\n\n"
            "说明：此邮件由容器内 cron 定时发送，用于验证定时任务与邮件通道是否正常。\n\n"
            f"- 触发时间：{local_time}\n"
            f"- 容器：{hostname}\n"
            f"- SMTP：{smtp.host}:{smtp.port} (ssl={smtp.use_ssl}, starttls={smtp.use_starttls})\n"
        )
        msg = build_message(mail=mail, body=text_body)
        msg.add_alternative(
            _build_html(
                title=mail.subject,
                lines=[
                    ("触发时间", local_time),
                    ("容器", hostname or "-"),
                    ("SMTP Host", smtp.host),
                    ("SMTP Port", str(smtp.port)),
                    ("SMTP SSL", "true" if smtp.use_ssl else "false"),
                    ("SMTP STARTTLS", "true" if smtp.use_starttls else "false"),
                    ("收件人", ", ".join(mail.mail_to)),
                ],
            ),
            subtype="html",
        )

        send_mail(
            smtp=smtp,
            msg=msg,
        )
        print(f"[cron] 邮件已发送：to={','.join(mail.mail_to)} subject={mail.subject!r}")
    except ValueError as e:
        # 配置缺失/格式错误：给出人类可读提示
        print(f"[cron] 邮件未发送：{e}")
    except Exception as e:
        # 连接被对方关闭是最常见的一类问题（端口/TLS/SSL 不匹配、被运营商拦截等）
        # 这里把关键配置打印出来，便于定位（不会输出密码）。
        host = os.getenv("SMTP_HOST", "").strip()
        port = _env_int("SMTP_PORT", 587)
        use_starttls = _env_bool("SMTP_STARTTLS", True)
        use_ssl = _env_bool("SMTP_SSL", False)
        user = os.getenv("SMTP_USER", "").strip()

        print(f"[cron] 邮件发送失败：{type(e).__name__}: {e}")
        print(
            "[cron] 调试信息："
            f" SMTP_HOST={host!r} SMTP_PORT={port} SMTP_USER={user!r}"
            f" SMTP_STARTTLS={use_starttls} SMTP_SSL={use_ssl}"
        )
        print(
            "[cron] 常见组合："
            " 587+STARTTLS(SSL=false, STARTTLS=true) 或 465+SSL(SSL=true)。"
        )


if __name__ == "__main__":
    main()

