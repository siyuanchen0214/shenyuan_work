#!/usr/bin/env python3
"""
每周任务提醒邮件发送脚本
- 任务配置在 tasks.json，直接修改即可增删任务
- 通过系统 crontab 每天定时触发，当天有任务才发邮件
- 邮件通过阿里云企业邮箱 SMTP 发送
"""

import json
import smtplib
import os
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# ── 配置区（填好后不用再改）──────────────────────────────────────
SMTP_HOST = "smtp.qiye.aliyun.com"
SMTP_PORT = 465
SENDER    = os.environ.get("REMINDER_SENDER", "")   # 发件人邮箱，如 xxx@yourdomain.com
PASSWORD  = os.environ.get("REMINDER_PASSWORD", "") # 邮箱密码或授权码
RECEIVER  = os.environ.get("REMINDER_RECEIVER", "") # 收件人，填你自己的邮箱
# ─────────────────────────────────────────────────────────────────

WEEKDAY_MAP = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

def load_tasks():
    tasks_path = Path(__file__).parent / "tasks.json"
    with open(tasks_path, encoding="utf-8") as f:
        return json.load(f)

def build_email_html(weekday: str, tasks: list) -> str:
    today = date.today().strftime("%Y年%m月%d日")
    items = ""
    for i, task in enumerate(tasks, 1):
        items += f"""
        <tr>
          <td style="padding:12px 16px;border-bottom:1px solid #f0f0f0;font-size:15px;color:#333;">
            {i}. <strong>{task['title']}</strong><br>
            <span style="font-size:13px;color:#666;">{task['detail']}</span>
          </td>
        </tr>"""

    return f"""
    <html><body style="font-family:PingFang SC,sans-serif;background:#f5f5f5;padding:24px;">
      <div style="max-width:560px;margin:auto;background:#fff;border-radius:8px;overflow:hidden;
                  box-shadow:0 2px 8px rgba(0,0,0,.08);">
        <div style="background:#4A90D9;padding:20px 24px;">
          <h2 style="margin:0;color:#fff;font-size:18px;">📋 今日任务提醒</h2>
          <p style="margin:4px 0 0;color:#d0e8ff;font-size:13px;">{today}（{weekday}）</p>
        </div>
        <table style="width:100%;border-collapse:collapse;">{items}</table>
        <div style="padding:16px 24px;background:#fafafa;font-size:12px;color:#aaa;">
          此邮件由自动脚本发送，任务列表见 tasks.json
        </div>
      </div>
    </body></html>
    """

def send_email(subject: str, html_body: str):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = SENDER
    msg["To"]      = RECEIVER
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(SENDER, PASSWORD)
        server.sendmail(SENDER, RECEIVER, msg.as_string())
    print(f"邮件已发送至 {RECEIVER}")

def main():
    if not SENDER or not PASSWORD or not RECEIVER:
        print("错误：请先设置环境变量 REMINDER_SENDER / REMINDER_PASSWORD / REMINDER_RECEIVER")
        return

    weekday = WEEKDAY_MAP[date.today().weekday()]
    tasks   = load_tasks().get(weekday, [])

    if not tasks:
        print(f"今天是{weekday}，没有待办任务，跳过发送。")
        return

    subject   = f"【任务提醒】{weekday} 有 {len(tasks)} 件事要做"
    html_body = build_email_html(weekday, tasks)
    send_email(subject, html_body)

if __name__ == "__main__":
    main()
