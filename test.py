import os, smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
load_dotenv()

msg = MIMEText("这是测试正文", "plain", "utf-8")
msg["Subject"] = "SMTP 测试"
msg["From"] = os.environ["SMTP_USER"]
msg["To"] = os.environ["SMTP_USER"]   # 先发给自己

with smtplib.SMTP_SSL(os.environ["SMTP_HOST"], int(os.environ["SMTP_PORT"])) as s:
    s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
    s.send_message(msg)
print("发送成功")