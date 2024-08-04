#!/usr/bin/python3

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header

# 第三方 SMTP 服务
mail_host = "smtp.163.com"  # 设置服务器
mail_user = "13621547067@163.com"  # 邮箱用户名
mail_pass = "WEOBGOISSPCKCXZY"  # 是邮箱授权口令，不是邮箱登录密码

sender = "13621547067@163.com"  # 发送邮件邮箱
receivers = ["773662313@qq.com"]  # 接收邮件，可添加多个邮箱

message = MIMEMultipart("alternative")
message['From'] = "13621547067@163.com"  # 邮件发信人，也可以自己定义，建议和发件人一致
message['To'] = "773662313@qq.com"  # 邮件收件人，可自己定义，建议和收件人一致

subject = 'Python SMTP 邮件测试'
message['Subject'] = Header(subject, 'utf-8')

# 创建HTML内容
html_content = """
<html>
<head></head>
<body>
    <p>你好！</p>
    <p>这是一封<b style="color: red;">HTML格式的邮件</b>。</p>
</body>
</html>
"""

# 将HTML内容添加到邮件中
html_part = MIMEText(html_content, "html")
message.attach(html_part)

try:
    smtpObj = smtplib.SMTP()
    smtpObj.connect(mail_host, 25)  # SMTP端口号25，pop3端口号110
    smtpObj.login(mail_user, mail_pass)
    smtpObj.sendmail(sender, receivers, message.as_string())
    print("邮件发送成功")
except smtplib.SMTPException:
    print("Error: 无法发送邮件")