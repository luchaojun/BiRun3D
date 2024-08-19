#!/usr/bin/python3

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
import pyodbc

"""
發送郵件給對應的郵件地址
"""
def send_email(html_content):
    # 第三方 SMTP 服务
    mail_host = "ms4.clevo.com.tw"  # 设置服务器
    # mail_user = "13621547067@163.com"  # 邮箱用户名
    mail_user = "kapok_shopfloor@clevo.com.tw"  # 邮箱用户名
    # mail_pass = "WEOBGOISSPCKCXZY"  # 是邮箱授权口令，不是邮箱登录密码
    mail_pass = "Mis@202406"  # 是邮箱授权口令，不是邮箱登录密码
    sender = "kapok_shopfloor@clevo.com.tw"  # 发送邮件邮箱
    receivers = ["sky-lj@clevo.com.tw", "cheyagang@clevo.com.tw", "chaojun_Lu@clevo.com.tw"]  # 接收邮件，可添加多个邮箱
    # receivers = ["chaojun_Lu@clevo.com.tw"]  # 接收邮件，可添加多个邮箱
    message = MIMEMultipart("alternative")
    message['From'] = "kapok_shopfloor@clevo.com.tw"  # 邮件发信人，也可以自己定义，建议和发件人一致
    message['To'] = ','.join(receivers)
    # 邮件收件人，可自己定义，建议和收件人一致
    subject = 'BI RUN 3D 自動化報表'
    message['Subject'] = Header(subject, 'utf-8')

    # 将HTML内容添加到邮件中
    html_part = MIMEText(html_content, "html")
    message.attach(html_part)
    try:
        # smtpObj = smtplib.SMTP_SSL(mail_host,465)
        # smtpObj.connect(mail_host, 465)  # SMTP端口号25，pop3端口号110
        smtpObj = smtplib.SMTP('ms4.clevo.com.tw', 465)  # SMTP服务器地址和端口号
        smtpObj.starttls()
        smtpObj.login(mail_user, mail_pass)
        smtpObj.sendmail(sender, receivers, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException as e:
        print("e="+str(e))
        print("Error: 无法发送邮件")


"""
創建需要展示的頁面HTML
"""
def create_html(data_rows):
    # 创建HTML内容
    html_content_head = """
     <html>
        <head>
            <meta charset="UTF-8">
            <title>机种测试汇总表</title>
            <style>
                th{
                    background: #98f5ff;
                }
                td, th{
                    width: 120px;
                    height: 10px;
                    border: 2px solid #3366ff; /* 设置边框宽度和颜色 */
                }
                table {
                  border-collapse: collapse; /* 合并边框 */
                  text-align: center;
                  border: 2px solid #3366ff; /* 设置边框宽度和颜色 */
                }
            </style>
        </head>
        <body>
            <center>
                <table border="1">
                    <tr>
                        <th>序號</th>
                        <th>機種</th>
                        <th>累計run 3D數量</th>
                        <th>狀態</th>
                    </tr>
    """
    html_content_body = ""
    order_number = 1
    for data_row in data_rows:
        html_content_body += """
            <tr>
                <td>{order_number}</td>
                <td>{device_name}</td>
                <td><a href="http://ksintranet.clevo.com.tw/Shop_Floor/test/ljc/run3dautoreport_index.asp?device_name={device_name}">{count}</a></td>
                <td>{status}</td>
            </tr>
        """.format(order_number=order_number, device_name=data_row[0], count=data_row[1], status='Close' if int(data_row[1])>=5000 else 'Open')
        order_number += 1
    html_content_end = """
            </table>
            <center>
        </body>
        </html>
    """
    return html_content_head + html_content_body + html_content_end


"""
操作數據庫查詢出需要的數據
"""
def operate_db_select_info():
    connectionString = 'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'.format(
        SERVER='BIOS2', DATABASE='SWPRODUCE', USERNAME='sa', PASSWORD='P@ssw0rd')
    print(connectionString)
    conn = pyodbc.connect(connectionString)
    cursor = conn.cursor()
    cursor.execute(
        "select substring(t1.nbsno,3, 7) as device_name, count(t1.i) as count from (select A.nbsno,count(B.item_no) as i from NBSNOCHKLOG A join pd_test_data B on B.mac=A.MAC where B.mac=A.MAC AND B.create_date>'2024-01-01'  AND B.item_no IN ('firestrikecombinedscorep','firestrikegraphicsscorep','firestrikeoverallscorep','firestrikephysicsscorep') group by A.nbsno having count(B.item_no) > 3) AS t1 group by substring(t1.nbsno,3, 7) order by count(t1.i) desc")
    data_rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return data_rows


if __name__ == "__main__":
    data_rows = operate_db_select_info()
    html_content = create_html(data_rows)
    print(html_content)
    send_email(html_content)
