# coding=utf-8
__author__ = 'Administrator'
import pymysql
import xlwt
import xlrd
from datetime import date, datetime
import time
import uniout
import  smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

file_path = "C:/Users/Administrator/Desktop/debt_" + datetime.strftime(datetime.now(), '%Y%m%d_%H-%M-%S') + ".xls"


def execute_sql(dbargs):
    con = pymysql.connect('xxx', 'admin', passwd',
                          'db', port, charset='utf8')
    cur = con.cursor()
    sql = """
        SELECT
            p.id,
            p.userId,
            p. NAME,
            p.reality_name,
            p.originAmount,
            p.match_avaliable_amount,
            p.title,
            p.time
        FROM
            (
                SELECT
                    a.id,
                    a.originAmount,
                    b.match_avaliable_amount,
                    a.userId,
                    c. NAME,
                    c.reality_name,
                    a.time,
                    b.title
                FROM
                    t_invest_redeed AS a
                LEFT JOIN t_bill_invests AS b ON a.originInvestId = b.invest_id
                LEFT JOIN t_users AS c ON a.userId = b.user_id
                AND b.user_id = c.id
                WHERE
                    a.zero_type = 2
                AND a.match_status = 1
                AND b.match_avaliable_amount > 0
                AND a.userId NOT IN (204199, 204861, 63429)
                GROUP BY
                    a.id
            ) p
        WHERE
            p.title = %s  
        ORDER BY
            p.time ASC
        """

    cur.execute(sql, dbargs)
    count = cur.execute(sql, dbargs)
    result = cur.fetchall()
    print("一共有 " + str(count) + " 条记录")
    return result
    cur.close()
    con.close()


def write_excel(wbk, result, sheet_name):
    worksheet = wbk.add_sheet(sheet_name, cell_overwrite_ok=True)

    fileds = ['id', 'userId', 'name', 'reality_name', 'originAmount', 'match_avaliable_amount', 'title', 'time']

    datestyle = xlwt.XFStyle()
    datestyle.num_format_str = 'YYYY-MM-DD HH:MM:SS'

    style1 = xlwt.XFStyle()
    font1 = xlwt.Font()
    font1.bold = True
    style1.font = font1
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    style1.alignment = alignment

    for i in range(0, len(fileds)):
        worksheet.write(0, i, fileds[i])
        worksheet.write(0, i, fileds[i], style1)

    for row in range(1, len(result) + 1):
        for col in range(0, len(fileds)):
            worksheet.write(row, col, result[row - 1][col])
        worksheet.write(row, col, result[row - 1][7], datestyle)

def send_email():
    # 发送邮件服务器
    smtpserver = 'smtp.exmail.qq.com'
    # 发送邮箱用户名和授权码
    user = 'xxx@xgqq.com'
    password = 'xxx'
    # 发送邮箱
    sender = 'xxx@xgqq.com'
    # 接受邮箱
    receiver = 'xxx@xgqq.com'

    # 创建一个带附件的实例
    message = MIMEMultipart()
    message['From'] = user
    message['To'] = receiver
    subject = '债转数据'
    message['Subject'] = Header(subject, 'utf-8')


    # 邮件正文内容
    message.attach(MIMEText(u'最新债转数据,请查收!', 'plain', 'utf-8'))

    # 构造附件，传送当前目录下的文件
    att1 = MIMEText(open(file_path, 'rb').read(), 'base64', 'utf-8')
    att1['Content-Type'] = 'application/octet-stream'
    att1['Content-Disposition'] = 'attachment;filename="debt.xls"'
    message.attach(att1)
    try:
        smtp_mail = smtplib.SMTP_SSL()
        smtp_mail.connect(smtpserver,465)
        smtp_mail.login(user, password)
        smtp_mail.sendmail(sender, receiver, message.as_string())
        print "邮件发送成功!"
    except smtplib.SMTPException,e:
        print "Error:无法发送邮件!"
    finally:
        smtp_mail.quit()


if __name__ == '__main__':
    #file_path = "C:/Users/Administrator/Desktop/debt_" + datetime.strftime(datetime.now(), '%Y%m%d_%H-%M-%S') + ".xls"
    wbk = xlwt.Workbook(encoding='utf-8')
    write_excel(wbk, execute_sql('7天梦想储蓄罐'), u'7天梦想储蓄罐')
    write_excel(wbk, execute_sql('30天梦想储蓄罐'), u'30天梦想储蓄罐')
    write_excel(wbk, execute_sql('90天梦想储蓄罐'), u'90天梦想储蓄罐')
    write_excel(wbk, execute_sql('180天梦想储蓄罐'), u'180天梦想储蓄罐')
    wbk.save(file_path)
    time.sleep(2)
    send_email()
    print file_path
