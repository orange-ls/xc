"""
@Desc    : 
@Version : 1.0
@Time    : 2022/05/31 16:49:04
@Author  : 罗小北
@File    : mail_tools.py
@Software: VSCode
"""

import sqlite3
import datetime
from imbox import Imbox


def sqlite_execute(sqlfile, sql_s, sql_params=None):
    con = sqlite3.connect(sqlfile)
    try:
        if sql_params:
            ret = con.execute(sql_s, sql_params)
        else:
            ret = con.execute(sql_s)
        con.commit()
        ret = ret.fetchall()
        con.close()
    except Exception as err:
        con.close()
        raise err
    return ret


def get_mails(server, user, pd, sqlfile):
    # 检测sqlite中数据库是否存在，如果不存在，则创建
    create_sql = "create table email(id INTEGER PRIMARY KEY AUTOINCREMENT,账号 text, 标题 text, 发送时间 text, 正文 text, 项目名称 text, 二级经销商名称 text, 产品信息 text, 是否优选CSP项目 text, 收件人 text, 备注, 是否发送 text)"
    email_exist = "select count(*) from sqlite_master where type='table' and name='email'"
    insert_sql = "insert into email(账号, 标题, 发送时间, 正文) values(?,?,?,?)"

    if not sqlite_execute(sqlfile, email_exist)[0][0]:
        sqlite_execute(sqlfile, create_sql)
    with Imbox(server, user, pd, ssl=True) as imbox:
        all_messages = imbox.messages(unread=True)
        for uid, message in all_messages:
            subject = message.subject
            received_date = datetime.datetime.strptime(message.date[:24],
                                                       "%a, %d %b %Y %H:%M:%S")
            received_date = received_date.strftime('%Y-%m-%d %H:%M:%S')
            body = message.body["html"]
            if len(body) < 1:
                continue
            sqlite_execute(sqlfile, insert_sql,
                           (user, subject, received_date, body[0]))
            # 已读标记，最后再开启
            imbox.mark_seen(uid)


if __name__ == "__main__":
    pass
    get_mails("imap.feishu.cn", "zhengws@digitalchina-hw.com", "LRmLn2AAD2GmQnSj", r'C:\Users\user\Desktop\email.db')