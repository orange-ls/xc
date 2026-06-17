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
            if isinstance(sql_params[0], (list, tuple)):
                # 批量插入
                ret = con.executemany(sql_s, sql_params)
            else:
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
        mail_list = []
        messages_iter = iter(all_messages)
        while True:
            try:
                uid, message = next(messages_iter)
                subject = message.subject
                received_date = datetime.datetime.strptime(message.date[:24],
                                                           "%a, %d %b %Y %H:%M:%S")
                received_date_str = received_date.strftime('%Y-%m-%d %H:%M:%S')
                body = message.body["html"]
            except StopIteration:
                break
            except Exception as e:
                print(f"跳过一封邮件，解析失败: {e}")
                continue

            # 安全获取正文内容：body可能是list[str]、bytes或str
            if isinstance(body, list):
                if len(body) < 1:
                    continue
                body_text = body[0]
            elif isinstance(body, bytes):
                body_text = body.decode('utf-8', errors='replace')
            elif isinstance(body, str):
                body_text = body
            else:
                continue
            if not isinstance(body_text, str):
                body_text = str(body_text)
            mail_list.append((user, subject, received_date_str, body_text, received_date, uid))

        # 批量标记已读
        for uid in [item[5] for item in mail_list]:
            imbox.mark_seen(uid)

        # 按时间新旧顺序排序（旧->新）
        mail_list.sort(key=lambda x: x[4], reverse=True)  # 按received_date排序（新->旧）

        # 批量插入数据
        batch_size = 100
        for i in range(0, len(mail_list), batch_size):
            batch = mail_list[i:i + batch_size]
            batch_data = [(item[0], item[1], item[2], item[3]) for item in batch]
            sqlite_execute(sqlfile, insert_sql, batch_data)


if __name__ == "__main__":
    pass
    get_mails("imap.feishu.cn", "huaweirpa-hefei@digitalchina.com", "oCldgBwjN1USCsJ3", r'D:\Workspace\xc\RPA\func_codes\swtzd\email.db')