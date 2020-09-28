#!/usr/bin/env python3
# -*- coding:UTF-8 -*-
#
# Author:   @saeziae
# Lisense:  GPL3
#
'''
    Auto E-mail Homework Collecter
    Copyright (C) 2020 Saeziae

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
'''
from email.header import Header
from email.mime.text import MIMEText
import email
import smtplib
import imaplib
import json
print(
'''
--------------------------------
 Auto E-mail Homework Collecter
 Author:  @saeziae
 Lisense: GPL3

 Copyright (C) 2020 Saeziae
 This program comes with ABSO-
 LUTELY NO WARRANTY; This is 
 free software, and you are 
 welcome to redistribute it
 under certain conditions.
-------------------------------
''')
PASS, MAIL, IMAPSERVER, SMTPSERVER, SIGN = None,None,None,None,None,
with open("config.json","r",encoding="utf-8") as c:
    config=json.loads(c.read())
    PASS = config["PASS"]
    MAIL = config["MAIL"]
    IMAPSERVER = config["IMAPSERVER"]
    SMTPSERVER = config["SMTPSERVER"]
    SIGN = config["SIGN"]

print("MAIL: ",MAIL)
print("IMAPSERVER: ",IMAPSERVER)
print("SMTPSERVER: ",SMTPSERVER)

conn = imaplib.IMAP4_SSL(IMAPSERVER, "993")  # 走ssl的
smtpObj = smtplib.SMTP()
try:
    conn.login(MAIL, PASS)
    smtpObj.connect(SMTPSERVER, 25)
    smtpObj.login(MAIL, PASS)
except Exception as e:
    print('Error: %s' % str(e.args[0], "utf-8"))
    exit()

print("CONNECTED!")

print("Fetching mail list")
conn.select()
tipo, data = conn.search(None, 'ALL')

workid = input("Today's Work's ID:")  # 作业ID

# Get Student List
students = []
with open("student.txt", "r", encoding="utf-8") as f:
    students = f.readlines()
students = map(lambda x: x.rstrip("\r\n").split(" "), students)  # 按照"学号 名字"处理
students = map(lambda x: [x[0], [x[1], False]], students)       # 添加是否提交的判断
students = dict(students)

newlist = data[0].split()[::-1]  # 列表化收到的邮件，时间倒序
for i in newlist:
    print("Fetching mail %d" % int(i))
    tipo, data = conn.fetch(i, '(RFC822)')
    print("Analysing mail %d" % int(i))
    msg = email.message_from_string(data[0][1].decode('utf-8'))
    sub = email.header.decode_header(msg.get('subject'))[0][0]  # 标题
    if type(sub) == bytes:
        sub = sub.decode('utf-8')
    sub_0 = sub.split("/")  # 标题格式 学号/作业号
    if len(sub_0) == 2:
        if sub_0[1] != workid:  # 判断是不是本次作业
            continue
    else:
        continue
    pass
    stu = sub_0[0]  # 前半截是学号
    print("[SUBJECT] ", sub)
    mailfrom = email.utils.parseaddr(msg.get('from'))[1]  # 发件人,回信要用
    print("[   FROM] ", mailfrom)
    foo = 0  # 多文件防重名
    for part in msg.walk():
        if not part.is_multipart():
            name = part.get_param("name")  # 附件的文件名
            if name:
                print("[ ATTACH] ", name)
                attach_data = part.get_payload(decode=True)  # 附件数据
                filename = stu + students[stu][0]
                with open(workid + "/" + filename + (str(foo) if foo else "") + "." + name.split(".")[-1], 'wb') as f:
                    f.write(attach_data)
                    foo += 1
            else:
                # 不是附件
                pass
    students[stu][1] = True  # 标记提交
    # 保存
    mail_msg = "<p>" + \
        students[stu][0] + \
        "同学：<br>你的作业"+workid + \
        "已接收。<br>生活愉快！</p><hr>"+SIGN
    message = email.mime.text.MIMEText(mail_msg, 'html', 'utf-8')
    message['Subject'] = Header('Re:'+sub, 'utf-8')
    message['From'] = Header(MAIL, 'utf-8')
    message['To'] = Header(students[stu][0], 'utf-8')

    try:
        smtpObj.sendmail(MAIL, [mailfrom], message.as_string())  # 发送回信
        print("Reply Sent")
    except smtplib.SMTPException:
        print("Error: Reply not sent")

for i, j in students.items():  # 输出提交情况
    print(i, "|", j[0], "\t|", "submitted" if j[1] else "NOT submitted")

conn.logout()
