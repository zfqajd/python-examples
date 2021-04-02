#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time     : 2021/4/2 11:06
# @Author   : Xo9
# @FileName : SendMail.py
# @Blog     : https://zfqajd.github.io

import os
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart


class SendMail:
    def __init__(self):
        self.SMTP_SERVER = 'smtp.126.com'
        self.SMTP_PORT = '465'
        self.SENDER = ''
        self.PASSWORD = ''

    def send_mail(self, subject, message="", recipients='XXXXX@qq.com', img=None, attachment=None):
        msg = MIMEMultipart("related")
        msg["From"] = self.SENDER
        msg["To"] = recipients
        msg["Subject"] = subject

        if img is not None:
            if type(img) is not list:
                img = img.split(';')
            
            for one_img in img:
                with open(one_img, 'rb') as f:
                    img_data = MIMEImage(f.read())
                img_data.add_header("Content-ID", "<%s>" % os.path.basename(one_img))
                msg.attach(img_data)
                message = message + """<p><img src="cid:%s"></p>""" % os.path.basename(one_img)

        msg.attach(MIMEText(message, "html"))

        if attachment is not None:
            if type(attachment) is not list:
                attachment = attachment.split(';')

            for one_attachment in attachment:
                with open(one_attachment, 'rb') as f:
                    file = MIMEApplication(f.read(), name=os.path.basename(one_attachment))

                file['Content-Disposition'] = f'attachment; filename="{os.path.basename(one_attachment)}"'
                msg.attach(file)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(self.SMTP_SERVER, self.SMTP_PORT, context=context) as server:
            server.login(self.SENDER, self.PASSWORD)
            server.sendmail(self.SENDER, recipients.split(','), msg.as_string())


# if __name__ == "__main__":
#    mail = SendMail()
#    mail.send_mail(sys.argv[2], sys.argv[3], sys.argv[1])
#    receiver = "XXXXX@qq.com,XXXXX@qq.com"
#    mail.send_mail("附件及图片测试", "图片如下", receiver, attachment='D:\\Desktop_bak\\图\\3.png;D:\\Desktop_bak\\图\\2.png')
