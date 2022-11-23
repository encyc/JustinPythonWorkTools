# -*- coding: utf-8 -*-
"""
Created on Fri Nov  4 09:48:01 2022

@author: Administrator
"""


# In[]:

import smtplib
from email.header import Header
from email.utils import formataddr
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart


# In[]:

# 发件人邮箱账号
my_sender='gaoyuan@zhizu001.com'  
# user登录邮箱的用户名，password登录邮箱的密码（授权码）
my_pass = 'kbz89F44JvMW5mvD' 
# 收件人邮箱账号

receiver_list = ['zhangkai@zhizu001.com','sunzheyi@zhizu001.com','chenqingxiang@zhizu001.com','gyuanpro@163.com']
#'zhangkai@zhizu001.com'
#'1085484391@qq.com'
#'gyuanpro@163.com'
#'sunzheyi@zhizu001.com'
# In[]:
    


def Multi_mail(receiver,excel_df):
    ret=True
    try:
        # 邮件内容

        # 创建多形式组合邮件
        msg = MIMEMultipart('mixed')
        msg['From'] = formataddr(['Justin高原', my_sender])  # 发件人邮箱昵称、发件人邮箱账号
        msg['To'] = formataddr(['',receiver])  # 收件人邮箱昵称、收件人邮箱账号
        # msg['To'] = ','.join(receivers)
        msg['Subject'] = "【风控】-近一个月各机构查询数量"  # 邮件主题
        
        #写邮件，读取excel文件内容作为邮件正文
        def mailWrite():
            #表格的标题和头
            header = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head>'
            th = '<body text="#000000" ><table border="1" cellspacing="0" cellpadding="3" bordercolor="#000000" width="350" align="left" ><tr bgcolor="#F79646" align="left" ><th>查询机构</th><th>是否查得</th><th>数量</th></tr>'
            #打开文件
            #filepath设置详细的文件地址
            #filepath = file_name_query
            #book = xlrd.open_workbook(filepath)
            #sheet = book.sheet_by_index(0)
            sheet = excel_df
            #获取行列的数目，并以此为范围遍历获取单元数据
            #nrows 行数，ncols 列数
            nrows = len(sheet.index)
            ncols = len(sheet.columns)
            body = ''
            cellData = 1
            for i in range(0,nrows):
                td = ''
                for j in range(0,ncols):
                    cellData = sheet.iloc[i,j]
            
                    #读取单元格数据，赋给cellData变量供写入HTML表格中
                    tip = '<td>' + str(cellData) + '</td>'
                    td = td + tip
                    tr = '<tr>' + td + '</tr>'
                    #tr = tr.encode('utf-8')
                body = body + tr
                tail = '</table></body></html>'
                mailcontent = header+th+body+tail
            #将excel文件的内容转换为html格式，后续在邮件中拼接
            return mailcontent
        
        #邮件正文内容
        content = mailWrite()
        cs = """
        <p>上月各机构查询数量如下：</p>
        """
        contents = cs + content
        html = MIMEText(contents,'html','utf-8')
        msg.attach(html)
        
        
        ######## 添加附件 - excel
        att_excel1 = MIMEText(open(file_name_query, 'rb').read(),'base64','utf-8')
        att_excel1["Content-Type"] = 'application/octet-stream'
        att_excel1["Content-Disposition"] = 'attachment; filename="QueryNum_%s.xlsx"' %(T)  
        msg.attach(att_excel1)

         
        # SMTP服务器，腾讯企业邮箱端口是465，腾讯邮箱支持SSL(不强制)， 不支持TLS
        # qq邮箱smtp服务器地址:smtp.qq.com,端口号：456
        # 163邮箱smtp服务器地址：smtp.163.com，端口号：25
        server=smtplib.SMTP_SSL("smtp.exmail.qq.com", 465)  
        # 登录服务器，括号中对应的是发件人邮箱账号、邮箱密码
        server.login(my_sender, my_pass)  
        # 发送邮件，括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.sendmail(my_sender,[receiver,],msg.as_string())  
        # 关闭连接
        server.quit() 
        # 如果 try 中的语句没有执行，则会执行下面的 ret=False 
    except Exception:  
        ret=False
    return ret


for i in receiver_list:
    print(i)
    ret=Multi_mail(i)
    if ret:
        print("邮件发送成功")
    else:
        print("邮件发送失败")
