# -*- coding: utf-8 -*-

from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.utils import parseaddr, formataddr
from email import Encoders

from string import Template

import smtplib
import os
import xlrd
import ConfigParser

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( \
        Header(name, 'utf-8').encode(), \
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))


config_file_path = ('./config.ini')
cf = ConfigParser.ConfigParser()
cf.read (config_file_path)
sections = cf.sections()
print sections



# settings
##server
from_addr = cf.get("server","from_addr")
password =  cf.get("server","password")
smtp_server = cf.get("server","smtp_server")
server_port = cf.get("server","server_port")

##mail
header_str = cf.get("mail","header_str")
BAK_DIR = cf.get("mail","BAK_DIR")
TXT_FILE = cf.get("mail","TXT_FILE")
name_flag = cf.getboolean("mail","name_flag")

##receiver
list_file = cf.get("receiver","list_file")


# 获得收件人姓名和地址
list_data  = xlrd.open_workbook(list_file)
list_table = list_data.sheets()[0]
list_nrow = list_table.nrows
list_ary  = []
for i in range(list_nrow):
	# 第一列是姓名 第二列是邮件地址
	tmp_dic ={
		'count':i,
		'to_name':list_table.cell(i,0).value,
		'to_addr':list_table.cell(i,1).value,
	}
	list_ary.append(tmp_dic)

print list_ary

# msg 正文
msg_file = open(TXT_FILE,'r')
try:
     msg_tpl = msg_file.read( )
finally:
     msg_file.close( )


#获得附件，避免之后重复读取，浪费效率
files_cache = []

for filename in os.listdir(BAK_DIR):
	part = MIMEBase('application', "octet-stream")
	part.set_payload(open(os.path.join(BAK_DIR, filename),"rb").read() )
	Encoders.encode_base64(part)
	part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(filename))
	files_cache.append(part)


for single  in list_ary:
	msg = MIMEMultipart()
	# 读取文字
	msg['From'] = _format_addr(u'ASES-Zhejiang <%s>' % from_addr)
	# 若读取错误 则发送到这个这个邮箱:583077757@qq.com(张一舟的QQ邮箱)
	msg['To'] = _format_addr( single.get('to_name','ERROR PERSON NAME') + u' <%s>' % single.get('to_addr','583077757@qq.com'))
	msg['Subject'] = Header(header_str, 'utf-8').encode()

	if name_flag:
		map={'name': single.get('to_name').encode('UTF-8')}
		# temp = NewTemplate(msg_tpl)
		temp = Template(msg_tpl)
		msg_txt = temp.safe_substitute(map)
	else:
		msg_txt = msg_tpl

	msg.attach(MIMEText(msg_txt, 'plain', 'utf-8'))

	for file in files_cache:
		msg.attach(file)


	# OUTLOOK SMTP SETTINGS:https://outlook.live.com/owa/?path=/options/popandimap
	server = smtplib.SMTP(smtp_server, server_port,timeout=120)
	server.set_debuglevel(1)
	# http://stackoverflow.com/questions/19765073/cant-send-email-via-python-using-gmail-smtplib-smtpexception-smtp-auth-extens
	try_count = 0
	try_flag  = False
	while not try_flag and try_count < 5:
		try:
			print '第 %s 次尝试登陆' % (try_count + 1)
			server.starttls()
			server.login(from_addr, password)
			server.sendmail(from_addr,[single.get('to_addr','583077757@qq.com')], msg.as_string())
			server.quit()
			try_flag = True
		except :
			try_count+= 1
			if try_count == 5:
				print  "网络异常，已中断本次发送"


	print "%s: %s_%s is sent!" % (single.get('count','00000') + 1,single.get('to_name','ERROR PERSON').encode('UTF-8'),single.get('to_addr','583077757@qq.com'))