from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import smtplib
from email import encoders
from time import gmtime,strftime
from pytz import timezone
from datetime import datetime
import time

India = timezone('Asia/Kolkata')
ind_time = datetime.now(India)
 
def setup_mail(sender,recevier,subject,message,filename,cc,bcc):
	# create message object instance
	msg = MIMEMultipart()
	 
	# setup the parameters of the message
	password = "7698073773Nj@"
	msg['From'] = sender
	msg['To'] = recevier
	msg['Subject'] = subject
	msg["Cc"] = cc
	msg["Bcc"] = bcc
	# add in the message body
	msg.attach(MIMEText(message, 'plain'))
	 
	# open the file to be sent  
	attachment = open(filename, "rb")
	 
	# instance of MIMEBase and named as p
	p = MIMEBase('application', 'octet-stream')
	 
	# To change the payload into encoded form
	p.set_payload((attachment).read())
	 
	# encode into base64
	encoders.encode_base64(p)
	  
	p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
	 
	# attach the instance 'p' to instance 'msg'
	msg.attach(p)
	 
	#create server
	server = smtplib.SMTP('smtp.outlook.com: 587')
	 
	server.starttls()
	 
	# Login Credentials for sending the mail
	server.login(msg['From'], password)
	 
	 
	# send the message via the server.
	server.sendmail(msg['From'], msg['To'].split(",")+msg['Cc'].split(",")+msg['Bcc'].split(",") ,msg.as_string())
	 
	server.quit()

	 
	print "successfully sent email to %s:" % (msg['To'])

date1 = ind_time.strftime("%d_%m_%Y")
sender = "neel.jotani@streebo.com"
receiver = "vikash.tiwary@vlinkinfo.com,abhishekks@bluestarindia.com"
cc = "dhvanil.desai@streebo.com,khushboo.kejriwal@streebo.com,jay.oza@streebo.com"
bcc = "neel.jotani@streebo.com"
subject = "Zabbix Hourly Report"
filename="Bluesr_New_Prod_Server_hourly_report_"+date1+".xlsx"
message = "Hello Abhishek,\n\nI am sending you Zabbix Report that we are monitoring Every Hour.\nThis is the Report with all new Fresh data on Production Server.\nCurrently no issue has been identified in Zabbix monitoring.\n\nPlease feel free to reach out in case of any queries.\n\nRegards\nNeel Jotani"
setup_mail(sender,receiver,subject,message,filename,cc,bcc)

