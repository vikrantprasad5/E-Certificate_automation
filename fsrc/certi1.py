import xlrd
import smtplib
import cv2
import os
from PIL import Image
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

workbook = xlrd.open_workbook("computer.xls")

worksheet = workbook.sheet_by_index(0)
a = worksheet.nrows
b = worksheet.ncols

print(a)
print(b)
for i in range(1,a):

	#to read from excel file	
	name1 = worksheet.cell(i,1).value
	group_name = worksheet.cell(i,2).value
	email_id = worksheet.cell(i,3).value
	attend1 = worksheet.cell(i,5).value
	post1 = worksheet.cell(i,6).value

	attend = str(attend1)
	post = str(post1)

	if post == '':

		if attend == 'P':
			name = str(name1)
			name2 = name.split()
			firstname = name2[0]
			print(name)


		#to make copy of original certificate
			os.system('sudo cp sample_in.jpg /home/roti/Desktop/certifiacte_testing/{}.jpg'.format(firstname))

		#to wrtie name on certificate
			#for writing name
			im =cv2.imread("sample_in.jpg",1)
			font = cv2.FONT_ITALIC
			font1 = cv2.FONT_HERSHEY_SCRIPT_SIMPLEX
			x=1750-(len(name)*25)
			cv2.putText(im, name, (x,1300), font1, 3, (0, 0, 250), 2, cv2.LINE_AA)
			cv2.imwrite('{}.jpg'.format(firstname), im)

			#for writing group name
			x=1750-(len(group_name)*25)
			cv2.putText(im, group_name, (x,1550), font, 3, (0, 0, 50), 2, cv2.LINE_AA)
			cv2.imwrite('{}.jpg'.format(firstname), im)

		#to mail email certificate to corresponding emailid
			print("Sending mail to {}".format(email_id))
			email_user = 'saket.17jccs099@jietjodhpur.ac.in'
			email_password = 'Resonance06'
			email_send = str(email_id)

			subject = 'Testing'

			msg = MIMEMultipart()
			msg['From'] = email_user
			msg['To'] = email_send
			msg['Subject'] = subject

			body = 'Hi there, sending this email from Python!'
			msg.attach(MIMEText(body,'plain'))

			filename='{}.jpg'.format(firstname)
			attachment  =open(filename,'rb')
			part = MIMEBase('application','octet-stream')
			part.set_payload((attachment).read())
			encoders.encode_base64(part)
			part.add_header('Content-Disposition',"attachment; filename= "+filename)

			msg.attach(part)
			text = msg.as_string()
			server = smtplib.SMTP('smtp.gmail.com',587)
			server.starttls()
			server.login(email_user,email_password)


			server.sendmail(email_user,email_send,text)
			server.quit()

		else:
			continue
	if post :

		if attend == 'P':

			name = str(name1)
			name2 = name.split()
			firstname = name2[0]
			print(name)

			#to make copy of original certificate
			os.system('sudo cp achievement.jpg /home/roti/Desktop/certifiacte_testing/{}.jpg'.format(firstname))

			#for writing name in achievement
			im =cv2.imread("achievement.jpg",1)
			font = cv2.FONT_ITALIC
			font1 = cv2.FONT_HERSHEY_SCRIPT_SIMPLEX

			cv2.putText(im, name, (320,310), font1, 2, (204,204, 0), 2, cv2.LINE_AA)
			cv2.imwrite('{}.jpg'.format(firstname), im)

			x=550-(3*25)
			cv2.putText(im, post, (x + 35,430), font, 1, (204,204,0), 2, cv2.LINE_AA)
			cv2.imwrite('{}.jpg'.format(firstname), im)

			print("Sending mail to {}".format(email_id))
			email_user = 'saket.17jccs099@jietjodhpur.ac.in'
			email_password = 'Resonance06'
			email_send = str(email_id)
			subject = 'Testing'

			msg = MIMEMultipart()
			msg['From'] = email_user
			msg['To'] = email_send
			msg['Subject'] = subject

			body = 'Hi there, sending this email from Python!'
			msg.attach(MIMEText(body,'plain'))

			filename='{}.jpg'.format(firstname)
			attachment  =open(filename,'rb')

			part = MIMEBase('application','octet-stream')
			part.set_payload((attachment).read())
			encoders.encode_base64(part)
			part.add_header('Content-Disposition',"attachment; filename= "+filename)

			msg.attach(part)
			text = msg.as_string()
			server = smtplib.SMTP('smtp.gmail.com',587)
			server.starttls()
			server.login(email_user,email_password)

			server.sendmail(email_user,email_send,text)
			server.quit()

		else:
			continue
