# import smtplib
# # First name : Sender
# # Last name : Sender
# # username: sendersen51 
# # password: spirit2626

# # First name : Receiver
# # Last name : Receiver
# # username: receiversen51 
# # password: spirit2626 
 

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

    
email_user = 'sendersen51@gmail.com'
email_password = 'spirit2626'
email_send = 'receiversen51@gmail.com'
subject = 'Testing'

msg = MIMEMultipart()
msg['From'] = email_user
msg['To'] = email_send
msg['Subject'] = subject

body = 'Hi there, sending this email from Python!'
msg.attach(MIMEText(body,'plain'))

filename='Certis/fsrc/spirit.jpg'
attachment  =open(filename,'rb')

part = MIMEBase('application','octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition','attachment', filename="abc.jpg")

msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(email_user,email_password)

server.sendmail(email_user,email_send,text)
server.quit()

# mail_content = '''Hello,
# This is a test mail.
# In this mail we are sending some attachments.
# The mail is sent using Python SMTP library.
# Thank You
# '''
# #The mail addresses and password
# sender_address = 'sendersen51@gmail.com'
# sender_pass = 'spirit2626'
# receiver_address = 'receiversen51@gmail.com'
# #Setup the MIME
# message = MIMEMultipart()
# message['From'] = sender_address
# message['To'] = receiver_address
# message['Subject'] = 'A test mail sent by Python. It has an attachment.'
# #The subject line
# #The body and the attachments for the mail
# message.attach(MIMEText(mail_content, 'plain'))
# attach_file_name = 'Certis/fsrc/spirit.jpg'
# attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
# payload = MIMEBase('application', 'octate-stream')
# payload.set_payload((attach_file).read())
# encoders.encode_base64(payload) #encode the attachment
# #add payload header with filename
# payload.add_header('Content-Decomposition', 'attachment', attach_file_name="abc.jpg")
# # part.add_header('Content-Disposition',"attachment; filename= "+filename)
# message.attach(payload)
# session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
# session.starttls() #enable security
# session.login(sender_address, sender_pass) #login with mail_id and password
# text = message.as_string()
# session.sendmail(sender_address, receiver_address, text)
# session.quit()
# print('Mail Sent')