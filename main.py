#E-Certificate Generation and Emailing Program by Vikrant Prasad, IIT G
import xlrd
import cv2 
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Opening the excel file and selecting sheet number
workbook = xlrd.open_workbook("Certis/fsrc/test.xls") 
worksheet=workbook.sheet_by_index(0)  

C = worksheet.ncols # columns in our sheet (only valid)
R = worksheet.nrows # rows in our sheet (only valid)

email_user = 'xyz@gmail.com' #email of sender
email_password = '*********' #email of sender

#connecting to gmail
server = smtplib.SMTP('smtp.gmail.com',587) 
server.starttls()
server.login(email_user,email_password) #allow access to less secure apps in google security


for i in range(1,R):

    #to read from excel file
    name = worksheet.cell(i,1).value
    college = worksheet.cell(i,3).value
    sport = worksheet.cell(i,2).value
    recipentEmail = worksheet.cell(i,4).value

    #Adding text to image
    path="Certis/fsrc/spirit.jpg" # relative path of certificate template
    img = cv2.imread(path) #reading the image using openCV

    # Parameters of text to be added
    font = cv2.FONT_ITALIC 
    certi_text = "{} from {}".format(name,college) 
    offset1= int((len(certi_text)/2)*16) #adjustent
    orgCerti = (640-offset, 485)
    orgCollege = (680-int((len(sport)/2)*16), 547) 
    fontScale = 1 
    color = (0, 0, 0)  
    thickness = 2 
    
    cv2.putText(img, certi_text, orgCerti, font, fontScale, color, thickness, cv2.LINE_AA)
    cv2.putText(img, sport   , orgCollege, font, fontScale, color, thickness, cv2.LINE_AA)
    cv2.imwrite('Certis/fdst/{}.jpg'.format(name), img)


    #Email automation
    email_send = recipentEmail
    subject = "Certificate of participation in SPIRIT 2019, IIT Guwahati"

    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject

    body = 'Thanks {} for participating in SPIRIT IITG Guwahati. PFA of your e-certificate for representing {} in {} during the sports fest'.format(name,college,sport)
    msg.attach(MIMEText(body,'plain'))

    filename='Certis/fdst/{}.jpg'.format(name)
    attachment  =open(filename,'rb')

    part = MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment', filename="{}.jpg".format(name))
    msg.attach(part)
    text = msg.as_string()
    server.sendmail(email_user,email_send,text) 
    print("Mail {} sent successfully".format(i))

server.quit()


    # fontList=[
    # cv2.FONT_HERSHEY_COMPLEX,
    # cv2.FONT_HERSHEY_COMPLEX,
    # cv2.FONT_HERSHEY_COMPLEX_SMALL,
    # cv2.FONT_HERSHEY_DUPLEX,
    # cv2.FONT_HERSHEY_PLAIN,
    # cv2.FONT_HERSHEY_SCRIPT_COMPLEX,
    # cv2.FONT_HERSHEY_SCRIPT_SIMPLEX,
    # cv2.FONT_HERSHEY_SIMPLEX,
    # cv2.FONT_HERSHEY_TRIPLEX,
    # cv2.FONT_ITALIC]