import smtplib
import excel
import excelPython
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from skimage.filters import threshold_local
from PIL import Image
import cv2
import pytesseract
import argparse
import xlwt
from xlwt import Workbook
import os

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
details=['NAME','USN','1A','1B','1C','2A','2B','2C','3A','3B','3C','total','CIE']
marks=[r'C:\\Users\\HP\\Desktop\\HACK_IEEE\\Dataset\\image_marks\\Untitled1.png',r'C:\\Users\\HP\\Desktop\\HACK_IEEE\\Dataset\\image_marks\\Untitled2.png',r'C:\\Users\\HP\\Desktop\\HACK_IEEE\\Dataset\\image_marks\\Untitled3.png',r'C:\\Users\\HP\\Desktop\\HACK_IEEE\\Dataset\\image_marks\\Untitled4.png',r'C:\\Users\\HP\\Desktop\\HACK_IEEE\\Dataset\\image_marks\\Untitled5.png']
for d in range(13):
    sheet1.write(0, d, details[d])
    d=d+1

student=input("Enter the number of student: ")
student=int(student)
for i in range(student):
    ch='x'
    while ch!='c':
        ch=input('Enter C to capture: ')
    

    image = Image.open(marks[i])
    im1 = image.crop((228, 220, 375, 238))
    image_to_text1 = pytesseract.image_to_string(im1, lang='eng')
    im2 = image.crop((423, 219, 524, 237))
    image_to_text2 = pytesseract.image_to_string(im2, lang='eng')
    sheet1.write((i+1), 0,image_to_text1)
    sheet1.write((i+1), 1, image_to_text2)
    print(image_to_text1)
    print(image_to_text2)
    z=2
    top = 462
    btm = 476
    total=0
    for j in range(3):
        left = 230
        right = 247
        for k in range(3):
            im3 = image.crop((left, top, right, btm))
            image_to_text3 = pytesseract.image_to_string(im3, lang='eng', config='--psm 13 --oem 3 -c tessedit_char_whitelist=0123456789')
            total=total+int(image_to_text3)
            print(image_to_text3)
            sheet1.write((i + 1), z, int(image_to_text3))
            z=z+1
            left=left+23
            right=right+23
        top=top+18
        btm=btm+18
        j=j+1
    sheet1.write((i + 1), 11, total)
    CIE=int((total/2.94)+1)
    sheet1.write((i + 1), 12, CIE)
    i=i+1
    print(total)
    print(CIE)
wb.save('xlwt example1.xls')


mail_content = '''Hello,
This is a test mail.
            k=k+1
In this mail we are sending some attachments.
Thank You
'''
#The mail addresses and password
sender_address = 'nikitharjain28@gmail.com'
sender_pass = 'nikerjain28'
receiver_address = 'anushant.2k16@gmail.com'
#Setup the MIME
message = MIMEMultipart()
message['From'] = sender_address
message['To'] = receiver_address
message['Subject'] = 'A test mail sent by Python. It has an attachment.'
#The subject line
#The body and the attachments for the mail
message.attach(MIMEText(mail_content, 'plain'))
attach_file_name = 'xlwt example1.xls'
attach_file = open('xlwt example1.xls','rb')
 # Open the file as binary mode
payload = MIMEBase('application', 'octate-stream')
payload.set_payload((attach_file).read())
encoders.encode_base64(payload) #encode the attachment
#add payload header with filename
payload.add_header('Content-Disposition', 'attachment;filename="xlwt example1.xls"')
message.attach(payload)
#attach_file.close()
#Create SMTP session for sending the mail
session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
session.starttls() #enable security
session.login(sender_address, sender_pass) #login with mail_id and password
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Mail Sent')