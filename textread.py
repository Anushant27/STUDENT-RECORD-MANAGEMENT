from transform import four_point_transform
from skimage.filters import threshold_local
from PIL import Image
import cv2
import pytesseract
import argparse
import xlwt 
from xlwt import Workbook
import os
import imutils
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
details=['NAME','USN','1A','1B','1C','2A','2B','2C','3A','3B','3C']
for d in range(11):
    sheet1.write(0, d, details[d])
    d=d+1
student=input("Enter the number of student: ")
student=int(student)
for i in range(student):
    ch='x'
    while ch!='c':
        ch=input('Enter C to capture: ')
    cap = cv2.VideoCapture(1)
    ret,img=cap.read()
    #img = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    #ret, thresh = cv2.threshold(img, 47, 255, cv2.THRESH_BINARY)
    output = "G:\\openCV\\image\\test.png"
    cv2.imwrite(output, img)
    image = cv2.imread(r'G:\\openCV\\image\\detail.jpg')
    ratio = image.shape[0] / 500.0
    orig = image.copy()
    image = imutils.resize(image, height=500)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(gray, 0, 200)
    cnts = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)
    cnts = sorted(cnts, key=cv2.contourArea, reverse=True)[:5]
    for c in cnts:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)
        if len(approx) == 4:
            screenCnt = approx
            break

    warped = four_point_transform(orig, screenCnt.reshape(4, 2) * ratio)
    warped = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
    #T = threshold_local(warped, 11, offset=10, method="gaussian")
    #warped = (warped > T).astype("uint8") * 255
    # cv2.imshow("Original", imutils.resize(orig, height = 650))
    cv2.imshow("Scanned", imutils.resize(warped, height=650))
    cv2.imwrite('G:\\openCV\\image\\check2.jpg', warped)
    cv2.waitKey(0)
    basewidth = 1600
    img = Image.open(r'G:\\openCV\\image\\check2.jpg')
    wpercent = (basewidth / float(img.size[0]))
    hsize = int((float(img.size[1]) * float(wpercent)))
    img = img.resize((basewidth, hsize), Image.ANTIALIAS)
    img.save(r'G:\\openCV\\image\\test3.jpg')
    image = Image.open(r'G:\\openCV\\image\\test3.jpg')
    im1 = image.crop((237, 933, 893, 985))
    image_to_text1 = pytesseract.image_to_string(im1, lang='eng')
    im2 = image.crop((1141, 909, 1581, 949))
    image_to_text2 = pytesseract.image_to_string(im2, lang='eng', config='--psm 12 --oem 3 ')
    sheet1.write((i+1), 0,image_to_text1)
    sheet1.write((i+1), 1, image_to_text2)
    print(image_to_text1)
    print(image_to_text2)
    z=2
    top = 2029
    btm = 2081
    for j in range(3):
        left = 249
        right = 333
        for k in range(3):
            im3 = image.crop((left, top, right, btm))
            image_to_text3 = pytesseract.image_to_string(im3, lang='eng', config='--psm 10 --oem 3 -c tessedit_char_whitelist=0123456789')
            print(image_to_text3)
            sheet1.write((i + 1), z, image_to_text3)
            z=z+1
            left=left+108
            right=right+108
            k=k+1
        top=top+75
        btm=btm+75
        j=j+1
    i=i+1
wb.save('xlwt example.xls')