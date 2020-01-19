import cv2
from openpyxl import Workbook
import pytesseract
import time
import math

original = cv2.imread('C:/Users/user/Pictures/imgLocation.png')#replace with path to image

height = original.shape[0]
column = 1
row = 1
workbook = Workbook()
sheet = workbook.active
copy = original
chg = 31.3#31.3684211 # 1192/38
rows = int((height - 55)/chg) # 1247-1192
startTime = time.clock()
prevTime = startTime
rectangleUpLeft = (10, 49)
rectW = 356
rectH = 30
left = 0
leftAdd = 9


while(column < 6):
    if column == 1:
        row = 1
        newImg = original[:, :359]
        while(row < rows + 1):
            copy = original
            left = int((row-1)*chg) + leftAdd
            rowImg = newImg[left:left + rectH, :]
            txt = pytesseract.image_to_string(rowImg)
            sheet['A' + str(row)] = txt
            rectangleLoc = (0, left)
            rectW = 356
            row +=1
            cv2.rectangle(copy, rectangleLoc, (rectangleLoc[0] + rectW, rectangleLoc[1] + rectH), (255, 0, 0), 1)
            cv2.imshow('Image', copy)
            cv2.waitKey(1)
        print('A Complete in ' + str(time.clock() - prevTime) + ' seconds')
        prevTime = time.clock()
    elif column == 2:
        row = 1
        newImg = original[:, 555:630]
        while(row < rows + 1):
            copy = original
            left = int((row-1)*chg) + leftAdd
            rowImg = newImg[left:left + rectH, :]
            txt = pytesseract.image_to_string(rowImg)
            sheet['B' + str(row)] = txt
            rectangleLoc = (555, left)
            rectW = 75
            row +=1
            cv2.rectangle(copy, rectangleLoc, (rectangleLoc[0] + rectW, rectangleLoc[1] + rectH), (255, 0, 0), 1)
            cv2.imshow('Image', copy)
            cv2.waitKey(1)
        print('B Complete in ' + str(time.clock() - prevTime) + ' seconds')
        prevTime = time.clock()
    elif column == 3:
        row = 1
        newImg = original[:, 685:895]
        while(row < rows + 1):
            copy = original
            left = int((row-1)*chg) + leftAdd
            rowImg = newImg[left:left + rectH, :]
            txt = pytesseract.image_to_string(rowImg)
            sheet['C' + str(row)] = txt
            rectangleLoc = (685, left)
            rectW = 210
            row +=1
            cv2.rectangle(copy, rectangleLoc, (rectangleLoc[0] + rectW, rectangleLoc[1] + rectH), (255, 0, 0), 1)
            cv2.imshow('Image', copy)
            cv2.waitKey(1)
        print('C Complete in ' + str(time.clock() - prevTime) + ' seconds')
        prevTime = time.clock()
    elif column == 4:
        row = 1
        newImg = original[:, 1110:1280]
        while(row < rows + 1):
            copy = original
            left = int((row-1)*chg) + leftAdd
            rowImg = newImg[left:left + rectH, :]
            txt = pytesseract.image_to_string(rowImg)
            sheet['D' + str(row)] = txt
            rectangleLoc = (1110, left)
            rectW = 170
            row +=1
            cv2.rectangle(copy, rectangleLoc, (rectangleLoc[0] + rectW, rectangleLoc[1] + rectH), (255, 0, 0), 1)
            cv2.imshow('Image', copy)
            cv2.waitKey(1)
        print('D Complete in ' + str(time.clock() - prevTime) + ' seconds')
        prevTime = time.clock()
    elif column == 5:
        row = 1
        newImg = original[:, 1290:1360]
        while(row < rows + 1):
            copy = original
            left = int((row-1)*chg) + leftAdd
            rowImg = newImg[left:left + rectH, :]
            txt = pytesseract.image_to_string(rowImg)
            sheet['E' + str(row)] = txt
            rectangleLoc = (1290, left)
            rectW = 70
            row +=1
            cv2.rectangle(copy, rectangleLoc, (rectangleLoc[0] + rectW, rectangleLoc[1] + rectH), (255, 0, 0), 1)
            cv2.imshow('Image', copy)
            cv2.waitKey(1)
        print('E Complete in ' + str(time.clock() - prevTime) + ' seconds')
        prevTime = time.clock()
    column += 1
sheet['E39'] = '=COUNTBLANK(A1:E38)'
print('Finished in ' + str(time.clock() - startTime) + ' seconds')
workbook.save('xlFile' + str(int(time.time() - 1579411333)) + '.xlsx')
cv2.destroyAllWindows()
