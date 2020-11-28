import pygame
import cv2
from openpyxl import Workbook
import pytesseract
import time
import math

name = 'music'
items = 20
listItems = []

def numberToLetter(number):
    alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    return alphabet[number-1]
def clamp(n, maxv, minv):
    return max(min(n, maxv), minv)

while len(listItems) < items:
    listItems.append(len(listItems) + 1)

for i in listItems:
    pygame.init()
    running = True
    selectedRowLine = 0
    selectedColumnLine = 0
    workbook = Workbook()
    sheet = workbook.active

    win = pygame.display.set_mode((850, 1000))
    background = pygame.image.load('C:/Users/wruet/Documents/GitHub/ImgExcel/Quiz-Bowl-Data-master/Quiz-Bowl-Data-master/'+name+'/'+name+str(i)+'.png')
    original = cv2.imread('C:/Users/wruet/Documents/GitHub/ImgExcel/Quiz-Bowl-Data-master/Quiz-Bowl-Data-master/'+name+'/'+name+str(i)+'.png')
    origDim = (background.get_width(), background.get_height())
    print(origDim)
    bg = pygame.transform.scale(background, (win.get_width(), win.get_height()))
    wMultiplier = origDim[0]/bg.get_width()
    hMultiplier = origDim[1]/bg.get_height()

    row = 1
    column = 1

    height = original.shape[0]

    copy = original
    startTime = time.clock()
    prevTime = startTime
    columnLines = [47, 251, 327, 394, 774, 805]
    rowLines = [85, 101, 116, 134, 151, 167, 187, 205, 222, 240, 258, 274, 292, 308, 326, 343, 362, 379, 396, 414, 431, 449, 466, 482, 500, 519, 535, 553, 571, 588,
    604, 623, 640, 658, 676, 691, 709, 728, 744, 762, 781, 797, 813, 831, 849, 866, 883, 900, 917]

    while running:
        columnLines.sort()
        rowLines.sort()
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:
                    running = False
                if event.key == pygame.K_KP_ENTER:
                    running = False
                if column:
                    if event.key == pygame.K_RIGHT:
                        if selectedColumnLine < len(columnLines)-1:
                            selectedColumnLine += 1
                        else:
                            selectedColumnLine = 0
                    elif event.key == pygame.K_LEFT:
                        if selectedColumnLine > 0:
                            selectedColumnLine -= 1
                        else:
                            selectedColumnLine = len(columnLines)-1
                    elif event.key == pygame.K_DELETE and len(columnLines) > 1:
                        columnLines.pop(selectedColumnLine)
                    elif event.key == pygame.K_UP and columnLines[selectedColumnLine > 0]:
                        columnLines[selectedColumnLine] -= 1
                    elif event.key == pygame.K_DOWN and columnLines[selectedColumnLine] < win.get_height()-1:
                        columnLines[selectedColumnLine] += 1
                    elif event.key == pygame.K_SPACE:
                        columnLines[selectedColumnLine] = pygame.mouse.get_pos()[0]
                    elif event.key == pygame.K_t:
                        column = False
                    elif event.key == pygame.K_c:
                        columnLines.append(columnLines[selectedColumnLine] + 15)
                else:
                    if event.key == pygame.K_UP:
                        if selectedRowLine > 0:
                            selectedRowLine -= 1
                        else:
                            selectedRowLine = len(rowLines)-1
                    elif event.key == pygame.K_DOWN:
                        if selectedRowLine < len(rowLines)-1:
                            selectedRowLine +=1
                        else:
                            selectedRowLine = 0
                    elif event.key == pygame.K_DELETE and len(rowLines) > 1:
                        rowLines.pop(selectedRowLine)
                        if selectedRowLine > 0:
                            selectedRowLine -= 1
                        elif selectedRowLine == 0:
                            selectedRowLine = len(rowLines)-1
                        elif selectedRowLine < len(rowLines)-1:
                            selectedRowLine += 1
                        else:
                            selectedRowLine = 0
                        
                    elif event.key == pygame.K_LEFT and rowLines[selectedRowLine > 0]:
                        rowLines[selectedRowLine] -= 1
                    elif event.key == pygame.K_RIGHT and rowLines[selectedRowLine] < win.get_height()-1:
                        rowLines[selectedRowLine] += 1
                    elif event.key == pygame.K_SPACE:
                        rowLines[selectedRowLine] = pygame.mouse.get_pos()[1]
                    elif event.key == pygame.K_t:
                        column = True
                    elif event.key == pygame.K_r:
                        rowLines.append(rowLines[selectedRowLine] + 15)
                    elif event.key == pygame.K_u:
                        for q in rowLines:
                            new = q - 2
                            rowLines.remove(q)
                            rowLines.append(new)
                            rowLines.sort()
                            new = 0
                    elif event.key == pygame.K_d:
                        for q in rowLines:
                            new = 2 + q
                            rowLines.remove(q)
                            rowLines.append(new)
                            rowLines.sort()
                            new = 0
            
        win.blit(bg, (0,0))
        if column:
            for line in columnLines:
                if columnLines.index(line) == selectedColumnLine:
                    pygame.draw.line(win, pygame.Color('blue'), (line, 0), (line, win.get_height()-1))
                else:
                    pygame.draw.line(win, pygame.Color('red'), (line, 0), (line, win.get_height()-1))
        else:
            for line in rowLines:
                if rowLines.index(line) == selectedRowLine:
                    pygame.draw.line(win, pygame.Color('blue'), (0, line), (win.get_width()-1, line))
                else:
                    pygame.draw.line(win, pygame.Color('red'), (0, line), (win.get_width()-1, line))
        

        pygame.display.flip()
    print(columnLines)
    print(rowLines)
    print('Done')
    pygame.quit()

    column = 1
    row = 1

    for clm in columnLines:
        row = 1
        if column == len(columnLines):
            break
        else:
            for rw in rowLines:
                    if rowLines.index(rw) == len(rowLines)-1:
                        break
                    else:
                        copy = original
                        newImg = original[int(rowLines[row - 1]*hMultiplier):int(rowLines[row]*hMultiplier), int(columnLines[column-1]*wMultiplier):int(columnLines[column]*wMultiplier)]
                        txt = pytesseract.image_to_string(newImg)
                        sheet[numberToLetter(column) + str(row)] = txt
                        row +=1
        print(column)
        column += 1
    columnLines = [47, 251, 327, 394, 774, 805]
    rowLines = [85, 101, 116, 134, 151, 167, 187, 205, 222, 240, 258, 274, 292, 308, 326, 343, 362, 379, 396, 414, 431,
    449, 466, 482, 500, 519, 535, 553, 571, 588,604, 623, 640, 658, 676, 691, 709, 728, 744, 762, 781, 797, 813, 831, 849, 866, 883, 900, 917]
    print('Finished in ' + str(time.clock() - startTime) + ' seconds')
    cv2.destroyAllWindows()
    workbook.save(name + str(i) + '.xlsx')
    workbook.remove_sheet(sheet)