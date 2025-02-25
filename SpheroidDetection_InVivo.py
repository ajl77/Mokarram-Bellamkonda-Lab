# GLIOBLASTOMA (IN-VIVO) SPHEROID DETECTION (PYTHON)
# DREW LAWLESS, JANEL ENNIN; ORIGINALLY CREATED BY AYUSH UPHADHYAY
# MOKARRAM LAB (EMORY UNIVERISTY)
# STARTED - 09/07/2023, LAST UPDATED - 10/11/2023

import os
import openpyxl
import numpy as np
import cv2 as cv
import scipy
from matplotlib import pyplot as plt
import shutil
from openpyxl import load_workbook
import xlwt
from xlwt import Workbook

#set list of directories for image folders throughout running of code
os.chdir("/Users/drew/Desktop/Research/Image Analysis")
batchProcess = "/Users/drew/Desktop/Research/Image Analysis/Batch Process"
processedImage = "/Users/drew/Desktop/Research/Image Analysis/Processed Image"
excelFiles = "/Users/drew/Desktop/Research/Image Analysis/Excel Files"
detectedImage = "/Users/drew/Desktop/Research/Image Analysis/Detected Image"

#desginate directory of Excel file
wb = openpyxl.load_workbook('/Users/drew/Desktop/Research/Image Analysis/Excel Files/ImageAnalysisData.xlsx')

#list of all counting variables
rowCounter = 2
contourCounter = 1
otherCounter = 0
index = 0
imageCounter = 0

#loop through each image and open file in opencv 
for filename in os.listdir(batchProcess):
    if '.tif' in filename:
        imageCounter += 1
        print("ANAlYZING IMAGE #" + str(imageCounter) + "...")
        file = os.path.join(batchProcess, filename)
        oG = cv.imread(file)
        image = oG
        gray = cv.cvtColor(image, cv.COLOR_BGR2GRAY)

        #criteria for grayscaling
        lower_white = np.array([95,95,95])
        upper_white = np.array([255,255,255])
        mask = cv.inRange(image, lower_white, upper_white)
        image[mask > 0] = (30,30,30)
        image[mask <= 0] = (255,255,255)

        #fine-tuning
        image = cv.dilate(image, np.ones((13,13), np.uint8))
        image = cv.erode(image, np.ones((13,13), np.uint8))
        image = cv.dilate(image, np.ones((13,13), np.uint8))
        image = cv.GaussianBlur(image, (5,5), 0)
        image = cv.Canny(image,100,200)

        #set style of lettering for bounding box number
        font = cv.FONT_HERSHEY_SIMPLEX
        fontScale = 1
        fontColor = (0,0,255)
        thickness = 5
        lineType = 2

        print("Image converting to grayscale...")
        
        #create bounding boxes
        contours, hierarchy = cv.findContours(image,cv.RETR_LIST,cv.CHAIN_APPROX_SIMPLE)
        print("Number of contours: " + str(len(contours)))
        boundXleft = []
        boundXright = []
        boundYtop = []
        boundYbottom = []
        boundCX = []
        boundCY = []
        boundW = []
        boundH = []
        boundA = []
        duplicateRemovalX = []
        duplicateRemovalY = []
        duplicateRemovalX.append(0)
        duplicateRemovalY.append(0)
        dCounter = 0
        for c in contours:
            x,y,w,h = cv.boundingRect(c)
            area = cv.contourArea(c)
            aspectRatio = w / h
            if cv.contourArea(c) > 10000  and cv.contourArea(c) < 300000:
                if aspectRatio > 0.5 and aspectRatio < 2.0:
                    if duplicateRemovalX[dCounter] != x and duplicateRemovalY[dCounter] != y:
                        duplicateRemovalX.append(x)
                        duplicateRemovalY.append(y)
                        cv.rectangle(oG, (x, y), (x + w, y + h), (0, 255, 0), 2)
                        cv.putText(oG, 'C' + str(contourCounter), (x, y), font, fontScale, fontColor, thickness, lineType)
                        area = cv.contourArea(c)
                        center = (x, y)
                        boundXleft.append(x - (w / 2))
                        boundXright.append(x + (w / 2))
                        boundYbottom.append(y - (h / 2))  
                        boundYtop.append(y + (h / 2)) 
                        boundW.append(w)
                        boundH.append(h)
                        boundCX.append(x)
                        boundCY.append(y)
                        boundA.append(area)
                        contourCounter += 1
                        otherCounter += 1
                        dCounter += 1
                        print("Contour Area: " + str(area))
                        print("Aspect Ratio: " + str(aspectRatio))
                        print("Center: " + str(center))
                        print("Width: " + str(w) + " Height: " + str(h))
        print("Number of validated contours: " + str(otherCounter))
        print("Showing image with bounding boxes...")
        print(image)

        #saving data to designated Excel file
        os.chdir(excelFiles)
        sheet1 = wb['10/2023 In-Vivo']
        while index < otherCounter:        
            xVleft = boundXleft[index]
            xVright = boundXright[index]
            yVtop = boundYtop[index]
            yVbottom = boundYbottom[index]
            width = boundW[index]
            height = boundH[index]
            centerPointX = boundCX[index]
            centerPointY = boundCY[index]
            area = boundA[index]
            sheet1['A' + str(rowCounter)] = str(filename + " contour " + str(spheroidCounter) + ")")
            sheet1['B' + str(rowCounter)] = width
            sheet1['C' + str(rowCounter)] = height
            sheet1['D' + str(rowCounter)] = centerPointX
            sheet1['E' + str(rowCounter)] = centerPointY
            sheet1['F' + str(rowCounter)] = area
            sheet1['G' + str(rowCounter)] = xVleft
            sheet1['H' + str(rowCounter)] = xVright
            sheet1['I' + str(rowCounter)] = yVtop
            sheet1['J' + str(rowCounter)] = yVbottom
            wb.save('ImageAnalysisData.xlsx')
            rowCounter += 1
            index += 1
            print("Contour " + str(index) + " data saved")
            if index == contourCounter - 1:
                contourCounter = 1
                otherCounter = 0
                index = 0
                wb.save('ImageAnalysisData.xlsx')
                print("All contour data added to excel")
                break
      
        #move original and altered image to new files
        shutil.move(file, os.path.join(processedImage, filename))
        os.chdir(detectedImage)
        cv.imwrite(filename, oG)
        os.chdir("/Users/drew/Desktop/Research/Image Analysis")

#save excel and finish running code
os.chdir(excelFiles)
wb.save('ImageAnalysisData.xlsx')
print("IMAGE PROCESSING COMPLETE!")
