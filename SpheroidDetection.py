# IN-VITRO/EX-VIVO SPHEROID DETECTION (PYTHON)
# DESIGNED BY DREW LAWLESS; DERIVED FROM PREVIOUS CODE BY AYUSH UPHADHYAY
# MOKARRAM/BELLAMKONDA LAB (EMORY UNIVERISTY)
# created: 08/23/2023 | last updated: 09/03/2024

import os
import openpyxl
import numpy as np
import cv2 as cv
from matplotlib import pyplot as plt
import shutil
from openpyxl import load_workbook
import xlwt
from xlwt import Workbook

#set list of directories for image folders throughout running of code
main_path = '/Users/drew/Desktop/Research/Image Analysis'
input_path = '/Users/drew/Desktop/Research/Image Analysis/Input'
processed_path = '/Users/drew/Desktop/Research/Image Analysis/Processed'
excel_path = '/Users/drew/Desktop/Research/Image Analysis/PythonData.xlsx'
output_path = '/Users/drew/Desktop/Research/Image Analysis/Output'
wb = openpyxl.load_workbook(excel_path)
sheet1 = wb['Testing']

os.chdir(main_path)

#list of all counting variables
rowCounter = 2
spheroidCounter = 1
otherCounter = 0
index = 0
imageCounter = 0

#loop through each image and open file in opencv
for filename in os.listdir(input_path):
    if '.tif' in filename:
        imageCounter += 1
        print("ANAlYZING IMAGE #" + str(imageCounter) + "...")
        file = os.path.join(input_path, filename)
        oG = cv.imread(file)
        image = oG
        gray = cv.cvtColor(image, cv.COLOR_BGR2GRAY)
        grayConverted = cv.convertScaleAbs(gray)

        #create criteria for grayscaling
        dark = np.array([0,0,0])
        lower_medium = np.array([70,70,70])
        upper_medium = np.array([105,105,105])
        light = np.array([255,255,255])
        maskA = cv.inRange(image, upper_medium, light)
        maskB = cv.inRange(image, lower_medium, upper_medium)
        maskC = cv.inRange(image, dark, lower_medium)
        image[maskC > 0] = (255,255,255)
        image [maskB > 0] = (125, 125, 125)
        image[maskA > 0] = (0,0,0)

        #second fine-tuning
        image = cv.dilate(image, np.ones((13,13), np.uint8))
        image = cv.erode(image, np.ones((13,13), np.uint8))
        image = cv.dilate(image, np.ones((13,13), np.uint8))
        image = cv.GaussianBlur(image, (5,5), 0)
        image = cv.Canny(image,100,200)

        """


        #set style of lettering for bounding box number
        font = cv.FONT_HERSHEY_SIMPLEX
        fontScale = 0.8
        fontColor = (0,0,255)
        thickness = 3
        lineType = 2

        print("Image converting to grayscale...")

        #create bounding boxes
        contours, hierarchy = cv.findContours(image,cv.RETR_LIST,cv.CHAIN_APPROX_SIMPLE)
        print("Number of contours detected: " + str(len(contours)))
        duplicateRemovalX = []
        duplicateRemovalY = []
        duplicateRemovalX.append(0)
        duplicateRemovalY.append(0)
        dCounter = 0
        
        for c in contours:
            x,y,w,h = cv.boundingRect(c)
            area = cv.contourArea(c)
            aspect_ratio = w / h

            if 40000000 <= cv.contourArea(c) <= 50000000 and rf_check and 0.9 <= aspect_ratio <= 1.1:
                origin_x = x
                origin_y = y + h
                shifted_contour = []
                for point in contour:
                    if point and len(point) > 0:
                        shifted_point = (point[0][0] - origin_x, origin_y - point[0][1])
                        shifted_contour.append(shifted_point)
                        rf_check == False
                cv.drawContours(image, [shifted_contour], 0, (0, 0, 255), 2)
            
            if 2000 <= cv.contourArea(c) <= 20000:
                if .65 <= aspect_ratio <= 1.5:
                    if duplicateRemovalX[dCounter] != x and duplicateRemovalY[dCounter] != y:
                        duplicateRemovalX.append(x)
                        duplicateRemovalY.append(y)
                        left_x = x - (w / 2)
                        right_x = x + (w / 2)
                        bottom_y = y - (h / 2)
                        top_y = y + (h / 2)
                        cv.rectangle(oG, (x, y), (x + w, y + h), (0, 255, 0), 2)
                        cv.putText(oG, str(spheroidCounter), (x, y), font, fontScale, fontColor, thickness, lineType)
                        center = (x, y)
                        sheet1['A' + str(rowCounter)] = str(filename + " (spheroid " + str(spheroidN) + ")")
                        sheet1['B' + str(rowCounter)] = area
                        sheet1['C' + str(rowCounter)] = left_x
                        sheet1['D' + str(rowCounter)] = right_x
                        sheet1['E' + str(rowCounter)] = bottom_y
                        sheet1['F' + str(rowCounter)] = top_y
                        wb.save(excel_path)
                        rowCounter += 1
                        index += 1
                        
                    if index == spheroidCounter - 1:
                        spheroidCounter = 1
                        otherCounter = 0
                        index = 0
                        wb.save(excel_path)
                        print("All spheroid data added to excel")
                        print("\n------------------------------\n")
                        break
        """

        #move original and altered image to new files
        shutil.move(file, os.path.join(processed_path, filename))
        os.chdir(output_path)
        cv.imwrite(filename, oG)
        os.chdir(main_path)

#save excel and finish running code
wb.save(excel_path)
print("IMAGE ANALYSIS COMPLETE!")
